#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build.py —— 把段落 JSON 与 LaTeX 表合成为最终 Markdown。

用法:
    python build.py <workdir> --latex <latex.json> --out <output.md>
                    [--suggest] [--use-suggested]

输入:
    <workdir>/paras.json        extract.py 的产出
    <workdir>/eq_index.json     extract.py 的产出；OMML 条目带 latex 字段，
                                这里自动预填（latex.json 给同名键时以它为准）
    <workdir>/headings.json     （可选）{"段落索引": 级别}，人工确认过的标题
    <workdir>/headings_suggested.json  --suggest 生成，供人工修订后改名生效
    <latex.json>                {"oleObject1.bin": "$...$", ...}
                                纯 OMML 文档可以是 {}（全部走预填）

输出:
    <output.md>                 最终 Markdown
    <workdir>/build_report.json 统计与告警

标题策略（重要）:
    自动识别只是**初稿**。中文文档的标题层级不规范（"1．" / "A）" / "i）" /
    无编号的"引言"混用），任何正则都猜不全。因此本脚本：
      1. 用尽量宽的规则生成 headings_suggested.json；
      2. 人工审阅、修订、改名为 headings.json 后再跑一次；
    --suggest 只产出建议，--use-suggested 直接用建议（跳过人工确认）。

公式位置策略:
    * 整段只有一个公式 → display block（$$ 独占行）
    * 公式与文字混排   → 默认 inline，避免把句子劈开
    * 例外：含块级环境（aligned/matrix/...），或源串长度超过
      LONG_TEX_CHARS，即使混排也独立成块

    说明：LONG_TEX_CHARS 是排版启发式，不是硬性规则。调大可让更多公式
    留在行内（句子更连贯），调小则更多独立成块（长式更好读）。

Markdown 六条铁律（每条都踩过坑，详见 SKILL.md）:
    1. $$ 必须独占一行
    2. 行内公式不能跨行（不能含 \\begin{aligned}）
    3. 相邻两个公式之间必须有空行或正文
    4. \\tag 紧跟 \\end{aligned}
    5. 超长公式不要塞进表格单元格
    6. 图片占位用 HTML 注释
"""
import argparse
import json
import os
import re
import sys
from docxfile import force_utf8_stdio

EQ_RE = re.compile(r"\u27e6EQ:([^\u27e7]+)\u27e7")   # ⟦EQ:name⟧

# 混排段里，LaTeX 源串超过这个长度就独立成块。纯排版启发式，可按需调整。
#
# 从 150 降到 80 的理由：实测这份真实文档里有 115–119 字符的公式留在行内，
# 它们在中文字里几乎必然折行或溢出，窄栏/移动端渲染时更容易出问题。用户
# 反馈的「只有长公式坏、短的正常」正是这个区间。80 让绝大部分长式独立成块，
# 同时仍把 a+b、\alpha 这类真正的短式留在句子里，不破坏句子连贯性。
LONG_TEX_CHARS = 80

# 只由标点组成的残片，应粘回前一个块
PUNCT_ONLY = re.compile(r"^[\s，。；：、,.;:）)】\]】\"'～~—\-]+$")

# CJK 字符与全角标点。行内公式与它们相邻时，不少渲染器不认定界符。
CJK = re.compile(r"[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]")


def _boundary_after_formula(text):
    """把「$公式$紧跟中文」改成「$公式$ 紧跟中文」。

    只在闭合 $ 后紧贴 CJK 时插入一个空格，其余情况原样返回。
    """
    return re.sub(r"(?<=\$)(?=" + CJK.pattern + r")", " ", text)


def _img_rel_src(out_path, workdir, fname):
    """图片公式在 md 里的引用路径：相对 md 文件定位 workdir/media/ 原件。

    用相对路径而不是绝对路径，md 连同目录整体拷贝/上传后仍然能显示。
    空格按 %20 转义，中文目录名保持原样（Markdown 允许）。
    """
    from urllib.parse import quote
    media = os.path.join(workdir, "media", fname)
    rel = os.path.relpath(media, os.path.dirname(os.path.abspath(out_path)) or ".")
    return quote(rel.replace(os.sep, "/"), safe="/._-~")

BLOCK_ENV = re.compile(
    r"\\begin\{(aligned|array|matrix|pmatrix|bmatrix|vmatrix|cases|split|"
    r"gather|align|eqnarray|equation)\}")

# 常见无编号标题
BARE_HEADINGS = {
    "引言", "前言", "序言", "摘要", "结论", "讨论", "致谢", "参考文献",
    "附录", "ABSTRACT", "Abstract", "INTRODUCTION", "Introduction",
    "CONCLUSION", "Conclusion", "REFERENCES", "References",
}


def load_json(path, default=None):
    if not os.path.exists(path):
        return default
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


def classify(latex, mixed, long_tex_chars=None):
    """判断一条公式该用 display 还是 inline。"""
    limit = LONG_TEX_CHARS if long_tex_chars is None else long_tex_chars
    has_block = bool(BLOCK_ENV.search(latex))
    long_tex = len(latex) > limit
    if mixed:
        # 与文字混排：只有超长或含块级环境才独立成块
        return has_block or long_tex
    # 独占一段：一律 display
    return True


def guess_headings(paras):
    """生成标题建议 {索引: 级别}。保守但覆盖常见中文排版。"""
    out = {}
    for i, raw in enumerate(paras):
        t = raw.strip()
        if not t or len(t) > 90:
            continue
        # 已被公式占满的段落不可能是标题
        if EQ_RE.search(t):
            continue
        # 句号结尾基本是正文
        if t.endswith("。") or t.endswith("；"):
            continue

        if t in BARE_HEADINGS:
            out[i] = 2
            continue

        # 1． / 1. / 2.3 / 2.3.1
        m = re.match(r"^(\d+(?:[.．]\d+)*)[\s、.．]+(\S.*)$", t)
        if m and len(m.group(2)) <= 45:
            depth = re.split(r"[.．]", m.group(1))
            out[i] = min(len(depth) + 1, 6)
            continue

        # 第一章
        if re.match(r"^第[一二三四五六七八九十百]+章", t) and len(t) <= 40:
            out[i] = 2
            continue
        if re.match(r"^第[一二三四五六七八九十百]+节", t) and len(t) <= 40:
            out[i] = 3
            continue

        # 附录 A
        if re.match(r"^附录\s*[A-Z一二三四五六七八九十]", t) and len(t) <= 40:
            out[i] = 2
            continue

        # A） / A)  —— 一级子标题
        if re.match(r"^[A-Z][）)]", t) and len(t) <= 45:
            out[i] = 3
            continue

        # i） / ii） / I. —— 二级子标题
        if re.match(r"^(?:[ivxlcdm]{1,6})[）)].", t, re.I) and len(t) <= 45:
            out[i] = 4
            continue

        # 一、
        if re.match(r"^[一二三四五六七八九十]+[、.．]", t) and len(t) <= 45:
            out[i] = 3
            continue

        # （一）
        if re.match(r"^[（(][一二三四五六七八九十]+[）)]", t) and len(t) <= 45:
            out[i] = 4
            continue
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workdir")
    ap.add_argument("--latex", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--suggest", action="store_true",
                    help="只生成 headings_suggested.json，不写 Markdown")
    ap.add_argument("--use-suggested", action="store_true",
                    help="直接用自动建议的标题（不做人工确认）")
    ap.add_argument("--long-tex-chars", type=int, default=LONG_TEX_CHARS,
                    help="混排段里 LaTeX 源串超过这个长度就独立成块（默认 %d）。"
                         "调大可让更多公式留在句内、句子更连贯；调小则更多独立成块、"
                         "长式更好读也更不容易在窄栏里溢出。" % LONG_TEX_CHARS)
    args = ap.parse_args()

    if args.long_tex_chars < 1:
        sys.stderr.write("build 输入错误: --long-tex-chars 必须是正整数\n")
        sys.exit(3)

    paras_path = os.path.join(args.workdir, "paras.json")
    if not os.path.exists(paras_path):
        sys.stderr.write(
            "build 输入错误: %s 不存在。请先运行 extract.py 生成提取结果。\n" % paras_path)
        sys.exit(3)
    if not os.path.exists(args.latex):
        sys.stderr.write("build 输入错误: LaTeX 表不存在: %s\n" % args.latex)
        sys.exit(3)

    paras = load_json(paras_path, [])
    index = load_json(os.path.join(args.workdir, "eq_index.json"), [])
    try:
        latex_raw = load_json(args.latex, {})
    except json.JSONDecodeError as exc:
        sys.stderr.write("build 输入错误: %s 不是合法 JSON: %s\n" % (args.latex, exc))
        sys.exit(3)
    latex = latex_raw.get("equations", latex_raw) if isinstance(latex_raw, dict) else {}
    if not isinstance(latex, dict):
        sys.stderr.write("build 输入错误: LaTeX 表应为 JSON 对象（公式名 → LaTeX 字符串）\n")
        sys.exit(3)

    # ---- 标题 ----
    suggested = guess_headings(paras)
    if args.suggest:
        path = os.path.join(args.workdir, "headings_suggested.json")
        save_json(path, {str(k): v for k, v in sorted(suggested.items())})
        print(json.dumps({
            "suggested_headings": len(suggested),
            "file": path,
            "note": ("请人工审阅：修订后改名 headings.json 再运行 build.py。"
                     "中文标题层级（1． / A） / i） / 无编号引言）无法完全自动识别。"),
        }, ensure_ascii=False, indent=1))
        return

    confirmed = load_json(os.path.join(args.workdir, "headings.json"))
    if confirmed is not None:
        headings = {int(k): v for k, v in confirmed.items()}
        heading_source = "headings.json"
        # 标题表是**按段落序号**索引的。段落序号会随提取口径变化（文本框里的
        # 段落现在被排除出正文），所以旧表可能已经对不上号。
        #
        # 对不上号的后果不是「标题挂错」这么简单：挂错的那一段如果恰好是
        # 一个独占的公式段，那条公式会被当成标题直接输出——公式凭空消失，
        # 校验器到 T8 才发现数对不上，而真正的错因（标题表过期）在报告里
        # 完全看不出来。
        #
        # 实测踩过：一份 145 公式的文档，段落数从 308 变 293，旧的
        # headings.json 把「段落 271」标成标题，而那正是唯一一条
        # `oleObject139` 公式所在的段——公式被当成标题吞掉，T8 报 144≠145。
        #
        # 所以：标题指向的段落里若含公式占位符，直接拒绝采用这张表。
        stale = [i for i, raw in enumerate(paras)
                 if headings.get(i) and EQ_RE.search(raw or "")]
        if stale:
            sys.stderr.write(
                "build 输入错误: headings.json 指向的段落里含公式占位符"
                "（段落 %s）——这张标题表与当前 paras.json 对不上号，"
                "会把公式当成标题输出，公式会凭空消失。\n"
                "请重新生成：先跑 --suggest，人工修订后改名 headings.json。\n"
                % ", ".join(str(i) for i in stale[:5]))
            sys.exit(3)
    elif args.use_suggested:
        headings = suggested
        heading_source = "auto-suggested"
    else:
        headings = suggested
        heading_source = "auto-suggested (未确认，建议先 --suggest 再人工修订)"

    known = {r["stream"] for r in index if r.get("kind") != "image"}
    # 图片公式不算 missing：Agent 不转写它也有合法出路（落成图片引用），
    # 不像 OLE 缺表那样是「正文丢了公式」。它的账记在 image_placeholders。

    # ---- OMML 公式预填 ----
    # extract.py 已把 Word 原生公式（OMML）从 XML 确定性转成 LaTeX，
    # 带在 eq_index.json 的 latex 字段里。这里预填进公式表；
    # Agent 的 latex.json 若给了同名键，以 Agent 为准（允许纠正转得不满
    # 意的条目）。不预填的话，OMML 公式会被记成 missing → T9 FAIL，
    # 而它们根本不需要读图。
    omml_prefilled = 0
    for r in index:
        if r.get("kind") == "omml" and r.get("latex") and not latex.get(r["stream"]):
            latex[r["stream"]] = r["latex"]
            omml_prefilled += 1

    missing = sorted([k for k in known if k not in latex])

    out = []
    stats = {"paragraphs": 0, "headings": 0, "formulas_inline": 0,
             "formulas_display": 0, "formulas_missing": len(missing),
             "unresolved_placeholders": 0, "heading_source": heading_source,
             # 回报实际生效的阈值：调用方据此确认参数有没有传进来，
             # 也便于事后复现「这一版用的是哪个值」。
             "long_tex_chars": args.long_tex_chars,
             "omml_prefilled": omml_prefilled,
             "image_placeholders": 0,   # 图片公式：落成注释+图片引用的处数
             "image_transcribed": 0}    # Agent 已转写成 LaTeX 的图片公式数

    prev_display = False

    def push_display(tex):
        """输出一个 display 公式块，保证前后有空行。"""
        nonlocal prev_display
        if out and out[-1] != "":
            out.append("")
        out.append("$$")
        out.append(tex)
        out.append("$$")
        out.append("")
        prev_display = True

    for i, raw in enumerate(paras):
        text = raw.strip()
        if not text:
            if out and out[-1] != "":
                out.append("")
            continue
        stats["paragraphs"] += 1

        # ---- 标题 ----
        level = headings.get(i, 0)
        if level:
            stats["headings"] += 1
            if out and out[-1] != "":
                out.append("")
            out.append("#" * min(int(level), 6) + " " + text)
            out.append("")
            prev_display = False
            continue

        # ---- 拆出公式 ----
        parts = EQ_RE.split(text)
        pieces = []      # [(kind, payload)]  kind: 'text' | 'formula'
        for k, seg in enumerate(parts):
            if k % 2 == 0:
                if seg:
                    pieces.append(("text", seg))
            else:
                name = seg
                tex_val = latex.get(name)
                if isinstance(tex_val, str) and tex_val.strip():
                    if name.startswith("img:"):
                        stats["image_transcribed"] += 1
                    pieces.append(("formula", tex_val))
                elif name.startswith("img:"):
                    # 图片公式且未转写：落成**可见**的图片引用 + 溯源注释。
                    # 不能用纯 HTML 注释——注释渲染后不可见，那就又回到
                    # 「静默丢公式」，正是本插件最忌讳的失败方式。
                    stats["image_placeholders"] += 1
                    fname = name[4:]
                    src = _img_rel_src(args.out, args.workdir, fname)
                    pieces.append(("text",
                                   "![图片公式 %s](%s)<!-- 图片公式原件: %s；"
                                   "在 latex.json 里用 \"img:%s\" 作键可替换成 LaTeX -->"
                                   % (fname, src, fname, fname)))
                else:
                    # 表里没有、或值不是字符串：明确标出来，不要静默跳过
                    stats["unresolved_placeholders"] += 1
                    pieces.append(("text", "<!-- 未解析公式: %s -->" % name))

        n_form = sum(1 for p in pieces if p[0] == "formula")
        real_text = "".join(p[1] for p in pieces
                            if p[0] == "text"
                            and not p[1].lstrip().startswith("<!--")).strip()
        alone = (n_form == 1 and not real_text)

        # ---- 单公式独占一段 → display ----
        if alone:
            tex = [p[1] for p in pieces if p[0] == "formula"][0].strip()
            push_display(tex)
            stats["formulas_display"] += 1
            continue

        # ---- 混排：拼回一句话，逐个判断 ----
        line = []
        pending = []     # 公式之后的标点残片，等下粘到末尾
        for kind, val in pieces:
            if kind == "text":
                if n_form and val.strip() and PUNCT_ONLY.match(val):
                    pending.append(val.rstrip())
                else:
                    if pending:
                        line.append("".join(pending))
                        pending = []
                    line.append(val)
            else:
                tex = str(val).strip()
                if classify(tex, mixed=True, long_tex_chars=args.long_tex_chars):
                    # 超长公式：先把已攒的文字冲出去，再独立成块
                    if pending:
                        line.append("".join(pending))
                        pending = []
                    head = "".join(line).strip()
                    if head:
                        # 这条 flush 路径同样要补公式边界：head 里可能已经含有
                        # 前面若干个行内公式，而它们后面紧跟着中文。
                        # 实测漏掉这一步时，只要段落里有个超长公式把句子劈开，
                        # 前半段里的公式就全都变成「$公式$中文」，阈值越小
                        # （劈开越多）越明显——阈值 40 时冒出 3 条 T10 警告。
                        out.append(_boundary_after_formula(head))
                        out.append("")
                        line = []
                    stats["formulas_display"] += 1
                    push_display(tex)
                else:
                    # 相邻两个内联公式不能直接相接：$a$$b$ 会被解析成
                    # 一个 $$ 定界符，导致其后整篇错位。补一个空格隔开。
                    if line and line[-1].endswith("$"):
                        line.append(" ")
                    # 行内公式与中文之间必须留出边界。
                    # 不少渲染器（Pandoc 的 tex_math_dollars、部分 KaTeX 集成）
                    # 只在 $ 与前后字符有明确边界时才识别为公式；紧贴中文时
                    # 会被当成普通文本，用户看到的就是「LaTeX 源码原样出现」。
                    # 只在紧贴 CJK 时补空格，纯 ASCII 上下文不动，避免把
                    # $a$b 这类正常写法也无谓地撑开。
                    # 注意要看「已攒文本的整体末尾」，不能只看 line[-1]，
                    # 因为最后一个元素可能是一整段文字而非单个字符。
                    if line and CJK.match("".join(line)[-1:] or ""):
                        line.append(" ")
                    line.append("$%s$" % tex)
                    stats["formulas_inline"] += 1

        rest = "".join(line).strip()
        if pending:
            rest = rest + "".join(pending)
        # 公式后紧跟中文时同样要补边界，否则渲染器不认定界结束。
        rest = _boundary_after_formula(rest)
        if rest:
            out.append(rest)
            out.append("")
            prev_display = False

    body = "\n".join(out)
    body = re.sub(r"\n{3,}", "\n\n", body).strip() + "\n"

    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(body)

    stats["lines"] = body.count("\n")
    stats["chars"] = len(body)
    if missing:
        stats["missing_formulas"] = missing
    save_json(os.path.join(args.workdir, "build_report.json"), stats)
    print(json.dumps(stats, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    force_utf8_stdio()
    main()
