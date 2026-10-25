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
    """图片在 md 里的引用路径：相对 md 文件定位 workdir 下的原件。

    用相对路径而不是绝对路径，md 连同目录整体拷贝/上传后仍然能显示。
    空格按 %20 转义，中文目录名保持原样（Markdown 允许）。

    EMF/WMF 优先指向 png/ 里渲染好的同名 PNG：浏览器与所有 Markdown
    预览器都不认 .emf，直接引用 media/xxx.emf 只会得到一个破图标。
    render 阶段已经把每张 EMF/WMF 渲成 png/xxx.png（Windows GDI+，或
    Linux 上的 ImageMagick/LibreOffice），没渲出来才退回原件。
    """
    from urllib.parse import quote
    stem, ext = os.path.splitext(fname)
    path = os.path.join(workdir, "media", fname)
    if ext.lower() in (".emf", ".wmf"):
        cand = os.path.join(workdir, "png", stem + ".png")
        if os.path.isfile(cand):
            path = cand
    rel = os.path.relpath(path, os.path.dirname(os.path.abspath(out_path)) or ".")
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


def _heading_text(i, text, struct):
    """标题行文字：剥掉行首插图占位，并把 Word 自动编号补回去。

    自动编号（`<w:numPr>`）在 docx 里不是一个字符，是渲染时现算的。段落
    文字因此是干净的 "位错密度的演化方程"，读者看到的光秃秃标题会丢失
    "iii)" 这一层信息——标题层级只剩 `####` 的井号数，而井号数还要跟
    同级条目的计数对上才看得出来。

    补回编号是**保真**而不是**编造**：marker 完全由 docx 的 numbering.xml
    推算（见 docstyle.py），Word 屏幕上显示的也正是这个字面值。
    原文已经带同样编号的（"i) 推导…"）不重复添加。
    """
    body = _strip_leading_figs(text).strip()
    marker = _auto_marker(struct, i)
    if marker and not body.startswith(marker):
        return marker + " " + body
    return body


def _strip_leading_figs(text):
    """标题文字去掉行首的插图占位，但要保留这些插图本身。

    真实现象：小节标题与它的图版被 Word 放在同一个段落里——
    `⟦EQ:fig:image135.emf⟧两组相反Burgers矢量的直线位错相互作用…`。
    这一段既该出标题（文字部分）又该出图（占位部分）。

    以前整段按标题输出，29 张图里有几张就这样被吞掉（标题行里的
    `⟦EQ:fig:…⟧` 不会被展开成图片引用，它只是字面文本）；而如果整段按
    正文走，标题就丢了。所以拆开：图在标题上方输出，文字做标题。
    """
    # 只看行首连续出现的插图占位；中间夹在文字里的不动。
    m = re.match(r"^(?:\u27e6EQ:fig:[^\u27e7]+\u27e7)+", text or "")
    return (text or "")[m.end():] if m else (text or "")


def _leading_figs(text):
    """取出标题行里被剥掉的那几个插图占位（按出现顺序）。"""
    m = re.match(r"^(?:\u27e6EQ:fig:[^\u27e7]+\u27e7)+", text or "")
    return re.findall(r"\u27e6EQ:fig:([^\u27e7]+)\u27e7", m.group(0)) if m else []


def _next_is_prose(paras, i):
    """下一个非空段落是不是纯文字（既非公式、也非图片）。

    用来把「章节标签」和「公式行标签」分开。实测中文技术文档里两种加粗短句
    长得一模一样（"线张力" vs "位移场"），唯一的区别在**后面跟着什么**：
      章节标签 → 后面是讲解该概念的正文（"采用虚功原理…"）
      行标签   → 后面直接是它标注的那条公式（`⟦EQ:oleObject59.bin⟧`）
    这条判据把 11 个候选里的 7 个假标题剔掉了，代价是漏掉少数紧跟公式的
    真标签（如 "程序流程图"）——那种情况留给 headings_suggested.json 人工确认。
    """
    for k in range(i + 1, len(paras)):
        t = (paras[k] or "").strip()
        if not t:
            continue
        # 段落必须以文字开头。`⟦EQ:…⟧或者` 这种「公式 + 残句」是公式的
        # 补充说明，不是标签后面的解释段落。
        if EQ_RE.match(t):
            return False
        return bool(EQ_RE.sub("", t).strip())
    return False


def _deepest_before(out, i):
    """取 i 之前最近 6 条已定标题里最深的级别，用来给无编号标签定深度。

    只看最近几条：往前看得太远会把「上一章深处的小节」当成当前深度。
    下限 3、上限 4——这类加粗标签实测都是小节内的分组名，深到 5、6 级
    会让 Markdown 目录出现一堆「位移场 / 应力场」并列的假标题。
    """
    lv = [(k, v) for k, v in out.items() if k < i]
    if not lv:
        return 3
    lv.sort()
    return max(3, min(max(v for _, v in lv[-6:]), 4))


def _style_level(struct, i):
    """段落自带的样式级别（Word 的 heading N / Title）。无则 None。"""
    if not struct or i >= len(struct):
        return None
    lv = (struct[i] or {}).get("style")
    return int(lv) if lv else None


def _auto_marker(struct, i):
    """段落由 Word 自动编号生成的标记（"iii)" / "i." 等）。无则空串。"""
    if not struct or i >= len(struct):
        return ""
    return (struct[i] or {}).get("marker") or ""


def _is_bold(struct, i):
    if not struct or i >= len(struct):
        return False
    return bool((struct[i] or {}).get("bold"))


def guess_headings(paras, struct=None, bold_headings=False):
    """生成标题建议 {索引: 级别}。保守但覆盖常见中文排版。

    struct 是 extract.py 落下的 docstruct.json（样式级别 / 自动编号 / 整段
    加粗），可缺省。它的作用是补上「文字里没有编号」的标题：
    中文 Word 文档的 "1．" / "A）" / "i）" 常由 Word 自动编号生成，
    段落文字里一个字符都没有，纯正则必然漏。

    bold_headings 控制「整段加粗的短标签也算标题」这条兜底，默认关。
    它确实能找回 "攀移力" / "线张力" / "程序流程图" 这类真标签，但中文
    文档里加粗同样大量用于强调正文（"位移场" / "应力场" 这类公式行标签
    就是加粗），一开就会在目录里堆出一批假标题。宁可少认，也不要脏目录：
    --no-bold-headings 可以关掉这条兜底。

    优先级：Word 样式 > 正则可识别的前缀 > 整段加粗 + 结构信号。
    样式最可信（outlineLvl 直接决定导航窗格），前缀次之，加粗只是旁证，
    所以加粗单独出现时要求更严：短、无句末标点、后面紧跟正文。
    """
    out = {}
    for i, raw in enumerate(paras):
        # 行首的插图占位先剥掉再判标题：段落常把「图 + 标题文字」放在一起
        # （`⟦EQ:fig:image135.emf⟧两组相反Burgers矢量…`）。占位符本身含
        # "EQ:"，不剥就会命中下面那条「含公式不可能是标题」，真标题被漏掉。
        t = _strip_leading_figs(raw.strip()).strip()
        # 样式标题不受长度限制：长标题也是标题。但公式占位仍要排除——
        # 含公式的段落被当成标题会把公式整条吞掉。
        style_lv = _style_level(struct, i)
        if style_lv and t and not EQ_RE.search(t):
            out[i] = style_lv
            continue
        if not t or len(t) > 90:
            continue
        # 已被公式占满的段落不可能是标题
        if EQ_RE.search(t):
            continue
        # 句号结尾基本是正文
        if t.endswith("。") or t.endswith("；"):
            continue

        # Word 自动编号：把合成出的 "iii)" 拼回段首再走同一套正则。
        # 不另写一套判断——两套判据迟早会打架。
        marker = _auto_marker(struct, i)
        probe = (marker + t) if marker else t
        if marker:
            m = re.match(r"^([ivxlcdm]{1,6})[）)]\s*(\S.*)$", probe, re.I)
            if m and len(t) <= 45:
                # 罗马数字编号从 iii) 起（这份文档的第 2 节），层级=4；
                # 从 i) 起的另一组编号是实验小节，层级同为 4。
                out[i] = 4
                continue
            m = re.match(r"^(\d+)[）)]\s*(\S.*)$", probe)
            if m and len(t) <= 45:
                out[i] = 3
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
        # 兜底：整段加粗的**短标签**。
        # 中文技术文档里 "1．/ A） / i）" 之外还有一批「只有加粗、无样式、
        # 无编号」的小标题（"攀移力" / "线张力" / "程序流程图"）。它们唯一
        # 的结构证据就是整段加粗。
        #
        # 但加粗也大量用于强调正文（"对各向同性弹性介质的张量 Green函数"
        # 就是一句正文），所以门槛必须收紧：长度 <= 12 且不以任何标点收尾。
        # 实测这份文档里 3~5 字的真标签能过、17 字以上的正文句子过不了。
        #
        # 级别取 5 而不是 3：这类标签总是出现在已由自动编号确定的小节
        # （#### = 2.3.x）之下，再深一级才符合原有层级。猜错级别只影响
        # 目录深度，猜错「是不是标题」会把正文变成标题——后者代价大得多。
        if bold_headings and _is_bold(struct, i) and len(t) <= 15 \
                and not re.search(r"[。；：，,:.！？、）)】]$", t) \
                and _next_is_prose(paras, i):
            out[i] = min(_deepest_before(out, i) + 1, 6)
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
    ap.add_argument("--no-bold-headings", dest="bold_headings",
                    action="store_false",
                    help="不把「整段加粗的短标签」当标题。默认把它们当标题："
                         "中文技术文档里这类小节标题（如「攀移力」「线张力」）"
                         "只有加粗、没有编号也没有样式，其它信号一概没有。"
                         "假标题已由两道判据挡住（不超过 15 字、且后面紧跟"
                         "正文而不是公式），仍嫌吵时用这个开关关掉。")
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
    struct = load_json(os.path.join(args.workdir, "docstruct.json"))
    if struct is not None and len(struct) != len(paras):
        # 索引对不上就不能用：标题会挂到错的段落。宁可退回纯正则。
        struct = None
    suggested = guess_headings(paras, struct, args.bold_headings)
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
             "image_transcribed": 0,    # Agent 已转写成 LaTeX 的图片公式数
             "figures": 0,              # 正文落成的示意图处数
             "orphan_punct_dropped": 0}  # 公式块后无处安放的孤立标点

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
            # 标题段落里的插图先落到标题上方，再输出标题文字。
            for fname in _leading_figs(text):
                ref = _img_rel_src(args.out, args.workdir, fname)
                if ref:
                    out.append("![图](%s)" % ref)
                    out.append("")
                    stats["figures"] += 1
            out.append("#" * min(int(level), 6) + " " + _heading_text(i, text, struct))
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
                if name.startswith("fig:"):
                    # 示意图：正文里必须留下可见的图片引用。它不参与公式
                    # 判定（不是公式，不该被拉出成块、也不该进 latex 表），
                    # 所以在这里单独接住，并标记成 figure 片让外层知道
                    # 这一段含有图（有图的段落不能当"公式独占段"处理）。
                    stats["figures"] += 1
                    fname = name[4:]
                    src = _img_rel_src(args.out, args.workdir, fname)
                    pieces.append(("figure",
                                   "![图](%s)<!-- 原文插图原件: media/%s -->"
                                   % (src, fname)))
                    continue
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
        n_fig = sum(1 for p in pieces if p[0] == "figure")
        real_text = "".join(p[1] for p in pieces
                            if p[0] in ("text", "figure")
                            and not p[1].lstrip().startswith("<!--")).strip()
        # 有插图的段落不能当作「公式独占段」：否则图会被 display 块挤掉。
        alone = (n_form == 1 and n_fig == 0 and not real_text)

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
            elif kind == "figure":
                # 插图：独立成行，前后留空行。混排进句子会让 Markdown
                # 把图片塞进行内上下文，多个公式/文字环绕时渲染成乱码。
                if pending:
                    line.append("".join(pending))
                    pending = []
                head = "".join(line).strip()
                if head:
                    out.append(_boundary_after_formula(head))
                    out.append("")
                    line = []
                out.append(val)
                out.append("")
                prev_display = False
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
        # 公式块后「无处安放」的孤立标点。
        #
        # 典型来源：源文档把一个公式单独放在一段、末尾还带一个「，」（
        # `⟦EQ:oleObject4.bin⟧，`），真正的句子在下一段继续（「其中…」）。
        # 公式已被判成 display 块、独立成段输出，这个「，」前面没有任何
        # 文字可粘——旧版会把它原样输出成一行单独的「，」，读者看到的是
        # 一个悬空的标点行（实测同一条段落里 5 处）。
        #
        # 中文排版里标点不能出现在行首，把「，」贴到下一段开头是错的；
        # 贴回 $$ 收盘行会破坏 T1（`$$` 必须独占一行）。所以只能丢弃，
        # 但要**记账**（orphan_punct_dropped），并在 build_report.json 里
        # 可见——本插件最忌讳的是静默丢东西，丢一个标点也一样。
        if not rest and pending:
            stats["orphan_punct_dropped"] += 1
            pending = []
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
