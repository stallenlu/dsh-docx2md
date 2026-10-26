#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""validate.py —— 校验生成的 Markdown 是否符合六条铁律。

用法:
    python validate.py <file.md> [--expect-formulas N] [--json]

退出码:
    0 = 全部通过
    1 = 有错误级问题

检查项（对应 METHODOLOGY 的六条铁律 + 结构校验）:
    T1  $$ 必须独占一行
    T2  行内公式不得跨行 / 不得含块级环境 / 不得含 \\ 与 &（见铁律 2）
    T3  相邻公式之间必须有空行或正文
    T4  \\tag 必须紧跟 \\end{aligned}（或 \\end{array} 等）
    T5  表格单元格内不得出现超长公式（T5_MAX_CELL_FORMULA 字符以上）
    T6  图片占位应为 HTML 注释
    T7  $$ 配对数为偶数（成对闭合）
    T8  公式编号连续、无重复（可选，--expect-formulas 时检查总数；
        display + inline + 图片公式引用 一起计入）
    T9  不得残留未解析的公式占位符或「未解析公式」注释
    T11 图片公式的引用必须能解析到真实文件（断链=破图=公式又消失了）
    T12 行内强调标记必须能闭合（`**a  **` 这类右定界符前带空白的写法不闭合）
    T13 「一行公式长得很高」= 图里原本是多行，转写时被并成了一行
        （依赖 eq_index.json + media/ 的真实像素，见 check_media_rows）
    T14 「图很宽、转写却几乎没字」= 转写漏了内容（中文与公式混排的长条
        最容易整段没转进来，见 check_media_density）
"""
import argparse
import json
import os
import re
import sys
from docxfile import force_utf8_stdio

BLOCK_ENVS = ("aligned", "array", "matrix", "pmatrix", "bmatrix",
              "cases", "split", "gather", "align", "eqnarray")

# 表格单元格里的公式字符数上限。SKILL.md 与 README 的铁律 5 写的是
# 「超过约 60 字符就移出表格」，这里就用 60 —— 早先实现是 400，与文档不符，
# 于是 100 字符的公式塞进单元格也能拿到 PASS。
T5_MAX_CELL_FORMULA = 60

# 行内公式里不允许出现的字符：换行由正则的 [^$\n] 挡住，这里管 \\ 与 &。
# 铁律 2 原文：`$...$` 里不能出现 `\begin{aligned}`、`\\`、`&`。
#   - `\\`（LaTeX 换行）会让行内公式在渲染器里断行
#   - `&`（对齐符）只在 aligned/array 环境里有意义，出现在行内就是结构错位
# 注意 `\{`、`\}`、`\&` 是转义的字面量，必须排除，否则会误报。
INLINE_BAD_LINEBREAK = re.compile(r"(?<!\\)\\\\")      # 未被转义的 \\
INLINE_BAD_ALIGN = re.compile(r"(?<!\\)&")              # 未被转义的 &

# CJK 字符与全角标点：公式与它们相邻时缺边界，部分渲染器就不认定界符。
CJK_CHAR = re.compile(r"[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]")

def _bad_emphasis(prose):
    """找出无法闭合的 ** / * 定界符对，返回 [(位置, 标记)]。

    CommonMark 的 flanking 规则：开定界符后面不能紧跟空白，闭定界符前面
    不能紧跟空白。违反时不是「渲染得难看」，而是星号按字面输出——正文里
    凭空多出 `**`，看起来像转义 bug，但校验器原来完全不查。

    不是假想问题：docx 里整段加粗的 run 常带尾随空格（"应用公式   "），
    早期实现直接 `**%s**` 就生成 `**应用公式   **`，真实文档里 4 处。

    只做**成对**判断：单个 `**` 是合法的左侧定界符（`正常 **加粗** 文本`
    里第一个 `**` 前面是空格，但它没坏）。成对扫描才不会误报。
    """
    bad = []
    for mark in ("**", "*"):
        pos = []
        i = 0
        while True:
            j = prose.find(mark, i)
            if j < 0:
                break
            # "**" 扫描时跳过 "***" 里的重叠，避免重复计数
            pos.append(j)
            i = j + len(mark)
        # 两两配对：0-1、2-3 … 奇数个则最后一个悬空（也是错误）
        for k in range(0, len(pos) - 1, 2):
            op, cl = pos[k], pos[k + 1]
            after_open = prose[op + len(mark):op + len(mark) + 1]
            before_close = prose[cl - 1:cl] if cl > 0 else ""
            if after_open.isspace() or before_close.isspace():
                bad.append((op, mark))
                break
    return bad


def _inline_bodies(line):
    """找出这一行里所有行内公式（$...$）的内容。"""
    return [m.group(1) for m in re.finditer(r"(?<!\$)\$([^$\n]+)\$(?!\$)", line)]


def _cell_formula_lengths(line):
    """表格行里每个单元格内公式的 LaTeX 长度（不含 $ 定界符本身）。"""
    out = []
    for cell in line.strip().strip("|").split("|"):
        for m in re.finditer(r"\$([^$]+)\$", cell):
            out.append((cell.strip(), len(m.group(1))))
    return out


def split_formula_rows(tex):
    """按「原本分行」的痕迹拆成若干行，返回 list（已是完整的一行 LaTeX）。

    转写的人把多行推导并成一行时，行间的 `\\quad` / `\\qquad` 会留在里面
    ——它们就是分行的痕迹，`\\qquad` 的间距更大（原文里通常对应更复杂的
    分行），两者混用时以出现次数更多的那个为准，避免只切一半。

    只在**深度 0** 处切：`{c,\\quad d}`、`\\left(a,\\quad b\\right)` 里的
    分隔符属于函数体内部，切了就把它拆成两个坏参数。
    """
    if not tex:
        return []
    n_q, n_qq = tex.count("\\quad"), tex.count("\\qquad")
    # \\qquad 里含 \\quad 子串，计数要先扣掉。
    n_q -= n_qq
    marker = "\\qquad" if n_qq > n_q else "\\quad"
    parts, depth, cur, i = [], 0, [], 0
    while i < len(tex):
        if tex.startswith(marker, i) and depth == 0:
            while i < len(tex) and tex[i] in " \t":
                i += 1
            i += len(marker)
            while i < len(tex) and tex[i] in " \t":
                i += 1
            parts.append("".join(cur))
            cur = []
            continue
        if tex.startswith("\\left", i) or tex.startswith("\\begin", i):
            depth += 1
        elif tex.startswith("\\right", i) or tex.startswith("\\end", i):
            depth -= 1
        elif tex[i] == "{":
            depth += 1
        elif tex[i] == "}":
            depth -= 1
        cur.append(tex[i])
        i += 1
    parts.append("".join(cur))
    # 切点前的逗号是原文的行尾标点，跟着分隔符一起去掉：留着会让读者读
    # 到「a,」独占一行，看着像断句错了。
    parts = [p.rstrip().rstrip(",").rstrip() for p in parts if p.strip()]
    return [p for p in parts if p]


def join_formula_rows(parts):
    """把拆好的若干行拼回一个 aligned 环境。"""
    if len(parts) < 2:
        return None
    return "\\begin{aligned}\n%s\n\\end{aligned}" % " \\\\\n".join(parts)


def _tex_rows(tex):
    """一段 LaTeX 排出来是几行：1 行 + 显式换行符数。

    不数 \\begin{aligned} 这些环境本身——它们内部一定也用 \\\\ 分行，数环
    境会把同一个换行算两遍。对,\\quad 连写的单行公式，这里老老实实给 1，
    于是「图很高但这里只有 1 行」就成了漏换行的证据。
    """
    if not tex:
        return 0
    return 1 + len(re.findall(r"(?<!\\)\\\\", tex))


def _sig(s):
    """抹掉空白与换行符后的「字形」：用来认 md 里是不是同一条公式。

    补好分行的公式在 md 里是同一串 LaTeX、只是多了 `\\\\` 与换行。按字符
    逐个比对会认不出来（间距不同），抹掉这些排版字符后两者完全一致。
    """
    return re.sub(r"[\s\\]+", "", s or "")


def _tex_as_built(md, tex):
    """这条公式在产物 md 里实际的样子；不在 md 里则返回 None。

    md 里没有公式名（⟦EQ:…⟧ 早已换成 LaTeX），只能按**内容**认。返回值要
    么是原样单行（说明漏了换行），要么已拆成多行（说明修好了）——后者让
    后面的 check_media_rows 直接跳过，不再报。
    """
    if not tex:
        return None
    sig_md = _sig(md)
    if _sig(tex) and _sig(tex) in sig_md:
        return tex                      # 原样单行进了产物
    parts = split_formula_rows(tex)
    if len(parts) >= 2:
        joined = " \\\\\n".join(parts)
        if _sig(joined) in sig_md:
            return joined               # 已拆成多行
    return None


def check_media_rows(eq_index, sizes, latex):
    """T13: 图里明明是多行、转写却成了单行的公式。

    这类错从 Markdown 自身看不出来：`a,\\quad b,\\quad c` 六条铁律全过、
    渲染也正常，但原文这一处是三行推导，读者拿到的是一行挤在一起的长式。

    主证据是**转写里残留的分行痕迹**：把多行并成一行的人几乎总是用
    `\\quad` / `\\qquad` 连接（Word 行间就是这么对齐的），split_formula_rows
    数出 ≥2 段就说明原文不止一行。光有痕迹还不够——"a=\\frac{1}{2}b,
    \\quad c=d" 这种真·单行也带 \\quad，所以再用像素高度把关：图的高度
    必须明显超过"一行"的高度。

    行高**自标定**：取本文档里那些又短、又没有大算符（分式/积分/求和会
    把单行撑高一倍）的单行公式的中位高度。实测 145 公式文档基线 28px，
    两行公式最低 49px，1.6 倍处干净分开；换个字号的文档基线自动跟着变。
    曾经用过"多行公式 h/rows"标定，实测被分式行抬高到 68px，漏掉了
    oleObject121（98×111 的三行公式，用户报的正是这条）。

    返回 [(公式名, 宽, 高, 行数, LaTeX)]。
    """
    if not sizes or not latex:
        return [], 0.0
    # 大算符：出现任何一条，这一行的高度就不是"普通行高"，不能进基线。
    TALL = ("frac", "int", "sum", "oint", "iint", "prod", "lim", "left",
            "right", "Big", "bigg", "sqrt")
    plain = []
    for eq in eq_index:
        key, img = eq.get("stream"), eq.get("img")
        if not key or not img or img not in sizes:
            continue
        tex = latex.get(key)
        if not tex or _tex_rows(tex) != 1 or len(tex) > 40:
            continue
        if any(("\\%s" % x) in tex for x in TALL):
            continue
        w, h = sizes[img]
        if w and h:
            plain.append(h)
    if not plain:
        # 没有可用的"普通单行"样本就无从标定：宁可不报，也不拿猜的阈值
        # 去判整篇文档。
        return [], 0.0
    plain.sort()
    row_h = plain[len(plain) // 2]

    hits = []
    for eq in eq_index:
        key, img = eq.get("stream"), eq.get("img")
        if not key or not img or img not in sizes:
            continue
        tex = latex.get(key)
        if not tex or _tex_rows(tex) != 1:
            continue
        seps = len(split_formula_rows(tex))
        if seps < 2:
            # 没有分行痕迹就不报：大 ∑/积分会把单行公式撑到两行高
            # （实测 oleObject116 "α=∑ρ_b b" 115×58 就是单行），猜不得。
            continue
        w, h = sizes[img]
        if not w or not h or h < 1.6 * row_h:
            continue
        hits.append((key, w, h, seps, tex))
    hits.sort(key=lambda r: -r[2])
    return hits, row_h


def formula_glyphs(tex):
    """数一条 LaTeX 里有多少个「会占宽度的字形」。

    T14 用它把「图有多宽」换算成「每字形多少像素」。\\text{中文} 里的每个
    汉字都算一个字形（汉字本来就宽）；\\alpha 这类命令算一个；花括号、
    定界、空白、运算符标点不算。
    """
    if not tex:
        return 0
    s = re.sub(r"\\(?:quad|qquad)", " ", tex)
    s = re.sub(r"\\(?:text|textrm|textnormal|mathrm|mathbf|boldsymbol|"
               r"mathit|textbf|mathsf|mathcal|hat|widehat|tilde|bar|vec|"
               r"overline|underline|dot|ddot|operatorname)\{([^{}]*)\}",
               r"\1", s)
    s = re.sub(r"\\[a-zA-Z]+", "G", s)          # \alpha、\nabla… 一个字形
    s = re.sub(r"\\.", "G", s)                  # \{ \& \% 等转义符号
    s = re.sub(r"[{}_$&^~ ,.;:!?（）()\[\]\-+=*/|]", "", s)
    return len(s)


def check_media_density(eq_index, sizes, latex):
    """T14: 图很宽、转写却几乎没字 —— 转写漏了内容。

    与 T13 是两类错：T13 是「内容都在、只是没分行」，T14 是「内容根本没
    转全」。实测踩过：一条 1109×43 的公式带（局部受力方向 b_xσ_xx，位错
    线方向…Burgers矢量 b=b(1,0,0)…单向拉伸σ_xx，中文与公式混排），转写
    只留下末尾的 `\\sigma_{zz}` —— 连下标都读错了（实际是 xx）。默认
    max-width=1600 渲染后这种细长条只有 62px 高，读图时整段中文直接糊掉。

    判据是**像素密度**：宽度 ÷ 字形数 = 每个字形占多少像素。本文档 125 条
    公式的中位数 16px，最宽的合法值是 40px（2.5 倍）；漏转的那条是 370px
    （23 倍）。取 6 倍中位数做门槛，两边都留了一倍以上的余量。

    返回 ([(公式名, 宽, 高, 字形数, 密度, LaTeX)], 中位密度)。
    """
    if not sizes or not latex:
        return [], 0.0
    dens = []
    for eq in eq_index:
        key, img = eq.get("stream"), eq.get("img")
        if not key or not img or img not in sizes or eq.get("in_canvas"):
            continue
        tex = latex.get(key)
        if not tex:
            continue
        w, h = sizes[img]
        g = formula_glyphs(tex)
        if w and h and g:
            dens.append(w / g)
    if not dens:
        return [], 0.0
    dens.sort()
    med = dens[len(dens) // 2]
    hits = []
    for eq in eq_index:
        key, img = eq.get("stream"), eq.get("img")
        if not key or not img or img not in sizes or eq.get("in_canvas"):
            continue
        tex = latex.get(key)
        w, h = sizes[img]
        g = formula_glyphs(tex) if tex else 0
        if not w or not h or w < 200:
            continue                            # 小图噪声大，不判
        # 一个字形都没有 = 转写是空的或全是符号，按 1 个字形算密度（必超
        # 门槛）。用 inf 会让"中位的多少倍"打印成 inf，读者看不懂。
        d = w / max(1, g)
        if d > 6 * med:
            hits.append((key, w, h, g, d, tex or ""))
    hits.sort(key=lambda r: -r[4])
    return hits, med


def check(md_lines, t5_max_cell_formula=T5_MAX_CELL_FORMULA):
    errors, warns = [], []

    n_dollar = sum(ln.count("$$") for ln in md_lines)
    if n_dollar % 2 != 0:
        errors.append("T7: $$ 出现 %d 次（奇数），存在未闭合的公式块" % n_dollar)

    for i, ln in enumerate(md_lines, 1):
        s = ln.strip()

        # T1: $$ 独占一行
        if "$$" in ln and s not in ("$$",):
            # 允许行尾紧跟 \tag 的收尾行？不：收尾 $$ 必须独占
            errors.append("T1: 第 %d 行 $$ 未独占一行: %s" % (i, s[:90]))

        # T2: 行内公式不得含块级环境，也不得含 \\ 与 &
        for body in _inline_bodies(ln):
            if "\\begin{" in body:
                errors.append("T2: 第 %d 行行内公式含块级环境: %s" % (i, body[:80]))
            elif INLINE_BAD_LINEBREAK.search(body):
                errors.append(
                    "T2: 第 %d 行行内公式含 \\\\（换行符），应改用 $$ 块: %s"
                    % (i, body[:80]))
            elif INLINE_BAD_ALIGN.search(body):
                errors.append(
                    "T2: 第 %d 行行内公式含未转义的 &（对齐符），应改用 $$ 块: %s"
                    % (i, body[:80]))

        # T10: 行内公式与中文之间必须有边界。
        # 不少渲染器（Pandoc 的 tex_math_dollars、部分 KaTeX 集成）只在 $ 与
        # 前后字符有明确边界时才识别为公式；紧贴中文时按普通文本处理，用户
        # 看到的就是「LaTeX 源码原样出现」或公式整段不显示。实测一份真实文档
        # 里有 21 处属于这种情况。
        for m in re.finditer(r"(?<!\$)\$([^$\n]+?)\$(?!\$)", ln):
            before = ln[m.start() - 1] if m.start() > 0 else ""
            after = ln[m.end()] if m.end() < len(ln) else ""
            if before and not before.isspace() and CJK_CHAR.match(before):
                warns.append(
                    "T10: 第 %d 行公式紧贴前面的中文（应为「中文 $公式$」）: %s"
                    % (i, ln[max(0, m.start() - 20):m.end() + 20]))
            if after and not after.isspace() and CJK_CHAR.match(after):
                warns.append(
                    "T10: 第 %d 行公式紧贴后面的中文（应为「$公式$ 中文」）: %s"
                    % (i, ln[max(0, m.start() - 20):m.end() + 20]))

        # T5: 表格行里的超长公式
        if s.startswith("|") and s.count("|") >= 2:
            for cell, texlen in _cell_formula_lengths(s):
                if texlen > t5_max_cell_formula:
                    warns.append(
                        "T5: 第 %d 行表格单元格内公式 %d 字符（上限 %d），建议移出表格: %s"
                        % (i, texlen, t5_max_cell_formula, cell[:60]))

        # T6: 图片占位
        if re.search(r"!\[[^\]]*\]\(\s*\)", ln):
            warns.append("T6: 第 %d 行是空图片链接，建议改用 HTML 注释占位" % i)

        # T12: 强调标记必须能闭合。先去掉行内公式与行内代码——里面的 *
        # 是公式/代码的一部分，不是强调。
        _prose = re.sub(r"\$[^$\n]*\$", "", ln)
        _prose = re.sub(r"`[^`\n]*`", "", _prose)
        _bad = _bad_emphasis(_prose)
        if _bad:
            errors.append(
                "T12: 第 %d 行 %s 无法闭合（定界符内侧有空白，会按字面输出）: %s"
                % (i, _bad[0][1], s[:90]))

    # T4: \tag 紧跟 \end{...}
    for i, ln in enumerate(md_lines, 1):
        if "\\tag" in ln:
            before = ln.split("\\tag")[0].rstrip()
            if not re.search(r"\\end\{[a-z]+\}$", before) and "\\\\" not in before:
                # 单行公式 \tag 是合法的，只警告
                if not re.match(r"^\s*\$?\$?.*\\tag\{", ln):
                    warns.append("T4: 第 %d 行 \\tag 未紧跟 \\end{}" % i)

    # T3: 相邻公式块之间要有空行
    in_block = False
    prev_close = None
    for i, ln in enumerate(md_lines, 1):
        s = ln.strip()
        if s == "$$":
            if not in_block:
                in_block = True
                if prev_close is not None and i - prev_close == 1:
                    errors.append("T3: 第 %d 行的公式块紧贴上一个块结尾" % i)
            else:
                in_block = False
                prev_close = i

    return errors, warns


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--expect-formulas", type=int, default=None)
    # 画布（流程图）内的公式已画进 PNG，不在 md 里单独出现。校验时要
    # 从期望数里扣掉它们，否则带流程图的文档永远差这么多个。
    ap.add_argument("--eq-index", default=None,
                    help="eq_index.json 路径；给了就能识别画布内公式。")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--media", default=None,
                    help="extract 的 media/ 目录（或仍含公式预览图的 docx）。"
                         "给了才跑 T13/T14：这两项要靠预览图的真实像素判断。")
    ap.add_argument("--latex", default=None,
                    help="latex.json 路径。T13/T14 要拿它与图片尺寸对照。")
    ap.add_argument("--t5-max-cell-formula", type=int, default=T5_MAX_CELL_FORMULA,
                    help="表格单元格内公式的字符上限（默认 %d）。"
                         "超过就警告，建议把公式移出表格。" % T5_MAX_CELL_FORMULA)
    args = ap.parse_args()

    if args.t5_max_cell_formula < 1:
        sys.stderr.write("validate 输入错误: --t5-max-cell-formula 必须是正整数\n")
        sys.exit(2)

    with open(args.file, "r", encoding="utf-8") as f:
        md = f.read()
    lines = md.split("\n")

    errors, warns = check(lines, t5_max_cell_formula=args.t5_max_cell_formula)

    # 空文档不算通过：没有内容就没有可验证的东西，PASS 会给出错误的信心。
    if not md.strip():
        errors.append("T0: 文件为空，没有可校验的内容")

    # 未替换的公式。build.py 遇到 latex 表里没有的公式时，会把它写成
    # <!-- 未解析公式: oleObject7.bin -->，而不是留下 ⟦EQ:...⟧ 字面量
    # （两种形态都要查，因为别的工具可能直接留下占位符）。
    # 这类文档的 Markdown 结构完全合规，六条铁律全过，但正文里少了公式——
    # 只查铁律会给出 PASS，用户却拿到一份读不了的文档。
    left = re.findall(r"\u27e6EQ:([^\u27e7]+)\u27e7", md)
    if left:
        uniq = sorted(set(left))
        errors.append(
            "T9: 有 %d 处公式占位符未替换（%d 个不同公式：%s）"
            % (len(left), len(uniq), ", ".join(uniq[:12])
               + ("…" if len(uniq) > 12 else "")))
    marked = re.findall(r"<!--\s*未解析公式:\s*([^\s>]+)\s*-->", md)
    if marked:
        uniq = sorted(set(marked))
        errors.append(
            "T9: 有 %d 处公式未解析（%d 个不同公式：%s）"
            % (len(marked), len(uniq), ", ".join(uniq[:12])
               + ("…" if len(uniq) > 12 else "")))

    # 图片公式（build 对未转写的 ⟦EQ:img:…⟧ 生成的可见引用）。
    # 它不算 T9——图片公式没有"未解析"问题，引用就是它的合法出路。
    img_refs = re.findall(r"!\[图片公式 ([^\]]+)\]\(([^)]+)\)", md)
    # T11: 图片公式的引用目标必须真实存在（相对 md 所在目录解析）。
    # 断链 = 用户点开是破图，等于公式换了种方式消失。
    if img_refs:
        base = os.path.dirname(os.path.abspath(args.file))
        for alt, src in img_refs:
            from urllib.parse import unquote
            p = os.path.normpath(os.path.join(base, unquote(src)))
            if not os.path.exists(p):
                errors.append("T11: 图片公式引用断链: %s（%s）" % (src, alt))

    # T13: 图里有好几行、转写却是一行的公式。
    # 证据来自 Word 存的公式预览图像素，self-calibrate 每行高度，所以不
    # 依赖写死的阈值。没有 media/ 或 latex.json 就跳过——不能凭 md 判断。
    row_hits = []
    if args.media and args.latex and args.eq_index:
        try:
            import json as _json
            import zipfile
            import imgsize
            with open(args.eq_index, encoding="utf-8") as f:
                _eq = _json.load(f)
            with open(args.latex, encoding="utf-8") as f:
                _lt = _json.load(f)
            sizes = {}
            media = args.media
            if os.path.isdir(media):
                for name in os.listdir(media):
                    p = os.path.join(media, name)
                    if not os.path.isfile(p):
                        continue
                    with open(p, "rb") as fp:
                        try:
                            sizes[name] = imgsize.read_size(fp.read())
                        except Exception:
                            continue
            elif zipfile.is_zipfile(media):
                with zipfile.ZipFile(media) as z:
                    for n in z.namelist():
                        if os.path.basename(n).startswith("image"):
                            try:
                                sizes[os.path.basename(n)] = \
                                    imgsize.read_size(z.read(n))
                            except Exception:
                                continue
            # 画布内的公式已经画进流程图 PNG，它的尺寸是整张画布的形状，
            # 不是公式行数——参与标定会把 px/行 拉到荒谬的量级。
            #
            # 对照的对象必须是**产物里那条公式**，不能是 latex.json 里的原
            # 文：build 已经把漏掉的分行补成了 aligned，拿原文比就是拿修复
            # 前的状态去判修复后的文件，永远报错。
            #
            # md 里没有公式名（⟦EQ:…⟧ 早已替换成 LaTeX），所以按**内容**认：
            # 补好的 aligned 出现在 md 里就是已修；原样单行出现在 md 里就是
            # 漏了。两者都不在（该公式没进产物）时不报，由 T8 管。
            _eq = [e for e in _eq if not e.get("in_canvas")]
            _cmp = {}
            for _k, _tex in _lt.items():
                _v = _tex_as_built(md, _tex)
                if _v is not None:
                    _cmp[_k] = _v
            row_hits, _pxr = check_media_rows(_eq, sizes, _cmp)
            if row_hits:
                # 转写时把多行公式并成一行是**内容错**（读者少看两个推导步
                # 骤），不是排版瑕疵：渲染照样正常、六条铁律全过，但推导中间
                # 少了一步。按「看得出来的错」与「静默丢内容」的分级，这是
                # 错误而不是警告，必须让人回头重转写，而不是"知道就好"。
                for name, w, h, rows, tex in row_hits[:20]:
                    errors.append(
                        "T13: %s 的预览图 %dx%d（估算 %d 行）"
                        "但转写成了一行，疑似漏了换行: %s"
                        % (name, w, h, rows, tex[:70]))
                if len(row_hits) > 20:
                    errors.append("T13: 另有 %d 处同类问题未列出"
                                  % (len(row_hits) - 20))
            # T14: 图很宽、转写却几乎没字（内容根本没转全）。
            # 对照的是 latex.json 的**转写本身**——T14 查的就是转写漏没漏，
            # 与产物 md 无关。
            dens_hits, dens_med = check_media_density(_eq, sizes, _lt)
            for name, w, h, g, d, tex in dens_hits[:10]:
                errors.append(
                    "T14: %s 的预览图 %dx%d 但转写只有 %d 个字形"
                    "（每字形 %.0fpx，全体中位的 %.0f 倍），疑似漏转内容: %s"
                    % (name, w, h, g, d, d / max(1.0, dens_med), tex[:50]))
            if len(dens_hits) > 10:
                errors.append("T14: 另有 %d 处同类问题未列出"
                              % (len(dens_hits) - 10))
        except (OSError, ValueError, ImportError):
            row_hits = []

    n_display = sum(1 for ln in lines if ln.strip() == "$$") // 2
    n_inline = len(re.findall(r"(?<!\$)\$(?!\$)[^$\n]+\$(?!\$)", md))
    n_tag = len(re.findall(r"\\tag\{", md))

    report = {
        "file": args.file,
        "lines": len(lines),
        "chars": len(md),
        "display_blocks": n_display,
        "inline_formulas": n_inline,
        "image_formulas": len(img_refs),
        "tags": n_tag,
        "errors": errors[:40],
        "warnings": warns[:40],
        "error_count": len(errors),
        "warning_count": len(warns),
        "status": "PASS" if not errors else "FAIL",
    }
    if args.expect_formulas is not None:
        report["expected_formulas"] = args.expect_formulas
        # expectFormulas 用的是 extract 的 formulas_total（含图片公式）。
        #
        # 画布（流程图）内的公式**不在 md 里单独出现**：它们已经画进那张
        # PNG 了。若仍按"md 里的公式数"去对全量公式数，带流程图的文档会
        # 永远差这么多个。所以把画布内公式数从期望值里扣掉——它们确实
        # 在文档里，只是以图片形态存在，和 image_formulas 是同一类。
        n_canvas = 0
        if args.eq_index:
            try:
                import json as _json
                with open(args.eq_index, encoding="utf-8") as f:
                    _rows = _json.load(f)
                n_canvas = sum(1 for r in _rows if r.get("in_canvas"))
            except (OSError, ValueError):
                n_canvas = 0
        expected_visible = args.expect_formulas - n_canvas
        got = n_display + n_inline + len(img_refs)
        if got != expected_visible:
            report["status"] = "FAIL"
            report["errors"].append(
                "T8: 公式数 %d（display %d + inline %d + image %d）与预期 %d 不符"
                % (got, n_display, n_inline, len(img_refs), expected_visible))
            report["error_count"] += 1
        if n_canvas:
            report["canvas_formulas"] = n_canvas

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=1))
    else:
        print("文件: %s" % args.file)
        print("行数: %d  字符: %d" % (len(lines), len(md)))
        print("公式: display %d / inline %d / image %d / tag %d"
              % (n_display, n_inline, len(img_refs), n_tag))
        print("错误: %d  警告: %d  状态: %s"
              % (len(errors), len(warns), report["status"]))
        for e in errors[:20]:
            print("  [ERROR] " + e)
        for w in warns[:20]:
            print("  [WARN ] " + w)

    sys.exit(0 if not errors else 1)


if __name__ == "__main__":
    force_utf8_stdio()
    main()
