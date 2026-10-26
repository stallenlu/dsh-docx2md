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
import imgsize
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


def render_runs(runs):
    """把带格式的 Run 列表还原成 Markdown 行内语法。

    为什么不能只拼文字：docx 里的加粗/斜体/下划线/删除线/上下标/超链接/
    脚注都是**明确记录的结构**，不是排版噪声。一个格式转换器丢掉它们，
    转出的 md 就只是"能读"，而不是"等价"。

    相邻的加粗与斜体在 Markdown 里会粘成 "***"，定界符解析歧义，渲染
    出来是字面星号。所以：一旦上一个输出以 "*" 结尾、这一段的标记又以
    "*" 开头，就改用 HTML 标签（<strong>/<em>）——渲染结果一致，且没有
    歧义。GFM 与绝大多数渲染器都支持行内 HTML。
    """
    out = []
    notes = []

    def wrap(text, tag):
        """加行内标记，但把首尾空白挪到标记**外面**。

        为什么必须挪：CommonMark 规定 `**a  **` 这种「右定界符前面是空白」
        的写法**不能闭合**，渲染出来是两个字面星号。docx 里加粗 run 后面
        常常带一个尾随空格（"由位错密度求塑性变形场，应用公式   " 整段加粗，
        末尾三个空格），直接 `**%s**` 就会踩这个坑——语法上无效、内容上看
        起来像 bug，而且校验器不查这个。

        把空白挪到标记外，`**a**  ` 语义等价、渲染正确。
        """
        lead = text[:len(text) - len(text.lstrip())]
        trail = text[len(text.rstrip()):]
        core = text.strip()
        if not core:
            return text
        if tag == "bold":
            body = "**%s**" % core
        elif tag == "italic":
            body = "*%s*" % core
        elif tag == "underline":
            body = "<u>%s</u>" % core
        elif tag == "strike":
            body = "~~%s~~" % core
        else:
            body = core
        return lead + body + trail

    def push(s):
        # 与上一段以 * 结尾、这一段以 * 开头 → 改 HTML，避免 *** 歧义
        if out and out[-1].endswith("*") and s.startswith("*"):
            if s.startswith("**") and s.endswith("**") and len(s) > 4:
                s = "<strong>%s</strong>" % s[2:-2]
            elif s.startswith("*") and s.endswith("*") and len(s) > 2:
                s = "<em>%s</em>" % s[1:-1]
        out.append(s)

    for r in runs:
        t = r.get("text") or ""
        if t:
            esc = t
            # 上下标在 Markdown 里没有标准语法，用 HTML 的 sup/sub：
            # 通用渲染器都认，且不破坏行内公式的 $ 定界。
            if r.get("vert") == "sup":
                esc = "<sup>%s</sup>" % esc
            elif r.get("vert") == "sub":
                esc = "<sub>%s</sub>" % esc
            if r.get("strike") and esc.strip():
                esc = wrap(esc, "strike")
            if r.get("underline") and esc.strip():
                esc = wrap(esc, "underline")
            if r.get("italic") and esc.strip():
                esc = wrap(esc, "italic")
            if r.get("bold") and esc.strip():
                esc = wrap(esc, "bold")
            if r.get("link"):
                esc = "[%s](%s)" % (esc, r["link"])
            push(esc)
        if r.get("footnote"):
            # 脚注正文没有标准 Markdown 语法，用 GFM 脚注的退化形式：
            # 保留可见的 [^fnN] 标记 + 文末集中列出。绝不静默丢。
            notes.append(r["footnote"])
            push("[^fn%d]" % len(notes))
    return "".join(out), notes


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
    # 标题行不能带 ** / * / <u> 这类行内标记：Markdown 的 "# " 后面
    # 再加粗会渲染成字面星号（标题本身已经是加粗了）。行内标记在正文
    # 段落里照常保留，标题里剥掉——这是渲染约束，不是丢格式。
    body = re.sub(r"\*\*(.+?)\*\*", r"\1", body)
    body = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"\1", body)
    body = re.sub(r"</?u>", "", body)
    body = re.sub(r"~~(.+?)~~", r"\1", body)
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


CAPTION_RE = re.compile(
    r"^(?:fig(?:ure)?\.?\s*\d|tab(?:le|les)?\.?\s*\d|exhibit\s*\d"
    r"|scheme\s*\d|eq(?:uation)?\.?\s*\d|图\s*\d|表\s*\d"
    r"|式\s*\d|公式\s*\d|算法\s*\d|程序\s*\d)", re.I)


def _is_caption(t):
    """题注（图1 / Table 2 / 式(3)…）不是标题：它标注的是旁边的图/表/公式，
    升进目录会让目录里混一堆"Figure 1. xxx"。英文论文与中文论文都有这个
    形态——而且题注惯例恰好就是居中加粗，与"居中标签"撞形。"""
    return bool(CAPTION_RE.match(t or ""))


def _para_jc(struct, i):
    """段落对齐（"center" / "right" / "both"）。左对齐与缺省都是 None。"""
    if not struct or i >= len(struct):
        return None
    return (struct[i] or {}).get("jc") or None


def _para_color(struct, i):
    """段落字体颜色（RRGGBB）。黑（000000）与 auto 都不算。"""
    if not struct or i >= len(struct):
        return None
    return (struct[i] or {}).get("color") or None


def _para_indent_chars(struct, i):
    """首行缩进折算成「几个字」。0 = 不缩进。

    Word 给两种口径，优先用 firstLineChars（它直接就是"几个字"的 100 倍，
    与字号无关，任何文档都成立）：
      firstLineChars=200 → 2 字；firstLineChars=50 → 0 字（那 0.5 字其实是
      编号悬挂，缩进量属于编号，不该复制成正文缩进）。
    段落只写了 firstLine（twip，绝对长度）时，必须除以**这篇文档自己的**
    全角字宽才能换算成字数——字宽由 extract 记在 charTwip 里（从
    docDefaults/Normal 的字号推得，见 docstyle.default_char_twip）。
    写死 210（10.5pt 的字宽）是文档特化：12pt 文档上 2 字缩进会被算成
    1.75 字而整个丢掉，16pt 文档上会被算成 3 字。charTwip 缺失（旧产物）
    才退回 210。

    上限 4：正文缩进超过 4 字基本都是排版事故或悬挂缩进，照抄会把段落顶
    到页面中间去。
    """
    if not struct or i >= len(struct):
        return 0
    rec = struct[i] or {}
    try:
        flc = int(rec.get("firstLineChars") or 0)
    except (TypeError, ValueError):
        flc = 0
    if flc > 0:
        n = flc // 100
    else:
        try:
            fl = int(rec.get("firstLine") or 0)
        except (TypeError, ValueError):
            fl = 0
        try:
            cw = int(rec.get("charTwip") or 0)
        except (TypeError, ValueError):
            cw = 0
        n = fl // (cw if cw > 0 else 210)
    return max(0, min(n, 4))


def _align_attr(jc):
    """w:jc → HTML align 值。center/right 才有意义；both(两端对齐)/左对齐
    是渲染器默认，写成 align="both" 是非法值，等于没写。"""
    return jc if jc in ("center", "right") else None


def wrap_para_format(text, jc=None, color=None, indent=0):
    """把「居中 / 非黑颜色 / 首行缩进」三项排版信号落成 Markdown 里的 HTML。

    这三样在纯 Markdown 里都没有对应语法：
      * 居中——Markdown 无居中语法，只能借 <p align="center">；
      * 颜色——同上，只能借 <span style="color:...">；
      * 首行缩进——行首空白会被 Markdown 解析器整段吃掉，只能借全角空格
        （U+3000）顶出来，它在任何渲染器里都是实打实的两个字符宽的空白。

    不还原的后果是「静默丢格式」：正文一字不差，但原文的层级提示（居中 +
    蓝色表示这是小节标题）与段落起始（首行缩进两格）全没了，正是用户报的
    三类问题里的两类。
    """
    if not text:
        return text
    body = text
    # 首行缩进：只在没有更强信号时补。居中/着色的段落通常是标题性文字，
    # 再补两个全角空格只会让居中看起来偏了。
    if indent and not jc and not color:
        body = "\u3000" * indent + body.lstrip()
    if color:
        # 行内已有的加粗标记要留在 span 外面，否则 ** 与 HTML 混排时部分
        # 渲染器不闭合。这里只包一层颜色，不动原有标记。
        body = '<span style="color:#%s">%s</span>' % (color, body)
    if _align_attr(jc):
        body = '<p align="%s">%s</p>' % (jc, body)
    # Markdown 的 **加粗** 在 HTML 块级标签里不再被当作行内 Markdown 解析
    # （<p align="center">**标题**</p> 会原样显示两个星号）。只要套了 HTML，
    # 就把 **…** 换成形如 <strong>…</strong> 的标签——任何渲染器都认。
    if color or _align_attr(jc):
        body = _md_bold_to_html(body)
    return body


def _md_bold_to_html(s):
    """把 **加粗** 换成 <strong>…</strong>（HTML 里 ** 不再被解析）。

    只处理成对出现的 **，单个落单的 ** 原样保留（那多半是普通星号）。
    """
    return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", s)


def regroup_rows(tex, nrows):
    """把被并成一行的多行公式拆回多行（aligned 环境）。

    Word 里一个公式对象常常是好几行推导（三行最常见），转写的人却用
    `\\quad` 连成一整行——六条铁律全过、渲染也正常，但读者少看两个推导
    步骤。check_media_rows 拿预览图像素估算出真行数，这里负责动手拆。

    切分行、拼环境两件事都在 validate.py 里（split_formula_rows /
    join_formula_rows），与校验器共用同一套规则——否则会出现「build 说已
    修好、validate 仍报错」的打架局面：两边对同一个公式的行数理解不同。

    分隔符太多（一段长式里嵌了好几个 \\quad）时按 nrows 均分，而不是硬
    塞 nrows-1 个切点——那样最后一行会长得离谱。切不出至少两行就返回
    None，让调用方保留原文：宁可少拆一个，也不要给出一个只有一行的
    aligned 环境。
    """
    if not tex or nrows < 2:
        return None
    try:
        from validate import split_formula_rows, join_formula_rows
    except ImportError:
        return None
    parts = split_formula_rows(tex)
    if len(parts) < 2:
        return None
    if len(parts) > nrows:
        grouped, per = [], (len(parts) + nrows - 1) // nrows
        for k in range(0, len(parts), per):
            grouped.append(", ".join(parts[k:k + per]))
        parts = grouped
    return join_formula_rows(parts)


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
    # 走过「居中+加粗」这条规则的段落索引。用来让**相邻的一组**居中标签
    # 保持同一层级（见下面那条规则里的说明），而不是一条比一条深。
    center_label = set()
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

        # 居中 + 整段加粗 = Word 里的小节标签（实测本文档 6 条蓝色居中标题
        # 全是这个形态）。不加这条的话，它们只能走下面的加粗兜底，而兜底
        # 要求「≤15 字」且「下一段是正文」——"一般线弹性介质中…公式"19 字、
        # "一般位错密度演化方程"后面紧跟公式，全被挡掉，退化成普通段落；
        # 偏偏 13 字的"各向同性情况下物理场表达式"过了兜底成了标题。同一
        # 类标签两种下场，用户在渲染稿里看到的就是"有的居中有的不居中、
        # 有的多一行空行"。居中本身就是强结构信号，不需要旁证。
        #
        # 两个防误伤的限定（合成语料探针实测出来的）：
        #   * 图/表题注不是标题：英文论文里 "Figure 1. xxx" / "Table 1. xxx"
        #     就是居中加粗的，套这条规则会把题注升进目录。题注一律以
        #     Figure/Fig./Table/图N/表N/式N 这类词开头，挡掉。
        #   * 同级不递增：一组相邻居的标签（上面 6 条蓝标题就是）层级必须
        #     一致。若最近的上一条标题本身就是居中标签，这一条与它同级，
        #     而不是 _deepest_before+1——否则一过三条标签就一路掉到 h6。
        if _para_jc(struct, i) == "center" and _is_bold(struct, i) \
                and not re.search(r"[。；]$", t) \
                and not _is_caption(t):
            prev = [k for k in out if k < i]
            if prev and prev[-1] in center_label:
                # 上一条标题也是居中标签 → 同一组，层级照抄，不递增
                out[i] = out[prev[-1]]
            else:
                out[i] = min(_deepest_before(out, i) + 1, 6)
            center_label.add(i)
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
                and not _is_caption(t) \
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

    # 画布内的公式已经画进流程图 PNG，它的尺寸是整张画布的形状，
    # 不是公式行数——参与标定会把每行像素拉到荒谬的量级。
    index_scan = [e for e in index if isinstance(e, dict)
                  and not e.get("in_canvas")]
    sizes = {}
    media = os.path.join(args.workdir, "media")
    if os.path.isdir(media):
        for name in os.listdir(media):
            p = os.path.join(media, name)
            if not os.path.isfile(p):
                continue
            try:
                with open(p, "rb") as fp:
                    sizes[name] = imgsize.read_size(fp.read())
            except Exception:
                continue
    row_hits = {}
    try:
        from validate import check_media_rows
        _hits, _px = check_media_rows(index_scan, sizes, latex)
        row_hits = {h[0]: h[3] for h in _hits}
    except Exception:
        row_hits = {}
    # 预览图是多行、转写却成一行的公式：趁还在 latex 表里就把它拆回去。
    # 放在这里（而不是输出时）改，是为了让下游所有路径——行内、display、
    # 混排——都拿到已经修好的那一份，不必各自重拆一遍。
    # n_regrouped 先攒着：stats 在这一段之后才建。
    n_regrouped = 0
    for name in list(row_hits):
        regrouped = regroup_rows(latex.get(name), row_hits[name])
        if regrouped and regrouped != latex.get(name):
            latex[name] = regrouped
            n_regrouped += 1
        else:
            row_hits.pop(name, None)

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

    # in_canvas 的公式不参与"缺 LaTeX"统计：它们已经画进流程图 PNG，
    # 正文里本来就没有占位符，也不需要转写。把它们算进来会让每一次带
    # 流程图的构建都报"缺 N 个公式"，而实际一个都不缺。
    known = {r["stream"] for r in index
             if r.get("kind") != "image" and not r.get("in_canvas")}
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
             # 画布内的公式：已随流程图整块渲染，正文里没有独立占位符。
             # 必须单独报出来——否则 T8 拿 md 里的公式数去对 expectFormulas
             # （它来自 inspect 的全量 OLE 计数）会永远差这么多个。
             "formulas_canvas": sum(1 for r in index if r.get("in_canvas")),
             "canvases": 0,             # 正文落成的画布（流程图）张数
             "image_placeholders": 0,   # 图片公式：落成注释+图片引用的处数
             "image_transcribed": 0,    # Agent 已转写成 LaTeX 的图片公式数
             "figures": 0,              # 正文落成的示意图处数
             "orphan_punct_dropped": 0,  # 公式块后无处安放的孤立标点
             "tables": 0,               # 正文里还原出的表格数
             "table_rows": 0,           # 表格的数据行数（不含表头）
             "lists": 0,                # 识别出的列表项数
             # 预览图是多行、转写却是一行的公式：已自动拆成多行。
             "formulas_row_regrouped": n_regrouped}

    prev_display = False
    all_notes = []          # 脚注正文，文末统一列出

    # ---- 块级结构：表格 / 段落按真实顺序输出 ----
    # paras.json 只有段落；表格的行列关系在那里已经丢失。blocks.json 记的
    # 是正文的真实结构（段落与表格各自的顺序），有它才能把表格还原成
    # Markdown 表格。没有 blocks.json 时退回纯段落模式（兼容旧产物）。
    block_seq = None
    try:
        _b = load_json(os.path.join(args.workdir, "blocks.json"))
        if isinstance(_b, dict) and isinstance(_b.get("blocks"), list):
            block_seq = _b["blocks"]
    except (json.JSONDecodeError, OSError):
        block_seq = None

    def render_table(blk):
        """把一张表渲染成 GFM 表格。

        合并单元格在 Markdown 里没有语法：横向合并的格子留空（GFM 语义
        就是"沿用上一格"），纵向合并同理。**不复制内容**——复制会让同一
        段文字凭空出现两次。
        """
        rows = blk.get("rows") or []
        if not rows:
            return []
        ncol = max(len(r) for r in rows) or 1
        grid = []
        for r in rows:
            cells = []
            for c in r:
                txt = (c.get("text") or "").replace("|", "\\|").replace("\n", " ")
                cells.append(txt)
                # 横向合并：补齐被合并掉的列，留空
                for _ in range(int(c.get("span") or 1) - 1):
                    cells.append("")
            while len(cells) < ncol:
                cells.append("")
            grid.append(cells[:ncol])
        header = grid[0]
        body = grid[1:]
        lines = ["| " + " | ".join(header) + " |",
                 "| " + " | ".join(["---"] * ncol) + " |"]
        for row in body:
            lines.append("| " + " | ".join(row) + " |")
        stats["tables"] += 1
        stats["table_rows"] += len(body)
        return lines

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

    # 按块结构走：段落与表格交替，顺序就是原文顺序。
    # 没有 blocks.json 时构造一个「全是段落」的等价序列（向后兼容）。
    if block_seq is None:
        seq = [{"kind": "para", "index": i} for i in range(len(paras))]
    else:
        seq = []
        para_cursor = 0
        for blk in block_seq:
            if blk.get("kind") == "table":
                # 表格之前若有未被 blocks.json 记录的段落，补上
                while para_cursor < blk.get("para_from", para_cursor):
                    seq.append({"kind": "para", "index": para_cursor})
                    para_cursor += 1
                seq.append(blk)
                para_cursor = blk.get("para_to", para_cursor)
            else:
                idx = blk.get("index")
                if isinstance(idx, int) and idx >= para_cursor:
                    while para_cursor < idx:
                        seq.append({"kind": "para", "index": para_cursor})
                        para_cursor += 1
                    seq.append({"kind": "para", "index": idx,
                                "runs": blk.get("runs"),
                                "list": blk.get("list")})
                    para_cursor = idx + 1
        while para_cursor < len(paras):
            seq.append({"kind": "para", "index": para_cursor})
            para_cursor += 1

    for item in seq:
        if item.get("kind") == "table":
            tbl = render_table(item)
            if tbl:
                if out and out[-1] != "":
                    out.append("")
                out.extend(tbl)
                out.append("")
                prev_display = False
            continue
        i = item["index"]
        raw = paras[i] if i < len(paras) else ""
        # 有 runs 时用带格式的版本（加粗/斜体/链接/脚注/上下标）；
        # 没有时退回纯文本（旧产物、或该段来自降级路径）。
        runs = item.get("runs")
        if runs:
            rendered, fnotes = render_runs(runs)
            if rendered:
                raw = rendered
            for fn in fnotes:
                all_notes.append(fn)
        # 列表项：正文文字来自 paragraphs，编号来自 numbering.xml。
        # 有序列表用 "1. "，无序用 "- "；Markdown 会自己重排序号。
        lk = item.get("list")
        text = raw.strip()
        if not text:
            if out and out[-1] != "":
                out.append("")
            continue
        stats["paragraphs"] += 1

        # ---- 排版信号：居中 / 颜色 / 首行缩进 ----
        # 这三样在 Markdown 里都没有原生语法，不还原就是静默丢格式：正文
        # 一字不差，但原文的层级提示（居中+蓝色 = 小节标题）与段落起始
        # （首行缩进两格）全没了。
        jc = _para_jc(struct, i)
        color = _para_color(struct, i)
        # 标题不加首行缩进：它的 w:ind 是 Word 编号悬挂，不是正文缩进；
        # 列表项加了缩进会让 Markdown 认不出列表。
        indent = 0 if (headings.get(i, 0) or lk) else _para_indent_chars(struct, i)

        # ---- 标题 ----
        # 居中的标题不能用 `#`：Markdown 的标题语法没有居中，走了 `#` 分支
        # 居中就丢了（实测文档大标题就是这样丢的——它的居中来自 Title 样式
        # 的 `<w:jc w:val="center"/>`，段落自己没写）。改用 `<hN align>`：
        # 既是标题（渲染器给出标题的字号与上下间距，不会在正文里凭空多出
        # 一行空白），又保住了居中。
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
            htext = _heading_text(i, text, struct)
            if color:
                htext = '<span style="color:#%s">%s</span>' % (color, htext)
            if _align_attr(jc):
                out.append('<h%d align="%s">%s</h%d>'
                           % (min(int(level), 6), jc, htext, min(int(level), 6)))
            else:
                out.append("#" * min(int(level), 6) + " " + htext)
            out.append("")
            prev_display = False
            continue

        # 含公式/插图的段落不整体包 HTML：混排路径会把这一段拆成多行输出，
        # 包在首尾的 <p> 会跨行裂开，渲染出半个标签。这类段落里的公式本身
        # （$$ 块、独立成行的图）已经是居中呈现的，丢了字符居中不影响可读性。
        if (jc or color) and not EQ_RE.search(text):
            text = wrap_para_format(text, jc=jc, color=color)
            raw = wrap_para_format(raw, jc=jc, color=color)

        # ---- 列表 ----
        # Word 的列表项在正文里只是普通段落，形状由 numbering.xml 决定。
        # 不还原它，转出的 md 里列表就是一串裸段落——读者看不出这是列表，
        # 原文的分组/并列关系整个丢掉。Markdown 会自己重排序号，所以
        # 有序列表统一写 "1. "，无序写 "- "。
        if lk in ("bullet", "ordered"):
            marker = "- " if lk == "bullet" else "1. "
            stats["lists"] += 1
            text = marker + text
            raw = marker + raw

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
                    #
                    # 画布（流程图）走同一条路，但要单独计数：它承载的是
                    # 整块结构，与普通插图不是一回事，报告里必须能分开看。
                    stats["figures"] += 1
                    if name.startswith("fig:canvas"):
                        stats["canvases"] += 1
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
        # line_formula_flushed: 这一段里已有的文字是否作为**独立行**输出过
        # （被长公式或插图劈开）。一旦发生过，"rest" 只是原段落的最后一句，
        # 再给它加首行缩进就会让它看起来比同段的其它行更靠右——读者会以
        # 为那是新段落。
        line = []
        line_formula_flushed = False
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
                    line_formula_flushed = True
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
                        line_formula_flushed = True
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
            # 首行缩进在这里补，不能提前：混排路径输出的每一步都 strip()，
            # 提前加的全角空格会被当成普通空白吃掉。放在最后 append 之前
            # 才保得住。
            if indent and not line_formula_flushed:
                rest = "\u3000" * indent + rest.lstrip()
            out.append(rest)
            out.append("")
            prev_display = False

    # 脚注：正文里的 [^fnN] 标记必须有对应的定义，否则渲染器把它当字面
    # 文本。Word 的脚注正文写在 footnotes.xml 里，不读它等于把注释丢掉。
    if all_notes:
        out.append("")
        for k, note in enumerate(all_notes, 1):
            out.append("[^fn%d]: %s" % (k, note.replace("\n", " ")))
        out.append("")
        stats["footnotes"] = len(all_notes)

    body = "\n".join(out)
    body = re.sub(r"\n{3,}", "\n\n", body).strip("\n") + "\n"
    # 注意是 strip("\n") 不是 strip()：Python 的 str.strip() 把 U+3000 也算
    # 空白，一篇**以缩进正文开头**的文档（没有大标题打头）会在最后这一步
    # 被削掉首段的全角空格——16pt 合成文档的端到端测试实测撞上了这个。

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
