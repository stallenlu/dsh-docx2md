#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docstyle.py —— 从 docx 的样式表与编号表还原段落的结构信号。

为什么需要它：中文 Word 文档里大量标题**不写编号文本**，编号由 Word 自动
生成（`<w:numPr>`），或者干脆套用 `heading N` 样式。这两种情况下段落文字
里根本没有 "1．" / "A）" / "i）" 这类前缀，正则一个字也认不出来。

实测（145 公式的真实文档，293 段）：纯正则只认出 12 条标题，而文档真实
的标题结构是 32 条。漏掉的 20 条全部是「自动编号 + 样式标题」：
`<w:numPr>` 段落（"位错密度的演化方程" 等 10 条）与 `pStyle`（"引言"），
正文文字干净得像普通句子。

本模块只做**结构读取**，不猜语义：
    read_styles(z)     styleId → 标题级别（由 outlineLvl 或 "heading N" 名字）
    read_numbering(z)  numId → {ilvl: (numFmt, lvlText, start)}
    paragraph_marks(z, paras_elems)
                       与正文段落同序的结构信号：样式级别、合成出的编号
                       （"iii)"）、是否整段加粗

编号是**合成**出来的：按 (numId, ilvl) 在文档中出现的先后累加 start 值。
Word 的编号本就如此渲染，只是文件里没存成文字。合成之后交给 build.py
已有的那套「按前缀判标题」的正则，等于把两条通路合并成一条——不新增
一套判断逻辑，也就不会出现「正则说一套、结构说另一套」的分歧。

本模块**不改写**段落正文。编号只是判标题与小节深度的依据，是否把它补进
输出由 build.py 决定。
"""
import re

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

_ROMAN = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x",
          "xi", "xii", "xiii", "xiv", "xv", "xvi", "xvii", "xviii", "xix", "xx"]


def _roman(n):
    """小写罗马数字（1..20 够用；超出回退十进制，绝不编造）。"""
    return _ROMAN[n - 1] if 1 <= n <= len(_ROMAN) else str(n)


def _alpha(n, upper=False):
    if 1 <= n <= 26:
        c = chr(ord("a") + n - 1)
        return c.upper() if upper else c
    return str(n)


def _cn(n):
    """中文计数（一…二十）。超出回退十进制。"""
    cn = "一二三四五六七八九十"
    if 1 <= n <= 10:
        return cn[n - 1]
    if 11 <= n <= 19:
        return "十" + cn[n - 11]
    if n == 20:
        return "二十"
    return str(n)


def _format(n, num_fmt):
    """按 numFmt 把序号 n 变成 Word 会显示的样子。"""
    if num_fmt == "lowerRoman":
        return _roman(n)
    if num_fmt == "upperRoman":
        return _roman(n).upper()
    if num_fmt == "lowerLetter":
        return _alpha(n)
    if num_fmt == "upperLetter":
        return _alpha(n, upper=True)
    if num_fmt in ("chineseCounting", "chineseCountingThousand",
                   "japaneseCounting", "taiwaneseCounting"):
        return _cn(n)
    if num_fmt == "none":
        return ""
    # decimal / decimalZero / decimalEnclosedCircle 及其余未覆盖格式一律十进制：
    # 合成标记只需能喂给正则判深度，不要求与 Word 屏幕显示逐字一致。
    return str(n)


def _render(lvl_text, num_fmt, n):
    """把 lvlText 里的 %1 换成实际编号；其余 %k 留空（本模块只合成一级）。"""
    def sub(m):
        return _format(n, num_fmt) if int(m.group(1)) == 1 else ""
    return re.sub(r"%(\d)", sub, lvl_text or "")


def _parse(z, part):
    from xml.etree import ElementTree as ET
    try:
        return ET.fromstring(z.read(part))
    except (KeyError, ET.ParseError):
        return None


def read_styles(z):
    """styleId → 标题级别（1..6）。非标题样式不出现在结果里。

    判定优先级：
      1. `<w:pPr><w:outlineLvl w:val="N"/>` → N+1
         （outlineLvl 是 0 基的，Word 用它给导航窗格排序，最可信）
      2. 样式名匹配 `heading N` / `标题 N` → N
      3. 样式名为 `Title` / `标题` → 1
    """
    out = {}
    root = _parse(z, "word/styles.xml")
    if root is None:
        return out
    for st in root.iter(W + "style"):
        if st.get(W + "type") not in (None, "paragraph"):
            continue
        sid = st.get(W + "styleId")
        if not sid:
            continue
        level = None
        ppr = st.find(W + "pPr")
        if ppr is not None:
            o = ppr.find(W + "outlineLvl")
            if o is not None:
                try:
                    level = int(o.get(W + "val")) + 1
                except (TypeError, ValueError):
                    level = None
        if level is None:
            nm = st.find(W + "name")
            name = ((nm.get(W + "val") if nm is not None else "") or "").strip()
            m = re.match(r"^(?:heading|标题)\s*(\d+)$", name, re.I)
            if m:
                level = int(m.group(1))
            elif name.lower() in ("title", "标题"):
                level = 1
        if level:
            out[sid] = max(1, min(int(level), 6))
    return out


def default_char_twip(z):
    """文档默认字号下的全角字宽（twip）。取不到时退回 210（10.5pt 中文）。

    为什么不能写死 210：`w:ind w:firstLine` 是绝对长度（twip），换算成
    「几个字」必须除以**这篇文档自己的**字宽。实测本文档 Normal 样式
    `<w:sz w:val="21"/>`（半磅）= 10.5pt，全角字宽 = 10.5 × 20 = 210twip；
    同一份文档里 56 个同时写了 firstLine 与 firstLineChars 的段落反推
    出的字宽正是 210/240——写死 210 在 12pt 文档上就会把 2 字缩进算成
    1.75 字、在 16pt 上算成 3 字。

    查找顺序：docDefaults/rPrDefault/rPr/sz → Normal 样式的 sz → 21。
    """
    root = _parse(z, "word/styles.xml")
    if root is None:
        return 210
    sz = None
    dd = root.find(W + "docDefaults")
    if dd is not None:
        el = dd.find(W + "rPrDefault/" + W + "rPr/" + W + "sz")
        if el is not None:
            sz = el.get(W + "val")
    if sz is None:
        for st in root.iter(W + "style"):
            nm = st.find(W + "name")
            if nm is not None and (nm.get(W + "val") or "").strip().lower() in ("normal", "默认段落字体"):
                el = st.find(W + "rPr/" + W + "sz")
                if el is not None:
                    sz = el.get(W + "val")
                break
    try:
        half_pt = int(sz)
    except (TypeError, ValueError):
        return 210
    if half_pt < 1:
        return 210
    # 半磅 → 磅(÷2) → twip(×20)；全角字宽 = 1em = 字号
    return half_pt * 10


def read_style_jc(z):
    """styleId → 段落对齐（沿 basedOn 链继承）。只收非默认的 center/right。

    为什么必须查样式链：Word 里「对齐」经常写在样式上而不是段落上。实测
    本文档的标题段（"基于矢量位错密度的位错动力学及其数理方案"）段落
    pPr 里根本没有 <w:jc>，居中来自 Title 样式（styleId=7）的
    `<w:pPr><w:jc w:val="center"/>`。只读段落级 w:jc，标题就丢了居中。
    """
    root = _parse(z, "word/styles.xml")
    if root is None:
        return {}
    raw, based = {}, {}
    for st in root.iter(W + "style"):
        if st.get(W + "type") not in (None, "paragraph"):
            continue
        sid = st.get(W + "styleId")
        if not sid:
            continue
        ppr = st.find(W + "pPr")
        jc_el = ppr.find(W + "jc") if ppr is not None else None
        val = jc_el.get(W + "val") if jc_el is not None else None
        # 只认 center/right：both(两端对齐)/left 是排版默认，既写不进
        # HTML 的 align（align="both" 是非法值），也不影响读者看到的层级。
        if val in ("center", "right"):
            raw[sid] = val
        b = st.find(W + "basedOn")
        if b is not None:
            based[sid] = b.get(W + "val")

    def resolve(sid, depth=0):
        if sid in raw:
            return raw[sid]
        if depth > 10:                       # 样式链成环的保险
            return None
        parent = based.get(sid)
        return resolve(parent, depth + 1) if parent else None

    return {sid: resolve(sid)
            for sid in set(list(raw) + list(based)) if resolve(sid)}


def read_numbering(z):
    """numId → {ilvl: (numFmt, lvlText, start)}。"""
    out = {}
    root = _parse(z, "word/numbering.xml")
    if root is None:
        return out

    abstract = {}
    for a in root.iter(W + "abstractNum"):
        aid = a.get(W + "abstractNumId")
        if aid is None:
            continue
        levels = {}
        for lvl in a.findall(W + "lvl"):
            try:
                ilvl = int(lvl.get(W + "ilvl") or 0)
            except (TypeError, ValueError):
                ilvl = 0
            fmt = lvl.find(W + "numFmt")
            txt = lvl.find(W + "lvlText")
            sta = lvl.find(W + "start")
            try:
                start = int((sta.get(W + "val") if sta is not None else 1) or 1)
            except (TypeError, ValueError):
                start = 1
            levels[ilvl] = (
                (fmt.get(W + "val") if fmt is not None else "decimal"),
                (txt.get(W + "val") if txt is not None else "%1."),
                start,
            )
        abstract[aid] = levels

    for num in root.iter(W + "num"):
        nid = num.get(W + "numId")
        ref = num.find(W + "abstractNumId")
        if nid is None or ref is None:
            continue
        levels = dict(abstract.get(ref.get(W + "val"), {}))
        # num 级覆盖（<w:lvlOverride>）优先，Word 用它做「从此处重新编号」。
        for ov in num.findall(W + "lvlOverride"):
            try:
                ilvl = int(ov.get(W + "ilvl") or 0)
            except (TypeError, ValueError):
                ilvl = 0
            so = ov.find(W + "startOverride")
            if so is not None and ilvl in levels:
                fmt, txt, _ = levels[ilvl]
                try:
                    levels[ilvl] = (fmt, txt, int(so.get(W + "val") or 1))
                except (TypeError, ValueError):
                    pass
        out[nid] = levels
    return out


def paragraph_marks(z, paras_elems):
    """给每个正文段落取出结构信号，顺序与 paras_elems 严格一致。

    paras_elems 必须是 extract.py 认定的**正文段落**（已排除文本框里的段落）。
    索引对不上会让标题挂到错误的段落上，那比不识别更糟，所以调用方必须传
    与 paras.json 同序、同口径的那一组元素。

    返回与 paras_elems 同长的列表，每项：
        {"style": int|None, "marker": str, "bold": bool,
         "jc": str|None, "color": str|None,
         "firstLine": int|None, "firstLineChars": int|None,
         "charTwip": int}

    后五项是**排版**信号，不是结构信号，但同样是原文里明确记录的东西：
      * jc / color: Word 里居中且非黑色的段（"一般位错密度演化方程"）在纯
        Markdown 里没有任何对应语法，不读它们，这段就退化成普通一行。
      * firstLine: 首行缩进（中文正文缩进两字符 = 2 个字宽）。Markdown
        解析器会把行首空白整段吃掉，不还原就彻底消失。
      * charTwip: 本文档默认字号的全角字宽，firstLine 换算成「几个字」
        要除以它（见 default_char_twip）。
      """
    styles = read_styles(z)
    style_jc = read_style_jc(z)
    numbering = read_numbering(z)
    char_twip = default_char_twip(z)
    counters = {}
    out = []
    for p in paras_elems:
        style = None
        marker = ""
        bold = False
        jc = None
        first_line = None
        first_line_chars = None
        ppr = p.find(W + "pPr")
        if ppr is not None:
            ps = ppr.find(W + "pStyle")
            if ps is not None:
                style = styles.get(ps.get(W + "val"))
                # 段落自己没写 w:jc 时，用样式链上的（Title 样式的居中就
                # 是这么来的；段落级覆盖样式级）。
                jc = style_jc.get(ps.get(W + "val"))
            jc_el = ppr.find(W + "jc")
            if jc_el is not None:
                cand = jc_el.get(W + "val")
                if cand in ("center", "right"):
                    jc = cand
                elif cand in ("left", "start", "both", "justify",
                              "distribute"):
                    jc = None               # 段落显式默认对齐，覆盖样式
            ind = ppr.find(W + "ind")
            if ind is not None:
                cand = ind.get(W + "firstLine")
                if cand:
                    try:
                        first_line = int(cand)
                    except ValueError:
                        pass
                cand = ind.get(W + "firstLineChars")
                if cand:
                    try:
                        first_line_chars = int(cand)
                    except ValueError:
                        pass
            npr = ppr.find(W + "numPr")
            if npr is not None:
                n_el = npr.find(W + "numId")
                l_el = npr.find(W + "ilvl")
                nid = n_el.get(W + "val") if n_el is not None else None
                try:
                    ilvl = int((l_el.get(W + "val") if l_el is not None else 0) or 0)
                except (TypeError, ValueError):
                    ilvl = 0
                lv = numbering.get(nid) or {}
                if ilvl in lv:
                    key = (nid, ilvl)
                    counters[key] = counters.get(key, 0) + 1
                    fmt, lvl_text, start = lv[ilvl]
                    marker = _render(lvl_text, fmt, start + counters[key] - 1)
        # 整段加粗：段落标记本身带 <w:b/>，或所有有字的 run 都带 <w:b/>。
        # 中文技术文档里「1．/ A） / i）」这类小标题常只有加粗、无样式，
        # 这是唯一能认出它们的结构信号。
        rpr = ppr.find(W + "rPr") if ppr is not None else None
        if rpr is not None and rpr.find(W + "b") is not None:
            b = rpr.find(W + "b")
            bold = b.get(W + "val") not in ("0", "false")
        if not bold:
            runs = [r for r in p.iter(W + "r") if r.find(W + "t") is not None]
            if runs:
                flags = []
                for r in runs:
                    if not (r.find(W + "t").text or "").strip():
                        continue
                    rp = r.find(W + "rPr")
                    b = rp.find(W + "b") if rp is not None else None
                    flags.append(b is not None
                                 and b.get(W + "val") not in ("0", "false"))
                bold = bool(flags) and all(flags)
        color = None
        rprs = [rpr] if rpr is not None else []
        rprs += [r.find(W + "rPr") for r in p.iter(W + "r")]
        for rp in rprs:
            if rp is None:
                continue
            c = rp.find(W + "color")
            if c is None:
                continue
            val = c.get(W + "val")
            if val and val.upper() not in ("000000", "AUTO"):
                # 蓝色标题（"一般位错密度演化方程" 0000FF）在纯 Markdown 里
                # 没有任何语法能表达，不读它，颜色就静默消失。
                color = val.upper()
                break
        out.append({"style": style, "marker": marker, "bold": bold,
                    "jc": jc, "color": color,
                    "firstLine": first_line,
                    "firstLineChars": first_line_chars,
                    # 每条都带上：docstruct.json 是「与段落同序的列表」，
                    # 多一个顶层对象会让所有按下标取记录的调用方都要改。
                    # 值全相同，293 条多几 KB，换 build 端零改动。
                    "charTwip": char_twip})
    return out


def list_kind(z, num_id, ilvl):
    """这个自动编号段落是项目符号列表还是有序列表？都不是返回 None。

    Word 的列表在正文里只有一个 <w:numPr>，真正的形状（•、1.、i)）由
    numbering.xml 决定。不读它，转出的 md 里列表就是一串裸段落——
    读者看不出这是列表，也就丢掉了原文的格式信息。

    lowerRoman / upperLetter 这类多用于**标题编号**（"i)" / "A）"），
    标题识别已经把它们补回标题里了，这里不再重复加列表标记，返回 None。
    """
    if not num_id:
        return None
    numbering = read_numbering(z)
    lv = numbering.get(str(num_id)) or {}
    try:
        il = int(ilvl or 0)
    except (TypeError, ValueError):
        il = 0
    if il not in lv:
        return None
    fmt = lv[il][0]
    if fmt == "bullet":
        return "bullet"
    if fmt in ("decimal", "decimalZero", "decimalEnclosedCircle"):
        return "ordered"
    return None
