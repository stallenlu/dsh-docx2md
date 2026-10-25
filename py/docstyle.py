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
        {"style": int|None, "marker": str, "bold": bool}
    """
    styles = read_styles(z)
    numbering = read_numbering(z)
    counters = {}
    out = []
    for p in paras_elems:
        style = None
        marker = ""
        bold = False
        ppr = p.find(W + "pPr")
        if ppr is not None:
            ps = ppr.find(W + "pStyle")
            if ps is not None:
                style = styles.get(ps.get(W + "val"))
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
        out.append({"style": style, "marker": marker, "bold": bold})
    return out
