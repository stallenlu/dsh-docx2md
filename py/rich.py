#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""rich.py —— 从 docx 正文里读出**块级结构**（表格 / 列表 / 行内格式 / 链接 / 脚注）。

为什么需要它：extract.py 原来只认「正文 = 一串 w:p」。这是把 docx 当纯文本来
读的假设，只在**没有表格、没有列表、没有行内格式**的文档上成立。真实文档只要
出现一样，信息就没了：

  * 表格：`w:tbl` 里的每个 `w:tc` 都是一个 `w:p`，被 `body.iter(w:p)` 平铺
    成散段——"表头A / 表头B / 单元格1 / 单元格2" 各占一行，行列关系彻底消失。
    不报错、不告警，读起来就是一段碎句。
  * 列表：`w:numPr` 的段落只有编号在 numbering.xml 里，正文文字干干净净，
    转出来是一串裸段落，读者看不出这是列表。
  * 行内格式：`w:rPr/w:b`（加粗）、`w:i`（斜体）、`w:u`（下划线）、
    `w:strike`（删除线）、`w:vertAlign`（上下标）全被丢掉。
  * 超链接 `w:hyperlink`：只留下锚文本，URL 消失。
  * 脚注/尾注 `w:footnoteReference`：锚点消失，注释正文从未被读取。

这些都不是「这份文档特有」的问题，而是任何 docx 都可能有的标准结构。一个
格式转换器必须按 OOXML 的**实际结构**遍历，而不是假设「正文只有段落」。

本模块只做读取与结构建模，输出成中立的数据结构交给 extract.py。
"""
import re

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"


def _ln(tag):
    return tag.split("}")[-1] if "}" in tag else tag


def _on(el):
    """`<w:b/>` 这类开关属性：元素存在且 val 不是 0/false 才算真。"""
    if el is None:
        return False
    return el.get(W + "val") not in ("0", "false", "none", "off")


def _rpr_flags(rpr):
    """从一个 w:rPr 读出全部行内格式标记。"""
    out = {"bold": False, "italic": False, "underline": False,
           "strike": False, "vert": None}
    if rpr is None:
        return out
    out["bold"] = _on(rpr.find(W + "b"))
    out["italic"] = _on(rpr.find(W + "i"))
    u = rpr.find(W + "u")
    out["underline"] = u is not None and u.get(W + "val") not in ("none", "0", "false")
    out["strike"] = (_on(rpr.find(W + "strike"))
                     or _on(rpr.find(W + "dstrike")))
    va = rpr.find(W + "vertAlign")
    if va is not None:
        v = va.get(W + "val")
        if v in ("superscript", "subscript"):
            out["vert"] = "sup" if v == "superscript" else "sub"
    return out


def read_rels(z, part="word/_rels/document.xml.rels"):
    """rId → Target（超链接、图片、脚注都靠它）。"""
    try:
        raw = z.read(part).decode("utf-8", "replace")
    except KeyError:
        return {}
    return dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', raw))


def read_notes(z):
    """脚注/尾注 id → 纯文字。"""
    from xml.etree import ElementTree as ET
    out = {}
    for part, key in (("word/footnotes.xml", "footnote"),
                      ("word/endnotes.xml", "endnote")):
        try:
            root = ET.fromstring(z.read(part))
        except Exception:                  # noqa: BLE001 — 缺失/损坏都当没有
            continue
        for fn in root.iter(W + key):
            fid = fn.get(W + "id")
            if fid is None or fid.startswith("-"):   # -1/0 是分隔符，不是内容
                continue
            txt = "".join(t.text or "" for t in fn.iter(W + "t")).strip()
            if txt:
                out[fid] = txt
    return out


def _new_run(text="", bold=False, italic=False, underline=False,
             strike=False, vert=None, link=None):
    return {"text": text, "bold": bold, "italic": italic,
            "underline": underline, "strike": strike, "vert": vert,
            "link": link, "footnote": None, "code": False}


def runs_of(p, ctx):
    """把一个 w:p 拆成 Run 列表（保留行内格式、链接、脚注锚点）。

    ctx 含：
        media  回调 (node) -> 占位符字符串或 None
               （图片 / 公式 / 表格外的一切非文字内容都由它决定怎么占位）
        notes  脚注 id → 文字
    """
    media = ctx.get("media")
    notes = ctx.get("notes") or {}
    out = []

    def emit(run):
        if run["text"] or run["footnote"]:
            out.append(run)

    def walk(el, fmt, link):
        """fmt = (bold, italic, underline, strike, vert)"""
        for c in el:
            tag = c.tag
            if tag == W + "r":
                rpr = c.find(W + "rPr")
                f = _rpr_flags(rpr)
                cur = (f["bold"] or fmt[0], f["italic"] or fmt[1],
                       f["underline"] or fmt[2], f["strike"] or fmt[3],
                       f["vert"] or fmt[4])
                buf = []
                for rc in c:
                    if rc.tag == W + "t":
                        buf.append(rc.text or "")
                    elif rc.tag == W + "tab":
                        buf.append("\t")
                    elif rc.tag in (W + "br", W + "cr"):
                        buf.append("\n")
                    elif rc.tag == W + "noBreakHyphen":
                        buf.append("-")
                    elif rc.tag in (W + "footnoteReference",
                                    W + "endnoteReference"):
                        if buf:
                            emit(_new_run("".join(buf), *cur, link=link))
                            buf = []
                        fid = rc.get(W + "id") or ""
                        note = notes.get(fid)
                        r = _new_run("", *cur, link=link)
                        if note:
                            r["footnote"] = note
                            emit(r)
                        # 没有正文的脚注锚点（分隔符等）不占位
                    elif rc.tag == MC + "AlternateContent":
                        # 兼容性包装常常**长在 w:r 里面**（Word 2010+ 插图
                        # 的标准写法：<w:r><mc:AlternateContent><mc:Choice>
                        # <w:drawing>…）。只在段落层处理 AlternateContent
                        # 会漏掉这一大类，图就整张消失。
                        if buf:
                            emit(_new_run("".join(buf), *cur, link=link))
                            buf = []
                        picked = rc.find(MC + "Choice")
                        if picked is None:
                            picked = rc.find(MC + "Fallback")
                        if picked is not None:
                            walk(picked, cur, link)
                    elif rc.tag in (W + "drawing", W + "pict", W + "object"):
                        if buf:
                            emit(_new_run("".join(buf), *cur, link=link))
                            buf = []
                        ph = media(rc) if media else None
                        if ph:
                            emit(_new_run(ph, *cur, link=link))
                    elif rc.tag == M + "oMath":
                        if buf:
                            emit(_new_run("".join(buf), *cur, link=link))
                            buf = []
                        ph = media(rc) if media else None
                        if ph:
                            emit(_new_run(ph, *cur, link=link))
                if buf:
                    emit(_new_run("".join(buf), *cur, link=link))
            elif tag == W + "hyperlink":
                target = None
                if ctx.get("rels"):
                    target = ctx["rels"].get(c.get(R + "id"))
                walk(c, fmt, target or link)
            elif tag == MC + "AlternateContent":
                # mc:AlternateContent 是 Word 的兼容性包装：Choice 是新的
                # DrawingML 写法，Fallback 是老的 VML 写法，**两者指向同一份
                # 媒体**。两个分支都读会得到两个占位符（同一张图出现两次）；
                # 只读 Fallback 会漏掉只有 Choice 的图。
                # 规范做法：优先 Choice，没有才退 Fallback。
                picked = c.find(MC + "Choice")
                if picked is None:
                    picked = c.find(MC + "Fallback")
                if picked is not None:
                    walk(picked, fmt, link)
            elif tag in (W + "ins", W + "smartTag", W + "sdt", W + "sdtContent",
                         W + "fldSimple", W + "customXml", W + "dir", W + "bdo"):
                # 修订插入 / 智能标记 / 内容控件 / 简单域：结构外壳，递归不丢字
                walk(c, fmt, link)
            elif tag == W + "del":
                # 修订删除：文字还在 XML 里，但不属于当前版本。保留成删除线
                # Run（可见、不冒充正文），而不是直接丢掉。
                for r in c.iter(W + "r"):
                    f = _rpr_flags(r.find(W + "rPr"))
                    txt = "".join(t.text or "" for t in r.iter(W + "t"))
                    if txt:
                        emit(_new_run(txt, f["bold"], f["italic"],
                                      f["underline"], True, f["vert"], link))
            elif tag in (W + "drawing", W + "pict", W + "object"):
                ph = media(c) if media else None
                if ph:
                    emit(_new_run(ph, *fmt, link=link))
            elif tag == M + "oMath":
                ph = media(c) if media else None
                if ph:
                    emit(_new_run(ph, *fmt, link=link))
            elif tag in (M + "oMathPara",):
                walk(c, fmt, link)

    walk(p, (False, False, False, False, None), None)
    return out


def plain(runs):
    return "".join(r["text"] for r in runs)


def table_rows(tbl, ctx):
    """读一张表的行列。支持横向/纵向合并（gridSpan / vMerge）。

    合并单元格在 Markdown 里没有对应语法，只能降级：横向合并把内容放进
    第一格、其余格留空（GFM 语义就是「同上一格」）；纵向合并同理。绝不
    把内容复制到每个格子里——那会凭空多出重复文字。
    """
    rows = []
    for tr in tbl.findall(W + "tr"):
        cells = []
        for tc in tr.findall(W + "tc"):
            tcpr = tc.find(W + "tcPr")
            span, vmerge = 1, None
            if tcpr is not None:
                gs = tcpr.find(W + "gridSpan")
                if gs is not None:
                    try:
                        span = max(1, int(gs.get(W + "val") or 1))
                    except (TypeError, ValueError):
                        span = 1
                vm = tcpr.find(W + "vMerge")
                if vm is not None:
                    vmerge = vm.get(W + "val") or "continue"
            cell_runs = []
            for p in tc.findall(W + "p"):
                cell_runs.extend(runs_of(p, ctx))
                cell_runs.append(_new_run("\n"))
            # 去掉末尾那个分隔换行
            while cell_runs and cell_runs[-1]["text"] == "\n":
                cell_runs.pop()
            cells.append({"runs": cell_runs, "span": span, "vmerge": vmerge})
        rows.append(cells)
    header = False
    trs = tbl.findall(W + "tr")
    if trs:
        trpr = trs[0].find(W + "trPr")
        if trpr is not None and trpr.find(W + "tblHeader") is not None:
            header = True
    return rows, header


def blocks(z, body, ctx):
    """把 body 的块级子元素按**原始顺序**读成 Block 列表。

    关键：遍历 body 的**直接子元素**，不是 `body.iter(w:p)`。只有这样，
    表格才是一个块、段落才是一个块，顺序才不会乱——真实文档里「图 + 表 +
    图」的交错顺序只有按子元素走才保得住。
    """
    out = []
    for child in body:
        tag = child.tag
        if tag == W + "p":
            ppr = child.find(W + "pPr")
            num = None
            if ppr is not None:
                npr = ppr.find(W + "numPr")
                if npr is not None:
                    nid = npr.find(W + "numId")
                    ilvl = npr.find(W + "ilvl")
                    num = {
                        "numId": nid.get(W + "val") if nid is not None else None,
                        "ilvl": int((ilvl.get(W + "val") if ilvl is not None
                                     else 0) or 0),
                    }
            out.append({"kind": "para", "runs": runs_of(child, ctx),
                        "num": num, "el": child})
        elif tag == W + "tbl":
            rows, header = table_rows(child, ctx)
            out.append({"kind": "table", "rows": rows, "header": header,
                        "el": child})
        elif tag == W + "sdt":
            content = child.find(W + "sdtContent")
            if content is not None:
                out.extend(blocks(z, content, ctx))
        elif tag in (W + "bookmarkStart", W + "bookmarkEnd", W + "proofErr",
                     W + "commentRangeStart", W + "commentRangeEnd",
                     W + "sectPr"):
            continue
        else:
            inner = list(child.iter(W + "p"))
            if inner:
                for p in inner:
                    out.append({"kind": "para", "runs": runs_of(p, ctx),
                                "el": p})
    return out
