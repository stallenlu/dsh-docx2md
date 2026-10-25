#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""extract.py —— 从 .docx 提取正文段落、公式索引与公式内嵌预览图。

用法:
    python extract.py <input.docx> <outdir>

产出（全部落在 <outdir>）:
    paras.json      段落列表，公式位置标为 ⟦EQ:oleObjectN.bin⟧（OLE）、
                    ⟦EQ:ommlK⟧（Word 原生 OMML）或 ⟦EQ:img:文件名⟧（图片公式）
    eq_index.json   公式索引：stream / img / bytes / variant / pairing；
                    OMML 条目带 kind="omml" 与 latex（已确定性转出）
    media/          从 docx 抽出的公式预览图（EMF/WMF）与图片公式原件
    image_inventory.json
                    正文全部行内图片的形状证据与上下文（公式候选的判定依据）
    summary.json    统计摘要（供工具层直接返回给模型）

设计要点:
    * 两族公式都要占位。旧版只认 <o:OLEObject>，OMML 文档的公式会连同
      结构一起静默消失（正文照常输出，不报错）——那是最坏的一类失败。
    * OMML（m:oMath）是结构化 XML，LaTeX 由 py/omml.py 在这里直接转出，
      写进 eq_index.json 的 latex 字段；build 自动预填，无需读图。
    * OLE 公式按出现顺序替换为占位符
    * 预览图按 OLEObject 的 ShapeID 精确配对到带 o:ole="t" 的 <v:shape>，
      再从该 shape 的 <v:imagedata> 取 rId。不依赖顺序，不受同段落插图干扰。
      ShapeID 缺失时才退回「同段落第 k 个公式配第 k 张图」，并把结果标为
      pairing="ordinal"，让调用方知道这张图未经确认。
    * MTEF 变体由 Equation Native 流前几字节判定：
        A = 起始 0x03，可解出字符（但会丢结构，只能做交叉校验）
        B = 起始 05 00 00 06 07 "DSMT6"，不可解，必须走 EMF→内嵌PDF 或看图
    * 本脚本不解码 MTEF 文本，也没有解码开关：变体 A 的解码结果丢结构严重
      （实测 151 字符的公式只解出 2 个字符），把它当"答案"会产出错误文档。
      变体判定只用于告诉调用方「这个公式能不能只靠字符」——答案是不能。
"""
import io
import json
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

from docxfile import BadDocument, fail, force_utf8_stdio, open_docx
import canvas
import docstyle
import imgsize
import imgformula as imgf
import omml
import rich

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _ln(tag):
    return tag.split("}")[-1] if "}" in tag else tag

EQ_OPEN = "\u27e6EQ:"      # ⟦EQ:
EQ_CLOSE = "\u27e7"        # ⟧

# 提取自检拦下「静默丢失」时的退出码。与「输入不是 docx」（3）区分开：
# 那个是用户的问题，这个是工具的问题，排查方向完全不同。
EXIT_SILENT_LOSS = 4


def natkey(name):
    """按文件名里的数字自然排序（oleObject2 在 oleObject10 之前）。"""
    return int(re.sub(r"\D", "", name) or 0)


def parent_map(root):
    """{子元素: 父元素}。

    标准库 ElementTree 没有 getparent()（那是 lxml 的），而判断「某节点是不是
    位于某个段落内部」必须往上走。这里一次性把父子关系建出来。
    """
    return {c: p for p in root.iter() for c in p}


def in_nested_p(node, host_p, parents):
    """node 是否归属于 host_p 内部的**另一个** w:p（文本框的段落）。

    文本框（w:txbxContent）里的段落自己也是一个 w:p。若外层段落把它展开，
    同一段内容会被记两遍：框里的文字在外层再出现一次，框里的公式占位符
    也出现两次。

    判定：从 node 往上走，遇到的第一个 w:p 若是 host_p → 归 host 管（False）；
    若是别的 w:p → 归那个嵌套段落管（True）。
    """
    anc = parents.get(node)
    while anc is not None:
        if anc.tag == W + "p":
            return anc is not host_p
        anc = parents.get(anc)
    return False


def outermost_p(node, parents):
    """node 最外层的 w:p 祖先（没有则 None）。

    **公式的归属按这个算**：公式画在流程图的文本框里时，它在 XML 上属于
    那个嵌套段落，但在语义上属于「放着这张图的那一段」。若按最近的 w:p
    归属，公式就会跟着嵌套段落一起被排除出正文——公式就这么丢了。

    所以：文字归最近的段落（嵌套段落排除出正文，标签另存），公式归最外层
    的段落（留在正文里，位置不丢）。两者各归其位，同一份内容只被记一次。
    """
    result = None
    anc = parents.get(node)
    while anc is not None:
        if anc.tag == W + "p":
            result = anc
        anc = parents.get(anc)
    return result


def is_nested_paragraph(p, parents):
    """这个 w:p 本身是不是位于某个文本框（另一个 w:p）内部。

    文本框里的段落不是正文段落：它是流程图的标签。混进正文会让段落数
    与原文对不上。
    """
    anc = parents.get(p)
    while anc is not None:
        if anc.tag == W + "p":
            return True
        anc = parents.get(anc)
    return False


def find_lost_placeholders(paras, rows):
    """哪些公式「索引里有、正文里却没有占位符」。

    返回排序后的 stream 名列表，空列表表示账目相符。

    这是防「静默丢失」的最后一道网。背景：w:pict 容器既可能是公式的
    预览图，也可能是**装着公式的画布**（Word 流程图：v:group 里既有
    o:ole="t" 的预览 shape，也有 o:OLEObject 本身）。两者混为一谈时，
    容器被整棵跳过，里面的公式占位符全部消失——正文照常输出，不报错
    不告警，直到 build 阶段 T8 才对不上数，而那时原因已无从查证。
    """
    placeholders = set()
    for x in paras:
        placeholders.update(re.findall(r"⟦EQ:([^⟧]+)⟧", x))
    # 画布内的公式不算"丢失"：它们已经画进画布 PNG，是**图的一部分**，
    # 不再单独占位。把它们算成丢失会让每次带流程图的文档都报假警。
    indexed = {r["stream"] for r in rows
               if r.get("kind") != "image" and not r.get("in_canvas")}
    return sorted(indexed - placeholders)


def find_duplicated_placeholders(paras, rows):
    """哪些公式的占位符**出现了不止一次**。

    同样属于「账目对不上」，但方向相反：不是少，是多。

    成因：同一份 XML 被遍历了两遍。实测踩过——公式画在流程图的文本框里，
    而段落收集用 `body.iter(w+"p")`，框里的段落既被当作独立段落收进来、
    又在外层段落里被访问一次，公式占位符因此出现两次。

    多出来的占位符和少掉的一样坏：build 会拿公式数去对，数字永远对不上，
    而且「多出来的那个」长得很像正文，排查时极易被当成「文档本来就这样」。
    """
    seen = {}
    for x in paras:
        for s in re.findall(r"⟦EQ:([^⟧]+)⟧", x):
            seen[s] = seen.get(s, 0) + 1
    indexed = {r["stream"] for r in rows
               if r.get("kind") != "image" and not r.get("in_canvas")}
    return sorted(s for s, n in seen.items() if n > 1 and s in indexed)


def get_eq_stream(raw_bin):
    """从 oleObjectN.bin（OLE2 复合文档）中取出 Equation Native 流。"""
    import olefile
    ole = olefile.OleFileIO(io.BytesIO(raw_bin))
    try:
        for s in ole.listdir():
            if s[-1] == "Equation Native":
                return ole.openstream("/".join(s)).read()
        return b""
    finally:
        ole.close()


def mtef_variant(eq):
    """判定 MTEF 变体：'A' 可解字符，'B' 需走 EMF/PDF 或看图。"""
    if len(eq) < 33:
        return "EMPTY"
    body = eq[28:]
    if body[:5] == b"\x05\x00\x00\x06\x07":
        return "B"
    if body[:1] == b"\x03":
        return "A"
    return "?"


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    src, outdir = sys.argv[1], sys.argv[2]

    # 先校验文档再建目录：坏输入不该在磁盘上留下一个空目录。
    try:
        z, root = open_docx(src)
    except BadDocument as exc:
        fail(str(exc))
    names = z.namelist()

    os.makedirs(outdir, exist_ok=True)

    rels_path = "word/_rels/document.xml.rels"
    rel_map = {}
    if rels_path in names:
        rels = z.read(rels_path).decode("utf-8", "replace")
        rel_map = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))

    body = root.find(W + "body")

    # 清掉上一次转换留下的产物。不复用旧目录内容：换一份文档复用同一个
    # outDir 时，残留的 media/ 会混进这次的图，导致公式与预览图错配。
    #
    # headings.json 也必须清：它是上一份文档**人工确认过**的段落索引 → 标题级别。
    # 段落号对另一份文档毫无意义，而 build.py 只要看到它存在就直接采用
    # （heading_source="headings.json"），于是新文档的标题会挂到错误的段落上，
    # 且不会有任何提示。宁可清掉让人重新确认一次。
    for stale in ("paras.json", "eq_index.json", "summary.json",
                  "build_report.json", "headings_suggested.json", "headings.json",
                  "image_inventory.json", "textbox_labels.json", "docstruct.json"):
        p_stale = os.path.join(outdir, stale)
        if os.path.exists(p_stale):
            os.remove(p_stale)
    for sub in ("media", "png"):
        d = os.path.join(outdir, sub)
        if os.path.isdir(d):
            for name in os.listdir(d):
                fp = os.path.join(d, name)
                if os.path.isfile(fp):
                    os.remove(fp)

    # ---- 1. 块级结构 + 行内格式 + 公式占位 ----
    # 三种公式载体都要占位：
    #   OLE（MTEF 二进制）  → ⟦EQ:oleObjectN.bin⟧，LaTeX 靠读图转写
    #   OMML（Word 原生）   → ⟦EQ:ommlK⟧，LaTeX 在这里就确定性转好
    #   图片（截图/扫描）    → ⟦EQ:img:media/xxx.png⟧，无法转换，
    #                          build 落成可见注释 + 图片引用，绝不静默丢
    # 漏掉任何一族，正文都会「安静地丢公式」——段落文字照常输出，公式无影
    # 无踪，校验器也查不出来（占位符本来就没生成）。
    #
    # 图片占位只对**公式候选**生成（扁条状行内图，见 imgformula.py）。
    # 示意图走同一占位会把正文搅成注释堆；它们进 image_inventory 供人核对。
    ole_ids = imgf.ole_shape_ids(body)
    ole_container = imgf.ole_preview_skip(body)
    parents = parent_map(body)
    imgf.set_parent_map(parents)
    # 图片自身的像素尺寸：判「扁条公式还是示意图」只认它（理由见 imgsize.py）。
    media_px = imgsize.load_media_sizes(z)
    paras = []
    omml_latex = []             # 与 ⟦EQ:ommlK⟧ 的 K 同序：转出的 LaTeX
    img_hits = []               # 正文里落成占位的公式图（按出现顺序）
    fig_hits = []               # 正文里落成占位的示意图（按出现顺序）
    box_text = []               # 文本框里的文字（流程图标签），单独留痕
    p_elems = []                # 与 paras 同序同口径的段落元素（供 docstyle 取结构）
    blocks = []                 # 块级结构（段落 / 表格 / 列表），build 按它输出
    notes = rich.read_notes(z)
    rels = rich.read_rels(z)

    # media/ 必须在**画布渲染之前**就存在：画布 PNG 直接写进去。
    # 早先它到第 3 步（抽预览图）才 makedirs，而画布在第 1 步就要写文件。
    media_dir = os.path.join(outdir, "media")
    os.makedirs(media_dir, exist_ok=True)

    # ---- 画布（流程图）识别 ----
    # v:group 画布在 docx 里不是一张图：它是并排放着的几十个形状（框、
    # 菱形、箭头、标签）+ 若干 OLE 公式。早先只认 imagedata 和 OLE，画布
    # 的骨架全部消失，正文里只剩一串互不相干的公式——读者完全看不出那
    # 是一张流程图，而且不报错不告警。
    #
    # 画布必须**整块**变成一张图：文字层无法表达"这些框怎么连"。所以
    # 画布内的公式不再各自占位（那样会和画布图重复），改由画布图承载。
    canvas_hits = []            # 已渲染的画布
    canvas_of = {}              # 画布元素 → 输出文件名

    # 画布 SVG 要把框内的公式预览图**内嵌**成 data URI（EMF/WMF 浏览器不认，
    # 必须先转成 PNG）。而媒体抽取和 EMF→PNG 渲染原本排在第 3、4 步，
    # 画布渲染在第 1 步——顺序反了，画布里的公式框就会是空的（只剩绿框）。
    #
    # 所以先把画布**自己**引用到的那些媒体抽出来并转好，再渲染画布。
    # 只处理画布用到的（通常十几张），不为此把全部 175 个媒体提前渲染。
    _canvas_media = set()
    for _g in canvas.find_canvases(body):
        for _im in _g.iter(V + "imagedata"):
            _rid = _im.get(R + "id")
            _t = rel_map.get(_rid) if _rid else None
            if _t:
                _canvas_media.add(_t.split("/")[-1])
    # DrawingML 画布（wpc）里的图也要提前备好——它是另一套表示，
    # 图片不在 v:imagedata 里，而在 a:blip 上。
    for _el in body.iter(canvas.WPC + "wpc"):
        for _b in _el.iter(A + "blip"):
            _rid = _b.get(R + "embed")
            _t = rel_map.get(_rid) if _rid else None
            if _t:
                _canvas_media.add(_t.split("/")[-1])
    if _canvas_media:
        os.makedirs(media_dir, exist_ok=True)
        for _fn in _canvas_media:
            _zp = "word/media/" + _fn
            if _zp in names and not os.path.exists(os.path.join(media_dir, _fn)):
                with open(os.path.join(media_dir, _fn), "wb") as _f:
                    _f.write(z.read(_zp))
        _wmf = [f for f in _canvas_media
                if os.path.splitext(f)[1].lower() in (".emf", ".wmf")]
        if _wmf:
            try:
                import render as _render
                _png_dir = os.path.join(outdir, "png")
                os.makedirs(_png_dir, exist_ok=True)
                for _eng in _render.ENGINES:
                    try:
                        if _eng(media_dir, _png_dir, 1600):
                            break
                    except Exception:                     # noqa: BLE001
                        continue
            except Exception:                             # noqa: BLE001
                pass

    # 画布里的形状 id 集合：画布内的 OLE 公式，其 OLEObject 是画布 group 的
    # **兄弟节点**（同在 w:pict 下），但 ShapeID 指向画布内部的 v:shape。
    # 这些公式已经画进画布 PNG 了，如果还按 OLEObject 逐个占位，读者会在
    # 同一段里看到「一张流程图 + 同一批公式再来一遍」，公式数也会虚高。
    canvas_shape_ids = set()
    for g in canvas.find_canvases(body):
        for sh in g.iter(V + "shape"):
            sid = sh.get("id")
            if sid:
                canvas_shape_ids.add(sid)

    def _canvas_placeholder(g, extent=None):
        """把一个画布渲染成 media/ 里的 PNG，返回占位符文本。

        两种画布都走这里：
          * VML `v:group`（流程图、VML 图版）
          * DrawingML `wpc`（Word 2010+ 的绘图画布，带每张图的绝对坐标）
        """
        if g in canvas_of:
            return canvas_of[g]
        fname = "canvas%d.png" % (len(canvas_hits) + 1)
        out_png = os.path.join(media_dir, fname)
        if canvas.is_drawing_canvas(g):
            ok, why = canvas.render_drawing_canvas(g, rel_map, outdir,
                                                   out_png, extent)
        else:
            ok, why = canvas.render_canvas(g, rel_map, outdir, out_png)
        canvas_of[g] = (fname, ok, why)
        canvas_hits.append({"file": fname, "ok": ok, "reason": why})
        if ok:
            return "%sfig:%s%s" % (EQ_OPEN, fname, EQ_CLOSE)
        return None

    def _emit_media(node, host_p):
        """把容器里的图片 / 公式按**文档顺序**变成占位符文本。

        这是全流程唯一的「非文字内容 → 文本」出口。规则：
          * OLE 公式       → ⟦EQ:oleObjectN.bin⟧
          * OMML 公式      → ⟦EQ:ommlK⟧
          * 公式候选图片    → ⟦EQ:img:文件名⟧
          * 示意图 / 认不出 → ⟦EQ:fig:文件名⟧
          * 公式的预览图    → 跳过（由 OLE 那条路负责，不能重复占位）

        必须按 node.iter() 的**文档顺序**走，不能先收公式再收图片：
        一个 v:group 里可以既有公式预览、又有插图（实测 p281：6 张三维图 +
        6 个公式预览），顺序就是读者看到的顺序，打乱它等于改内容。
        """
        def owned(el):
            """这个 OLE/OMML 是否归 host_p 这个段落认领。

            host_p 为 None 表示「不按段落归属」（rich 层的回调不传段落）：
            此时一律认领，因为每个块只会被遍历一次，不存在重复计数。
            """
            if host_p is None:
                return True
            return outermost_p(el, parents) is host_p

        out = []
        seen = set()
        # 纯 OLE 预览图容器：里面没有别的内容，整棵跳过。
        # 但容器自己装着 OLEObject 时 is_ole_preview 会返回 False（见其
        # docstring），所以这里不会吞掉公式。
        skip_preview = imgf.is_ole_preview(node, ole_ids, ole_container)

        def _walk(el):
            """按文档顺序遍历，画布子树整体短路。

            用递归而不是 `node.iter()`：iter 没法在画布处停下。画布一旦
            命中就必须**不再向下走**——否则画布里的公式/图片又各占一次位，
            和画布图重复，读者会看到同一批公式两遍。
            """
            if el is not node and canvas.is_canvas(el):
                ph = _canvas_placeholder(el)
                if ph:
                    out.append(ph)
                return
            # DrawingML 画布（wpc）：同一份图版在 mc:Fallback 里还有一个
            # v:group 副本。这里认 wpc 就整块短路，Fallback 那份不会再被
            # 走到（rich.py 只挑 Choice），不会重复出图。
            if el is not node and canvas.is_drawing_canvas(el):
                ext = None
                cur = parents.get(el)
                while cur is not None:
                    if cur.tag == canvas.WP + "extent":
                        try:
                            ext = (int(cur.get("cx")), int(cur.get("cy")))
                        except (TypeError, ValueError):
                            ext = None
                        break
                    cur = parents.get(cur)
                ph = _canvas_placeholder(el, ext)
                if ph:
                    out.append(ph)
                return
            tag = el.tag
            if tag == O + "OLEObject":
                if owned(el) and el.get("ShapeID") not in canvas_shape_ids:
                    rid = el.get(R + "id")
                    tgt = (rel_map.get(rid) or "?").split("/")[-1]
                    out.append("%s%s%s" % (EQ_OPEN, tgt, EQ_CLOSE))
            elif tag == M + "oMath":
                # m:oMathPara 内部就是 m:oMath；只认 oMath，避免双份。
                if owned(el):
                    tex = omml.to_latex(el)
                    omml_latex.append(tex)
                    out.append("%somml%d%s" % (EQ_OPEN, len(omml_latex), EQ_CLOSE))
            elif tag in (V + "imagedata", A + "blip"):
                _emit_picture(el)
            # 叶子节点没有子节点，不必递归（也是热路径）
            for ch in el:
                _walk(ch)

        def _emit_picture(el):
            """一张图片 → 占位符（或按规则跳过）。"""
            if skip_preview:
                return
            rid = el.get(R + "id") or el.get(R + "embed")
            if not rid or rid in seen:
                return
            # 这一张是不是某个 OLE 公式的预览（逐张判，不能按容器判）
            if imgf.leaf_is_ole_preview(el, ole_ids, node, ole_container):
                return
            seen.add(rid)
            target = rel_map.get(rid) or ""
            if not target:
                return
            fname = target.split("/")[-1]
            w_px, h_px = media_px.get(fname, (None, None))
            _lw, _lh = _leaf_wh(el, node)
            shape = (imgf.classify_px(w_px, h_px)
                     or imgf.classify_picture(_lw, _lh))
            if shape == "formula":
                img_hits.append({"file": fname, "target": target})
                out.append("%simg:%s%s" % (EQ_OPEN, fname, EQ_CLOSE))
            elif (w_px and h_px and not imgf.is_illustration_px(w_px, h_px)):
                # 已知尺寸、且小到像排版占位条（实测真实文档标题下有一张
                # 2×2 的 image1.png，125 字节）——它不是内容，落成插图只会
                # 插一行破图。仍然记进清单留痕，只是不进正文；这是**唯一**
                # 允许"不进正文"的图，而且它没有任何可见内容可丢。
                fig_hits.append({"file": fname, "target": target,
                                 "skipped": "排版占位条"})
            else:
                # 示意图（以及尺寸读不出来、小到像排版占位条的）：**必须
                # 进正文**。旧版在这里跳过、只记进 image_inventory 就算
                # "留痕"，于是正文里一张不剩、md 里连占位都没有——本工具
                # 最忌讳的静默丢失。
                fig_hits.append({"file": fname, "target": target})
                out.append("%sfig:%s%s" % (EQ_OPEN, fname, EQ_CLOSE))

        _walk(node)
        return "".join(out) or None

    def _leaf_wh(el, boundary):
        """一张图自己的 (宽, 高) EMU；取不到返回 (None, None)。"""
        cur = el
        while cur is not None:
            nm = cur.tag.split("}")[-1] if "}" in cur.tag else cur.tag
            if nm in ("extent", "ext"):
                try:
                    return int(cur.get("cx")), int(cur.get("cy"))
                except (TypeError, ValueError):
                    pass
            if nm == "shape":
                st = cur.get("style") or ""
                mw = re.search(r"width\s*:\s*([^;]+)", st)
                mh = re.search(r"height\s*:\s*([^;]+)", st)
                if mw and mh:
                    return imgf._emu_len(mw.group(1)), imgf._emu_len(mh.group(1))
            if cur is boundary:
                break
            cur = parents.get(cur)
        return None, None

    def _media_cb(node):
        """给 rich.runs_of 用的回调：返回占位文本或 None。"""
        return _emit_media(node, None)

    ctx = {"rels": rels, "notes": notes, "media": _media_cb}
    raw_blocks = rich.blocks(z, body, ctx)

    # 把 rich 的 Block 展开成「正文段落序列」+「块结构记录」。
    #
    # 为什么还要扁平成 paras：整条下游（公式占位计数、标题表按段落索引、
    # 校验器 T8 数占位符）都建立在「一段一行」上。表格单元格是块级结构，
    # 但它的文字同样必须进入 paras 参与占位/计数——否则表里的公式会被判
    # 成"没出现过"。两者用 block_map 对齐：每个 paras 下标属于哪个块。
    block_map = []              # paras 下标 → 块序号
    for bi, b in enumerate(raw_blocks):
        if b["kind"] == "para":
            if is_nested_paragraph(b["el"], parents):
                continue
            runs = b["runs"]
            # 空段落也要保留：它是段落分隔，build 用它控制空行
            txt = rich.plain(runs)
            paras.append(txt)
            p_elems.append(b["el"])
            block_map.append(len(blocks))
            num = b.get("num")
            lk = None
            if num and num.get("numId"):
                try:
                    lk = docstyle.list_kind(z, num["numId"], num.get("ilvl"))
                except Exception:                        # noqa: BLE001
                    lk = None
            blocks.append({"kind": "para", "index": len(paras) - 1,
                           "runs": runs, "num": num, "list": lk})
        elif b["kind"] == "table":
            rows = []
            first_cell_para = len(paras)     # 表格占用的 paras 区间起点
            for row in b["rows"]:
                cell_txt = []
                for cell in row:
                    runs = cell["runs"]
                    t = rich.plain(runs).strip()
                    if t:
                        # 单元格文字进 paras（参与占位与计数），但**不**进
                        # block_map——它属于表格块，不是独立正文段落。
                        paras.append(t)
                        p_elems.append(b["el"])
                    cell_txt.append({"text": t, "span": cell["span"],
                                     "vmerge": cell["vmerge"]})
                rows.append(cell_txt)
            blocks.append({"kind": "table", "rows": rows,
                           "header": b["header"], "index": len(blocks),
                           # 表格单元格的文字也进了 paras（参与公式占位与
                           # 计数），build 必须跳过这段区间，否则同一段文字
                           # 会既在表格里出现、又作为普通段落再出现一次。
                           "para_from": first_cell_para,
                           "para_to": len(paras)})
        else:
            # 未知块：降级成段落，不丢
            paras.append(rich.plain(b.get("runs") or []))
            p_elems.append(b.get("el"))
            block_map.append(len(blocks))
            blocks.append({"kind": "para", "index": len(paras) - 1,
                           "runs": b.get("runs") or []})

    # 文本框里的段落不是正文段落——它们是流程图的标签。留在 paras 里会
    # 让读者看到一堆莫名其妙的碎片（"VDDD3D程序结构图" / "结束" / "、位错密度等"
    # 各自成段），而且会让段落数与原文对不上。
    #
    # 但它们也不能丢：那是文档里真实存在的文字。单独进 textbox_labels.json，
    # 留给人核对。**框里的公式已经占位**（上面的循环），这里只收纯文字。
    for p in body.iter(W + "p"):
        if not is_nested_paragraph(p, parents):
            continue
        t = "".join(x.text or "" for x in p.iter(W + "t")).strip()
        if t:
            box_text.append(t)

    # 图片清单：正文里所有行内/浮动图片（排除 OLE 预览），带形状证据与
    # 所在段落文字。公式候选只是启发式，最终判断交给人/Agent——这份清单
    # 就是给它看的。
    inventory = []
    for pi, p in enumerate(body.iter(W + "p")):
        # 只统计**正文段落**：文本框里的副本（v:group 里同一张图又被
        # 框内段落引用一次）不是独立的插图，否则 p271 会在 271/273/276
        # 各出现一次，清单里同一张 image135.emf 记三遍。
        if is_nested_paragraph(p, parents):
            continue
        has_text = bool(imgf.para_text(p))
        listed = set()
        for node in p.iter():
            if _ln(node.tag) not in ("drawing", "pict"):
                continue
            if imgf.is_ole_preview(node, ole_ids, ole_container):
                continue
            # 逐张登记：一个容器里可以并排放十几张图（实测 p281 有 12 张）。
            # 只记第一张会让其余的原件既不在清单里、也不被抽出。
            for leaf in imgf.picture_leaves(node):
                if imgf.leaf_is_ole_preview(leaf["el"], ole_ids, node,
                                            ole_container):
                    continue
                if leaf["rid"] in listed:
                    continue
                listed.add(leaf["rid"])
                target = rel_map.get(leaf["rid"]) or ""
                if not target:
                    continue
                fname = target.split("/")[-1]
                w_px, h_px = media_px.get(fname, (None, None))
                w_emu, h_emu = leaf["w"], leaf["h"]
                inventory.append({
                    "file": fname,
                    "target": target,
                    "paragraph": pi,
                    # 形状口径与正文判定完全一致（像素优先），否则清单说
                    # "插图"、正文当公式处理，两边各说各话。
                    "shape": (imgf.classify_px(w_px, h_px)
                              or imgf.classify_picture(w_emu, h_emu)),
                    "inline_with_text": has_text,
                    "ratio": (round(w_px / h_px, 2) if w_px and h_px
                              else (round(w_emu / h_emu, 2) if w_emu and h_emu
                                    else None)),
                    "size_px": ([w_px, h_px] if w_px and h_px else None),
                    "context": imgf.para_text(p)[:80],
                })

    # 画布（流程图）也登记进清单：它是一整块结构，用户核对"图齐不齐"时
    # 必须能看到它，而不是只看到画布内部的零散图片。
    for ci, ch in enumerate(canvas_hits, 1):
        inventory.append({
            "file": ch["file"],
            "target": "canvas",
            "paragraph": None,
            "kind": "canvas",
            "shape": "canvas",
            "rendered": bool(ch["ok"]),
            "reason": ch.get("reason") or "",
            "context": "VML v:group 画布（流程图/示意图），整块渲染",
        })

    with open(os.path.join(outdir, "paras.json"), "w", encoding="utf-8") as f:
        json.dump(paras, f, ensure_ascii=False, indent=1)

    # ---- 1a. 块级结构（表格 / 段落顺序 / 行内格式）----
    # paras.json 是"一串段落"，它天然表达不了表格：表格的行列关系一旦被
    # 平铺成段落就永久丢失。blocks.json 记的是正文的**真实结构**——
    # 哪些是段落、哪些是表格、表格几行几列、每个格子归哪一行哪一列。
    #
    # 同时记下行内格式（加粗/斜体/下划线/删除线/上下标）、超链接目标、
    # 脚注正文。这些都在 docx 里明明白白写着，丢掉它们不是"做不到"，
    # 只是没读。格式转换器的正确性包含格式。
    try:
        with open(os.path.join(outdir, "blocks.json"), "w",
                  encoding="utf-8") as f:
            json.dump({"blocks": blocks, "paragraphs": len(paras)},
                      f, ensure_ascii=False, indent=1)
    except (TypeError, ValueError) as exc:                # noqa: BLE001
        sys.stderr.write("警告: 块结构落盘失败（表格将退化为段落）: %s\n" % exc)

    # ---- 1b. 段落结构信号（样式级别 / 自动编号 / 整段加粗）----
    # 为什么单独落一个文件而不是塞进 paras.json：paras.json 是「正文」，
    # 结构信号是判标题用的旁证，两者的消费方不同。放进 paras.json 会让
    # build.py 的文本处理函数多一个非字符串元素要防，得不偿失。
    #
    # 这些信号正是「自动编号标题」的唯一线索：中文文档里 "iii)" 这类编号
    # 从不写进文字，由 Word 按 numbering.xml 现算。不还原它们，正则只能
    # 认出带字面前缀的少数标题（实测 12/32）。
    try:
        marks = docstyle.paragraph_marks(z, p_elems)
    except Exception as exc:                              # noqa: BLE001
        # 结构读取只是加分项：任何异常都不能让整次提取失败，退回纯正则。
        sys.stderr.write("警告: 段落结构读取失败（退回纯正则判标题）: %s\n" % exc)
        marks = []
    if marks:
        with open(os.path.join(outdir, "docstruct.json"), "w", encoding="utf-8") as f:
            json.dump(marks, f, ensure_ascii=False, indent=1)

    # ---- 2. 公式 → 预览图映射 ----
    # rel_map 的 Target 形如 "media/image49.wmf"（相对 word/）。
    #
    # 主路径：按 ShapeID 精确配对。
    #   公式的预览图是一个带 o:ole="t" 的 v:shape，它的 id 等于 OLEObject 的
    #   ShapeID。这是唯一可靠的联系——不依赖顺序、不受同段落其他插图干扰。
    #
    # 早期版本按「同段落第 k 个 OLEObject ↔ 第 k 张 v:imagedata」配对，实测在
    # 真实文档上会错位：段落里公式之前夹着程序截图时，整个下标右移，把插图
    # 的位图当成公式预览图交给模型去读，而且不报错。实测 145 个公式里有 10 个
    # 被配成 VDDD3D 程序截图/位错环三维图。
    #
    # 回退：ShapeID 缺失或对不上时（少数生成器不写 ShapeID），退回下标配对，
    # 但把结果标记为 unverified，让 summary 能报出来。
    shape_preview = {}          # shape id -> media target
    for p in body.iter(W + "p"):
        for sh in p.iter(V + "shape"):
            if sh.get(O + "ole") != "t":
                continue
            ims = list(sh.iter(V + "imagedata"))
            if ims:
                shape_preview[sh.get("id")] = rel_map.get(ims[0].get(R + "id"))

    pairs = {}
    pairing = {}                # stream -> "shape-id" | "ordinal"
    ole_shape_id = {}           # stream -> OLEObject@ShapeID（判"是否在画布内"）
    for p in body.iter(W + "p"):
        imgs = [i.get(R + "id") for i in p.iter(V + "imagedata")]
        for k, obj in enumerate(p.iter(O + "OLEObject")):
            oid = (rel_map.get(obj.get(R + "id")) or "?").split("/")[-1]
            sid = obj.get("ShapeID")
            if sid:
                ole_shape_id[oid] = sid
            target = shape_preview.get(sid)
            if target:
                pairs[oid] = target
                pairing[oid] = "shape-id"
            elif k < len(imgs):
                pairs[oid] = rel_map.get(imgs[k]) or None      # "media/image49.wmf"
                pairing[oid] = "ordinal"
            else:
                pairs[oid] = None
                pairing[oid] = "none"

    # ---- 3. 抽取预览图到 media/ ----
    os.makedirs(media_dir, exist_ok=True)
    extracted = set()
    for rel_target in list(pairs.values()):
        if not rel_target:
            continue
        zpath = rel_target if rel_target.startswith("word/") else "word/" + rel_target
        if zpath in names and zpath not in extracted:
            with open(os.path.join(media_dir, os.path.basename(rel_target)), "wb") as f:
                f.write(z.read(zpath))
            extracted.add(zpath)
    # 公式候选的位图原件也抽出来：它本身就是"预览图"，Agent 可直接读它转写
    for hit in img_hits:
        t = hit["target"]
        zpath = t if t.startswith("word/") else "word/" + t
        if zpath in names and zpath not in extracted:
            with open(os.path.join(media_dir, os.path.basename(t)), "wb") as f:
                f.write(z.read(zpath))
            extracted.add(zpath)

    # 示意图（inventory 里 shape=illustration 的那些）也要抽出来。
    #
    # 清单里写着「image135.emf 在段落 271」，但 media/ 里根本没有这个文件——
    # 清单指着一个**不存在**的东西，人去核对时无从下手。清单与原件必须
    # 同时到位，否则「留痕供人核对」就是一句空话。
    for e in inventory:
        target = e.get("target") or ""
        fname = e.get("file") or ""
        if not fname or not target:
            continue
        zpath = target if target.startswith("word/") else "word/" + target
        if zpath in names and zpath not in extracted:
            with open(os.path.join(media_dir, fname), "wb") as f:
                f.write(z.read(zpath))
            extracted.add(zpath)

    # ---- 4. 公式索引 ----
    def img_of(base):
        """pairs 里存的是 'media/image49.wmf'，索引里统一用文件名。"""
        target = pairs.get(base)
        return os.path.basename(target) if target else None

    rows = []
    for name in sorted([n for n in names
                        if n.startswith("word/embeddings/") and n.endswith(".bin")],
                       key=natkey):
        base = name.split("/")[-1]
        try:
            eq = get_eq_stream(z.read(name))
        except Exception as exc:            # 损坏的 OLE 不应中断整批提取
            row = {
                "stream": base,
                "img": img_of(base),
                "bytes": 0,
                "variant": "ERROR",
                "pairing": pairing.get(base, "none"),
                "error": str(exc),
            }
            # 异常分支也必须标 in_canvas：读不出 MTEF 与「在不在画布里」
            # 是两件独立的事。漏标会让丢失自检把画布内公式报成"丢失"。
            _sid = ole_shape_id.get(base)
            if _sid and _sid in canvas_shape_ids:
                row["in_canvas"] = True
            rows.append(row)
            continue
        row = {
            "stream": base,
            "img": img_of(base),
            "bytes": len(eq),
            "variant": mtef_variant(eq),
            "pairing": pairing.get(base, "none"),
        }
        # 画布内的公式：已经画进画布 PNG，不再单独占位。
        # 必须标记出来，否则下游（丢失自检、T8 计数、latex.json 覆盖率）
        # 会把它当成"应该有却找不到"的公式，天天报假警。
        sid = ole_shape_id.get(base)
        if sid and sid in canvas_shape_ids:
            row["in_canvas"] = True
            row["note"] = ("该公式位于流程图/画布内，已随画布整块渲染成图片，"
                           "不再单独占位，无需转写。")
        rows.append(row)

    # OMML 公式：LaTeX 已经确定性转出，直接带在索引里（kind=omml）。
    # build.py 会把它预填进公式表；Agent 的 latex.json 若给了同名键则以
    # Agent 为准——留了人工/模型纠正的余地。
    for k, tex in enumerate(omml_latex, 1):
        rows.append({
            "stream": "omml%d" % k,
            "kind": "omml",
            "img": None,
            "bytes": 0,
            "variant": "OMML",
            "pairing": "xml",          # 结构直接来自 XML，不涉及图片配对
            "latex": tex,
        })

    # 图片公式：正文落成 ⟦EQ:img:文件名⟧ 占位。build 不替换它时会落成
    # 可见注释 + 图片引用（绝不静默丢）；Agent 读 media/ 里的原件转写后，
    # 用 "img:文件名" 作键写进 latex.json 即可替换成 LaTeX。
    for hit in img_hits:
        rows.append({
            "stream": "img:%s" % hit["file"],
            "kind": "image",
            "img": hit["file"],
            "bytes": 0,
            "variant": "IMG",
            "pairing": "n/a",          # 图片本身就是内容，不存在配对问题
            "note": "以图片形式存在的公式候选（扁条状行内图）。读 media/%s "
                    "转写 LaTeX，键用 \"img:%s\"；判断它其实是插图时，"
                    "从正文删掉该占位即可。" % (hit["file"], hit["file"]),
        })

    with open(os.path.join(outdir, "eq_index.json"), "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)

    with open(os.path.join(outdir, "image_inventory.json"), "w", encoding="utf-8") as f:
        json.dump(inventory, f, ensure_ascii=False, indent=1)

    # 文本框（流程图）里的标签文字：不是正文段落，但也不能丢。
    with open(os.path.join(outdir, "textbox_labels.json"), "w", encoding="utf-8") as f:
        json.dump(box_text, f, ensure_ascii=False, indent=1)

    # ---- 5. 摘要 ----
    n_empty = sum(1 for x in paras if not x.strip())
    counts = {}
    for r in rows:
        counts[r["variant"]] = counts.get(r["variant"], 0) + 1
    ole_rows = [r for r in rows if r.get("kind") not in ("omml", "image")]
    omml_rows = [r for r in rows if r.get("kind") == "omml"]
    image_rows = [r for r in rows if r.get("kind") == "image"]
    # 「缺图」只对 OLE 公式有意义：OMML 的 LaTeX 直接从 XML 转出，不需要图。
    no_media = [r["stream"] for r in ole_rows if not r.get("img")]
    # 靠下标回退配上的公式：图可能是插图的位图，必须让人知道。
    by_ordinal = [r["stream"] for r in ole_rows if r.get("pairing") == "ordinal"]

    # ---- 自检：公式索引里的每一条，正文里都必须有对应的占位符 ----
    # 把「静默」变成「显式」：索引有、正文没有 → 直接报错退出。
    lost = find_lost_placeholders(paras, rows)
    if lost:
        # 不用 fail()：那是「输入不对」，这里是「工具自己的账对不上」，
        # 退出码要能区分（4 = 静默丢失已拦下）。
        force_utf8_stdio()
        sys.stderr.write(
            "提取自检失败：正文里少了 %d 个公式占位符（%s…）。\n"
            "索引里有、正文里没有，说明公式在提取时被整段跳过了——"
            "这属于静默丢失，必须查清是哪种容器把公式吞了，"
            "不能带着缺口往下走。\n"
            % (len(lost), ", ".join(lost[:5])))
        sys.exit(EXIT_SILENT_LOSS)

    # ---- 自检（另一方向）：同一个公式的占位符不能出现多次 ----
    # 少和多同样坏：多的会让公式数与 LaTeX 表永远对不上，而「多出来的
    # 那个」长得很像正文，排查时极易被当成「文档本来就这样」。
    dup = find_duplicated_placeholders(paras, rows)
    if dup:
        force_utf8_stdio()
        sys.stderr.write(
            "提取自检失败：%d 个公式的占位符在正文里出现了多次（%s…）。\n"
            "同一份内容被收了两遍——通常是文本框（流程图）里的内容既被当作"
            "独立段落、又在外层段落里被访问了一次。必须查清，不能带着重复往下走，"
            "否则公式数永远对不上 LaTeX 表。\n"
            % (len(dup), ", ".join(dup[:5])))
        sys.exit(EXIT_SILENT_LOSS)

    summary = {
        "paragraphs": len(paras),
        "paragraphs_nonempty": len(paras) - n_empty,
        # 文本框（流程图）里的文字：不是正文段落，但真实存在，单独留痕。
        "textbox_labels": len(box_text),
        # 口径与 inspect 一致：formulas 只数 OLE；OMML / 图片单独报。
        # 两个工具对「formulas」给两个数，混排文档里 expectFormulas 就会传错。
        "formulas": len(ole_rows),
        "formulas_omml": len(omml_rows),
        "formulas_image": len(image_rows),
        "formulas_total": len(rows),
        "variants": counts,
        "media_extracted": len(extracted),
        "formulas_without_media": len(no_media),
        "formulas_paired_by_shape_id": sum(1 for r in ole_rows if r.get("pairing") == "shape-id"),
        "formulas_paired_by_ordinal": len(by_ordinal),
        "next": (
            "全部公式都有按 ShapeID 确认的预览图，可逐张读图转写 LaTeX。"
            if not no_media and not by_ordinal else
            ("有 %d 个公式没有预览图（%s），只能靠 MTEF 或原始 Word 补齐。"
             % (len(no_media), ", ".join(no_media[:5])))
            if no_media else
            ("有 %d 个公式的预览图是按出现顺序回退配上的（%s），"
             "无法确认它们不是同段落插图的位图；读图前建议先核对。"
             % (len(by_ordinal), ", ".join(by_ordinal[:5])))
        ),
    }
    if image_rows:
        summary["next_image"] = (
            "正文里有 %d 处以图片形式存在的公式候选（按形状判定，见 "
            "image_inventory.json 的完整清单）。它们已落成 ⟦EQ:img:…⟧ 占位、"
            "原件在 media/；读图转写后以 \"img:文件名\" 为键写进 latex.json。"
            "不处理也不丢——build 会把它们落成可见注释与图片引用。"
            % len(image_rows))
    # 示意图：不是公式，但同样不能丢。它们在正文里落成 ⟦EQ:fig:…⟧ 占位，
    # build 直接转成 ![图](…)，无需任何后续人工动作。
    if fig_hits:
        summary["figures"] = len(fig_hits)
        summary["next_figure"] = (
            "正文里有 %d 张示意图，已落成 ![图](…) 图片引用（原件在 media/，"
            "EMF/WMF 引用 png/ 里渲染好的 PNG）。无需转写；若要调整某张的"
            "位置或删掉，直接编辑 md 即可。" % len(fig_hits))
    # 画布（流程图）必须报出来：它是一整块结构，用户最关心「图还在不在」。
    # 渲染失败时**不能静默**——必须说清是哪一块、为什么失败，因为一块
    # 流程图消失比一张插图消失严重得多（它承载的是整章的逻辑关系）。
    if canvas_hits:
        ok_n = sum(1 for c in canvas_hits if c["ok"])
        summary["canvases"] = len(canvas_hits)
        summary["canvases_rendered"] = ok_n
        if ok_n == len(canvas_hits):
            summary["next_canvas"] = (
                "文档里有 %d 块流程图/画布（VML v:group），已整块渲染成 "
                "media/canvas*.png 并在正文落成 ![图](…) 引用。画布内的公式已"
                "画进那张图，不再单独占位。" % ok_n)
        else:
            summary["canvas_warning"] = (
                "有 %d 块画布渲染失败（%s），内容尚未还原：%s。"
                "流程图里的形状与公式会整块丢失，请检查渲染环境。"
                % (len(canvas_hits) - ok_n,
                   ", ".join(c["file"] for c in canvas_hits if not c["ok"]),
                   canvas_hits[0]["reason"] or "未知原因"))
    if omml_rows:
        summary["next_omml"] = (
            "另有 %d 个 Word 原生公式（OMML），LaTeX 已直接写入 eq_index.json "
            "的 latex 字段，build 会自动预填，无需读图；不满意时可在 latex.json "
            "里用同名键覆盖。" % len(omml_rows))
    if by_ordinal:
        summary["pairing_warning"] = (
            "这些公式的文档里找不到 o:ole=\"t\" 的 v:shape（或 ShapeID 不匹配），"
            "退回了「同段落第 k 个公式配第 k 张图」。若段落内有插图，配到的可能是插图。"
        )
    with open(os.path.join(outdir, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=1)

    print(json.dumps(summary, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    force_utf8_stdio()
    main()
