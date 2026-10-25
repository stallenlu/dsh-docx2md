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
import imgformula as imgf
import omml

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


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
    indexed = {r["stream"] for r in rows if r.get("kind") != "image"}
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
    indexed = {r["stream"] for r in rows if r.get("kind") != "image"}
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
                  "image_inventory.json", "textbox_labels.json"):
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

    # ---- 1. 段落文本 + 公式占位 ----
    # 三种载体都要占位：
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
    paras = []
    omml_latex = []             # 与 ⟦EQ:ommlK⟧ 的 K 同序：转出的 LaTeX
    img_hits = []               # 正文里落成占位的公式图（按出现顺序）
    box_text = []               # 文本框里的文字（流程图标签），单独留痕
    for p in body.iter(W + "p"):
        # 文本框里的段落不产生正文段落：它是流程图的标签。标签进
        # textbox_labels.json（见下面），公式由最外层段落认领。
        if is_nested_paragraph(p, parents):
            continue
        buf = []
        consumed = set()        # 已作为图片容器处理过的子树元素 id
        for node in p.iter():
            if id(node) in consumed:
                continue
            if node.tag == W + "t":
                # 文本框里的文字是流程图的标签，不是正文。归那个嵌套
                # 段落管，否则同一段内容会被记两遍；框里的文字另存到
                # textbox_labels.json，不丢。
                if in_nested_p(node, p, parents):
                    continue
                buf.append(node.text or "")
            elif node.tag == W + "tab":
                buf.append("\t")
            elif node.tag == W + "br":
                buf.append("\n")
            elif node.tag == O + "OLEObject":
                # 公式占位：**无论公式住在正文还是文本框里，都只占一次**。
                # 框里的公式也必须占位——公式不能丢；但它不能既算框外的、
                # 又算框内的（同一条 XML 会被外层段落与框内段落各访问一次）。
                # 归属规则：归**最外层**那个段落（放着这张图的那一段），
                # 嵌套段落不认领它——否则公式会跟着标签一起被排除出正文。
                if outermost_p(node, parents) is not p:
                    continue
                rid = node.get(R + "id")
                tgt = (rel_map.get(rid) or "?").split("/")[-1]
                buf.append("%s%s%s" % (EQ_OPEN, tgt, EQ_CLOSE))
            elif node.tag == M + "oMath":
                # m:oMathPara（独立成段的公式）内部就是 m:oMath，
                # iter 会先进 oMathPara 再进 oMath，这里只认 oMath。
                if outermost_p(node, parents) is not p:
                    continue
                tex = omml.to_latex(node)
                omml_latex.append(tex)
                buf.append("%somml%d%s" % (EQ_OPEN, len(omml_latex), EQ_CLOSE))
            elif _ln(node.tag) in ("drawing", "pict"):
                # 整棵子树一次性处理，防止嵌套元素被重复访问。
                # 但**不能整棵跳过**：Word 画流程图时，公式可以嵌在同一个
                # w:pict 的 v:group 里（实测：145 公式文档里有 3 个 pict 各
                # 装着 6–8 个公式）。整棵标记为已消费会把里面的
                # o:OLEObject 一并吞掉，公式占位符静默消失。
                # 正确做法：只消费「图片本体」那部分，公式节点留给自己处理。
                if imgf.is_ole_preview(node, ole_ids, ole_container):
                    for d in node.iter():
                        consumed.add(id(d))
                    continue                    # OLE 预览图，走 OLE 路径
                if imgf.holds_ole_formula(node):
                    # 容器里既有公式、又有图：只把图的部分标记为已消费，
                    # 公式的 o:OLEObject 留给上面的分支生成占位符。
                    for d in node.iter():
                        if d.tag == O + "OLEObject":
                            continue
                        consumed.add(id(d))
                else:
                    for d in node.iter():
                        consumed.add(id(d))
                rid = imgf.picture_rid(node)
                if not rid:
                    continue
                w_emu, h_emu = imgf.picture_size(node)
                target = rel_map.get(rid) or ""
                fname = target.split("/")[-1]
                shape = imgf.classify_picture(w_emu, h_emu)
                if shape != "formula":
                    continue                    # 示意图：进清单，不进正文
                img_hits.append({"file": fname, "target": target})
                buf.append("%simg:%s%s" % (EQ_OPEN, fname, EQ_CLOSE))
        paras.append("".join(buf))

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
        has_text = bool(imgf.para_text(p))
        listed = set()
        for run in p.iter(W + "r"):
            for child in run:
                if _ln(child.tag) not in ("drawing", "pict"):
                    continue
                if imgf.is_ole_preview(child, ole_ids, ole_container):
                    continue
                rid = imgf.picture_rid(child)
                if not rid or (pi, rid) in listed:
                    continue
                listed.add((pi, rid))
                w_emu, h_emu = imgf.picture_size(child)
                target = rel_map.get(rid) or ""
                inventory.append({
                    "file": target.split("/")[-1],
                    "paragraph": pi,
                    "shape": imgf.classify_picture(w_emu, h_emu),
                    "inline_with_text": has_text,
                    "ratio": (round(w_emu / h_emu, 2) if w_emu and h_emu else None),
                    "h_px": (round(h_emu / imgf.EMU_PER_PX) if h_emu else None),
                    "context": imgf.para_text(p)[:80],
                })

    with open(os.path.join(outdir, "paras.json"), "w", encoding="utf-8") as f:
        json.dump(paras, f, ensure_ascii=False, indent=1)

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
    for p in body.iter(W + "p"):
        imgs = [i.get(R + "id") for i in p.iter(V + "imagedata")]
        for k, obj in enumerate(p.iter(O + "OLEObject")):
            oid = (rel_map.get(obj.get(R + "id")) or "?").split("/")[-1]
            target = shape_preview.get(obj.get("ShapeID"))
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
    media_dir = os.path.join(outdir, "media")
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
            rows.append({
                "stream": base,
                "img": img_of(base),
                "bytes": 0,
                "variant": "ERROR",
                "pairing": pairing.get(base, "none"),
                "error": str(exc),
            })
            continue
        rows.append({
            "stream": base,
            "img": img_of(base),
            "bytes": len(eq),
            "variant": mtef_variant(eq),
            "pairing": pairing.get(base, "none"),
        })

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
