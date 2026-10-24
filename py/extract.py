#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""extract.py —— 从 .docx 提取正文段落、公式索引与公式内嵌预览图。

用法:
    python extract.py <input.docx> <outdir>

产出（全部落在 <outdir>）:
    paras.json      段落列表，公式位置标为 ⟦EQ:oleObjectN.bin⟧
    eq_index.json   公式索引：stream / img / media / bytes / variant
    media/          从 docx 抽出的公式预览图（EMF/WMF 原格式）
    summary.json    统计摘要（供工具层直接返回给模型）

设计要点:
    * 公式在 document.xml 中以 <o:OLEObject> 出现，按出现顺序替换为占位符
    * 预览图通过同段落内 <v:imagedata> 的 rId 关联到 media 文件
    * MTEF 变体由 Equation Native 流前几字节判定：
        A = 起始 0x03，可解出字符（但会丢结构，只能做交叉校验）
        B = 起始 05 00 00 06 07 "DSMT6"，不可解，必须走 EMF→内嵌PDF 或看图
    * 故意不解码 MTEF 文本：变体 A 的解码结果丢结构严重（实测 151 字符的
      公式只解出 2 个字符），把它当"答案"会产出错误文档。字符解码仅作
      交叉校验用，由 --mtef 开关单独触发。
"""
import io
import json
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

from docxfile import BadDocument, fail, force_utf8_stdio, open_docx

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"

EQ_OPEN = "\u27e6EQ:"      # ⟦EQ:
EQ_CLOSE = "\u27e7"        # ⟧


def natkey(name):
    """按文件名里的数字自然排序（oleObject2 在 oleObject10 之前）。"""
    return int(re.sub(r"\D", "", name) or 0)


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
    for stale in ("paras.json", "eq_index.json", "summary.json",
                  "build_report.json", "headings_suggested.json"):
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
    paras = []
    for p in body.iter(W + "p"):
        buf = []
        for node in p.iter():
            if node.tag == W + "t":
                buf.append(node.text or "")
            elif node.tag == W + "tab":
                buf.append("\t")
            elif node.tag == W + "br":
                buf.append("\n")
            elif node.tag == O + "OLEObject":
                rid = node.get(R + "id")
                tgt = (rel_map.get(rid) or "?").split("/")[-1]
                buf.append("%s%s%s" % (EQ_OPEN, tgt, EQ_CLOSE))
        paras.append("".join(buf))

    with open(os.path.join(outdir, "paras.json"), "w", encoding="utf-8") as f:
        json.dump(paras, f, ensure_ascii=False, indent=1)

    # ---- 2. 公式 → 预览图映射（逐段配对，保持出现顺序）----
    # rel_map 的 Target 形如 "media/image49.wmf"（相对 word/），
    # 这里分别保存完整相对路径（用于取 zip 条目）和文件名（用于输出）。
    pairs = {}
    for p in body.iter(W + "p"):
        imgs = [i.get(R + "id") for i in p.iter(V + "imagedata")]
        for k, obj in enumerate(p.iter(O + "OLEObject")):
            oid = (rel_map.get(obj.get(R + "id")) or "?").split("/")[-1]
            if k < len(imgs):
                pairs[oid] = rel_map.get(imgs[k]) or None      # "media/image49.wmf"
            else:
                pairs[oid] = None

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

    # ---- 4. 公式索引 ----
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
                "img": pairs.get(base),
                "bytes": 0,
                "variant": "ERROR",
                "error": str(exc),
            })
            continue
        rows.append({
            "stream": base,
            "img": os.path.basename(pairs[base]) if pairs.get(base) else None,
            "bytes": len(eq),
            "variant": mtef_variant(eq),
        })

    with open(os.path.join(outdir, "eq_index.json"), "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)

    # ---- 5. 摘要 ----
    n_empty = sum(1 for x in paras if not x.strip())
    counts = {}
    for r in rows:
        counts[r["variant"]] = counts.get(r["variant"], 0) + 1
    no_media = [r["stream"] for r in rows if not r.get("img")]

    summary = {
        "paragraphs": len(paras),
        "paragraphs_nonempty": len(paras) - n_empty,
        "formulas": len(rows),
        "variants": counts,
        "media_extracted": len(extracted),
        "formulas_without_media": len(no_media),
        "next": (
            "全部公式都有预览图，可逐张读图转写 LaTeX。"
            if not no_media else
            "有 %d 个公式没有预览图（%s），只能靠 MTEF 或原始 Word 补齐。"
            % (len(no_media), ", ".join(no_media[:5]))
        ),
    }
    with open(os.path.join(outdir, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=1)

    print(json.dumps(summary, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    force_utf8_stdio()
    main()
