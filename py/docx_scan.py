#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx_scan.py —— 只读扫描 .docx，判断可转换性与需要走哪条路。

用法:
    python docx_scan.py <input.docx>

输出（stdout，单个 JSON），主要字段:
    paragraphs / formulas / embeddings
    formulas            OLE 公式数（与 extract 的 formulas 同口径）
    formulas_omml       Word 原生公式（OMML）数——extract 阶段确定性转出
                        LaTeX，不需要读图；不计入 formulas
    formulas_total      两族之和（expectFormulas 用它）
    variants            公式的 MTEF 变体分布（A 可解字符 / B 必须看图）
    formulas_with_preview, formulas_without_preview
    media_formats       预览图的实际格式分布（wmf/emf/png/jpeg）
    missing_preview_context
                        缺图公式的所在段落与段落文字。只报告事实，
                        不作「这是图注还是正文公式」的判断。
    render_engines      本机可用的渲染引擎

本脚本只输出事实，不输出判断（如「该走哪条路径」「这是图注还是正文」）。
判断需要文档上下文，由调用方来做。

不做任何写操作，不渲染，不依赖 pypdf。
"""
import io
import json
import os
import platform
import re
import shutil
import sys
import zipfile
from xml.etree import ElementTree as ET

from docxfile import BadDocument, fail, force_utf8_stdio, open_docx

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def natkey(name):
    return int(re.sub(r"\D", "", name) or 0)


def mtef_variant(eq):
    if len(eq) < 33:
        return "EMPTY"
    body = eq[28:]
    if body[:5] == b"\x05\x00\x00\x06\x07":
        return "B"
    if body[:1] == b"\x03":
        return "A"
    return "?"


def eq_stream(raw_bin):
    import olefile
    ole = olefile.OleFileIO(io.BytesIO(raw_bin))
    try:
        for s in ole.listdir():
            if s[-1] == "Equation Native":
                return ole.openstream("/".join(s)).read()
        return b""
    finally:
        ole.close()


def engines():
    """可用渲染引擎探测。"""
    def have(names):
        for n in names:
            p = shutil.which(n)
            if p:
                return p
        return None
    out = {
        "platform": platform.system(),
        "powershell": have(["powershell.exe", "powershell", "pwsh"]),
        "libreoffice": have(["soffice", "libreoffice"]),
    }
    try:
        import PIL  # noqa: F401
        out["pillow"] = True
    except ImportError:
        out["pillow"] = False
    out["can_render"] = bool(out["powershell"] or out["libreoffice"] or out["pillow"])
    return out


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    src = sys.argv[1]

    try:
        z, root = open_docx(src)
    except BadDocument as exc:
        fail(str(exc))
    names = z.namelist()

    rels_path = "word/_rels/document.xml.rels"
    rel_map = {}
    if rels_path in names:
        rels = z.read(rels_path).decode("utf-8", "replace")
        rel_map = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))

    body = root.find(W + "body")

    paras = list(body.iter(W + "p"))
    n_para = len(paras)

    # 公式 → 预览图
    pairs = {}
    n_eq = 0
    for p in paras:
        imgs = [i.get(R + "id") for i in p.iter(V + "imagedata")]
        objs = list(p.iter(O + "OLEObject"))
        n_eq += len(objs)
        for k, obj in enumerate(objs):
            oid = (rel_map.get(obj.get(R + "id")) or "?").split("/")[-1]
            if k < len(imgs):
                pairs[oid] = rel_map.get(imgs[k]) or None
            else:
                pairs[oid] = None

    # Word 原生公式（OMML）：新版 Word 的默认格式，结构化 XML，
    # extract 阶段可确定性转成 LaTeX，不需要读图。inspect 必须报出来，
    # 否则纯 OMML 文档会显示 formulas: 0 —— 用户以为「没有公式」，
    # 实际是「有一族本工具此前不认识」。
    n_omml = sum(1 for _ in body.iter(M + "oMath"))

    # 变体统计
    variants = {}
    conv_errors = []
    embeds = sorted([n for n in names
                     if n.startswith("word/embeddings/") and n.endswith(".bin")],
                    key=natkey)
    try:
        for nm in embeds:
            base = nm.split("/")[-1]
            try:
                v = mtef_variant(eq_stream(z.read(nm)))
            except Exception as exc:
                v = "ERROR"
                conv_errors.append("%s: %s" % (base, str(exc)[:80]))
            variants[v] = variants.get(v, 0) + 1
    except ImportError:
        variants = {"UNKNOWN": len(embeds)}

    # 预览图格式
    media_formats = {}
    for tgt in pairs.values():
        if tgt:
            ext = os.path.splitext(tgt)[1].lstrip(".").lower()
            media_formats[ext] = media_formats.get(ext, 0) + 1

    with_preview = sum(1 for t in pairs.values() if t)
    without = len(pairs) - with_preview

    # 缺图的公式：只报告事实——它出现在哪个段落、那段文字是什么、
    # 同段落里还有几个公式、前后段落是什么。不写任何「这意味着什么」。
    # 工具看不到图，也不知道作者的排版意图，判断应交给调用方。
    missing_context = []
    if without:
        para_texts = ["".join(t.text or "" for t in p_.iter(W + "t")).strip()
                      for p_ in paras]
        para_objs = [[(rel_map.get(o.get(R + "id")) or "?").split("/")[-1]
                      for o in p_.iter(O + "OLEObject")] for p_ in paras]

        def near(idx, step):
            """取前后第 step 个非空段落，作为上下文。"""
            j = idx + step
            while 0 <= j < len(para_texts):
                if para_texts[j] or para_objs[j]:
                    return {"paragraph": j, "text": para_texts[j][:120]}
                j += step
            return None

        for idx, objs in enumerate(para_objs):
            no_img = [o for o in objs if not pairs.get(o)]
            if not no_img:
                continue
            missing_context.append({
                "paragraph": idx,
                "objects": no_img,
                "formulas_in_paragraph": len(objs),
                "paragraph_text": para_texts[idx][:200],
                "prev": near(idx, -1),
                "next": near(idx, +1),
            })

    eng = engines()

    result = {
        "paragraphs": n_para,
        # formulas 保持「OLE 公式数」的语义不变（现有测试与调用方依赖它）；
        # OMML 单独报，避免把两个族混进一个数字里掩盖差异。
        "formulas": n_eq,
        "formulas_omml": n_omml,
        "formulas_total": n_eq + n_omml,
        "embeddings": len(embeds),
        "variants": variants,
        "formulas_with_preview": with_preview,
        "formulas_without_preview": without,
        "media_formats": media_formats,
        "has_ole_objects": n_eq > 0,
        "has_omml": n_omml > 0,
        "render_engines": eng,
    }
    if n_omml:
        result["omml_note"] = (
            "有 %d 个 Word 原生公式（OMML）。它们是结构化 XML，extract 阶段"
            "会确定性转成 LaTeX（eq_index.json 的 latex 字段），无需读图。"
            % n_omml)
    if n_eq == 0 and n_omml == 0:
        result["no_formulas_note"] = (
            "两族公式（OLE / OMML）都没有检测到。若文档肉眼可见有公式，"
            "可能是图片式公式或第三方编辑器私有格式，本工具不覆盖。")
    if without:
        result["missing_preview_count"] = without
        result["missing_preview_context"] = missing_context
    if conv_errors:
        result["embedding_errors"] = conv_errors[:10]

    # OLEObject 引用数与 embeddings 文件数本应一致；不一致说明文档结构异常，
    # 如实报出来，而不是用 max() 悄悄抹平。
    if n_eq != len(embeds):
        result["ole_object_count_mismatch"] = {
            "ole_objects_in_document": n_eq,
            "embedding_files": len(embeds),
            "note": "两者不一致，说明有 OLE 对象未落到 embeddings/，或有未被引用的 embeddings 文件。",
        }
    print(json.dumps(result, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    force_utf8_stdio()
    main()
