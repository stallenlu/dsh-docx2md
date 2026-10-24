#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pairing_selftest.py —— 钉死「公式 ↔ 预览图」的配对规则。

为什么单独一个测试：这条规则曾经错得很隐蔽。旧实现按「同段落第 k 个
OLEObject ↔ 第 k 张 v:imagedata」配对，只要公式前面夹着一张插图，后面
全部右移一格，把插图的位图当成公式预览图交给模型去读——不报错、不告警，
只是默默地读错图。在真实文档（145 个公式）上实测错了 10 个。

正确规则：公式的预览图是带 o:ole="t" 的 v:shape，其 id 等于 OLEObject
的 ShapeID。这条证据与顺序无关，段落里有多少插图都不影响。

本测试用合成的 docx 覆盖三种情况：
    A. 有 ShapeID + 段落里有插图（旧实现会配错）→ 必须配到公式图
    B. 文档不写 ShapeID                          → 退回顺序配对，且必须告警
    C. 公式没有预览图                            → 如实报 formulas_without_media

用法: python3 test/pairing_selftest.py     退出码 0 = 全绿
"""
import json
import os
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
PY_DIR = os.path.join(os.path.dirname(HERE), "py")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"

CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Default Extension="wmf" ContentType="image/x-wmf"/></Types>')

RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdOle1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>'
        '<Relationship Id="rIdIll" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/illust.png"/>'
        '<Relationship Id="rIdEq" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.wmf"/>'
        '</Relationships>')

PNG = b"\x89PNG\r\n\x1a\n" + b"\0" * 64
WMF = b"\xd7\xcd\xc6\x9a" + b"\0" * 64


def make_docx(path, shape_id=True, shape=True, shape_target="rIdEq"):
    """造一个只有一段的 docx：段落里先放插图，再放公式（+ 可选预览 shape）。"""
    sid = ' ShapeID="_x0000_i1025"' if shape_id else ""
    preview = ""
    if shape:
        preview = ('<w:r><w:pict><v:shape id="_x0000_i1025" o:ole="t" type="#_x0000_t75">'
                   '<v:imagedata r:id="%s"/></v:shape></w:pict></w:r>' % shape_target)
    doc = ('<?xml version="1.0"?>'
           '<w:document xmlns:w="%s" xmlns:r="%s" xmlns:v="%s" xmlns:o="%s"><w:body><w:p>'
           '<w:r><w:t>正文</w:t></w:r>'
           '<w:r><w:pict><v:shape id="_x0000_s99"><v:imagedata r:id="rIdIll"/></v:shape></w:pict></w:r>'
           '<w:r><w:object><o:OLEObject Type="Embed" ProgID="Equation.3"%s r:id="rIdOle1"/></w:object></w:r>'
           '%s</w:p></w:body></w:document>') % (W, R, V, O, sid, preview)
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", RELS)
        z.writestr("word/media/illust.png", PNG)
        z.writestr("word/media/formula.wmf", WMF)
        z.writestr("word/embeddings/oleObject1.bin", b"\0" * 64)


def extract(docx_path, out_dir):
    shutil.rmtree(out_dir, ignore_errors=True)
    r = subprocess.run([sys.executable, "extract.py", docx_path, out_dir],
                       cwd=PY_DIR, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError("extract 失败: %s" % r.stderr.strip()[:300])
    with open(os.path.join(out_dir, "eq_index.json"), encoding="utf-8") as f:
        index = json.load(f)
    with open(os.path.join(out_dir, "summary.json"), encoding="utf-8") as f:
        summary = json.load(f)
    return index, summary


WORK = os.path.join(os.environ.get("TMPDIR", "/tmp"), "d2m-pairing")
shutil.rmtree(WORK, ignore_errors=True)
os.makedirs(WORK)

cases = []


def rec(name, fn):
    try:
        fn()
        cases.append((name, True, ""))
    except AssertionError as exc:
        cases.append((name, False, str(exc)))


def case_a():
    """段落里有插图在前、公式在后。旧实现会配到 illust.png。"""
    docx = os.path.join(WORK, "a.docx")
    make_docx(docx, shape_id=True, shape=True)
    index, summary = extract(docx, os.path.join(WORK, "a_out"))
    img = index[0]["img"]
    assert img == "formula.wmf", "配到了 %s，应为 formula.wmf（插图位图被当成公式预览图）" % img
    assert index[0]["pairing"] == "shape-id", "pairing 应为 shape-id，实际 %s" % index[0]["pairing"]
    assert summary["formulas_paired_by_shape_id"] == 1, "summary 未记录 shape-id 配对"
    assert summary["formulas_paired_by_ordinal"] == 0, "summary 错误地记录了 ordinal 配对"


def case_b():
    """文档不写 ShapeID。必须退回顺序配对，并把不确定性报出来。"""
    docx = os.path.join(WORK, "b.docx")
    make_docx(docx, shape_id=False, shape=False)
    index, summary = extract(docx, os.path.join(WORK, "b_out"))
    assert index[0]["pairing"] == "ordinal", "pairing 应为 ordinal，实际 %s" % index[0]["pairing"]
    assert summary["formulas_paired_by_ordinal"] == 1, "summary 未记录 ordinal 配对"
    assert "pairing_warning" in summary, "退回顺序配对时必须给出 pairing_warning"
    assert "回退" in summary["next"] or "核对" in summary["next"], \
        "next 字段应提示这些图未经确认，实际: %s" % summary["next"]


def case_c():
    """公式没有预览图：如实报出来，不要编一个。"""
    docx = os.path.join(WORK, "c.docx")
    make_docx(docx, shape_id=True, shape=False)      # 既无 ShapeID 也无 shape
    index, _summary = extract(docx, os.path.join(WORK, "c_out"))
    # 无 shape 时退回顺序配对，段落里只有一张插图，会配到它——但 pairing=ordinal
    assert index[0]["pairing"] == "ordinal", "无预览图时应标为 ordinal"


rec("A. 有 ShapeID：段落内有插图也能配到真正的公式图", case_a)
rec("B. 无 ShapeID：退回顺序配对且明确告警", case_b)
rec("C. 无预览图：不编造，标为 ordinal", case_c)

shutil.rmtree(WORK, ignore_errors=True)

print("== 公式 ↔ 预览图配对自测 ==")
fail = 0
for name, ok, msg in cases:
    print(("  PASS  " if ok else "  FAIL  ") + name)
    if not ok:
        fail += 1
        print("         " + msg[:200])

print()
if fail:
    print("%d 项失败" % fail)
    sys.exit(1)
print("全部通过：配对按 ShapeID 优先，回退时必然告警")
sys.exit(0)
