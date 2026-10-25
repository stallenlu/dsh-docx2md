#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""canvas_e2e.py —— 流程图（VML 画布）必须整块还原的端到端测试。

背景（真实事故）：
    文档里"程序流程图"那一章，正文里只剩一串互不相干的公式，图没了。
    原因是流程图在 docx 里不是一张图，而是一个 v:group 画布，里面并排放着
    几十个形状（框、菱形、箭头、标签）+ 若干 OLE 公式。旧实现只认
    v:imagedata 和 OLE，画布的骨架一律不看，于是框、箭头、标签全消失，
    画布里的公式散落成正文——不报错、不告警，直到读者发现"图不见了"。

这个测试锁住三件事：
    1. 画布被识别出来（不是普通插图组）；
    2. 画布整块渲染成一张 PNG，并在正文里落成图片引用；
    3. 画布内的公式**不再**单独占位（否则同一批公式出现两遍）。
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PY = os.path.join(ROOT, "py")
sys.path.insert(0, PY)

import canvas  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

FAILS = []


def check(cond, msg):
    if cond:
        print("  ok   %s" % msg)
    else:
        print("  FAIL %s" % msg)
        FAILS.append(msg)


# ------------------------------------------------------------------ 合成文档
# 手写一个最小 docx：一个画布，里面有 2 个圆角矩形（步骤框）+ 1 个菱形
# （判断）+ 1 个箭头（连线）+ 1 个文本框（标签）+ 1 个 OLE 公式。
# 不依赖 python-docx，避免测试自己引入依赖。
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:o="urn:schemas-microsoft-com:office:office" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def shape(id_, spt, x, y, w, h, text=None, fill="#00CC99"):
    box = ""
    if text:
        box = ('<v:textbox><w:txbxContent><w:p><w:r><w:t>%s</w:t></w:r>'
               '</w:p></w:txbxContent></v:textbox>' % text)
    return ('<v:shape id="%s" o:spt="%s" fillcolor="%s" stroked="t" '
            'strokecolor="#000000" style="position:absolute;left:%d;'
            'top:%d;width:%d;height:%d;mso-position-horizontal-relative:'
            'page;mso-position-vertical-relative:page">%s</v:shape>'
            % (id_, spt, fill, x, y, w, h, box))


CANVAS_XML = (
    '<w:p><w:pPr/><w:r><w:pict>'
    '<v:group id="grp1" coordsize="2000,1000" coordorigin="0,0" '
    'style="position:absolute;left:0;top:0;width:400pt;height:200pt">'
    + shape("s1", "176", 100, 100, 500, 200, "开始")
    + shape("s2", "13", 700, 80, 400, 240)                 # 菱形
    + shape("s3", "176", 1200, 100, 500, 200, "计算应力")
    + '<v:line id="l1" from="600,200" to="700,200" '
      'style="position:absolute;left:600;top:200;width:100;height:0" '
      'strokecolor="#FF0000" strokeweight="2pt"/>'
    + '<v:shape id="t1" o:spt="202" style="position:absolute;left:100;'
      'top:600;width:1500;height:200"><v:textbox><w:txbxContent><w:p><w:r>'
      '<w:t>VDDD3D程序结构图</w:t></w:r></w:p></w:txbxContent></v:textbox>'
      '</v:shape>'
    + '</v:group>'
    # OLE 是 group 的**兄弟节点**，ShapeID 指向画布内的 v:shape。
    # 这是 Word 的真实结构：公式的宿主对象在外层，预览 shape 在画布内。
    '<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="s3" '
    'r:id="rIdEq1"/>'
    '</w:pict></w:r></w:p>'
)

DOC_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document %s><w:body>'
    '<w:p><w:r><w:t>正文第一段。</w:t></w:r></w:p>'
    % NS + CANVAS_XML +
    '<w:p><w:r><w:t>正文第二段。</w:t></w:r></w:p>'
    '<w:sectPr/></w:body></w:document>'
)

RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
    'relationships">'
    '<Relationship Id="rIdEq1" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/oleObject" '
    'Target="embeddings/oleObject1.bin"/>'
    '</Relationships>'
)

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
    'package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="bin" ContentType="application/vnd.openxmlformats-'
    'officedocument.oleObject"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.'
    'openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>'
)

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
    'relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    '</Relationships>'
)


def build_docx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", DOC_XML)
        z.writestr("word/_rels/document.xml.rels", RELS_XML)
        # 一个假的 OLE 容器：extract 读不出 MTEF 会走异常分支，
        # 那正好也验证了"异常分支仍要标 in_canvas"。
        z.writestr("word/embeddings/oleObject1.bin", b"\x00" * 64)


def main():
    tmp = tempfile.mkdtemp(prefix="canvas_e2e_")
    doc = os.path.join(tmp, "canvas.docx")
    out = os.path.join(tmp, "out")
    try:
        build_docx(doc)

        # ---- 1. 画布识别 ----
        print("== A. 画布识别 ==")
        with zipfile.ZipFile(doc) as z:
            root = ET.fromstring(z.read("word/document.xml"))
        body = root.find(W + "body")
        groups = canvas.find_canvases(body)
        check(len(groups) == 1, "识别出 1 块画布（实际 %d）" % len(groups))
        if groups:
            structural, media, has_text = canvas._content_shapes(groups[0])
            check(structural >= 3, "画布里有结构形状（框/菱形/箭头）")
            check(has_text, "画布里有文字标签")

        # ---- 2. SVG 生成 ----
        print("== B. SVG 生成 ==")
        if groups:
            svg, w, h = canvas.canvas_svg(groups[0], {}, tmp)
            check(bool(svg), "能生成 SVG")
            check("<polygon" in (svg or ""), "菱形被画成 polygon")
            check("<rect" in (svg or ""), "步骤框被画成 rect")
            check("<line" in (svg or ""), "箭头/连线被画成 line")
            check("VDDD3D程序结构图" in (svg or ""), "文本框标签进了 SVG")
            check(w > 0 and h > 0, "画布有物理尺寸（%.1f×%.1f pt）" % (w, h))

        # ---- 3. 端到端：extract 把画布落成图 ----
        print("== C. extract 端到端 ==")
        env = dict(os.environ)
        env["PYTHONPATH"] = PY
        r = subprocess.run([sys.executable, os.path.join(PY, "extract.py"),
                            doc, out], capture_output=True, text=True,
                           env=env, timeout=300)
        if r.returncode != 0:
            check(False, "extract 失败: %s" % (r.stderr or r.stdout)[-400:])
        else:
            summary = json.load(open(os.path.join(out, "summary.json"),
                                     encoding="utf-8"))
            check(summary.get("canvases") == 1,
                  "summary 报出 1 块画布（实际 %s）" % summary.get("canvases"))
            check(summary.get("canvases_rendered") == 1,
                  "画布渲染成功")
            paras = json.load(open(os.path.join(out, "paras.json"),
                                   encoding="utf-8"))
            joined = "\n".join(paras)
            check(re.search(r"⟦EQ:fig:canvas\d+\.png⟧", joined) is not None,
                  "正文里有画布图片占位")
            # 画布内的公式不能再单独占位
            check("⟦EQ:oleObject1.bin⟧" not in joined,
                  "画布内的公式**没有**单独占位（否则同一批公式出现两遍）")
            rows = json.load(open(os.path.join(out, "eq_index.json"),
                                  encoding="utf-8"))
            eq = [x for x in rows if x["stream"] == "oleObject1.bin"]
            check(bool(eq) and eq[0].get("in_canvas") is True,
                  "eq_index 里标了 in_canvas（读 MTEF 失败也要标）")
            check(os.path.isfile(os.path.join(out, "media", "canvas1.png")),
                  "media/canvas1.png 确实写出来了")
            # 不留临时文件
            leftovers = [n for n in os.listdir(os.path.join(out, "media"))
                         if n.startswith("_canvas_tmp")]
            check(not leftovers, "没有留下临时 SVG")

        # ---- 4. build 落成图片引用 ----
        print("== D. build 端到端 ==")
        if r.returncode == 0:
            latex = os.path.join(out, "latex.json")
            with open(latex, "w", encoding="utf-8") as f:
                json.dump({}, f)
            md = os.path.join(out, "out.md")
            r2 = subprocess.run([sys.executable, os.path.join(PY, "build.py"),
                                 out, "--latex", latex, "--out", md,
                                 "--use-suggested"],
                                capture_output=True, text=True, env=env,
                                timeout=300)
            if r2.returncode != 0:
                check(False, "build 失败: %s" % (r2.stderr or r2.stdout)[-400:])
            else:
                text = open(md, encoding="utf-8").read()
                check("canvas1.png" in text, "md 里引用了画布图")
                check("![图](media/canvas1.png)" in text,
                      "画布落成标准 Markdown 图片引用")
                rep = json.load(open(os.path.join(out, "build_report.json"),
                                     encoding="utf-8"))
                check(rep.get("canvases") == 1,
                      "build_report 里 canvases=1（实际 %s）"
                      % rep.get("canvases"))
                check(rep.get("formulas_canvas") == 1,
                      "build_report 里 formulas_canvas=1")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print()
    if FAILS:
        print("FAILED  %d 项未通过" % len(FAILS))
        return 1
    print("画布端到端全部通过：流程图整块还原，画布内公式不重复占位")
    return 0


if __name__ == "__main__":
    sys.exit(main())
