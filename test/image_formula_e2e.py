#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""image_formula_e2e.py —— 图片形式公式的端到端回归。

钉死四件事（图片公式是第三种载体：像素没有结构可解，工具的唯一职责是
**绝不静默丢** + **如实报告**）：

  1. inspect 报 formulas_image 与候选清单（ratio/上下文），不报结论。
  2. extract 给公式候选生成 ⟦EQ:img:文件名⟧ 占位、原件抽进 media/；
     **示意图落成 ⟦EQ:fig:文件名⟧ 占位**（正文里必须有它——旧版只记
     清单不进正文，转出的 md 里插图全部消失，是本插件最忌讳的静默丢失）。
  3. build 未转写时落成**可见**的图片引用（不是纯注释——注释渲染后
     看不见，等于公式换了种方式消失）；latex.json 给 "img:文件名" 键
     后替换成 LaTeX。
  4. validate：图片引用计入 T8 总数；断链报 T11；未转写不报 T9。

用法: python3 test/image_formula_e2e.py     退出码 0 = 全绿
"""
import json
import os
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
PY_DIR = os.path.join(os.path.dirname(HERE), "py")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"


def png(w, h, color):
    import zlib

    def chunk(t, d):
        c = t + d
        return struct.pack(">I", len(d)) + c + struct.pack(">I", zlib.crc32(c))
    raw = b"".join(b"\x00" + bytes(color) * w for _ in range(h))
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def drawing(rid, name, cx, cy, did):
    return ('<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="%s"/>'
            '<a:graphic><a:graphicData uri="%s">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
            % (cx, cy, did, name, P_NS, did, name, rid, cx, cy))


P_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')
ROOT_RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
             '</Relationships>')
WRELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rIdEq" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/eq_formula.png"/>'
         '<Relationship Id="rIdFig" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/figure.png"/>'
         '</Relationships>')

EMU_W = 1905000     # 200px
EMU_H = 317500      # ~33px  → ratio 6.0 → formula
FIG_W = 2857500     # 300px
FIG_H = 2857500     # ratio 1.0 → illustration

DOC = ('<?xml version="1.0"?>'
       '<w:document xmlns:w="%s" xmlns:r="%s" xmlns:wp="%s" xmlns:a="%s" xmlns:pic="%s" '
       'xmlns:v="%s" xmlns:o="%s"><w:body>'
       '<w:p><w:r><w:t>由定义</w:t></w:r>%s<w:r><w:t>可得。</w:t></w:r></w:p>'
       '<w:p>%s</w:p>'
       '</w:body></w:document>'
       % (W, R, WP, A, PIC, V, O,
          drawing("rIdEq", "公式截图", EMU_W, EMU_H, 101),
          drawing("rIdFig", "示意图", FIG_W, FIG_H, 201)))


def run(script, *args):
    r = subprocess.run([sys.executable, script, *args],
                       cwd=PY_DIR, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError("%s 失败(%d): %s" % (script, r.returncode,
                                                (r.stderr or r.stdout)[:300]))
    return r.stdout


work = tempfile.mkdtemp(prefix="d2m-imgf-")
fails = []
try:
    docx = os.path.join(work, "imgf.docx")
    with zipfile.ZipFile(docx, "w") as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", DOC)
        z.writestr("word/_rels/document.xml.rels", WRELS)
        z.writestr("word/media/eq_formula.png", png(240, 40, (10, 10, 10)))
        z.writestr("word/media/figure.png", png(300, 300, (80, 140, 80)))

    out = os.path.join(work, "out")

    # 1. inspect
    insp = json.loads(run("docx_scan.py", docx))
    if insp.get("formulas_image") != 1:
        fails.append("inspect: formulas_image=%s，应为 1" % insp.get("formulas_image"))
    cands = insp.get("image_formula_candidates") or []
    if not cands or cands[0]["file"] != "eq_formula.png":
        fails.append("inspect: 候选清单不对: %s" % cands)

    # 2. extract：占位 + 原件 + 示意图不进正文
    run("extract.py", docx, out)
    paras = json.load(open(os.path.join(out, "paras.json"), encoding="utf-8"))
    if "⟦EQ:img:eq_formula.png⟧" not in paras[0]:
        fails.append("extract: 公式图未落成占位: %r" % paras[0])
    # 示意图**必须**进正文：旧版在这里 continue（只记清单、正文里一个
    # 占位都不留），结果 145 公式文档转出的 md 里六张插图全部消失，连
    # 注释都没有——那是本工具最忌讳的静默丢失。现在落成 ⟦EQ:fig:…⟧，
    # build 再换成可见的图片引用。
    if not any("⟦EQ:fig:figure.png⟧" in p for p in paras):
        fails.append("extract: 示意图没有落成正文占位: %r" % (paras,))
    if any("figure.png" in p and "⟦EQ:img:figure.png⟧" in p for p in paras):
        fails.append("extract: 示意图被当成了图片公式")
    if not os.path.exists(os.path.join(out, "media", "eq_formula.png")):
        fails.append("extract: 公式图原件没抽进 media/")
    inv = json.load(open(os.path.join(out, "image_inventory.json"), encoding="utf-8"))
    shapes = {i["file"]: i["shape"] for i in inv}
    if shapes.get("eq_formula.png") != "formula" or shapes.get("figure.png") != "illustration":
        fails.append("extract: 清单形状判定错: %s" % shapes)
    # 清单里列出的每个文件，原件必须在 media/ 里。清单写着「figure.png 在
    # 段落 2」而 media/ 里没有这个文件，人去核对时无从下手——「留痕供人
    # 核对」就是一句空话。示意图不进正文是另一回事，它的原件要到位。
    for e in inv:
        if not os.path.exists(os.path.join(out, "media", e["file"])):
            fails.append("extract: 清单里的 %s 原件没抽进 media/" % e["file"])

    # 3a. build 未转写：可见的图片引用（不是纯注释）
    empty = os.path.join(work, "empty.json")
    json.dump({}, open(empty, "w"))
    md_a = os.path.join(work, "a.md")
    b = json.loads(run("build.py", out, "--latex", empty, "--out", md_a,
                       "--use-suggested"))
    text_a = open(md_a, encoding="utf-8").read()
    if "![图片公式 eq_formula.png](" not in text_a:
        fails.append("build: 未转写的图片公式没有落成可见引用: %r" % text_a[:120])
    if b.get("image_placeholders") != 1:
        fails.append("build: image_placeholders=%s，应为 1" % b.get("image_placeholders"))
    if b.get("unresolved_placeholders") != 0:
        fails.append("build: 图片公式不该计入 unresolved: %s" % b.get("unresolved_placeholders"))
    # 示意图必须落成**可见**的图片引用（不是注释）。
    if "![图](media/figure.png)" not in text_a and "![图](" not in text_a:
        fails.append("build: 示意图没有落成可见引用: %r" % text_a[:200])
    if b.get("figures") != 1:
        fails.append("build: figures=%s，应为 1" % b.get("figures"))

    # 3b. build 转写后：LaTeX 替换，计数迁移
    latex = os.path.join(work, "latex.json")
    json.dump({"img:eq_formula.png": "\\sigma_{xy}"}, open(latex, "w", encoding="utf-8"))
    md_b = os.path.join(work, "b.md")
    b2 = json.loads(run("build.py", out, "--latex", latex, "--out", md_b,
                        "--use-suggested"))
    text_b = open(md_b, encoding="utf-8").read()
    if "$\\sigma_{xy}$" not in text_b or "图片公式" in text_b:
        fails.append("build: 转写后仍留图片引用或未替换: %r" % text_b[:120])
    if b2.get("image_transcribed") != 1 or b2.get("image_placeholders") != 0:
        fails.append("build: 转写计数错: %s/%s" % (b2.get("image_transcribed"),
                                                   b2.get("image_placeholders")))

    # 4. validate：T8 计入、T11 断链
    v = json.loads(run("validate.py", md_a, "--expect-formulas", "1", "--json"))
    if v.get("status") != "PASS" or v.get("image_formulas") != 1:
        fails.append("validate(a): %s image_formulas=%s" % (v.get("status"),
                                                            v.get("image_formulas")))
    v2 = json.loads(run("validate.py", md_b, "--expect-formulas", "1", "--json"))
    if v2.get("status") != "PASS":
        fails.append("validate(b): %s errors=%s" % (v2.get("status"), v2.get("errors")[:2]))
    broken = text_a.replace("/media/eq_formula.png)", "/media/nope.png)")
    bp = os.path.join(work, "broken.md")
    open(bp, "w", encoding="utf-8").write(broken)
    # validate 判 FAIL 时退出码是 1（这是它的契约），run() 会把它当崩溃；
    # 这里直接跑并读 stdout。
    rv = subprocess.run([sys.executable, "validate.py", bp, "--json"],
                        cwd=PY_DIR, capture_output=True, text=True)
    v3 = json.loads(rv.stdout)
    if not any("T11" in e for e in v3.get("errors", [])):
        fails.append("validate: 断链没有报 T11: %s" % v3.get("errors")[:2])

    print("== 图片公式端到端 ==")
    if fails:
        for f in fails:
            print("  FAIL " + f)
        print("\n%d 项失败" % len(fails))
        sys.exit(1)
    print("  ok   inspect 报告 / extract 占位+原件+清单 / build 可见引用与转写替换 / T8+T11")
    print("\n全部通过：图片公式绝不静默丢——不转写就落成可见引用，转写就换 LaTeX")
finally:
    shutil.rmtree(work, ignore_errors=True)
