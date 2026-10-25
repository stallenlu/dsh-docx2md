#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""omml_e2e.py —— Word 原生公式（OMML）的端到端回归。

钉死三件事（都是这次支持 OMML 时最容易回退的点）：

  1. inspect 报 formulas_omml，纯 OMML 文档不再显示 formulas: 0 像「没公式」。
  2. extract 给 OMML 生成占位符并把 LaTeX 确定性转进 eq_index.json——
     正文里公式**绝不静默消失**（旧版会：段落文字照常、公式无影无踪）。
  3. build 用**空的** latex.json 也能产出完整 md（预填生效），validate PASS。

用法: python3 test/omml_e2e.py     退出码 0 = 全绿
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
PY_DIR = os.path.join(os.path.dirname(HERE), "py")

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CT = ('<?xml version="1.0"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" '
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>')

DOC = ('<?xml version="1.0"?>'
       '<w:document xmlns:w="%s" xmlns:m="%s"><w:body>'
       '<w:p><w:r><w:t>设比值</w:t></w:r>'
       '<m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num>'
       '<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>'
       '<w:r><w:t>，且平方项为</w:t></w:r></w:p>'
       '<w:p><m:oMathPara><m:oMath><m:sSup>'
       '<m:e><m:r><m:t>x</m:t></m:r></m:e>'
       '<m:sup><m:r><m:t>2</m:t></m:r></m:sup>'
       '</m:sSup></m:oMath></m:oMathPara></w:p>'
       '<w:p><w:r><w:t>其中</w:t></w:r>'
       '<m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>'
       '<w:r><w:t>为角度。</w:t></w:r></w:p>'
       '</w:body></w:document>') % (W, M)


def run(script, *args):
    r = subprocess.run([sys.executable, script, *args],
                       cwd=PY_DIR, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError("%s 失败(%d): %s" % (script, r.returncode,
                                                (r.stderr or r.stdout)[:300]))
    return r.stdout


work = tempfile.mkdtemp(prefix="d2m-omml-e2e-")
fails = []
try:
    docx = os.path.join(work, "omml.docx")
    with zipfile.ZipFile(docx, "w") as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", DOC)

    out = os.path.join(work, "out")

    # 1. inspect
    insp = json.loads(run("docx_scan.py", docx))
    if insp.get("formulas_omml") != 3:
        fails.append("inspect: formulas_omml=%s，应为 3" % insp.get("formulas_omml"))
    if insp.get("formulas_total") != 3:
        fails.append("inspect: formulas_total=%s，应为 3" % insp.get("formulas_total"))

    # 2. extract
    run("extract.py", docx, out)
    paras = json.load(open(os.path.join(out, "paras.json"), encoding="utf-8"))
    n_ph = sum(p.count("⟦EQ:omml") for p in paras)
    if n_ph != 3:
        fails.append("extract: 正文里 omml 占位符 %d 个，应为 3（公式静默丢失）" % n_ph)
    idx = json.load(open(os.path.join(out, "eq_index.json"), encoding="utf-8"))
    got = {r["stream"]: r.get("latex") for r in idx if r.get("kind") == "omml"}
    want = {"omml1": "\\frac{a}{b}", "omml2": "x^{2}", "omml3": "\\alpha"}
    for k, v in want.items():
        if (got.get(k) or "").strip() != v:
            fails.append("extract: %s latex=%r，应为 %r" % (k, got.get(k), v))

    # 3. build（空 latex 表：全靠预填）
    empty = os.path.join(work, "empty.json")
    json.dump({}, open(empty, "w"))
    md_path = os.path.join(work, "out.md")
    b = json.loads(run("build.py", out, "--latex", empty, "--out", md_path,
                       "--use-suggested"))
    md = open(md_path, encoding="utf-8").read()
    if b.get("omml_prefilled") != 3:
        fails.append("build: omml_prefilled=%s，应为 3" % b.get("omml_prefilled"))
    if b.get("formulas_missing") != 0:
        fails.append("build: formulas_missing=%s，应为 0" % b.get("formulas_missing"))
    for frag in ("$\\frac{a}{b}$", "$$", "x^{2}", "\\alpha"):
        if frag not in md:
            fails.append("build: 产物缺 %r" % frag)
    if "设比值 $\\frac{a}{b}$" not in md and "设比值 $\\frac{a}{b}$ " not in md.replace("\n", " "):
        # 中文边界规则：紧贴中文的公式前后应有空格
        import re
        if not re.search(r"设比值 \$\\frac\{a\}\{b\}\$ ", md):
            fails.append("build: 行内公式与中文之间没有边界: %r" % md[:80])

    # 4. validate
    v = json.loads(run("validate.py", md_path, "--expect-formulas", "3", "--json"))
    if v.get("status") != "PASS":
        fails.append("validate: %s errors=%s" % (v.get("status"), v.get("errors")[:2]))

    print("== OMML 端到端 ==")
    print("  inspect formulas_omml=%s / extract 占位符=%d / build 预填=%s / validate=%s"
          % (insp.get("formulas_omml"), n_ph, b.get("omml_prefilled"), v.get("status")))
    for f in fails:
        print("  FAIL " + f)
    if fails:
        print("\n%d 项失败" % len(fails))
        sys.exit(1)
    print("\n全部通过：Word 原生公式从 inspect 到 validate 全链路可用")
finally:
    shutil.rmtree(work, ignore_errors=True)
