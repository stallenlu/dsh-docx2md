#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""corpus_probe.py —— 用一批**形态各异**的合成 docx 探通用性。

目的不是断言"必须怎样"，而是把每种排版惯例跑一遍，看新加的三条规则
（居中→标题、颜色→span、firstLine→全角空格）和 T13/T14 在**不是**
"中文位错文档"这种形态上会不会误伤。输出是观察报告，不是 pass/fail。
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
PY_DIR = os.path.join(os.path.dirname(HERE), "py")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" '
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>')


def styles(sz_halfpt=None, extra=""):
    dd = ('<w:docDefaults><w:rPrDefault><w:rPr>%s</w:rPr></w:rPrDefault></w:docDefaults>'
          % ('<w:sz w:val="%d"/>' % sz_halfpt if sz_halfpt else ""))
    return ('<?xml version="1.0"?><w:styles xmlns:w="%s">%s%s</w:styles>') % (W, dd, extra)


def p(text, jc=None, color=None, first_line=None, first_line_chars=None,
      highlight=None):
    inner = ""
    if jc:
        inner += '<w:jc w:val="%s"/>' % jc
    if first_line or first_line_chars:
        a = ""
        if first_line:
            a += ' w:firstLine="%d"' % first_line
        if first_line_chars:
            a += ' w:firstLineChars="%d"' % first_line_chars
        inner += '<w:ind%s/>' % a
    rpr = ""
    if color:
        rpr += '<w:color w:val="%s"/>' % color
    if highlight:
        rpr += '<w:highlight w:val="%s"/>' % highlight
    if text.startswith("<b>"):
        rpr += "<w:b/>"
        text = text[3:]
    rpr = "<w:rPr>%s</w:rPr>" % rpr if rpr else ""
    ppr = "<w:pPr>%s</w:pPr>" % inner if inner else ""
    return '<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ppr, rpr, text)


def doc(body):
    return ('<?xml version="1.0"?><w:document xmlns:w="%s"><w:body>%s</w:body></w:document>') % (W, body)


CASES = {
    # 英文论文：居中加粗是**图/表题注**，不是标题；缩进是物理 0.5 英寸
    "english_paper": (
        styles(sz_halfpt=22),
        doc(p("<b>DISLOCATION DYNAMICS: A REVIEW") + p("")
            + p("Abstract", jc="center")
            + p("This paper studies dislocations and their motion under stress. "
                "The model is validated against experiments.", first_line=720)
            + p("<b>Figure 1. Ring shrinkage", jc="center")
            + p("<b>Table 1. Material constants", jc="center")
            + p("The elastic constants are listed below.", first_line=720))),
    # 中文公文/学位论文：标题居中不加粗、正文首行缩进两字
    "cn_thesis": (
        styles(sz_halfpt=32),   # 16pt
        doc(p("某某研究报告", jc="center")
            + p("一、研究背景", first_line=640)
            + p("本报告围绕某项研究展开讨论，首先说明研究背景与意义，然后给出方法。",
                first_line=640)
            + p("二、研究方法", first_line=640))),
    # 无 styles.xml（最裸的 docx）
    "no_styles": (None, doc(p("只有一段话的文档。"))),
    # 右对齐 + 高亮 + 纯强调居中（居中加粗但其实是强调句）
    "mixed_emphasis": (
        styles(sz_halfpt=24),
        doc(p("正文开始。")
            + p("<b>注意：这句话只是强调，不是标题", jc="center")
            + p("右对齐的落款", jc="right")
            + p("高亮词", highlight="yellow"))),
    # 全角缩进用 firstLineChars（Word「按字符」口径）
    "char_units": (
        styles(),
        doc(p("按字符缩进两段。", first_line_chars=200)
            + p("按字符缩进三段。", first_line_chars=300))),
}


def run(script, *a):
    r = subprocess.run([sys.executable, script, *a], cwd=PY_DIR,
                       capture_output=True, text=True)
    return r


def main():
    tmp = tempfile.mkdtemp(prefix="corpus_")
    for name, (st, body) in CASES.items():
        d = os.path.join(tmp, name)
        os.makedirs(d)
        path = os.path.join(d, "in.docx")
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/document.xml", body)
            if st:
                z.writestr("word/styles.xml", st)
        work = os.path.join(d, "work")
        r = run("extract.py", path, work)
        if r.returncode != 0:
            print("== %-16s EXTRACT FAILED rc=%d %s" % (name, r.returncode,
                                                        r.stderr[:200]))
            continue
        open(os.path.join(d, "latex.json"), "w").write("{}")
        out = os.path.join(d, "out.md")
        r2 = run("build.py", work, "--latex", os.path.join(d, "latex.json"),
                 "--out", out, "--use-suggested")
        if r2.returncode != 0:
            print("== %-16s BUILD FAILED %s" % (name, r2.stderr[:300]))
            continue
        md = open(out, encoding="utf-8").read()
        st_json = json.load(open(os.path.join(work, "docstruct.json"),
                                 encoding="utf-8")) if os.path.exists(
            os.path.join(work, "docstruct.json")) else []
        print("== %s" % name)
        print("   charTwip:", sorted({x.get("charTwip") for x in st_json}) or "n/a")
        for line in md.split("\n"):
            if line.strip():
                print("   | " + line[:78])
        print()


if __name__ == "__main__":
    main()
