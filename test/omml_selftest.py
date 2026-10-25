#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""OMML→LaTeX 转换器的质量测试：每题给出 OMML 输入与期望 LaTeX 输出。
期望值按 LaTeX 语义写死；比对时忽略空白差异。
"""
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "py"))
from xml.etree import ElementTree as ET
import omml

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

def omml_xml(inner):
    return ('<m:oMath xmlns:m="%s" '
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">%s</m:oMath>'
            % (M, inner))

def r(t):
    return '<m:r><m:t>%s</m:t></m:r>' % t

def mt(t):
    return '<m:t>%s</m:t>' % t

def sty_r(t, sty):
    return '<m:r><m:rPr><m:sty m:val="%s"/></m:rPr><m:t>%s</m:t></m:r>' % (sty, t)

CASES = [
    ("分式", '<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>' % (r("a"), r("b")),
     r"\frac{a}{b}"),
    ("linear 分式",
     '<m:f><m:fPr><m:type m:val="lin"/></m:fPr><m:num>%s</m:num><m:den>%s</m:den></m:f>' % (r("a"), r("b")),
     "a/b"),
    ("linear 分式带和",
     '<m:f><m:fPr><m:type m:val="lin"/></m:fPr><m:num>%s</m:num><m:den>%s</m:den></m:f>'
     % (r("a+b"), r("c")),
     "(a+b)/c"),
    ("上标", '<m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup>' % (r("x"), r("2")),
     "x^{2}"),
    ("下标", '<m:sSub><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>' % (r("u"), r("i")),
     "u_{i}"),
    ("上下标",
     '<m:sSubSup><m:e>%s</m:e><m:sub>%s</m:sub><m:sup>%s</m:sup></m:sSubSup>' % (r("σ"), r("x"), r("P")),
     "\\sigma_{x}^{P}"),
    ("根号", '<m:rad><m:e>%s</m:e></m:rad>' % r("a+b"),
     "\\sqrt{a+b}"),
    ("n 次根",
     '<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg>%s</m:deg><m:e>%s</m:e></m:rad>' % (r("3"), r("x")),
     "\\sqrt[3]{x}"),
    ("定界符",
     '<m:d><m:e>%s</m:e></m:d>' % r("a+b"),
     "\\left( a+b \\right)"),
    ("定制界符",
     '<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>%s</m:e></m:d>' % r("x"),
     "\\left[ x \\right]"),
    ("求和带界",
     '<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr>'
     '<m:sub>%s</m:sub><m:sup>%s</m:sup><m:e>%s</m:e></m:nary>' % (r("i=1"), r("n"), r("a")),
     "\\sum_{i=1}^{n}a"),
    ("函数",
     '<m:func><m:funcName>%s</m:funcName><m:e>%s</m:e></m:func>' % (r("sin"), r("x")),
     "\\sin x"),
    ("重音 hat",
     '<m:acc><m:accPr><m:chr m:val="ˆ"/></m:accPr><m:e>%s</m:e></m:acc>' % r("k"),
     "\\hat{k}"),
    ("重音 vec",
     '<m:acc><m:accPr><m:chr m:val="⃗"/></m:accPr><m:e>%s</m:e></m:acc>' % r("b"),
     "\\vec{b}"),
    ("横线", '<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>%s</m:e></m:bar>' % r("x"),
     "\\overline{x}"),
    ("矩阵",
     '<m:m><m:mr><m:e>%s</m:e><m:e>%s</m:e></m:mr><m:mr><m:e>%s</m:e><m:e>%s</m:e></m:mr></m:m>'
     % (r("a"), r("b"), r("c"), r("d")),
     "\\begin{matrix} a & b \\\\ c & d \\end{matrix}"),
    ("方程组",
     '<m:eqArr><m:e>%s</m:e><m:e>%s</m:e></m:eqArr>'
     % ('<m:r><m:t>x=1</m:t></m:r>', '<m:r><m:t>y=2</m:t></m:r>'),
     "\\begin{gathered} x=1 \\\\ y=2 \\end{gathered}"),
    ("underset",
     '<m:limLow><m:e>%s</m:e><m:lim>%s</m:lim></m:limLow>' % (r("max"), r("x")),
     "\\underset{x}{max}"),
    ("纯文本样式",
     '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>初始化</m:t></m:r>',
     "\\text{初始化}"),
    ("希腊字母", r("α+β"), "\\alpha +\\beta "),
    ("符号映射", r("∇×β=0"), "\\nabla \\times \\beta =0"),
    ("嵌套", '<m:f><m:num><m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup></m:num><m:den>%s</m:den></m:f>'
     % (r("σ"), r("2"), r("2")),
     "\\frac{\\sigma^{2}}{2}"),
]

def norm(s):
    return " ".join(s.split())

fails = 0
for name, inner, want in CASES:
    el = ET.fromstring(omml_xml(inner))
    got = omml.to_latex(el)
    ok = norm(got) == norm(want)
    print(("  PASS  " if ok else "  FAIL  ") + name)
    if not ok:
        fails += 1
        print("        want: %r" % want)
        print("        got : %r" % got)

print()
print("%d/%d 通过" % (len(CASES) - fails, len(CASES)))
sys.exit(1 if fails else 0)
