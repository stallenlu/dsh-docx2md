#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""parafmt_e2e.py —— 三项排版信号的端到端回归：居中 / 颜色 / 首行缩进。

为什么要有这个测试：它们和 docstruct_e2e.py 盯的「标题识别」是同一份
结构信号（docstruct.json），但那一版只验了 style/marker/bold。真实文档
里被丢掉的三样东西恰好都在剩下的字段上，而且**丢得静悄悄**——正文一字
不差，人眼只有对着原文才会发现「这行本来是居中的、本来是蓝的、本来要
空两格」。用户报的三类问题里有两类就是这么没的。

不还原的难点在于 Markdown 根本没有这三种语法，只能借 HTML：

  * 居中 → `<p align="center">`（Markdown 无居中语法）
  * 颜色 → `<span style="color:#RRGGBB">`
  * 首行缩进 → 行首空白会被 Markdown 解析器整段吃掉，只能借全角空格
    （U+3000）顶出实打实的两格

本测试钉死四件事：

  1. extract 读出了 jc / color / firstLine / firstLineChars，且是**段落
     级**信号（十段里只有第 2 段居中，不能串到第 3 段）。
  2. build 如实落成 HTML + 全角空格。
  3. 首行缩进的折算：`firstLineChars=50`（Word 的「50/100 字符」口径）是
     **半个字**，属于编号悬挂而不是正文缩进。真实文档里 "i) 位错密度与其他
     物理场的关系" 就带着 firstLineChars=50，照抄成两个全角空格会让一个
     带编号的标题看着像错位了的正文开头。必须落成 0。
  4. 「标题段不该被补缩进」：」「注释，在真实文档里它在 iii) 编号下，
     加了缩进就会让读者以为它是正文。

用法: python3 test/parafmt_e2e.py     退出码 0 = 全绿
"""
import json
import re
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PY_DIR = os.path.join(ROOT, "py")

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

STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="%s">' % W
          + '<w:style w:type="paragraph" w:styleId="Heading2">'
          '<w:name w:val="heading 2"/>'
          '<w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>'
          + '</w:styles>')

NUMBERING = ('<?xml version="1.0"?><w:numbering xmlns:w="%s">' % W
             + '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
             '<w:start w:val="1"/><w:numFmt w:val="lowerRoman"/>'
             '<w:lvlText w:val="%%1)"/></w:lvl></w:abstractNum>'
             '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
             + '</w:numbering>')


def para(text, jc=None, color=None, indent=None, indent_chars=None,
         style=None, numid=None, bold=False):
    """造一个带指定排版信号的段落。"""
    ppr_inner = ""
    if style:
        ppr_inner += '<w:pStyle w:val="%s"/>' % style
    if numid:
        ppr_inner += ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/>'
                      '</w:numPr>' % numid)
    if jc:
        ppr_inner += '<w:jc w:val="%s"/>' % jc
    if indent is not None or indent_chars is not None:
        attrs = ""
        if indent is not None:
            attrs += ' w:firstLine="%d"' % indent
        if indent_chars is not None:
            attrs += ' w:firstLineChars="%d"' % indent_chars
        ppr_inner += '<w:ind%s/>' % attrs
    # rPr 里的 jc/color 与 pPr 里的同义；真实文档两种都出现，这里走 pPr。
    ppr = "<w:pPr>%s</w:pPr>" % ppr_inner if ppr_inner else ""
    rpr = "<w:rPr><w:b/>" + ('<w:color w:val="%s"/>' % color if color else "")
    rpr += "</w:rPr>" if (bold or color) else ""
    if not (bold or color):
        rpr = ""
    return '<w:p>%s<w:r>%s<w:t>%s</w:t></w:r></w:p>' % (ppr, rpr, text)


# 段落顺序刻意安排过： authentic 文档里「一般位错密度演化方程」这种居中+蓝
# 色+加粗的段落，后面紧跟的是它标注的那条公式（见基于矢量位错密度的位错
# 动力学及其数理方案_docx2md/paras.json 的 188/189）。合成 docx 里造不出真
# 公式（OLE 需要二进制流），所以把它放在**最后一段**——_next_is_prose 对
# 「后面没有段落」和「后面是公式」都返回 False，走同一条分支，不会掉进
# 「整段加粗短标签」的标题兜底。
#
# 同理，「位错密度与其他物理场的关系」在真机里带自动编号（i)），这里去掉
# numid 只留 firstLineChars=50：带上编号它会被认成四级标题，于是「有没有
# 被加缩进」这条断言就变得永远为真，测不出东西。
DOC = ('<?xml version="1.0"?><w:document xmlns:w="%s"><w:body>' % W
       + para("普通正文，什么信号都没有。")
       # 蓝色但不居中（真机里的「标量Green函数g:」属于这类）
       + para("标量格林函数 g：", color="0000FF", bold=True)
       # 居中但不蓝
       + para("输出应力、应变", jc="center")
       # 首行缩进 420 twip = 2 字 → 落成两个全角空格
       + para("辐照金属材料的强度变化是武器库存课题关心的关键问题。", indent=420)
       # firstLineChars=50 → 半字，编号悬挂，不该给正文缩进
       + para("位错密度与其他物理场的关系", indent_chars=50)
       # 普通黑色正文不居中 → 什么都不该加
       + para("这一段是普通正文，末尾也没有任何信号。")
       # 样式标题（有 outlineLvl）带着悬挂缩进：标题不该被补缩进
       + para("位错动力学总论", style="Heading2", indent_chars=50)
       # 居中 + 蓝色 + 加粗 → 小节标签，三样都要还原
       + para("一般位错密度演化方程", jc="center", color="0000FF", bold=True)
       + '</w:body></w:document>')


def write_docx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", DOC)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING)


def run(script, *args):
    r = subprocess.run([sys.executable, script, *args],
                       cwd=PY_DIR, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError("%s 失败(%d): %s\n%s"
                           % (script, r.returncode, r.stdout, r.stderr))
    return r.stdout


def main():
    fails = []
    tmp = tempfile.mkdtemp(prefix="parafmt_")
    docx = os.path.join(tmp, "sample.docx")
    write_docx(docx)
    work = os.path.join(tmp, "work")
    run("extract.py", docx, work)

    paras = json.load(open(os.path.join(work, "paras.json"), encoding="utf-8"))
    struct = json.load(open(os.path.join(work, "docstruct.json"), encoding="utf-8"))
    if len(struct) != len(paras):
        fails.append("docstruct.json 长度 %d 与 paras.json %d 不一致"
                     % (len(struct), len(paras)))

    # ---- 1. extract 读出的信号是段落级的 ----
    def sig(i):
        return struct[i] if i < len(struct) else {}

    cases = [
        (0, "jc", None),
        (1, "color", "0000FF"),
        (1, "jc", None),
        (2, "jc", "center"),
        (2, "color", None),
        (3, "firstLine", 420),
        (4, "firstLineChars", 50),
        (5, "jc", None),
        (5, "color", None),
        (5, "firstLine", None),
        (7, "jc", "center"),
        (7, "color", "0000FF"),
    ]
    for i, key, want in cases:
        got = sig(i).get(key)
        if got != want:
            fails.append("段落 %d 的 %s: 期望 %r 实得 %r" % (i, key, want, got))

    # ---- 2. build 落成 HTML + 全角空格 ----
    blank = os.path.join(tmp, "latex.json")
    with open(blank, "w", encoding="utf-8") as f:
        f.write("{}")
    md = os.path.join(tmp, "out.md")
    run("build.py", work, "--latex", blank, "--out", md, "--use-suggested")
    lines = [l.rstrip("\n") for l in open(md, encoding="utf-8")]
    body = "\n".join(lines)

    # 居中+蓝+加粗 = Word 里的小节标签 → 标题路径，但标题语法 `#` 没有
    # 居中，所以落成 `<hN align="center">`（N 由上下文决定，不断言具体级）。
    # 颜色用 span 保住；** 绝不允许残留——Markdown 强调在 HTML 标签里不
    # 被解析，留着会让读者看到两个星号（真实文档里就是这样被发现的）。
    lab = [l for l in lines if "一般位错密度演化方程" in l]
    if not lab or not re.match(r'^<h\d align="center">'
                               r'<span style="color:#0000FF">'
                               r'一般位错密度演化方程</span></h\d>$', lab[0]):
        fails.append("居中+蓝色+加粗的标签没有落成 <hN align=center> + "
                     "<span color>；实际 %r" % lab)
    if "**一般位错密度演化方程**" in body:
        fails.append("HTML 里残留 ** 加粗标记（渲染器不解析，会原样显示）")
    # 蓝色不居中
    bl = [l for l in lines if "标量格林函数" in l]
    if not bl or 'style="color:#0000FF"' not in bl[0]:
        fails.append("非居中的蓝色段落没保留颜色；实际 %r" % bl)
    elif bl[0].startswith("<p align="):
        fails.append("非居中段落被套了居中标签；实际 %r" % bl[0])
    # 居中不蓝
    cl = [l for l in lines if "输出应力" in l]
    if not cl or '<p align="center">' not in cl[0]:
        fails.append("居中段落没落成 <p align=center>；实际 %r" % cl)
    elif 'color:' in cl[0]:
        fails.append("非蓝段落被套了颜色标签；实际 %r" % cl[0])
    # 首行缩进
    ind = [l for l in lines if "辐照金属材料的强度变化" in l]
    if not ind or not ind[0].startswith("\u3000\u3000"):
        fails.append("firstLine=420 的段落没有落成两个全角空格（U+3000）；"
                     "实际行首 %r" % (ind[0][:6] if ind else None))
    # firstLineChars=50 → 编号悬挂，不该给正文缩进
    hang = [l for l in lines if "位错密度与其他物理场的关系" in l]
    if hang and hang[0].lstrip().startswith("\u3000"):
        fails.append("firstLineChars=50（编号悬挂）被错误地加了首行缩进；"
                     "实际 %r" % hang[0][:20])
    # 普通正文不该有任何包装
    plain = [l for l in lines if "这一段是普通正文" in l]
    if not plain or ("<p" in plain[0] or "color:" in plain[0]
                     or plain[0].startswith("\u3000")):
        fails.append("普通正文段落被误加了排版包装；实际 %r" % (plain[0][:40] if plain else None))

    # ---- 3. 缩进折算：直接在单元层喂 record，覆盖各种边界 ----
    sys.path.insert(0, PY_DIR)
    import build as buildmod
    fake = [{"style": None, "marker": "", "bold": False,
             "jc": None, "color": None, "firstLine": None,
             "firstLineChars": None}]
    recs = [
        # firstLineChars=200 → 2 字
        ({"firstLineChars": 200}, 2),
        # firstLineChars=50 → 半字 → 0
        ({"firstLineChars": 50}, 0),
        # firstLine=420 twip → 2 字
        ({"firstLine": 420}, 2),
        # firstLine=105 twip → 0 字（编号悬挂）
        ({"firstLine": 105}, 0),
        # 上限 4：排版事故不该照抄
        ({"firstLine": 5000}, 4),
        # 空 signals
        ({}, 0),
        # —— 字宽换算：同样 640twip，字宽 320（16pt）是 2 字、210 是 3 字。
        # 写死 210 的实现在 16pt 文档上会把 2 字缩进算成 3 字。
        ({"firstLine": 640, "charTwip": 320}, 2),
        ({"firstLine": 640, "charTwip": 210}, 3),
        # charTwip 缺失（旧产物）退回 210
        ({"firstLine": 640}, 3),
    ]
    for rec, want in recs:
        full = dict(fake[0])
        full.update(rec)
        got = buildmod._para_indent_chars([full], 0)
        if got != want:
            fails.append("_para_indent_chars(%r): 期望 %d 实得 %d"
                         % (rec, want, got))

    # ---- 3b. 端到端：默认字号 16pt 的文档（charTwip=320）----
    # 上一节的单元测能算对，但 charTwip 是 extract 写进 docstruct 的；
    # 真正要验的是「extract 从 styles.xml 读出 16pt → build 用它换算」。
    # 640twip 缩进在 16pt 文档里是 2 字；若 build 端仍按 210 硬算会得到
    # 3 个字（渲染上缩进明显过深）。这是本工具通用性的关键一条。
    doc16 = ('<?xml version="1.0"?><w:document xmlns:w="%s"><w:body>' % W
             + para("十六磅文档里的正文缩进应当是两个全角空格。", indent=640)
             + '</w:body></w:document>')
    styles16 = ('<?xml version="1.0"?><w:styles xmlns:w="%s">' % W
                + '<w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:sz w:val="32"/></w:rPr></w:rPrDefault></w:docDefaults>'
                + '</w:styles>')
    tmp2 = os.path.join(tmp, "w16")
    os.makedirs(tmp2)
    docx2 = os.path.join(tmp2, "s16.docx")
    with zipfile.ZipFile(docx2, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc16)
        z.writestr("word/styles.xml", styles16)
        z.writestr("word/numbering.xml", NUMBERING)
    work2 = os.path.join(tmp2, "work")
    run("extract.py", docx2, work2)
    st2 = json.load(open(os.path.join(work2, "docstruct.json"),
                         encoding="utf-8"))
    if st2 and st2[0].get("charTwip") != 320:
        fails.append("extract 应从 docDefaults 的 w:sz=32(16pt) 推出字宽 "
                     "320twip，实得 %r" % (st2[0].get("charTwip") if st2 else None,))
    blank2 = os.path.join(tmp2, "latex.json")
    with open(blank2, "w", encoding="utf-8") as f:
        f.write("{}")
    md2 = os.path.join(tmp2, "out.md")
    run("build.py", work2, "--latex", blank2, "--out", md2, "--use-suggested")
    body2 = open(md2, encoding="utf-8").read()
    line2 = [l for l in body2.split("\n") if "十六磅文档" in l]
    if not line2 or line2[0].count("\u3000") != 2:
        fails.append("16pt 文档里 firstLine=640 应落成 2 个全角空格（640/320），"
                     "实得 %r" % (line2[0][:12] if line2 else None))

    # ---- 3c. 居中+加粗规则的防误伤（合成语料探针实测出的两条）----
    # (a) 图/表题注在英文论文里就是「居中+加粗」，但它们不是标题，不能
    #     升进目录；(b) 一组相邻居的居中标签必须**同一层级**，否则一过
    #     三条就一路掉到 h6（_deepest_before+1 会逐条累加）。
    cap_paras = ["Figure 1. Ring shrinkage", "Table 1. Constants",
                 "图 2 位错环", "标签甲", "", "标签乙"]
    cap_struct = []
    for t in cap_paras:
        bold = t[:1] in ("标",) or t.startswith(("Figure", "Table", "图 "))
        cap_struct.append({"style": None, "marker": "", "bold": bold,
                           "jc": "center" if bold else None, "color": None,
                           "firstLine": None, "firstLineChars": None,
                           "charTwip": 210})
    got_h = buildmod.guess_headings(cap_paras, cap_struct, True)
    for i in (0, 1, 2):
        if i in got_h:
            fails.append("题注 %r 被误认成标题（题注是正文，不进目录）"
                         % cap_paras[i])
    if got_h.get(3) is not None and got_h.get(5) is not None \
            and got_h[3] != got_h[5]:
        fails.append("相邻两个居中标签层级不一致：%r（应同级）" % got_h)

    if fails:
        print("FAIL (%d)" % len(fails))
        for f in fails:
            print("  -", f)
        return 1
    print("parafmt_e2e: OK（三项排版信号全部还原）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
