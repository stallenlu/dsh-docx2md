#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docstruct_e2e.py —— 标题识别的结构信号回归。

为什么要有这个测试：真实的 145 公式文档里，标题识别曾经只认出 12 条，
而人工整理过的参考稿有 32 条。漏掉的 20 条全是「文字里没有编号」的标题：

  - Word 自动编号（`<w:numPr>`）：文档里写的是 "位错密度的演化方程"，
    屏幕上显示的却是 "iii) 位错密度的演化方程"——编号由 numbering.xml
    现算，段落文字里一个字符都没有。
  - 样式标题（`pStyle=heading N` / `outlineLvl`）：文字跟正文一模一样。
  - 只有加粗、连编号都没有的小标题（"攀移力" / "线张力"）。

纯正则对这三类全都无能为力。本测试钉死三件事：

  1. extract 落出 docstruct.json，且与 paras.json 同长同序。
  2. 自动编号被正确合成（`lowerRoman` + start=3 → "iii)"），并且
     自成一组的第二份列表从 "i)" 重新起算——不做全局累加。
  3. build 把它们认成标题；样式标题的层级来自 outlineLvl，而不是猜。
  4. 整段加粗的「公式行标签」不能被误认成标题（"位移场" / "应力场"
     这类后面紧跟公式的加粗短句），而「章节标签」要认得（后面跟正文）。

用法: python3 test/docstruct_e2e.py     退出码 0 = 全绿
"""
import json
import os
import shutil
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

# 两份列表：numId=1 用 lowerRoman 且 start=3（首项 iii），
# numId=2 同样 lowerRoman 但 start=1（首项 i）。真实文档就是这种形态——
# 第二组必须从 i) 重新起算，若按 numId 全局累加会变成 "vii)"。
STYLES = ('<?xml version="1.0"?>'
          '<w:styles xmlns:w="%s">'
          '<w:style w:type="paragraph" w:styleId="Heading2">'
          '<w:name w:val="heading 2"/>'
          '<w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Heading3">'
          '<w:name w:val="heading 3"/>'
          '<w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>'
          '</w:styles>') % W

NUMBERING = ('<?xml version="1.0"?>'
             '<w:numbering xmlns:w="%s">'
             '<w:abstractNum w:abstractNumId="0">'
             '<w:lvl w:ilvl="0"><w:start w:val="3"/>'
             '<w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%%1)"/></w:lvl>'
             '</w:abstractNum>'
             '<w:abstractNum w:abstractNumId="1">'
             '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
             '<w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%%1)"/></w:lvl>'
             '</w:abstractNum>'
             '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
             '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
             '</w:numbering>') % W


def para(text, style=None, numid=None, bold=False, emph=False):
    """造一个段落。bold=整段加粗；emph=只有行内强调（不加粗整段）。"""
    ppr = ""
    if style or numid:
        ppr = "<w:pPr>"
        if style:
            ppr += '<w:pStyle w:val="%s"/>' % style
        if numid:
            ppr += ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>'
                    % numid)
        ppr += "</w:pPr>"
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    if emph:
        return ('<w:p>%s<w:r><w:rPr><w:b/></w:rPr><w:t>%s</w:t></w:r>'
                '<w:r><w:t>是正文的普通字。</w:t></w:r></w:p>' % (ppr, text))
    return '<w:p>%s<w:r>%s<w:t>%s</w:t></w:r></w:p>' % (ppr, rpr, text)


DOC = ('<?xml version="1.0"?>'
       '<w:document xmlns:w="%s"><w:body>'
       # 样式标题：文字跟正文一模一样，只有 outlineLvl 能认出来
       + para("位错动力学总论", style="Heading2")
       + para("这一段是普通正文，用来隔开标题。")
       # 自动编号段落 1：numId=1 start=3 → iii)
       + para("位错密度的演化方程", numid=1)
       + para("先说明这一节的背景。")
       # 自动编号段落 2：同一列表，累加到 iv)
       + para("位错间的短程作用", numid=1)
       + para("短程作用的物理图像。")
       # 第二份列表：numId=2 start=1 → i)，必须重新起算而不是接成 v)
       + para("数值求解方法", numid=2)
       + para("离散化之后迭代求解。")
       # 加粗章节标签，后面跟正文 → 应认成标题
       + para("攀移力", bold=True)
       + para("采用虚功原理推导攀移力。")
       # 加粗短句，后面紧跟正文 → 章节标签，认标题
       + para("线张力", bold=True)
       + para("先计算短程自相互作用能。")
       # 行内强调（不是整段加粗）→ 绝不应认成标题
       + para("规范不变性", emph=True)
       + '</w:body></w:document>') % W


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
    tmp = tempfile.mkdtemp(prefix="docstruct_")
    try:
        docx = os.path.join(tmp, "sample.docx")
        write_docx(docx)
        work = os.path.join(tmp, "work")
        run("extract.py", docx, work)

        paras = json.load(open(os.path.join(work, "paras.json"), encoding="utf-8"))
        sp = os.path.join(work, "docstruct.json")
        if not os.path.exists(sp):
            fails.append("extract 未生成 docstruct.json")
            struct = []
        else:
            struct = json.load(open(sp, encoding="utf-8"))
            if len(struct) != len(paras):
                fails.append("docstruct.json 长度 %d 与 paras.json %d 不一致"
                             % (len(struct), len(paras)))

        # ---- 1. 自动编号合成 ----
        if len(struct) == len(paras):
            got = [(paras[i].strip(), struct[i]["marker"])
                   for i in range(len(paras))]
            want = [("位错动力学总论", ""),
                    ("这一段是普通正文，用来隔开标题。", ""),
                    ("位错密度的演化方程", "iii)"),
                    ("先说明这一节的背景。", ""),
                    ("位错间的短程作用", "iv)"),
                    ("短程作用的物理图像。", ""),
                    ("数值求解方法", "i)"),
                    ("离散化之后迭代求解。", ""),
                    ("攀移力", ""),
                    ("采用虚功原理推导攀移力。", ""),
                    ("线张力", ""),
                    ("先计算短程自相互作用能。", ""),
                    ("规范不变性是正文的普通字。", "")]
            for i, (wt, wm) in enumerate(want):
                if i >= len(got):
                    break
                if got[i][0] != wt:
                    fails.append("段落 %d 文字不对: 期望 %r 实得 %r"
                                 % (i, wt, got[i][0]))
                elif got[i][1] != wm:
                    fails.append("段落 %d (%r) 的自动编号: 期望 %r 实得 %r"
                                 % (i, wt, wm, got[i][1]))

        # ---- 2. 样式级别 ----
        if len(struct) == len(paras):
            if struct[0]["style"] != 2:
                fails.append("样式标题的级别应由 outlineLvl 决定："
                             "Heading2(outlineLvl=1) 期望 2，实得 %r"
                             % struct[0]["style"])
            if struct[8]["style"] is not None:
                fails.append("普通段落不该有样式级别，实得 %r"
                             % struct[8]["style"])

        # ---- 3. build 认标题 ----
        # 这个合成文档里没有公式，空表就够——本测试只关心标题。
        blank = os.path.join(tmp, "latex.json")
        with open(blank, "w", encoding="utf-8") as f:
            f.write("{}")
        md = os.path.join(tmp, "out.md")
        run("build.py", work, "--latex", blank, "--out", md, "--use-suggested")
        lines = [l.rstrip("\n") for l in open(md, encoding="utf-8")]
        heads = [l for l in lines if l.startswith("#")]

        def has(text):
            return any(text in h for h in heads)

        if not has("位错动力学总论"):
            fails.append("build 没把样式标题认成标题（只靠 outlineLvl 才能认出）")
        elif not any(h.startswith("## 位错动力学总论") for h in heads):
            fails.append("样式标题的层级不对，期望 '## '，实得 %r"
                         % [h for h in heads if "位错动力学总论" in h])
        if not has("iii) 位错密度的演化方程"):
            fails.append("build 没把自动编号段落认成标题，或编号没补回去；"
                         "实际标题行 %r" % heads)
        if not has("iv) 位错间的短程作用"):
            fails.append("同一列表的第二项编号没累加到 iv)；实际标题行 %r" % heads)
        if not has("i) 数值求解方法"):
            fails.append("第二份列表没有从 i) 重新起算（按 numId 全局累加会成 v)）；"
                         "实际标题行 %r" % heads)
        if not has("攀移力"):
            fails.append("加粗章节标签（后面跟正文）没被认成标题；"
                         "实际标题行 %r" % heads)
        if not has("线张力"):
            fails.append("第二个加粗章节标签（线张力）没被认成标题")
        if has("规范不变性"):
            fails.append("行内强调（非整段加粗）被误认成标题了")

        # ---- 3b. 加粗的「公式行标签」不能认成标题 ----
        # 这一条只能在单元层验：它取决于**下一段是不是公式**，而合成 docx
        # 里造不出一条真公式（OLE 需要二进制流）。直接喂 paras 给
        # guess_headings，把公式占位写成 extract 会写的样子。
        sys.path.insert(0, PY_DIR)
        import build as buildmod
        fake_paras = ["位移场",
                      "\u27e6EQ:oleObject59.bin\u27e7",
                      "应力场",
                      "\u27e6EQ:oleObject61.bin\u27e7或者",
                      "攀移力",
                      "采用虚功原理推导。"]
        fake_struct = [{"style": None, "marker": "", "bold": True},
                       {"style": None, "marker": "", "bold": False},
                       {"style": None, "marker": "", "bold": True},
                       {"style": None, "marker": "", "bold": False},
                       {"style": None, "marker": "", "bold": True},
                       {"style": None, "marker": "", "bold": False}]
        got_h = buildmod.guess_headings(fake_paras, fake_struct, True)
        if 0 in got_h:
            fails.append("加粗的公式行标签「位移场」（下一段是公式）被误认成标题")
        if 2 in got_h:
            fails.append("加粗短句「应力场」（下一段是公式+残句）被误认成标题")
        if 4 not in got_h:
            fails.append("加粗短句「攀移力」（下一段是正文）没被认成标题")

        # ---- 4. --no-bold-headings 关掉兜底 ----
        md2 = os.path.join(tmp, "out2.md")
        run("build.py", work, "--latex", blank, "--out", md2,
            "--use-suggested", "--no-bold-headings")
        heads2 = [l.rstrip("\n") for l in open(md2, encoding="utf-8")
                  if l.startswith("#")]
        if any("攀移力" in h for h in heads2):
            fails.append("--no-bold-headings 没关掉加粗兜底")
        if not any("iii) 位错密度的演化方程" in h for h in heads2):
            fails.append("--no-bold-headings 关掉了不该关的结构标题（自动编号）")

        # ---- 5. 结构文件对不上时必须整体退回纯正则，不能错位挂标题 ----
        struct2 = list(struct)
        if struct2:
            struct2 = struct2[:-1]        # 故意弄短：模拟口径漂移
            json.dump(struct2, open(sp, "w", encoding="utf-8"),
                      ensure_ascii=False)
            md3 = os.path.join(tmp, "out3.md")
            r = subprocess.run([sys.executable, "build.py", work, "--latex",
                                blank, "--out", md3, "--use-suggested"],
                               cwd=PY_DIR, capture_output=True, text=True)
            if r.returncode != 0:
                fails.append("docstruct 长度不符时 build 崩了（应静默退回纯正则）: %s"
                             % r.stderr.strip())
            else:
                heads3 = [l for l in open(md3, encoding="utf-8")
                          if l.startswith("#")]
                if any("位错动力学总论" in h for h in heads3):
                    fails.append("docstruct 长度不符时仍采用了结构信号"
                                 "（会挂到错误段落上，比不识别更糟）")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    if fails:
        for f in fails:
            sys.stderr.write("FAIL  %s\n" % f)
        print("FAIL  %d 项不合格" % len(fails))
        return 1
    print("PASS  标题结构信号（样式/自动编号/加粗）全部符合")
    return 0


if __name__ == "__main__":
    sys.exit(main())
