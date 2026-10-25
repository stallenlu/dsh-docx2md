#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx 通用结构 e2e：表格 / 列表 / 行内格式 / 超链接 / 脚注 / 合并单元格。

为什么要这个测试：插件曾经假设「正文 = 一串 w:p」。这个假设在**只有段落**
的文档上成立，所以全部测试都绿；一旦文档里有表格，`body.iter(w:p)` 会把
每个单元格平铺成独立段落，行列关系彻底消失，而且不报错、不告警。

本测试用**合成 docx**（zipfile 直接拼 OOXML，不依赖 python-docx，也不依赖
任何真实文档）钉死这几件事：
  1. 表格转成 GFM 表格，行列对齐，单元格内容不串行；
  2. 表头（tblHeader 或整行加粗）被识别；
  3. 横向合并 gridSpan 不复制内容、不撑破列数；
  4. 列表项按 numPr 转成 Markdown 列表，有序/无序分开；
  5. 行内加粗/斜体/删除线/上下标/超链接/脚注都落成对应 Markdown；
  6. 段落与表格的**原始交错顺序**保持不变。

全部用合成文档：真实文档没有表格，拿它测不出这类回归。
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
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

FAILS = []
OKS = []


def ok(msg):
    OKS.append(msg)


def bad(msg):
    FAILS.append(msg)


CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOCRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/x" TargetMode="External"/>
<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
<Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
<Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
<w:style w:type="paragraph" w:styleId="1"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/>
  <w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
</w:styles>"""

NUMBERING = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="{W}">
<w:abstractNum w:abstractNumId="10">
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="11">
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="10"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="11"/></w:num>
</w:numbering>"""

FOOTNOTES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{W}">
<w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:t>sep</w:t></w:r></w:p></w:footnote>
<w:footnote w:id="1"><w:p><w:r><w:t>这是脚注正文</w:t></w:r></w:p></w:footnote>
</w:footnotes>"""


def p(text="", style=None, num=None, runs=None, ilvl=0):
    ppr = ""
    if style or num:
        inner = ""
        if style:
            inner += f'<w:pStyle w:val="{style}"/>'
        if num:
            inner += (f'<w:numPr><w:ilvl w:val="{ilvl}"/>'
                      f'<w:numId w:val="{num}"/></w:numPr>')
        ppr = f"<w:pPr>{inner}</w:pPr>"
    if runs is None:
        runs = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{ppr}{runs}</w:p>"


def r(text, b=False, i=False, strike=False, vert=None, link=None, fn=None):
    rpr = ""
    if b:
        rpr += "<w:b/>"
    if i:
        rpr += "<w:i/>"
    if strike:
        rpr += "<w:strike/>"
    if vert:
        rpr += f'<w:vertAlign w:val="{vert}"/>'
    rpr = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    out = f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'
    if fn:
        out += (f'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                f'<w:footnoteReference w:id="{fn}"/></w:r>')
    if link:
        out = f'<w:hyperlink r:id="{link}">{out}</w:hyperlink>'
    return out


def tc(runs_xml, span=1, vmerge=None, text=None):
    if runs_xml is None:
        runs_xml = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    tcpr = ""
    if span > 1 or vmerge:
        inner = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
        if vmerge:
            inner += f'<w:vMerge w:val="{vmerge}"/>'
        tcpr = f"<w:tcPr>{inner}</w:tcPr>"
    return f"<w:tc>{tcpr}{p(runs=runs_xml)}</w:tc>"


def tr(cells, header=False):
    trpr = "<w:trPr><w:tblHeader/></w:trPr>" if header else ""
    return f"<w:tr>{trpr}{''.join(cells)}</w:tr>"


def build_docx(path):
    body = []
    # 1) 标题 + 普通段落
    body.append(p("文档标题", style="2"))
    body.append(p("这是正文第一段。"))
    # 2) 表格：表头 + 两行 + 横向合并
    body.append(
        "<w:tbl><w:tblPr/>"
        + tr([tc(None, text="表头A"), tc(None, text="表头B"),
              tc(None, text="表头C")], header=True)
        + tr([tc(None, text="a1"), tc(None, text="b1"), tc(None, text="c1")])
        + tr([tc(None, text="跨两列", span=2), tc(None, text="c2")])
        + "</w:tbl>"
    )
    # 3) 表格后的段落：顺序必须保住
    body.append(p("表格之后的段落。"))
    # 4) 有序列表 3 项 + 无序列表 2 项
    body.append(p("第一列表项", num=1))
    body.append(p("第二列表项", num=1))
    body.append(p("第三列表项", num=1))
    body.append(p("无序一", num=2))
    body.append(p("无序二", num=2))
    # 5) 行内格式
    body.append(p(runs=(
        r("加粗", b=True) + r("普通") + r("斜体", i=True)
        + r("删除", strike=True) + r("上标", vert="superscript")
        + r("下标", vert="subscript")
    )))
    # 6) 超链接 + 脚注
    body.append(p(runs=r("点这里", link="rId10") + r("，脚注", fn="1")))

    doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document xmlns:w="{W}" '
           f'xmlns:r="{R}"><w:body>'
           + "".join(body)
           + '<w:sectPr/></w:body></w:document>')

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", DOCRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING)
        z.writestr("word/footnotes.xml", FOOTNOTES)


def run(cmd, cwd=None):
    r = subprocess.run(cmd, cwd=cwd or ROOT, capture_output=True, text=True)
    return r.returncode, r.stdout, r.stderr


def main():
    tmp = tempfile.mkdtemp(prefix="docgen_")
    try:
        docx = os.path.join(tmp, "generic.docx")
        build_docx(docx)
        out = os.path.join(tmp, "out")
        rc, so, se = run([sys.executable, os.path.join(ROOT, "py", "extract.py"),
                          docx, out])
        if rc != 0:
            bad(f"extract 失败({rc}): {(se or so).strip()[:400]}")
            return report()

        bp = os.path.join(out, "blocks.json")
        if not os.path.isfile(bp):
            bad("extract 没有产出 blocks.json")
            return report()
        with open(bp, encoding="utf-8") as f:
            blocks = json.load(f)["blocks"]

        kinds = [b["kind"] for b in blocks]
        if "table" not in kinds:
            bad(f"blocks 里没有表格：{kinds}")
        else:
            t = blocks[kinds.index("table")]
            if len(t["rows"]) != 3:
                bad(f"表格行数应为 3，实际 {len(t['rows'])}")
            elif [len(x) for x in t["rows"]] != [3, 3, 2]:
                bad(f"每行列数应为 [3,3,2]，实际 {[len(x) for x in t['rows']]}")
            else:
                ok("表格 3 行、合并行 2 格")
            if not t.get("header"):
                bad("表头未被识别（tblHeader）")
            else:
                ok("表头识别（tblHeader）")
            first = t["rows"][0][0]["text"]
            if first.strip() != "表头A":
                bad(f"表头首格内容错：{first!r}")
            merged = t["rows"][2][0]["text"]
            if "跨两列" not in merged:
                bad(f"合并单元格内容错：{merged!r}")
            else:
                ok("合并单元格内容正确")

        # 顺序：段落 → 表格 → 段落
        seq = "".join("T" if k == "table" else "P" for k in kinds)
        if "PTP" not in seq:
            bad(f"段落/表格交错顺序丢了：{seq}")
        else:
            ok("段落与表格的原始顺序保住")

        # 列表
        nums = [b.get("num") for b in blocks if b["kind"] == "para" and b.get("num")]
        if len(nums) < 5:
            bad(f"列表项应至少 5 个，实际 {len(nums)}")
        else:
            ok(f"列表项读到 {len(nums)} 个")

        # 行内格式
        flat = []
        for b in blocks:
            if b["kind"] == "para":
                flat.extend(b["runs"])
            elif b["kind"] == "table":
                for row in b["rows"]:
                    for c in row:
                        flat.append({"text": c.get("text") or ""})
        if not any(x.get("bold") for x in flat):
            bad("没有读到任何加粗 run")
        else:
            ok("加粗 run 读到")
        if not any(x.get("italic") for x in flat):
            bad("没有读到任何斜体 run")
        else:
            ok("斜体 run 读到")
        if not any(x.get("strike") for x in flat):
            bad("没有读到任何删除线 run")
        else:
            ok("删除线 run 读到")
        if not any(x.get("vert") == "sup" for x in flat):
            bad("没有读到上标 run")
        else:
            ok("上标 run 读到")
        if not any(x.get("link") for x in flat):
            bad("没有读到超链接 URL")
        else:
            ok("超链接 URL 读到")
        if not any(x.get("footnote") for x in flat):
            bad("没有读到脚注正文")
        else:
            ok("脚注正文读到")

        # ---- build ----
        md = os.path.join(tmp, "out.md")
        rc, so, se = run([sys.executable, os.path.join(ROOT, "py", "build.py"),
                          out, "--latex", os.path.join(tmp, "latex.json"),
                          "--out", md])
        if rc != 0:
            # 无公式的文档不需要 latex 表；给个空表重试
            with open(os.path.join(tmp, "latex.json"), "w", encoding="utf-8") as f:
                f.write("{}")
            rc, so, se = run([sys.executable, os.path.join(ROOT, "py", "build.py"),
                              out, "--latex", os.path.join(tmp, "latex.json"),
                              "--out", md])
        if rc != 0:
            bad(f"build 失败({rc}): {(se or so).strip()[:400]}")
            return report()

        text = open(md, encoding="utf-8").read()
        checks = [
            ("| 表头A | 表头B | 表头C |", "GFM 表头行"),
            ("| --- | --- | --- |", "GFM 分隔行"),
            ("| a1 | b1 | c1 |", "表格数据行"),
            ("**加粗**", "加粗 → **"),
            ("*斜体*", "斜体 → *"),
            ("~~删除~~", "删除线 → ~~"),
            ("<sup>上标</sup>", "上标 → <sup>"),
            ("<sub>下标</sub>", "下标 → <sub>"),
            ("[点这里](https://example.com/x)", "超链接 → [](url)"),
            ("1. 第一列表项", "有序列表 → 1."),
            ("- 无序一", "无序列表 → -"),
            ("表格之后的段落。", "表格后段落未丢"),
        ]
        for needle, label in checks:
            if needle in text:
                ok(label)
            else:
                bad(f"{label} 缺失，期望含 {needle!r}")
        if "脚注" in text and "这是脚注正文" in text:
            ok("脚注正文落进 md")
        else:
            bad("脚注正文未落进 md")

        # 合并单元格不该复制内容
        if text.count("跨两列") != 1:
            bad(f"合并单元格内容被复制了 {text.count('跨两列')} 次")
        else:
            ok("合并单元格内容不复制")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return report()


def report():
    for m in OKS:
        print(f"  ok   {m}")
    for m in FAILS:
        print(f"  FAIL {m}")
    print(f"\n{len(OKS)} 通过 / {len(FAILS)} 失败")
    if FAILS:
        print("\n通用结构 e2e 未通过")
        return 1
    print("通用结构 e2e 全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
