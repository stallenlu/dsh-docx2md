#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pict_ole_e2e.py —— 「画布里装着公式」不能静默丢公式。

这个测试钉死的是一次真实的 P0 回归：

    Word 画流程图时，公式会嵌在同一个 w:pict 的 v:group 里。那个 pict
    既是「装着公式的画布」，又长得像「公式的预览图」（里面有 o:ole="t"
    的 v:shape）。实现若按「见 o:ole="t" 就算预览图」判断，整棵 pict 被
    当成一张预览图跳过，里面的 o:OLEObject 一并被吞掉——

        正文照常输出，不报错、不告警，只是公式消失了。

    实测在真实文档（145 个公式）里，3 个 pict 各装着 6–8 个公式，
    共 20 个公式无声消失。校验器直到 build 阶段才对不上数（T8 FAIL），
    而那时原因已经无从查证。

本测试覆盖四件事：

    1. extract：pict 内的公式必须生成占位符，**一个都不能少**。
    2. extract：画布里的示意图仍按原规则处理（进清单，不进正文）。
    3. 提取自检：索引有、正文没有占位符时，extract 必须**报错退出**
       （退出码 4），而不是带着缺口往下走。
    4. build + validate：占位符齐了，整条链路的公式数才对得上。

用法: python3 test/pict_ole_e2e.py     退出码 0 = 全绿
"""
import json
import os
import re
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
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"

CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Default Extension="wmf" ContentType="image/x-wmf"/></Types>')

PNG = b"\x89PNG\r\n\x1a\n" + b"\0" * 64
WMF = b"\xd7\xcd\xc6\x9a" + b"\0" * 64


def rels(n_ole, n_fig=1):
    out = ['<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i in range(1, n_ole + 1):
        out.append('<Relationship Id="rIdOle%d" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject%d.bin"/>'
                   % (i, i))
        out.append('<Relationship Id="rIdPrev%d" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/image" Target="media/prev%d.wmf"/>'
                   % (i, i))
    for i in range(1, n_fig + 1):
        out.append('<Relationship Id="rIdFig%d" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/image" Target="media/fig%d.png"/>' % (i, i))
    out.append("</Relationships>")
    return "".join(out)


def ole_in_pict(shape_id, rid_ole, rid_prev, group_id="grp1"):
    """一个 v:group 里的公式：o:ole="t" 的预览 shape + o:OLEObject 本体。

    这正是 Word 流程图里公式的真实结构（真实文档里 ShapeID 形如
    "Object 78"、shape id 也叫 "Object 78"）。
    """
    return (
        '<v:group id="%s" style="position:absolute;width:459.6pt;height:346.55pt">'
        '<v:shape id="%s" o:ole="t" type="#_x0000_t75" style="position:absolute;left:100;top:100">'
        '<v:imagedata r:id="%s"/></v:shape>'
        '<o:OLEObject Type="Embed" ProgID="Equation.3" ShapeID="%s" r:id="%s"/>'
        '</v:group>' % (group_id, shape_id, rid_prev, shape_id, rid_ole))


def fig_shape(rid_fig, sid="_x0000_s2050"):
    """画布里的一张示意图（不是公式）：只有 shape + imagedata，没有 OLEObject。"""
    return ('<v:shape id="%s" type="#_x0000_t75" style="position:absolute;left:10;top:10">'
            '<v:imagedata r:id="%s"/></v:shape>' % (sid, rid_fig))


def make_docx(path, n_ole=3, in_pict=True):
    """造一个 docx：段落里的公式装在 w:pict 的 v:group 里（+ 一张示意图）。"""
    parts = []
    for i in range(1, n_ole + 1):
        sid = "Object %d" % (100 + i)
        if in_pict:
            parts.append(ole_in_pict(sid, "rIdOle%d" % i, "rIdPrev%d" % i))
        else:
            # 对照组：公式在正常的 w:object 里
            parts.append('<w:r><w:object><o:OLEObject Type="Embed" ProgID="Equation.3" '
                         'ShapeID="%s" r:id="rIdOle%d"/></w:object></w:r>' % (sid, i))
    doc = ('<?xml version="1.0"?>'
           '<w:document xmlns:w="%s" xmlns:r="%s" xmlns:v="%s" xmlns:o="%s"><w:body>'
           '<w:p><w:r><w:t>流程如下</w:t></w:r>'
           '<w:r><w:pict>%s%s</w:pict></w:r>'
           '<w:r><w:t>结束。</w:t></w:r>'
           '</w:p></w:body></w:document>'
           % (W, R, V, O, "".join(parts), fig_shape("rIdFig1")))

    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", rels(n_ole))
        for i in range(1, n_ole + 1):
            z.writestr("word/embeddings/oleObject%d.bin" % i, b"\0" * 64)
            z.writestr("word/media/prev%d.wmf" % i, WMF)
        z.writestr("word/media/fig1.png", PNG)


def make_textbox_docx(path, n_ole=2):
    """造一个 docx：公式装在流程图的**文本框**（w:txbxContent）里。

    文本框是另一种「长得像正文、其实不是」的东西：它的段落自己也是 w:p，
    段落遍历会把它顶出来，于是框里的内容既被当作独立段落收一次、
    又在外层段落里被访问一次——公式占位符因此出现两次。

    框里的文字是流程图的标签（"开始"、"结束"），不是正文段落。
    """
    parts = []
    for i in range(1, n_ole + 1):
        sid = "Object %d" % (200 + i)
        parts.append(
            '<v:shape id="tb%d" type="#_x0000_t202" style="width:120pt;height:40pt">'
            '<v:textbox><w:txbxContent><w:p><w:r><w:t>框内标签</w:t></w:r>'
            '<w:r><w:object><o:OLEObject Type="Embed" ProgID="Equation.3" '
            'ShapeID="%s" r:id="rIdOle%d"/></w:object></w:r>'
            '</w:p></w:txbxContent></v:textbox></v:shape>' % (i, sid, i))
    doc = ('<?xml version="1.0"?>'
           '<w:document xmlns:w="%s" xmlns:r="%s" xmlns:v="%s" xmlns:o="%s"><w:body>'
           '<w:p><w:r><w:t>正文段落</w:t></w:r></w:p>'
           '<w:p><w:r><w:pict><v:group id="grp1">%s</v:group></w:pict></w:r></w:p>'
           '<w:p><w:r><w:t>结束。</w:t></w:r></w:p>'
           '</w:body></w:document>'
           % (W, R, V, O, "".join(parts)))

    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", rels(n_ole))
        for i in range(1, n_ole + 1):
            z.writestr("word/embeddings/oleObject%d.bin" % i, b"\0" * 64)
            z.writestr("word/media/prev%d.wmf" % i, WMF)


def run(script, *args, expect_ok=True):
    r = subprocess.run([sys.executable, script, *args],
                       cwd=PY_DIR, capture_output=True, text=True)
    if expect_ok and r.returncode != 0:
        raise RuntimeError("%s 失败(%d): %s" % (script, r.returncode,
                                                (r.stderr or r.stdout)[:400]))
    return r


def placeholders(out):
    with open(os.path.join(out, "paras.json"), encoding="utf-8") as f:
        paras = json.load(f)
    return re.findall(r"⟦EQ:([^⟧]+)⟧", "".join(paras))


work = tempfile.mkdtemp(prefix="d2m-pictole-")
fails = []
try:
    docx = os.path.join(work, "pictole.docx")
    out = os.path.join(work, "out")
    make_docx(docx, n_ole=3)

    # 1. extract：pict 内的公式必须一个不少地生成占位符
    r = run("extract.py", docx, out)
    ph = placeholders(out)
    if len(ph) != 3:
        fails.append("extract: pict 内 3 个公式只生成了 %d 个占位符 %s"
                     % (len(ph), ph))
    with open(os.path.join(out, "eq_index.json"), encoding="utf-8") as f:
        index = json.load(f)
    if len(index) != 3:
        fails.append("extract: eq_index 应有 3 条，实际 %d" % len(index))
    streams = {r_["stream"] for r_ in index}
    missing = streams - set(ph)
    if missing:
        fails.append("extract: 这些公式在正文里没有占位符: %s" % sorted(missing))

    # 2. 画布里的示意图不因修复而混进正文
    with open(os.path.join(out, "paras.json"), encoding="utf-8") as f:
        paras = json.load(f)
    if any("fig1.png" in p for p in paras):
        fails.append("extract: 示意图混进了正文占位")

    # 3. 自检：把公式从正文里拿掉，extract 必须报错退出（退出码 4）
    #    —— 用 monkeypatch 模拟「容器把公式吞了」的场景。
    probe = os.path.join(work, "probe.py")
    with open(probe, "w", encoding="utf-8") as f:
        f.write(
            "import sys, os\n"
            "sys.path.insert(0, %r)\n"
            "import extract, imgformula\n"
            "orig = imgformula.holds_ole_formula\n"
            "imgformula.holds_ole_formula = lambda node: False\n"
            "extract.holds_ole_formula = imgformula.holds_ole_formula\n"
            "sys.argv = ['extract.py', %r, %r]\n"
            "extract.main()\n" % (PY_DIR, docx, os.path.join(work, "out2")))
    rp = subprocess.run([sys.executable, probe], cwd=PY_DIR,
                        capture_output=True, text=True)
    if rp.returncode != 4:
        fails.append("自检: 公式被吞时 extract 应以退出码 4 拦下，实际 %d"
                     % rp.returncode)
    if "静默丢失" not in (rp.stderr or ""):
        fails.append("自检: 错误信息应说明是静默丢失，实际: %s"
                     % (rp.stderr or "")[:200])

    # 4. 端到端：build + validate 的公式数要对得上
    latex = os.path.join(work, "latex.json")
    json.dump({s: "x" for s in streams}, open(latex, "w", encoding="utf-8"))
    md = os.path.join(work, "out.md")
    b = json.loads(run("build.py", out, "--latex", latex, "--out", md,
                       "--use-suggested").stdout)
    if b.get("formulas_missing") != 0:
        fails.append("build: formulas_missing=%s，应为 0" % b.get("formulas_missing"))
    if b.get("unresolved_placeholders") != 0:
        fails.append("build: unresolved_placeholders=%s，应为 0"
                     % b.get("unresolved_placeholders"))
    v = json.loads(run("validate.py", md, "--expect-formulas", "3", "--json").stdout)
    if v.get("status") != "PASS":
        fails.append("validate: %s errors=%s" % (v.get("status"), v.get("errors")[:2]))

    # 5. 文本框（w:txbxContent）里的公式：占位符只出现一次，标签不进正文
    tb_docx = os.path.join(work, "textbox.docx")
    tb_out = os.path.join(work, "out_tb")
    make_textbox_docx(tb_docx, n_ole=2)
    run("extract.py", tb_docx, tb_out)
    tb_ph = placeholders(tb_out)
    if len(tb_ph) != 2:
        fails.append("文本框: 2 个公式应生成 2 个占位符，实际 %d %s"
                     % (len(tb_ph), tb_ph))
    if len(set(tb_ph)) != len(tb_ph):
        fails.append("文本框: 占位符重复（公式被数两遍）: %s" % tb_ph)
    with open(os.path.join(tb_out, "paras.json"), encoding="utf-8") as f:
        tb_paras = json.load(f)
    if any("框内标签" in p for p in tb_paras):
        fails.append("文本框: 框内标签混进了正文段落")
    labels_path = os.path.join(tb_out, "textbox_labels.json")
    if not os.path.exists(labels_path):
        fails.append("文本框: 没有生成 textbox_labels.json")
    else:
        with open(labels_path, encoding="utf-8") as f:
            labels = json.load(f)
        if "框内标签" not in labels:
            fails.append("文本框: 标签没进 textbox_labels.json: %s" % labels)

    # 6. inspect 的段落数必须与 extract 的正文段落数一致（同一口径）
    scan = json.loads(run("docx_scan.py", tb_docx).stdout)
    if scan.get("paragraphs") != len(tb_paras):
        fails.append("inspect 报 %s 段，extract 产出 %s 段——口径不一致"
                     % (scan.get("paragraphs"), len(tb_paras)))
    if scan.get("formulas_total") != 2:
        fails.append("inspect: 文本框里的公式没被计入，formulas_total=%s"
                     % scan.get("formulas_total"))

    # 6b. inspect 报的「段落文字」必须与 extract 的 paras.json 一致。
    #     实测踩过：inspect 的 para_texts 直接遍历 w:t，把**文本框里**的
    #     标签也捞进来，于是 inspect 说「这一段的文字是『框内标签』」，
    #     而 extract 的同一段只有公式占位符——同一个索引，两个工具各说各话，
    #     排查时会被当成「文档本来就这样」。
    #
    #     比的是**文字**：extract 的段落里还带着 ⟦EQ:…⟧ 占位符，那是公式
    #     的位置标记，inspect 的 paragraph_text 里当然没有——先剥掉再比。
    def strip_ph(s):
        return re.sub(r"⟦EQ:[^⟧]*⟧", "", s or "").strip()

    ctx = scan.get("missing_preview_context") or []
    for c in ctx:
        idx = c.get("paragraph")
        if idx is None or idx >= len(tb_paras):
            continue
        if (c.get("paragraph_text") or "") != strip_ph(tb_paras[idx]):
            fails.append("inspect 上下文与 extract 不一致: 段 %s inspect=%r "
                         "extract=%r" % (idx, c.get("paragraph_text"),
                                         strip_ph(tb_paras[idx])))
        if "框内标签" in (c.get("paragraph_text") or ""):
            fails.append("inspect: 上下文里混进了文本框标签 %r"
                         % c.get("paragraph_text"))

    # 7. 过期 headings.json 必须被拒：它按旧段落序号索引，序号一变就指错。
    #    指错的后果特别隐蔽——若被指的段落恰好是公式段，公式会被当成标题
    #    直接输出，公式凭空消失，直到 T8 才发现数对不上。
    stale_out = os.path.join(work, "out_stale")
    shutil.copytree(tb_out, stale_out)
    with open(os.path.join(stale_out, "headings.json"), "w", encoding="utf-8") as f:
        json.dump({"1": 2}, f)          # 段落 1 = 装着公式的那段
    rs = run("build.py", stale_out, "--latex", os.path.join(work, "latex.json"),
             "--out", os.path.join(work, "stale.md"), expect_ok=False)
    if rs.returncode != 3:
        fails.append("过期标题表: build 应以退出码 3 拒绝，实际 %d"
                     % rs.returncode)
    if "公式" not in (rs.stderr or ""):
        fails.append("过期标题表: 错误信息应说明会吞掉公式，实际: %s"
                     % (rs.stderr or "")[:200])

    print("== 画布内公式（w:pict）端到端 ==")
    if fails:
        for f_ in fails:
            print("  FAIL " + f_)
        print("\n%d 项失败" % len(fails))
        sys.exit(1)
    print("  ok   pict 内公式占位符一个不少 / 示意图仍不进正文 / "
          "自检拦下静默丢失 / 端到端计数相符 / "
          "文本框内公式只占一次且标签不进正文 / inspect 与 extract 段落口径一致 / "
          "过期标题表被拒（不会把公式当标题吞掉）")
    print("\n全部通过：画布里的公式不会被当成预览图整棵跳过，"
          "也不会被文本框或过期标题表吞掉")
finally:
    shutil.rmtree(work, ignore_errors=True)
