#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""display_selftest.py —— 钉死「生成的 md 里公式能显示」的几条判据。

背景：用户反馈「md 文档存在部分公式不能正确显示」。排查中发现
「KaTeX 能渲染」和「Markdown 渲染器会认出这是公式」是两层独立判定，
只测前者会给出虚假的信心。本测试把两层都测，并且只用 py/ 的脚本，
不依赖 KaTeX 是否安装——没有 KaTeX 时退化为纯结构断言。

判据：
    D1  行内公式前后与中文之间必须有空格（T10）
    D2  行内公式不得含块级环境 / \\ / &（T2）
    D3  行内公式长度不超过 LONG_TEX_CHARS（否则应已升格为 display）
    D4  $$ 块必须独占成段，且块内无裸 $
    D5  公式总数与输入一致，无未解析占位

用法: python3 test/display_selftest.py <真实.docx> <latex.json>
      不给参数时只用合成段落做结构断言。
退出码 0 = 全绿。
"""
import json
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PY_DIR = os.path.join(os.path.dirname(HERE), "py")

CJK = re.compile(r"[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]")
INLINE = re.compile(r"(?<!\$)\$([^$\n]+?)\$(?!\$)")
BLOCK_ENV = re.compile(r"\\begin\{(aligned|array|matrix|pmatrix|bmatrix|cases|split|gather|align)")


def run_build(workdir, latex_path, out_path, long_tex_chars=None):
    cmd = [sys.executable, "build.py", workdir, "--latex", latex_path,
           "--out", out_path, "--use-suggested"]
    if long_tex_chars is not None:
        cmd += ["--long-tex-chars", str(long_tex_chars)]
    r = subprocess.run(cmd, cwd=PY_DIR, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError("build 失败: %s" % r.stderr.strip()[:400])
    return json.loads(r.stdout)


def analyse(md_path, long_tex_chars):
    """对生成的 md 做全部结构断言，返回 (问题列表, 统计)。"""
    lines = open(md_path, encoding="utf-8").read().split("\n")
    problems = []
    stats = {"inline": 0, "display": 0, "max_inline_len": 0}

    # D4: $$ 配对与独占
    delims = [i for i, ln in enumerate(lines) if ln.strip() == "$$"]
    if len(delims) % 2:
        problems.append("D4: $$ 行数为奇数（%d），有未闭合块" % len(delims))
    blocks = []
    for a, b in zip(delims[0::2], delims[1::2]):
        body = "\n".join(lines[a + 1:b])
        blocks.append(body)
        if "$" in body:
            problems.append("D4: $$ 块内含裸 $（会让块提前结束）: %s" % body[:60])
    stats["display"] = len(blocks)

    # D1/D2/D3: 逐行检查行内公式
    for i, ln in enumerate(lines, 1):
        if ln.strip() == "$$":
            continue
        for m in INLINE.finditer(ln):
            body = m.group(1)
            stats["inline"] += 1
            stats["max_inline_len"] = max(stats["max_inline_len"], len(body))

            before = ln[m.start() - 1] if m.start() > 0 else ""
            after = ln[m.end()] if m.end() < len(ln) else ""
            if before and CJK.match(before):
                problems.append("D1: 行%d 公式紧贴前一个中文（渲染器可能不认）: %s"
                                % (i, ln[max(0, m.start() - 16):m.end() + 16][:90]))
            if after and CJK.match(after):
                problems.append("D1: 行%d 公式紧贴后一个中文（渲染器可能不认）: %s"
                                % (i, ln[max(0, m.start() - 16):m.end() + 16][:90]))
            if BLOCK_ENV.search(body):
                problems.append("D2: 行%d 行内公式含块级环境: %s" % (i, body[:70]))
            if re.search(r"(?<!\\)\\\\", body):
                problems.append("D2: 行%d 行内公式含 \\\\: %s" % (i, body[:70]))
            if len(body) > long_tex_chars:
                problems.append(
                    "D3: 行%d 行内公式 %d 字符，超过阈值 %d 却没升格为 display"
                    % (i, len(body), long_tex_chars))

    return problems, stats


def build_stats_ok(workdir, reported):
    """D5: 公式总数与输入一致。"""
    problems = []
    idx_path = os.path.join(workdir, "eq_index.json")
    total = len(json.load(open(idx_path, encoding="utf-8")))
    got = reported["formulas_display"] + reported["formulas_inline"]
    if got != total:
        problems.append("D5: 公式数 %d 与输入 %d 不符" % (got, total))
    if reported.get("unresolved_placeholders"):
        problems.append("D5: 有 %d 处占位符未解析" % reported["unresolved_placeholders"])
    if reported.get("formulas_missing"):
        problems.append("D5: 有 %d 个公式缺 LaTeX" % reported["formulas_missing"])
    return problems


def make_synthetic(workdir):
    """合成一组典型形态，覆盖会被判错的写法。"""
    shutil.rmtree(workdir, ignore_errors=True)
    os.makedirs(workdir)
    paras = [
        "纯公式⟦EQ:e1.bin⟧",                       # 独占 → display
        "其中⟦EQ:e2.bin⟧是位错密度。",                # 混排短式 → inline + 需补边界
        "应力场⟦EQ:e3.bin⟧",                        # 尾部紧贴中文
        "⟦EQ:e4.bin⟧，",                            # 公式 + 标点
        "长公式混排⟦EQ:e5.bin⟧收尾",                  # 超长 → 应升格 display
    ]
    json.dump(paras, open(os.path.join(workdir, "paras.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    json.dump([{"stream": "e%d.bin" % i, "img": None, "bytes": 8, "variant": "A"}
               for i in range(1, 6)],
              open(os.path.join(workdir, "eq_index.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    latex = {
        "e1.bin": "\\alpha",
        "e2.bin": "\\rho_{b}",
        "e3.bin": "\\sigma_{xy}",
        "e4.bin": "t=0",
        "e5.bin": "x" * 90,
    }
    path = os.path.join(workdir, "latex.json")
    json.dump(latex, open(path, "w", encoding="utf-8"), ensure_ascii=False)
    return path


def main():
    failures = []

    print("== A. 合成段落：结构与边界 ==")
    work = os.path.join(os.environ.get("TMPDIR", "/tmp"), "d2m-display-syn")
    latex_path = make_synthetic(work)
    reported = run_build(work, latex_path, os.path.join(work, "out.md"))
    # 阈值现在是用户可配的（插件配置 longTexChars 或单次调用参数）。
    # 必须用 build 回报的实际阈值，不能在这里写死：写死会让用户一改配置、
    # 这个测试就误报 D3——那是测试的错，不是产物的错。
    limit = reported.get("long_tex_chars", 80)
    problems, stats = analyse(os.path.join(work, "out.md"), limit)
    problems += build_stats_ok(work, reported)
    print("  阈值 %d  行内 %d / display %d / 最长行内 %d 字符"
          % (limit, stats["inline"], stats["display"], stats["max_inline_len"]))
    for p in problems:
        print("  FAIL " + p)
    if not problems:
        print("  ok   全部判据通过")
    else:
        failures += problems
    shutil.rmtree(work, ignore_errors=True)

    # 真实文档：只在给了参数时跑
    if len(sys.argv) >= 3:
        docx, latex = sys.argv[1], sys.argv[2]
        print("\n== B. 真实文档：%s ==" % os.path.basename(docx))
        work = os.path.join(os.environ.get("TMPDIR", "/tmp"), "d2m-display-real")
        shutil.rmtree(work, ignore_errors=True)
        os.makedirs(work)
        r = subprocess.run([sys.executable, "extract.py", docx, work],
                           cwd=PY_DIR, capture_output=True, text=True)
        if r.returncode != 0:
            print("  FAIL extract 失败: %s" % r.stderr[:300])
            return 1
        reported = run_build(work, latex, os.path.join(work, "out.md"))
        limit = reported.get("long_tex_chars", 80)
        problems, stats = analyse(os.path.join(work, "out.md"), limit)
        problems += build_stats_ok(work, reported)
        print("  阈值 %d  行内 %d / display %d / 最长行内 %d 字符"
              % (limit, stats["inline"], stats["display"], stats["max_inline_len"]))
        for p in problems[:20]:
            print("  FAIL " + p)
        if not problems:
            print("  ok   全部判据通过")
        else:
            failures += problems
        shutil.rmtree(work, ignore_errors=True)
    else:
        print("\n== B. 真实文档：跳过（未提供 docx 与 latex.json）==")

    print()
    if failures:
        print("%d 项失败" % len(failures))
        return 1
    print("全部通过：生成的 md 里公式的定界、边界、长度都合规")
    return 0


if __name__ == "__main__":
    sys.exit(main())
