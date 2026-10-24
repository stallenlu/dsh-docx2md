#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""validate.py —— 校验生成的 Markdown 是否符合六条铁律。

用法:
    python validate.py <file.md> [--expect-formulas N] [--json]

退出码:
    0 = 全部通过
    1 = 有错误级问题

检查项（对应 METHODOLOGY 的六条铁律 + 结构校验）:
    T1  $$ 必须独占一行
    T2  行内公式不得跨行 / 不得含 \\begin{aligned} 等块级环境
    T3  相邻公式之间必须有空行或正文
    T4  \\tag 必须紧跟 \\end{aligned}（或 \\end{array} 等）
    T5  表格单元格内不得出现超长公式
    T6  图片占位应为 HTML 注释
    T7  $$ 配对数为偶数（成对闭合）
    T8  公式编号连续、无重复（可选，--expect-formulas 时检查总数）
"""
import argparse
import json
import re
import sys
from docxfile import force_utf8_stdio

BLOCK_ENVS = ("aligned", "array", "matrix", "pmatrix", "bmatrix",
              "cases", "split", "gather", "align", "eqnarray")


def check(md_lines):
    errors, warns = [], []

    n_dollar = sum(ln.count("$$") for ln in md_lines)
    if n_dollar % 2 != 0:
        errors.append("T7: $$ 出现 %d 次（奇数），存在未闭合的公式块" % n_dollar)

    for i, ln in enumerate(md_lines, 1):
        s = ln.strip()

        # T1: $$ 独占一行
        if "$$" in ln and s not in ("$$",):
            # 允许行尾紧跟 \tag 的收尾行？不：收尾 $$ 必须独占
            errors.append("T1: 第 %d 行 $$ 未独占一行: %s" % (i, s[:90]))

        # T2: 行内公式不得含块级环境
        for m in re.finditer(r"(?<!\$)\$([^$\n]+)\$(?!\$)", ln):
            body = m.group(1)
            if "\\begin{" in body:
                errors.append("T2: 第 %d 行行内公式含块级环境: %s" % (i, body[:80]))

        # T5: 表格行里的超长公式
        if s.startswith("|") and s.count("|") >= 2:
            if len(ln) > 400 and ("$" in ln):
                warns.append("T5: 第 %d 行表格单元格含长公式（%d 字符）" % (i, len(ln)))

        # T6: 图片占位
        if re.search(r"!\[[^\]]*\]\(\s*\)", ln):
            warns.append("T6: 第 %d 行是空图片链接，建议改用 HTML 注释占位" % i)

    # T4: \tag 紧跟 \end{...}
    for i, ln in enumerate(md_lines, 1):
        if "\\tag" in ln:
            before = ln.split("\\tag")[0].rstrip()
            if not re.search(r"\\end\{[a-z]+\}$", before) and "\\\\" not in before:
                # 单行公式 \tag 是合法的，只警告
                if not re.match(r"^\s*\$?\$?.*\\tag\{", ln):
                    warns.append("T4: 第 %d 行 \\tag 未紧跟 \\end{}" % i)

    # T3: 相邻公式块之间要有空行
    in_block = False
    prev_close = None
    for i, ln in enumerate(md_lines, 1):
        s = ln.strip()
        if s == "$$":
            if not in_block:
                in_block = True
                if prev_close is not None and i - prev_close == 1:
                    errors.append("T3: 第 %d 行的公式块紧贴上一个块结尾" % i)
            else:
                in_block = False
                prev_close = i

    return errors, warns


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--expect-formulas", type=int, default=None)
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    with open(args.file, "r", encoding="utf-8") as f:
        md = f.read()
    lines = md.split("\n")

    errors, warns = check(lines)

    # 空文档不算通过：没有内容就没有可验证的东西，PASS 会给出错误的信心。
    if not md.strip():
        errors.append("T0: 文件为空，没有可校验的内容")

    # 未替换的公式。build.py 遇到 latex 表里没有的公式时，会把它写成
    # <!-- 未解析公式: oleObject7.bin -->，而不是留下 ⟦EQ:...⟧ 字面量
    # （两种形态都要查，因为别的工具可能直接留下占位符）。
    # 这类文档的 Markdown 结构完全合规，六条铁律全过，但正文里少了公式——
    # 只查铁律会给出 PASS，用户却拿到一份读不了的文档。
    left = re.findall(r"\u27e6EQ:([^\u27e7]+)\u27e7", md)
    if left:
        uniq = sorted(set(left))
        errors.append(
            "T9: 有 %d 处公式占位符未替换（%d 个不同公式：%s）"
            % (len(left), len(uniq), ", ".join(uniq[:12])
               + ("…" if len(uniq) > 12 else "")))
    marked = re.findall(r"<!--\s*未解析公式:\s*([^\s>]+)\s*-->", md)
    if marked:
        uniq = sorted(set(marked))
        errors.append(
            "T9: 有 %d 处公式未解析（%d 个不同公式：%s）"
            % (len(marked), len(uniq), ", ".join(uniq[:12])
               + ("…" if len(uniq) > 12 else "")))

    n_display = sum(1 for ln in lines if ln.strip() == "$$") // 2
    n_inline = len(re.findall(r"(?<!\$)\$(?!\$)[^$\n]+\$(?!\$)", md))
    n_tag = len(re.findall(r"\\tag\{", md))

    report = {
        "file": args.file,
        "lines": len(lines),
        "chars": len(md),
        "display_blocks": n_display,
        "inline_formulas": n_inline,
        "tags": n_tag,
        "errors": errors[:40],
        "warnings": warns[:40],
        "error_count": len(errors),
        "warning_count": len(warns),
        "status": "PASS" if not errors else "FAIL",
    }
    if args.expect_formulas is not None:
        report["expected_formulas"] = args.expect_formulas
        if n_display + n_inline != args.expect_formulas:
            report["status"] = "FAIL"
            report["errors"].append(
                "T8: 公式数 %d（display %d + inline %d）与预期 %d 不符"
                % (n_display + n_inline, n_display, n_inline,
                   args.expect_formulas))
            report["error_count"] += 1

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=1))
    else:
        print("文件: %s" % args.file)
        print("行数: %d  字符: %d" % (len(lines), len(md)))
        print("公式: display %d / inline %d / tag %d"
              % (n_display, n_inline, n_tag))
        print("错误: %d  警告: %d  状态: %s"
              % (len(errors), len(warns), report["status"]))
        for e in errors[:20]:
            print("  [ERROR] " + e)
        for w in warns[:20]:
            print("  [WARN ] " + w)

    sys.exit(0 if not errors else 1)


if __name__ == "__main__":
    force_utf8_stdio()
    main()
# x
