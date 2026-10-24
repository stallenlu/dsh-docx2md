#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""validate_selftest.py —— 验证 validate.py 真的能抓到每一类问题。

一个校验器最危险的情况是「对坏输入也返回 PASS」。这里为 T0–T8 每条规则
各造一个反例，断言它们都被判为 FAIL；再造一个合规样本，断言它 PASS。

用法:
    python validate_selftest.py
退出码: 0 = 校验器工作正常，1 = 校验器有漏判/误判
"""
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
VALIDATOR = os.path.join(os.path.dirname(HERE), "py", "validate.py")

# (名称, 内容, 期望状态)
CASES = [
    ("合规样本", """# 标题

正文含行内公式 $a+b$ 与 $c$ 两个。

$$
E = mc^2
$$
""", "PASS"),

    ("T0 空文件", "", "FAIL"),

    ("T1 $$ 未独占一行", """正文。

$$E = mc^2$$
""", "FAIL"),

    ("T2 行内公式含块级环境", r"""正文 $\begin{aligned} a &= b \end{aligned}$ 结束。
""", "FAIL"),

    # 铁律 2 原文是「不能出现 \begin{aligned}、\\、&」。只查 \begin{} 会漏掉后两者，
    # 于是带 \\ 或 & 的行内公式拿到 PASS。这里为每条各钉一个反例。
    ("T2 行内公式含 \\\\（换行符）", "正文 $a \\\\ b$ 结束。\n", "FAIL"),

    ("T2 行内公式含未转义的 &", "正文 $a & b$ 结束。\n", "FAIL"),

    # 转义后的字面量是合法 LaTeX，不能误报。
    ("转义的 \\& 应放行", "正文 $a \\& b$ 结束。\n", "PASS"),

    ("转义的花括号应放行", "正文 $\\{x\\}$ 结束。\n", "PASS"),

    ("T3 相邻公式块紧贴", """$$
a
$$
$$
b
$$
""", "FAIL"),

    ("T7 $$ 奇数个（未闭合）", """正文。

$$
E = mc^2
""", "FAIL"),

    # 最容易漏的一类：Markdown 结构完全合规，但正文里还留着未替换的占位符。
    # 只查六条铁律会给出 PASS，用户却拿到一份读不了的文档。
    ("T9 残留公式占位符", """# 标题

正文 $\alpha$ 正常。

$$
\beta = 2
$$

但这里没换：⟦EQ:oleObject7.bin⟧。
""", "FAIL"),

    ("T9 多个占位符都要报出来", """正文。

$$
a = b
$$

一处 ⟦EQ:oleObject1.bin⟧ 与另一处 ⟦EQ:oleObject2.bin⟧。
""", "FAIL"),

    # build.py 的真实产物形态：未解析的公式被写成 HTML 注释。
    # 这比残留占位符更隐蔽——注释在渲染后不可见，文档看着「干净」。
    ("T9 build.py 生成的未解析注释", """# 标题

正文 $\alpha$ 正常。

$$
\beta = 2
$$

<!-- 未解析公式: oleObject7.bin -->
""", "FAIL"),

    ("正常多行公式 + tag", r"""正文。

$$
\begin{aligned}
a &= b \\
c &= d
\end{aligned}\tag{1}
$$
""", "PASS"),

    ("正常表格与长文本", """| 量 | 含义 |
|---|---|
| $\\alpha$ | 位错密度 |

这是一段普通文字，没有公式，长度也正常。
""", "PASS"),

    # 铁律 5：表格单元格里的公式超过约 60 字符就该移出表格。
    # 早先实现把阈值写成 400，于是 100 字符的公式塞进单元格也拿到 PASS——
    # 文档说 60、代码查 400，这种不一致比没有检查更糟。
    ("T5 表格单元格内 100 字符公式", "| 量 | 含义 |\n|---|---|\n| $%s$ | 说明 |\n"
     % ("A" * 100), "PASS"),
]

# T5 是警告级：它不该把文档判死，但必须真的报出来。
WARN_CASES = [
    ("T5 表格单元格内 100 字符公式应产生警告",
     "| 量 | 含义 |\n|---|---|\n| $%s$ | 说明 |\n" % ("A" * 100),
     "T5"),
    ("T5 表格单元格内 10 字符公式不应警告",
     "| 量 | 含义 |\n|---|---|\n| $%s$ | 说明 |\n" % ("A" * 10),
     None),
    # T10：行内公式紧贴中文。不少渲染器只在 $ 与前后字符有明确边界时
    # 才识别为公式，紧贴中文会按普通文本处理——用户看到的就是
    # 「LaTeX 源码原样出现」。实测真实文档里有 21 处属于这种情况。
    ("T10 公式紧贴前面的中文",
     "其中$\\rho_{b}$是位错密度。\n",
     "T10"),
    ("T10 公式紧贴后面的中文",
     "应力场$\\sigma_{xy}$\n",
     "T10"),
    ("T10 公式与中文留有空格时不应警告",
     "其中 $\\rho_{b}$ 是位错密度。\n",
     None),
]


def run_case(content):
    with tempfile.NamedTemporaryFile("w", suffix=".md", delete=False,
                                     encoding="utf-8") as f:
        f.write(content)
        path = f.name
    try:
        r = subprocess.run([sys.executable, VALIDATOR, path, "--json"],
                           capture_output=True, text=True)
        import json
        try:
            report = json.loads(r.stdout)
            return report.get("status"), report.get("errors", []), report.get("warnings", [])
        except json.JSONDecodeError:
            return "CRASH", [r.stderr.strip()[:200]], []
    finally:
        os.remove(path)


def main():
    failures = 0
    print("== validate.py 自测 ==")
    for name, content, expect in CASES:
        got, errors, _warns = run_case(content)
        ok = got == expect
        print(("  PASS  " if ok else "  FAIL  ") + f"{name}: 期望 {expect}，实际 {got}")
        if not ok:
            failures += 1
            for e in errors[:3]:
                print("         " + str(e)[:120])

    print("\n== T5 警告级检查 ==")
    for name, content, expect_marker in WARN_CASES:
        _status, _errors, warns = run_case(content)
        hit = any(w.startswith(expect_marker) for w in warns) if expect_marker else not warns
        print(("  PASS  " if hit else "  FAIL  ") + name)
        if not hit:
            failures += 1
            print("         实际 warnings: " + str(warns[:2])[:160])

    print()
    if failures:
        print(f"{failures} 项不符合预期——校验器需要检查")
        return 1
    print("全部符合预期：校验器对好样本放行，对每类反例都能报错")
    return 0


if __name__ == "__main__":
    sys.exit(main())
