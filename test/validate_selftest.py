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
            return report.get("status"), report.get("errors", [])
        except json.JSONDecodeError:
            return "CRASH", [r.stderr.strip()[:200]]
    finally:
        os.remove(path)


def main():
    failures = 0
    print("== validate.py 自测 ==")
    for name, content, expect in CASES:
        got, errors = run_case(content)
        ok = got == expect
        print(("  PASS  " if ok else "  FAIL  ") + f"{name}: 期望 {expect}，实际 {got}")
        if not ok:
            failures += 1
            for e in errors[:3]:
                print("         " + str(e)[:120])
    print()
    if failures:
        print(f"{failures} 项不符合预期——校验器需要检查")
        return 1
    print("全部符合预期：校验器对好样本放行，对每类反例都能报错")
    return 0


if __name__ == "__main__":
    sys.exit(main())
