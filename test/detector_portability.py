#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""detector_portability.py —— T13/T14 是通用探测器，不是对某一份文档调出来的阈值。

为什么要有这个测试：T13（多行公式被并成一行）与 T14（转写漏内容）的判据
是在一份 145 公式的真实文档上定下来的——1.6×行高、6×密度中位、40 字符、
200px。看着像"为这份文档量身裁衣"。它们其实全部**相对于文档自身的标定
基线**（普通单行的中位行高、全体公式的中位像素密度），基线随文档字号/
排版缩放，判据就跟着缩放。这个测试把同一组公式**等比放大/缩小**喂给探测
器，验证命中集合不变；再把阈值附近的情形摆出来，证明它不是靠运气分开。

用法: python3 test/detector_portability.py     退出码 0 = 全绿
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "py"))

from validate import check_media_rows, check_media_density   # noqa: E402

# 一份"文档"的骨架：普通单行（决定基线）、真多行（决定行高）、
# 该被 T13 抓的（多行并成一行）、不该抓的（大算符撑高的单行）、
# 该被 T14 抓的（宽图 + 几乎没字）。
LATEX = {
    "e_plain1.bin": "\\sigma_{xx}=0",
    "e_plain2.bin": "\\nabla\\cdot\\mathbf{C}=0",
    "e_plain3.bin": "a=b+c",
    "e_plain4.bin": "\\mathbf{u}=\\mathbf{A}*",
    "e_plain5.bin": "g=(-4\\pi r)^{-1}",
    "e_true_multi.bin": "\\begin{aligned}\na=b \\\\\nc=d\n\\end{aligned}",
    # 三行被并成一行（\\qquad 是分行痕迹）
    "e_squashed.bin": ("\\boldsymbol{\\sigma}=\\mathbf{C}:\\boldsymbol{\\beta},"
                       "\\qquad \\nabla\\cdot\\boldsymbol{\\sigma}=0,"
                       "\\qquad \\nabla\\times\\boldsymbol{\\beta}=\\boldsymbol{\\alpha}"),
    # 大 ∑ 把单行撑到两行高：没有分行痕迹，绝不能报
    "e_bigsum.bin": "\\alpha=\\sum_{b}\\rho_{b}b",
    # 又宽又矮的真·单行长式（宽 646）：没有分行痕迹，不能报
    "e_wide_single.bin": ("E=\\frac{1}{2}<\\boldsymbol{\\sigma}:\\boldsymbol{\\beta}>"
                          "=\\frac{1}{2}\\int d^{3}r\\,C_{ijkl}\\beta_{ij}\\beta_{kl}"),
    # 图很宽但转写只剩 3 个字形：T14 该抓
    "e_lost.bin": "\\sigma_{zz}",
}
# 基准像素（模拟一份 10.5pt 文档的预览图）
SIZES0 = {
    "e_plain1.bin": (60, 24), "e_plain2.bin": (70, 25), "e_plain3.bin": (40, 22),
    "e_plain4.bin": (66, 24), "e_plain5.bin": (80, 26),
    "e_true_multi.bin": (120, 100),          # 2 行
    "e_squashed.bin": (98, 111),             # 真 3 行，转写 1 行
    "e_bigsum.bin": (115, 58),               # 真 1 行，∑ 撑高
    "e_wide_single.bin": (646, 137),         # 真 1 行
    "e_lost.bin": (1109, 43),                # 转写漏了内容
}


def scaled(factor):
    """整份文档等比缩放（换字号/换渲染 DPI），结构不变。"""
    return {k: (max(1, int(w * factor)), max(1, int(h * factor)))
            for k, (w, h) in SIZES0.items()}


def run_detectors(sizes):
    eq = [{"stream": k, "img": k + ".png"} for k in LATEX]
    sizes = {k + ".png": v for k, v in sizes.items()}
    rows, _ = check_media_rows(eq, sizes, LATEX)
    dens, _ = check_media_density(eq, sizes, LATEX)
    return {k for k, *_ in rows}, {k for k, *_ in dens}


def main():
    fails = []
    base_rows, base_dens = run_detectors(SIZES0)
    if base_rows != {"e_squashed.bin"}:
        fails.append("基准命中不对：T13 应只抓 e_squashed，实得 %r" % base_rows)
    if base_dens != {"e_lost.bin"}:
        fails.append("基准命中不对：T14 应只抓 e_lost，实得 %r" % base_dens)

    # 0.5×（小字号）到 4×（大字号/高 DPI）：命中集合必须一模一样
    for f in (0.5, 0.75, 2.0, 4.0):
        r, d = run_detectors(scaled(f))
        if r != base_rows:
            fails.append("缩放 %.2g× 后 T13 命中漂移：期望 %r 实得 %r"
                         % (f, base_rows, r))
        if d != base_dens:
            fails.append("缩放 %.2g× 后 T14 命中漂移：期望 %r 实得 %r"
                         % (f, base_dens, d))

    # 密度对照：合法公式（含真多行）与漏转公式之间必须留有余量，
    # 不能是"刚好卡在这个文档上"的贴边阈值。
    sizes = {k + ".png": v for k, v in SIZES0.items()}
    eq = [{"stream": k, "img": k + ".png"} for k in LATEX]
    import re
    from validate import formula_glyphs
    dens = []
    for k, tex in LATEX.items():
        w, h = SIZES0[k]
        g = formula_glyphs(tex)
        if g and w >= 200:
            dens.append((k, w / g))
    dens.sort(key=lambda x: x[1])
    ok_max = max((d for k, d in dens if k != "e_lost.bin"), default=0)
    bad_min = dict(dens).get("e_lost.bin", 0)
    if not (bad_min > 3 * ok_max):
        fails.append("T14 的合法/漏转之间余量不足：合法最大 %.0fpx/字形，"
                     "漏转 %.0fpx/字形（要求 >3×）" % (ok_max, bad_min))

    if fails:
        print("FAIL (%d)" % len(fails))
        for f in fails:
            print("  -", f)
        return 1
    print("detector_portability: OK（0.5×–4× 缩放下命中集合不变，"
          "T14 余量 %.0f 倍）" % (bad_min / max(1.0, ok_max)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
