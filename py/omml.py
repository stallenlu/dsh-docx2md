#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""omml.py —— Word 原生公式（OMML）→ LaTeX 的确定性转换。

背景：docx 里的公式有两族，格式完全不同——

  * OLE 公式（Equation Editor / MathType）：私有二进制 MTEF 包在 .bin 里。
    结构不在 XML 里，任何字符级解码都会丢排版（见 SKILL.md 的三条路径），
    只能渲染成图后靠视觉转写。本插件的主战场。
  * OMML（Word 2007+ 默认；2016 以后用 LaTeX 语法敲的公式存下来也是它）：
    结构化 XML（m:oMath），分式、上下标、根号、矩阵都在树里。
    **不需要看图**——可以确定性地转成 LaTeX。

所以这个转换器是纯代码路径：extract 阶段直接产出 LaTeX，预填进
eq_index.json 的 latex 字段；Agent 的 latex.json 只需要覆盖 OLE 公式，
或者用来纠正这里转得不满意的条目（Agent 的值优先）。

覆盖范围：分式（含 linear）、上下标、根式（含 n 次）、定界符、
大运算符（∑∏∫…）、函数名、重音、横线、矩阵、方程组、underset/overset、
希腊字母与常用符号的 Unicode→命令映射、纯文本样式（\\text{}）。
未知结构降级为「拼接子节点」——宁可少一层装饰，绝不丢内容、绝不抛异常。
"""

M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Unicode → LaTeX 命令。KaTeX/MathJax 对裸 Unicode 的处理不一致（strict 模式
# 会直接报错），所以常见数学符号一律映射成命令。
SYM = {
    # 希腊字母（小写）
    "α": "\\alpha", "β": "\\beta", "γ": "\\gamma", "δ": "\\delta",
    "ε": "\\varepsilon", "ζ": "\\zeta", "η": "\\eta", "θ": "\\theta",
    "ι": "\\iota", "κ": "\\kappa", "λ": "\\lambda", "μ": "\\mu",
    "ν": "\\nu", "ξ": "\\xi", "π": "\\pi", "ρ": "\\rho", "σ": "\\sigma",
    "τ": "\\tau", "υ": "\\upsilon", "φ": "\\varphi", "χ": "\\chi",
    "ψ": "\\psi", "ω": "\\omega", "ς": "\\varsigma", "ϑ": "\\vartheta",
    "ϕ": "\\varphi", "ϖ": "\\varpi",
    # 希腊字母（大写）
    "Γ": "\\Gamma", "Δ": "\\Delta", "Θ": "\\Theta", "Λ": "\\Lambda",
    "Ξ": "\\Xi", "Π": "\\Pi", "Σ": "\\Sigma", "Υ": "\\Upsilon",
    "Φ": "\\Phi", "Ψ": "\\Psi", "Ω": "\\Omega",
    # 运算符与关系
    "×": "\\times", "÷": "\\div", "±": "\\pm", "∓": "\\mp",
    "⋅": "\\cdot", "∘": "\\circ", "⊕": "\\oplus", "⊗": "\\otimes",
    "≤": "\\leq", "≥": "\\geq", "≠": "\\neq", "≈": "\\approx",
    "≡": "\\equiv", "∝": "\\propto", "∞": "\\infty", "∂": "\\partial",
    "∇": "\\nabla", "∈": "\\in", "∉": "\\notin", "⊂": "\\subset",
    "⊃": "\\supset", "⊆": "\\subseteq", "⊇": "\\supseteq",
    "∪": "\\cup", "∩": "\\cap", "∀": "\\forall", "∃": "\\exists",
    "∅": "\\emptyset", "∴": "\\therefore", "∵": "\\because",
    "¬": "\\neg", "∧": "\\wedge", "∨": "\\vee",
    "→": "\\rightarrow", "←": "\\leftarrow", "↔": "\\leftrightarrow",
    "⇒": "\\Rightarrow", "⇔": "\\Leftrightarrow", "↦": "\\mapsto",
    "↑": "\\uparrow", "↓": "\\downarrow",
    "⋯": "\\cdots", "…": "\\ldots", "′": "'", "‖": "\\|",
    "∑": "\\sum", "∏": "\\prod", "∫": "\\int", "∮": "\\oint",
    "⋃": "\\bigcup", "⋂": "\\bigcap",
    "ℓ": "\\ell", "ℏ": "\\hbar",
    "−": "-",   # 数学减号 → ASCII
}

# 上/下标形式的 Unicode（OMML 里偶尔直接出现在文本中）
SUPSYM = {"¹": "^{1}", "²": "^{2}", "³": "^{3}", "ⁿ": "^{n}"}
SUBSYM = {c: "_{%d}" % i for i, c in enumerate("₀₁₂₃₄₅₆₇₈₉")}

ESCAPE = {"&": "\\&", "%": "\\%", "#": "\\#", "_": "\\_",
          "{": "\\{", "}": "\\}"}

NARY = {"∑": "\\sum", "∏": "\\prod", "∐": "\\coprod", "∫": "\\int",
        "∬": "\\iint", "∭": "\\iiint", "∮": "\\oint", "⨍": "\\fint",
        "⋃": "\\bigcup", "⋂": "\\bigcap", "⋀": "\\bigwedge",
        "⋁": "\\bigvee", "⨁": "\\bigoplus", "⨂": "\\bigotimes"}

FUNCS = {"sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos",
         "arctan", "sinh", "cosh", "tanh", "log", "ln", "lg", "exp",
         "max", "min", "sup", "inf", "det", "dim", "gcd", "deg", "hom",
         "ker", "arg", "Pr"}

ACC = {"ˆ": "\\hat", "´": "\\acute", "`": "\\grave", "ˇ": "\\check",
       "˘": "\\breve", "˝": "\\ddot", "ˉ": "\\bar", "˜": "\\tilde",
       "˙": "\\dot", "⃗": "\\vec", "→": "\\overrightarrow"}


def _ln(tag):
    return tag.split("}")[-1] if "}" in tag else tag


def _val(el, name):
    """取 <m:name m:val="x"/> 的 x。"""
    if el is None:
        return None
    c = el.find(M + name)
    if c is None:
        return None
    return c.get(M + "val")


def _first(el, name):
    return el.find(M + name)


def _hidden(pr, name):
    v = _val(pr, name)
    return v in ("1", "true", "on")


def esc_text(s):
    """m:t 文本 → LaTeX 数学模式安全形式。"""
    out = []
    for idx, ch in enumerate(s):
        nxt = s[idx + 1] if idx + 1 < len(s) else ""
        if ch in ESCAPE:
            out.append(ESCAPE[ch])
        elif ch in SYM:
            # 命令后面要跟非字母才不会被吃掉下一个词，所以补空格；
            # 但紧跟 _ ^ 时不能补——\sigma _{x} 里的空格既难看，
            # 某些渲染器还会把它当断点。
            out.append(SYM[ch] + ("" if nxt in ("_", "^", "{", "}", " ", "") else " "))
        elif ch in SUPSYM:
            out.append(SUPSYM[ch])
        elif ch in SUBSYM:
            out.append(SUBSYM[ch])
        elif ch == "\u00a0":
            out.append(" ")
        else:
            out.append(ch)
    return "".join(out)


def _wrap_lin(s):
    return "(%s)" % s if (len(s) > 1 or any(c in s for c in "+-")) else s


# ---------------------------------------------------------------- 元素处理


def _run(el):
    """m:r → 文本。m:sty="p"（纯文本样式）包 \\text{}。"""
    sty = _val(_first(el, "rPr"), "sty")
    txt = []
    for c in el:
        t = c.tag
        if t == M + "t":
            txt.append(c.text or "")
        elif t == W + "t":
            txt.append(c.text or "")
        elif t == W + "r":
            txt.extend(cc.text or "" for cc in c if cc.tag == W + "t")
    s = "".join(txt)
    if not s:
        return ""
    if sty == "p":
        return "\\text{%s}" % esc_text(s).replace("\\%", "%")
    return esc_text(s)


def _frac(el):
    n = conv(_first(el, "num"))
    d = conv(_first(el, "den"))
    if _val(_first(el, "fPr"), "type") == "lin":
        return "%s/%s" % (_wrap_lin(n), _wrap_lin(d))
    return "\\frac{%s}{%s}" % (n, d)


def _ssup(el):
    # e 末尾可能是「命令 + 补的空格」（跨 run 拼接时无法预判后面接什么）；
    # 这里紧接 _ 或 ^，空格必须去掉：\sigma _{x} 应为 \sigma_{x}。
    e = conv(_first(el, "e")).rstrip()
    sub = _first(el, "sub")
    sup = _first(el, "sup")
    s = conv(sub) if sub is not None else ""
    u = conv(sup) if sup is not None else ""
    if s and u:
        return "%s_{%s}^{%s}" % (e, s, u)
    if s:
        return "%s_{%s}" % (e, s)
    if u:
        return "%s^{%s}" % (e, u)
    return e


def _rad(el):
    e = conv(_first(el, "e"))
    pr = _first(el, "radPr")
    if pr is None or _hidden(pr, "degHide"):
        return "\\sqrt{%s}" % e
    return "\\sqrt[%s]{%s}" % (conv(_first(el, "deg")), e)


def _delim(ch):
    return {"|": "\\|", "{": "\\{", "}": "\\}", "<": "\\langle",
            ">": "\\rangle", "": ".", " ": "."}.get(ch, ch)


def _d(el):
    pr = _first(el, "dPr")
    beg = _val(pr, "begChr") if pr is not None else None
    end = _val(pr, "endChr") if pr is not None else None
    beg = "(" if beg is None else beg
    end = ")" if end is None else end
    inner = ",".join(conv(e) for e in el.findall(M + "e"))
    if beg == "." or beg == "":
        beg = "."
    if end == "." or end == "":
        end = "."
    if beg == "." and end == ".":
        return "{%s}" % inner          # 无形定界符：只做分组
    return "\\left%s %s \\right%s" % (_delim(beg), inner, _delim(end))


def _nary(el):
    pr = _first(el, "naryPr")
    ch = _val(pr, "chr") or "∑"
    op = NARY.get(ch, ch)
    base = conv(_first(el, "e"))
    sub = _first(el, "sub")
    sup = _first(el, "sup")
    s = conv(sub) if sub is not None and not _hidden(pr, "subHide") else ""
    u = conv(sup) if sup is not None and not _hidden(pr, "supHide") else ""
    if s and u:
        return "%s_{%s}^{%s}%s" % (op, s, u, base)
    if s:
        return "%s_{%s}%s" % (op, s, base)
    if u:
        return "%s^{%s}%s" % (op, u, base)
    return "%s %s" % (op, base)


def _func(el):
    name_el = _first(el, "funcName")
    name = "".join(t.text or "" for t in name_el.iter() if t.tag == M + "t") \
        if name_el is not None else ""
    base = conv(_first(el, "e"))
    if name.lower() in FUNCS:
        return "\\%s %s" % (name.lower(), base)
    if not name:
        return base
    return "\\operatorname{%s} %s" % (name, base)


def _acc(el):
    ch = _val(_first(el, "accPr"), "chr") or "ˆ"
    cmd = ACC.get(ch, "\\hat")
    return "%s{%s}" % (cmd, conv(_first(el, "e")))


def _children(el, name):
    """取子容器 name 的**内容**（直接拼接其子节点），不走容器自身的 handler。

    m:lim 有双重身份：作为运算符元素（lim x→0 那种）它有 e/sub/sup/lim 结构；
    作为 limLow/limUpp 的界限容器时它只是普通包装。后者若交给 conv() 分发，
    会进 _lim 再找不存在的 e 而崩溃（实测）。
    """
    c = _first(el, name)
    if c is None:
        return ""
    return "".join(conv(g) for g in c)


def _bar(el):
    pos = _val(_first(el, "barPr"), "pos") or "top"
    cmd = "\\underline" if pos == "bot" else "\\overline"
    return "%s{%s}" % (cmd, conv(_first(el, "e")))


def _matrix(el):
    rows = []
    for mr in el.findall(M + "mr"):
        rows.append(" & ".join(conv(e) for e in mr.findall(M + "e")))
    return "\\begin{matrix} %s \\end{matrix}" % " \\\\ ".join(rows)


def _eqarr(el):
    lines = [conv(e) for e in el.findall(M + "e")]
    return "\\begin{gathered} %s \\end{gathered}" % " \\\\ ".join(lines)


def _limlow(el):
    return "\\underset{%s}{%s}" % (_children(el, "lim"),
                                   conv(_first(el, "e")))


def _limupp(el):
    return "\\overset{%s}{%s}" % (_children(el, "lim"),
                                  conv(_first(el, "e")))


def _lim(el):
    e = el.find(M + "e")
    if e is None:                      # 缺 e 的畸形 lim：退化为拼接内容
        return "".join(conv(c) for c in el if c.tag != M + "limPr")
    e = conv(e)
    sub = _first(el, "sub")
    lim = _first(el, "lim")
    sup = _first(el, "sup")
    under = conv(sub) if sub is not None else (conv(lim) if lim is not None else "")
    out = "%s_{%s}" % (e, under) if under else e
    if sup is not None:
        out += "^{%s}" % conv(sup)
    return out


def _lit(el):
    return esc_text(el.get(M + "val") or "")


HANDLERS = {
    "r": _run, "f": _frac,
    "sSub": _ssup, "sSup": _ssup, "sSubSup": _ssup,
    "rad": _rad, "d": _d, "nary": _nary, "func": _func,
    "acc": _acc, "bar": _bar, "m": _matrix, "eqArr": _eqarr,
    "limLow": _limlow, "limUpp": _limupp, "lim": _lim,
    "lit": _lit,
    "borderBox": lambda el: "\\boxed{%s}" % conv(_first(el, "e")),
    "phant": lambda el: "\\phantom{%s}" % conv(_first(el, "e")),
    "t": lambda el: esc_text(el.text or ""),
}


def conv(el):
    """递归转换一个 OMML 元素为 LaTeX 串。"""
    if el is None:
        return ""                      # 畸形/缺结构：降级为空，绝不抛异常
    ln = _ln(el.tag)
    h = HANDLERS.get(ln)
    if h is not None:
        return h(el)
    if ln.endswith("Pr"):
        return ""                      # 属性容器：只在上面按需读取
    # 容器（oMath/e/num/den/sub/sup/deg/mr/box/groupChr…）：拼接子节点
    return "".join(conv(c) for c in el)


def to_latex(omath_el):
    """m:oMath → LaTeX 串（无定界符；调用方决定 inline/display 包装）。"""
    return conv(omath_el).strip()
