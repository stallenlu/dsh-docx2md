#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""imgformula.py —— 识别「以图片形式存在的公式」，只做候选检测，不做决定。

背景：docx 里的公式有三种载体，此前工具只认前两种——
  1. OLE 对象（MTEF 二进制）      → 渲染预览图 + 读图转写
  2. OMML（Word 原生 XML）        → 确定性转 LaTeX
  3. **图片**（截图/扫描/PDF 粘贴） → 旧版：正文里静默消失，不报错

图片公式没有结构可解、没有 MTEF 可查，像素就是全部内容。工具无法「转换」它，
能做的只有两件事：**绝不丢**（生成占位符，build 落成可见的 HTML 注释），
以及**如实报告**（按形状启发式给出候选清单与上下文，判断交给人/Agent）。

形状启发式（EMU 尺寸）：
  * 宽高比 >= 2.5 且 高 <= 120px   → formula 候选——公式截图天然是扁条状。
    独立成段的扁条也算候选（显示公式截图+编号是常态）；是否行内、有无
    文字夹持作为证据一并给出，降级判断归调用方。
  * 其余（方形、大图）              → illustration——示意图不进正文流，只进清单
  * 读不到尺寸                      → illustration（宁可漏报走清单留痕，
    不误报成公式逼人工逐条确认）
  OLE 预览图整体排除（两个检查并跑，缺一个都实测漏过）：
    ShapeID / o:ole="t"（VML 预览，含与 w:object 平级 run 的写法）；
    w:object 容器嵌套（DrawingML 包裹的预览没有 v:shape，只能靠这个）。

候选不等于结论：每个候选都带 ratio / 尺寸 / 上下文文字，最终判断由调用方做。
"""
import re

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"

EMU_PER_PX = 9525
FORMULA_MIN_RATIO = 2.5      # 扁条判定
FORMULA_MAX_H_PX = 120       # 行内图片高度上限（px）


def is_ole_preview(node, ole_ids, container_skip=None):
    """这个 drawing/pict 容器是不是某个 OLE 公式的预览图。

    两个互补检查，缺一个都会漏（都实测过）：
    * ShapeID / o:ole="t"：VML 预览图与公式的可靠联系。实测有文档把预览
      shape 放在与 w:object 平级的 run 里，只按容器嵌套排除会漏。
    * container_skip（w:object 内部的元素 id 集）：DrawingML 包裹的 OLE
      预览没有 v:shape，只能靠"整个在 w:object 里"来认。
    """
    for sh in node.iter(V + "shape"):
        if sh.get(O + "ole") == "t":
            return True
        if sh.get("id") in ole_ids:
            return True
    if container_skip and any(id(d) in container_skip for d in node.iter()):
        return True
    return False


def ole_shape_ids(body):
    """所有 OLEObject 的 ShapeID。"""
    return {o.get("ShapeID") for o in body.iter(W + "OLEObject") if o.get("ShapeID")}


def ole_preview_skip(body):
    """w:object 内部的所有元素 id（DrawingML 包裹的 OLE 预览靠它排除）。"""
    skip = set()
    for obj in body.iter(W + "object"):
        for d in obj.iter():
            skip.add(id(d))
    return skip


def _emu_len(text):
    """VML style 里的长度 → EMU。支持 pt/cm/mm/in/px；裸数字按 px。"""
    m = re.search(r"(\d+(?:\.\d+)?)\s*(pt|cm|mm|in|px)?", text)
    if not m:
        return None
    v = float(m.group(1))
    unit = m.group(2) or "px"
    per = {"pt": 12700, "cm": 360000, "mm": 36000, "in": 914400, "px": EMU_PER_PX}
    return int(v * per[unit])


def picture_size(run_el):
    """(宽, 高) EMU，取不到返回 (None, None)。"""
    ext = run_el.find(".//" + WP + "extent")
    if ext is not None:
        try:
            return int(ext.get("cx")), int(ext.get("cy"))
        except (TypeError, ValueError):
            pass
    for sh in run_el.iter(V + "shape"):
        st = sh.get("style") or ""
        mw = re.search(r"width\s*:\s*([^;]+)", st)
        mh = re.search(r"height\s*:\s*([^;]+)", st)
        if mw and mh:
            w, h = _emu_len(mw.group(1)), _emu_len(mh.group(1))
            if w and h:
                return w, h
    return None, None


def picture_rid(run_el):
    # blip 在 a（drawingml/main）命名空间，不在 pic——按 pic 找会永远 None，
    # 导致所有 DrawingML 图片漏检（实测踩过）。
    b = run_el.find(".//" + A + "blip")
    if b is not None and b.get(R + "embed"):
        return b.get(R + "embed")
    for im in run_el.iter(V + "imagedata"):
        rid = im.get(R + "id")
        if rid:
            return rid
    return None


def classify_picture(w_emu, h_emu):
    """formula（公式候选）/ illustration（按形状更像插图）。"""
    if not w_emu or not h_emu:
        return "illustration"
    ratio = w_emu / h_emu
    h_px = h_emu / EMU_PER_PX
    if ratio >= FORMULA_MIN_RATIO and h_px <= FORMULA_MAX_H_PX:
        return "formula"
    return "illustration"


def local_name(tag):
    """命名空间限定的 tag → 本地名。"""
    return tag.split("}")[-1] if "}" in tag else tag


def para_text(p):
    return "".join(t.text or "" for t in p.iter(W + "t")).strip()
