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
  * 其余（方形、大图）              → illustration——示意图进正文（见 extract）
  * 读不到尺寸                      → illustration（宁可漏报走清单留痕，
    不误报成公式逼人工逐条确认）

判形状用**图片自身像素尺寸**（imgsize.read_size），不用 docx 排版属性：
后者三套口径互不兼容（DrawingML 的 wp:extent 不在 blip 祖先链上；VML
<v:shape style> 只覆盖 VML；VML 画布的裸数字是 EMU 不是 px），实测把
450px 高的三维图算成 2640px、把 3742055EMU 当成 3742055px。
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
MIN_FIGURE_H_PX = 8          # 低于此高度的图是占位块/分隔条，不是插图
MIN_FIGURE_W_PX = 8


def is_ole_preview(node, ole_ids, container_skip=None):
    """这个 drawing/pict 容器是不是某个 OLE 公式的预览图。

    两个互补检查，缺一个都会漏（都实测过）：
    * ShapeID / o:ole="t"：VML 预览图与公式的可靠联系。实测有文档把预览
      shape 放在与 w:object 平级的 run 里，只按容器嵌套排除会漏。
    * container_skip（w:object 内部的元素 id 集）：DrawingML 包裹的 OLE
      预览没有 v:shape，只能靠"整个在 w:object 里"来认。

    但**容器自己装着 OLE 公式时，以上检查一律不算预览图**。
    实测踩过：Word 画流程图时，公式可以嵌在同一个 w:pict 的 v:group 里，
    于是该 pict 内既有公式的 o:ole="t" 预览 shape，又有 o:OLEObject 本身。
    若按「见 o:ole="t" 就算预览」判断，整棵 pict 被当成一张预览图跳过，
    里面的公式占位符**全部消失**——正文照常输出，不报错不告警。
    这正是本工具最忌讳的那类静默丢失。
    """
    # 容器自己就是公式的宿主（o:OLEObject 在里面）→ 它不是预览图。
    if next(node.iter(O + "OLEObject"), None) is not None:
        return False
    for sh in node.iter(V + "shape"):
        if sh.get(O + "ole") == "t":
            return True
        if sh.get("id") in ole_ids:
            return True
    if container_skip and any(id(d) in container_skip for d in node.iter()):
        return True
    return False


def holds_ole_formula(node):
    """这个容器里是否直接装着 OLE 公式（o:OLEObject）。

    用来区分两种长得很像的 w:pict：
    * 公式的**预览图**（只有 v:shape + o:ole="t"，没有 o:OLEObject）
    * **装着公式的画布**（Word 流程图：v:group 里既有预览 shape，
      也有 o:OLEObject 本身）

    两者混为一谈，公式就会从正文里静默消失——见 is_ole_preview 的说明。
    """
    return next(node.iter(O + "OLEObject"), None) is not None


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


def _leaf_size(leaf, boundary):
    """从一张图（blip / v:imagedata）往上找它**自己**的尺寸。

    不能对整个容器调 picture_size——一个容器里可以并排放十几张图（实测：
    145 公式文档的 p281 有 6 张三维图 + 6 个公式预览共 12 张），容器级
    尺寸只会给出第一张，其余十张的宽高全部张冠李戴，逼得启发式把示意图
    误判成扁条公式。
    """
    el = leaf
    while el is not None:
        name = local_name(el.tag)
        if name == "shape":
            st = el.get("style") or ""
            mw = re.search(r"width\s*:\s*([^;]+)", st)
            mh = re.search(r"height\s*:\s*([^;]+)", st)
            if mw and mh:
                w, h = _emu_len(mw.group(1)), _emu_len(mh.group(1))
                if w and h:
                    return w, h
        elif name in ("extent", "ext"):
            try:
                return int(el.get("cx")), int(el.get("cy"))
            except (TypeError, ValueError):
                pass
        if el is boundary:
            break
        el = _parent.get(el) if _parent else None
    return None, None


# 父表由调用方（extract.py）注入一次，避免每张图都重建整棵树的索引。
_parent = None


def set_parent_map(parents):
    """注入 {子元素: 父元素}，供 _leaf_size 向上走。"""
    global _parent
    _parent = parents


def leaf_is_ole_preview(leaf, ole_ids, boundary=None, container_skip=None):
    """这一张图是不是某个 OLE 公式的预览（往上走，只看到 boundary 为止）。

    为什么必须逐张判、不能按整个容器判：Word 画「三维图 + 公式」混排的画布
    时，公式预览（v:shape o:ole="t"）和正文插图共处同一个 w:pict。容器级
    判断只能给出「这容器里有公式」这一个结论，既不能据它跳过整棵（插图
    全丢），也不能据它保留整棵（把公式预览当插图塞进正文）。实测 145 公式
    文档的 p281：6 张三维图 + 6 个公式预览在同一个 pict 里。
    """
    el = leaf
    while el is not None and el is not boundary:
        if local_name(el.tag) == "shape":
            if el.get(O + "ole") == "t" or el.get("id") in ole_ids:
                return True
        if container_skip and id(el) in container_skip:
            return True
        el = _parent.get(el) if _parent else None
    return False


def picture_leaves(container):
    """容器里的**每一张**图（按出现顺序，同 rId 只算一次）。

    返回 [{"rid", "w", "h"}]。同 rId 重复出现是常态：Word 的
    mc:AlternateContent 会同时写 Choice（DrawingML）与 Fallback（VML）
    两个分支指向同一份媒体，文本框里的副本还会再来一份。不按 rId 去重，
    同一张图会在正文里出现三次。
    """
    out, seen = [], set()
    for leaf in list(container.iter(V + "imagedata")) + \
            list(container.iter(A + "blip")):
        rid = leaf.get(R + "id") or leaf.get(R + "embed")
        if not rid or rid in seen:
            continue
        seen.add(rid)
        w, h = _leaf_size(leaf, container)
        out.append({"el": leaf, "rid": rid, "w": w, "h": h})
    return out


def classify_picture(w_emu, h_emu):
    """formula（公式候选）/ illustration（按形状更像插图）。"""
    if not w_emu or not h_emu:
        return "illustration"
    ratio = w_emu / h_emu
    h_px = h_emu / EMU_PER_PX
    if ratio >= FORMULA_MIN_RATIO and h_px <= FORMULA_MAX_H_PX:
        return "formula"
    return "illustration"


def classify_px(w_px, h_px):
    """按图片**自身**像素尺寸判定形状（首选口径，理由见 imgsize.py）。

    w_px/h_px 来自 imgsize.read_size，不受 docx 排版属性三套口径的干扰。
    取不到时返回 None，调用方退回 classify_picture(EMU)。
    """
    if not w_px or not h_px:
        return None
    if (w_px / h_px) >= FORMULA_MIN_RATIO and h_px <= FORMULA_MAX_H_PX:
        return "formula"
    return "illustration"


def is_illustration_px(w_px, h_px):
    """像素口径：够大、不是排版占位条，才值得作为正文插图。

    image1.png 这种 2×2 的 125 字节空图（标题下的排版块）必须排除——
    落成插图只会插入一行破图。
    """
    if not w_px or not h_px:
        return False
    return h_px >= MIN_FIGURE_H_PX and w_px >= MIN_FIGURE_W_PX


def is_illustration_emu(w_emu, h_emu):
    """这张图是否值得作为正文插图落进 md。

    排除两类不是插图的东西：
    * 读不到尺寸的（无法排版，也无从判断，宁可留清单不猜）
    * 高度 < MIN_FIGURE_H_PX 的像素条——实测真实文档 p0 标题下有一张
      2px 高的空图（image1.png，125 字节），它是排版占位块，落成插图
      只会插入一行破图。
    """
    if not w_emu or not h_emu:
        return False
    return (h_emu / EMU_PER_PX) >= MIN_FIGURE_H_PX


def local_name(tag):
    """命名空间限定的 tag → 本地名。"""
    return tag.split("}")[-1] if "}" in tag else tag


def para_text(p):
    return "".join(t.text or "" for t in p.iter(W + "t")).strip()
