#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""canvas.py —— 把 Word 的 VML 画布（v:group）渲染成 PNG。

为什么必须有这一层：
    流程图 / 示意图在 docx 里不是一张图，而是一个 **v:group 画布**，里面
    并排放着几十个形状：圆角矩形（步骤框）、菱形（判断）、箭头（连线）、
    文本框（标签）、以及若干 OLE 公式（框里的公式）。

    早先的实现只认 `<v:imagedata>`（真正的图片）和 OLE/OMML 公式，**所有
    非图片形状一律不看**。后果是：画布的骨架（框、菱形、箭头、标签）全部
    消失，只剩散落在正文里的公式——读者看到一串互不相干的公式，完全看不出
    那是一张流程图。而且不报错、不告警。

    文字层无法表达"这些框怎么连"，唯一忠实的表示是把整块画布画成一张图。
    所以这里把 VML 几何翻译成 SVG（形状 + 颜色 + 文字 + 内嵌图片），再用
    无头浏览器栅格化成 PNG。

坐标系：
    v:group 有 coordsize="W,H" 与 coordorigin="x,y"。子形状的 style 里
    left/top 是**画布坐标**，要先减去 coordorigin 才是 SVG 坐标。
"""
import base64
import os
import re
import shutil
import subprocess
import sys

V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W10 = "{urn:schemas-microsoft-com:office:word}"

# 这些 spt 是「画布边框」或纯图片，不算画布的结构元素。
FRAME_SPT = {"1", "75"}

# ---- o:spt → 形状语义 ----
#
# 这张表是**实测**得来的，不是照抄规范。VML 的 o:spt 编号与 MSO 的
# AutoShapeType 并不同源：同一个流程图框在 VML 里是 109，Word 对象模型
# 里报告 auto=61。直接按 MSO 常量翻译会画错形状。
#
# 取法：用 Word COM 打开真实文档，逐 GroupItem 读 AutoShapeType 与
# Left/Top/Width/Height，再按几何尺寸与 VML 子形状一一对上。实测对应：
#   spt=109 → auto=61  FlowchartProcess（矩形处理框）
#   spt=176 → auto=62  FlowchartAlternateProcess（圆角处理框）
#   spt=115 → auto=68  FlowchartMultidocument（多文档）
#   spt=116 → auto=69  FlowchartTerminator（起止圆头框）
#   spt=67  → auto=36  DownArrow（向下箭头）
#   spt=13  → auto=33  RightArrow（向右箭头）
#   spt=93  → auto=49  FlowchartOffpageConnector（折角框）
#   spt=202 → 文本框（Word 报告 auto=1）
#   spt=75  → 图片 / OLE 公式预览
LINE_SPT = {"20"}
TEXTBOX_SPT = {"202"}
PICTURE_SPT = {"75"}

# 箭头：按方向画多边形，不能画成方块
ARROW_DOWN_SPT = {"67"}
ARROW_RIGHT_SPT = {"13"}
# 圆角矩形（处理框 / 起止框 / 多文档 / 折角框）
ROUNDED_SPT = {"109", "176", "115", "116", "93"}
# 起止框是"胶囊形"：圆角半径取高度的一半
STADIUM_SPT = {"116"}


def _num(v, default=0.0):
    if v is None:
        return default
    m = re.match(r"\s*(-?[\d.]+)", str(v))
    return float(m.group(1)) if m else default


def _style_map(style):
    out = {}
    for part in (style or "").split(";"):
        if ":" in part:
            k, _, val = part.partition(":")
            out[k.strip().lower()] = val.strip()
    return out


def _style_pt(sm, key, default=None):
    """从 style 里取一个带单位的长度，换算成 pt。"""
    raw = sm.get(key)
    if raw is None:
        return default
    raw = raw.strip()
    m = re.match(r"(-?[\d.]+)\s*([a-z%]*)", raw)
    if not m:
        return default
    val, unit = float(m.group(1)), m.group(2).lower()
    if unit == "pt" or unit == "":
        return val
    if unit == "in":
        return val * 72.0
    if unit == "cm":
        return val * 72.0 / 2.54
    if unit == "mm":
        return val * 72.0 / 25.4
    if unit == "px":
        return val * 0.75
    return val


def _content_shapes(el):
    """统计画布里的"版面结构"：几何形状 / 连线 / 成组图片 + 标签。

    返回 (structural, media, has_text)。

    判定「是不是画布」不能用「有文字」——Word 经常把公式或一段内容塞进
    一个 `v:group` 壳里（文本框是最常见的容器），那既没有版面也没有图，
    整块渲染反而会把框里的公式吞掉（公式被画进一张图，正文就没有占位符
    了，自检直接报"公式丢失"）。真正的画布有**版面**：

        a) 连线 / 箭头（`v:line`）——流程图的方向
        b) 有填充色的几何形状（矩形、圆角矩形、菱形）——流程图的框
        c) 多张图片 + 文字标签——并排的图版（如六联图 + 图注）

    纯文本框壳子 a/b/c 都不满足，按普通容器处理，里面的公式照常占位。
    """
    structural = 0
    media = 0
    has_text = False
    for ch in el:
        tag = ch.tag
        if tag == V + "group":
            if is_canvas(ch):
                structural += 1
            continue
        if tag not in (V + "shape", V + "line", V + "rect", V + "oval",
                       V + "roundrect", V + "polyline", V + "curve"):
            continue
        if tag == V + "line":
            structural += 1          # (a) 连线
            continue
        txt = "".join(t.text or "" for t in ch.iter(W + "t")).strip()
        if txt:
            has_text = True
        spt = ch.get(O + "spt") or ""
        # Word 有时写 type="#_x0000_t202" 而不是 o:spt
        typ = (ch.get("type") or "").lstrip("#")
        is_textbox = spt in TEXTBOX_SPT or typ.endswith("t202")
        imgs = [i for i in ch.iter(V + "imagedata") if i.get(R + "id")]
        if spt in PICTURE_SPT or (imgs and spt not in ROUNDED_SPT):
            if imgs:
                media += 1
            continue
        if is_textbox:
            # 文本框只是标签，不算版面结构；它的文字记进 has_text
            continue
        # (b) 有填充色/描边的几何形状才是流程图的框
        if (ch.get("filled") or "t").lower() != "f" and ch.get("fillcolor"):
            structural += 1
    return structural, media, has_text


def is_canvas(el):
    """一个元素是不是需要**整块渲染**的画布。

    判据见 `_content_shapes`：有连线、有填充几何形状，或「多图 + 标签」
    的图版。单纯装内容的 `v:group` 壳（例如只放公式文本框）不算。
    """
    if el.tag != V + "group":
        return False
    structural, media, has_text = _content_shapes(el)
    if structural > 0:
        return True
    # (c) 并排图版：多张图 + 文字标签，整块渲染才能保住图注与排布
    return media >= 2 and has_text


def find_canvases(root):
    """文档里所有画布元素（文档顺序）。"""
    return [el for el in root.iter(V + "group") if is_canvas(el)]


# ---------------------------------------------------------------- SVG 生成


def _text_of(shape):
    """形状文本框里的文字，段落之间用换行分隔。"""
    paras = []
    for p in shape.iter(W + "p"):
        t = "".join(x.text or "" for x in p.iter(W + "t"))
        if t.strip():
            paras.append(t)
    return "\n".join(paras)


def _font_pt(shape):
    """文本框里最大的字号（pt）。取不到就给 10pt。"""
    best = 0.0
    for sz in shape.iter(W + "sz"):
        v = _num(sz.get(W + "val"), 0)
        if v > best:
            best = v
    return (best / 2.0) if best else 10.0


def _fill(shape):
    fc = shape.get("fillcolor")
    filled = (shape.get("filled") or "t").lower()
    if fc and filled != "f":
        return fc
    return None


def _stroke(shape):
    stroked = (shape.get("stroked") or "t").lower()
    if stroked == "f":
        return None
    sc = shape.get("strokecolor") or "#000000"
    return sc


def _esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def _text_svg(x, y, w, h, text, font, fill="#000000", bold=False, anchor="middle"):
    """把一段（可能多行的）文字居中放进矩形。"""
    lines = [ln for ln in text.split("\n")]
    if not lines:
        return ""
    lh = font * 1.25
    total = lh * len(lines)
    # 基线：垂直居中
    first = y + (h - total) / 2.0 + font
    cx = x + w / 2.0 if anchor == "middle" else x + 2
    out = []
    for i, ln in enumerate(lines):
        if not ln:
            continue
        out.append(
            '<text x="%.2f" y="%.2f" font-size="%.2f" fill="%s"%s '
            'text-anchor="%s" font-family="SimSun, \'Microsoft YaHei\', '
            'sans-serif">%s</text>'
            % (cx, first + i * lh, font, fill,
               ' font-weight="bold"' if bold else "", anchor, _esc(ln)))
    return "".join(out)


def _picture_data_uri(imagedata, rel_map, workdir):
    """v:imagedata 指向的图片，读成 data URI（SVG 内嵌，浏览器才认）。"""
    rid = imagedata.get(R + "id")
    target = rel_map.get(rid) if rid else None
    if not target:
        return None
    fname = target.split("/")[-1]
    stem, ext = os.path.splitext(fname)
    cands = []
    if ext.lower() in (".emf", ".wmf"):
        # 渲染阶段已经把 EMF/WMF 转成 png/<stem>.png
        cands.append(os.path.join(workdir, "png", stem + ".png"))
    cands.append(os.path.join(workdir, "media", fname))
    for path in cands:
        if os.path.isfile(path):
            mime = "image/png"
            low = path.lower()
            if low.endswith(".jpg") or low.endswith(".jpeg"):
                mime = "image/jpeg"
            elif low.endswith(".gif"):
                mime = "image/gif"
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            return "data:%s;base64,%s" % (mime, b64)
    return None


def _shape_geometry(shape, ox, oy):
    sm = _style_map(shape.get("style"))
    x = _num(sm.get("left")) - ox
    y = _num(sm.get("top")) - oy
    w = _num(sm.get("width"))
    h = _num(sm.get("height"))
    return x, y, w, h, sm


def canvas_svg(group, rel_map, workdir):
    """把一个 v:group 画布翻译成 SVG 字符串。

    返回 (svg, width_units, height_units)；拿不到尺寸时返回 (None, 0, 0)。
    """
    sm = _style_map(group.get("style"))
    coord = group.get("coordsize") or ""
    m = re.match(r"\s*(-?[\d.]+)\s*,\s*(-?[\d.]+)", coord)
    if m:
        cw, ch = float(m.group(1)), float(m.group(2))
    else:
        cw = _style_pt(sm, "width")
        ch = _style_pt(sm, "height")
    if not cw or not ch:
        return None, 0, 0

    origin = group.get("coordorigin") or ""
    mo = re.match(r"\s*(-?[\d.]+)\s*,\s*(-?[\d.]+)", origin)
    ox, oy = (float(mo.group(1)), float(mo.group(2))) if mo else (0.0, 0.0)

    # 画布单位 → pt 的换算（用于把字号从 pt 换成画布单位）
    w_pt = _style_pt(sm, "width")
    scale = (w_pt / cw) if w_pt else 1.0
    if scale <= 0:
        scale = 1.0

    body = []
    # 白底：Word 画布本身透明，但导出成 PNG 必须有底色，否则黑底看不清
    body.append('<rect x="0" y="0" width="%.2f" height="%.2f" fill="#FFFFFF"/>'
                % (cw, ch))

    for shape in group:
        tag = shape.tag
        if tag == V + "group":
            # 嵌套画布：把子画布当作一个整体平移到它的位置
            sub_svg, sw, sh = canvas_svg(shape, rel_map, workdir)
            if not sub_svg:
                continue
            sx, sy, _, _, _ = _shape_geometry(shape, ox, oy)
            inner = re.sub(r"^<\?xml[^>]*\?>\s*", "", sub_svg)
            inner = re.sub(r"^<svg[^>]*>", "", inner)
            inner = re.sub(r"</svg>\s*$", "", inner)
            body.append('<g transform="translate(%.2f,%.2f)">%s</g>'
                        % (sx, sy, inner))
            continue
        if tag not in (V + "shape", V + "line", V + "rect", V + "oval",
                       V + "roundrect", V + "polyline", V + "curve"):
            continue

        x, y, w, h, ssm = _shape_geometry(shape, ox, oy)
        if w <= 0 and h <= 0:
            continue
        spt = shape.get(O + "spt") or ""
        fill = _fill(shape)
        stroke = _stroke(shape)
        sw = _num(ssm.get("stroke-width"), 1.0) if stroke else 0

        # ---- 连线 ----
        if tag == V + "line" or spt in LINE_SPT:
            x2 = _num(ssm.get("margin-left"), 0)
            flip = (ssm.get("flip") or "")
            x1, y1 = x, y
            x2c, y2c = x + w, y + h
            if "x" in flip:
                x1, x2c = x2c, x1
            if "y" in flip:
                y1, y2c = y2c, y1
            body.append('<line x1="%.2f" y1="%.2f" x2="%.2f" y2="%.2f" '
                        'stroke="%s" stroke-width="%.2f"/>'
                        % (x1, y1, x2c, y2c, stroke or "#000000", max(sw, 1.0)))
            continue

        # ---- 图片 ----
        # 注意：文本框 / 自选图形里也有一个**空的** `<v:imagedata/>`（只写
        # o:title=""，没有 r:id），那是 VML 的占位符，不是图片。把它当图片
        # 处理会整块 continue，把框里的文字一并吞掉——实测流程图标题
        # "VDDD3D程序结构图" 和 "t=t+dt" 就是这样丢的。
        # 判据是「有没有真正指向媒体文件的 rId」，不是「有没有 imagedata」。
        imgs = [i for i in shape.iter(V + "imagedata") if i.get(R + "id")]
        if spt in ("75",) or (imgs and spt not in ROUNDED_SPT):
            uri = _picture_data_uri(imgs[0], rel_map, workdir) if imgs else None
            if uri:
                body.append('<image x="%.2f" y="%.2f" width="%.2f" '
                            'height="%.2f" preserveAspectRatio="none" '
                            'href="%s"/>' % (x, y, w, h, uri))
            continue

        # ---- 几何形状 ----
        stroke_attr = 'stroke="%s" stroke-width="%.2f"' % (
            stroke or "none", max(sw, 1.0))
        if spt in ARROW_DOWN_SPT:
            # 向下箭头：轴宽取 40%，箭头三角占高 40%
            ax = x + w * 0.30
            aw = w * 0.40
            hy = y + h * 0.60
            pts = ("%.2f,%.2f %.2f,%.2f %.2f,%.2f %.2f,%.2f %.2f,%.2f "
                   "%.2f,%.2f %.2f,%.2f"
                   % (ax, y, ax + aw, y, ax + aw, hy, x + w, hy,
                      x + w / 2, y + h, x, hy, ax, hy))
            body.append('<polygon points="%s" fill="%s" %s/>'
                        % (pts, fill or "none", stroke_attr))
        elif spt in ARROW_RIGHT_SPT:
            ay = y + h * 0.30
            ah = h * 0.40
            vx = x + w * 0.60
            pts = ("%.2f,%.2f %.2f,%.2f %.2f,%.2f %.2f,%.2f %.2f,%.2f "
                   "%.2f,%.2f %.2f,%.2f"
                   % (x, ay, vx, ay, vx, y, x + w, y + h / 2, vx, y + h,
                      vx, ay + ah, x, ay + ah))
            body.append('<polygon points="%s" fill="%s" %s/>'
                        % (pts, fill or "none", stroke_attr))
        elif spt in STADIUM_SPT:
            rx = h / 2.0
            body.append('<rect x="%.2f" y="%.2f" width="%.2f" height="%.2f" '
                        'rx="%.2f" ry="%.2f" fill="%s" %s/>'
                        % (x, y, w, h, rx, rx, fill or "none", stroke_attr))
        elif spt in ROUNDED_SPT:
            rx = min(w, h) * 0.10
            body.append('<rect x="%.2f" y="%.2f" width="%.2f" height="%.2f" '
                        'rx="%.2f" ry="%.2f" fill="%s" %s/>'
                        % (x, y, w, h, rx, rx, fill or "none", stroke_attr))
        elif spt in TEXTBOX_SPT:
            # 文本框自己通常无填充；有底色才画
            if fill:
                body.append('<rect x="%.2f" y="%.2f" width="%.2f" '
                            'height="%.2f" fill="%s"/>' % (x, y, w, h, fill))
        else:
            body.append('<rect x="%.2f" y="%.2f" width="%.2f" height="%.2f" '
                        'fill="%s" %s/>'
                        % (x, y, w, h, fill or "none", stroke_attr))

        # ---- 文字 ----
        text = _text_of(shape)
        if text:
            fpt = _font_pt(shape) * (1.0 / scale if scale else 1.0)
            fpt = max(fpt, ch * 0.012)      # 太小的字号在缩略图里不可读
            body.append(_text_svg(x, y, w, h, text, fpt))

    # 画布单位是 VML 坐标（常是 10000 量级），**不是像素**。SVG 的
    # viewBox 用画布单位，但 width/height 必须给 pt 尺寸——否则浏览器
    # 会按「1 单位 = 1 像素」截出一张两万像素宽的图，既慢又超限。
    h_pt = _style_pt(sm, "height")
    if not w_pt:
        w_pt = cw
    if not h_pt:
        h_pt = ch
    svg = ('<svg xmlns="http://www.w3.org/2000/svg" '
           'xmlns:xlink="http://www.w3.org/1999/xlink" '
           'data-docx2md="canvas" '
           'width="%.2fpt" height="%.2fpt" viewBox="0 0 %.2f %.2f">%s</svg>'
           % (w_pt, h_pt, cw, ch, "".join(body)))
    return svg, w_pt, h_pt


# ---------------------------------------------------------------- 栅格化


def _browsers():
    """可用的无头浏览器（按优先级）。"""
    cands = []
    for name in ("msedge", "microsoft-edge", "google-chrome", "chromium",
                 "chromium-browser", "chrome"):
        p = shutil.which(name)
        if p:
            cands.append(p)
    for p in (
        "/mnt/c/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
        "/mnt/c/Program Files/Microsoft/Edge/Application/msedge.exe",
        "/mnt/c/Program Files/Google/Chrome/Application/chrome.exe",
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    ):
        if os.path.isfile(p):
            cands.append(p)
    return cands


def _to_windows_path(path):
    """WSL 路径 → Windows 路径（浏览器是 Windows 程序，只认 C:\\…）。"""
    if not path.startswith("/"):
        return path
    for prefix, drive in (("/mnt/c/", "C:\\"), ("/mnt/d/", "D:\\"),
                          ("/mnt/e/", "E:\\")):
        if path.startswith(prefix):
            return drive + path[len(prefix):].replace("/", "\\")
    try:
        out = subprocess.run(["wslpath", "-w", path], capture_output=True,
                             text=True, timeout=20)
        if out.returncode == 0 and out.stdout.strip():
            return out.stdout.strip()
    except Exception:                                     # noqa: BLE001
        pass
    return path


def _file_url(path):
    """本地路径 → 浏览器可打开的 file:// URL。

    WSL 下这里有个必须踩一次的坑：画布文件在 Linux 侧，浏览器是 Windows
    程序，只能读 `\\\\wsl.localhost\\Ubuntu\\...`。直接把反斜杠换成斜杠再拼
    `file:///` 会得到 `file://////wsl.localhost/...`（五个斜杠），Edge 认不出，
    于是截下来的 PNG 是一张「未找到文件 / ERR_FILE_NOT_FOUND」错误页——
    **文件非空、尺寸正确、退出码 0，肉眼不看图根本发现不了**。

    正确形式：
        盘符  C:\\a\\b   → file:///C:/a/b
        UNC   \\\\host\\share\\x → file://host/share/x   （两斜杠，不是五斜杠）
    """
    import urllib.parse
    win = _to_windows_path(path)
    if win.startswith("\\\\"):
        # UNC：去掉前导两反斜杠，主机名放在 authority 位置
        rest = win[2:].replace("\\", "/")
        return "file://" + urllib.parse.quote(rest, safe="/:")
    drive = win.replace("\\", "/")
    return "file:///" + urllib.parse.quote(drive, safe="/:")


def rasterize(svg, out_png, width_pt, height_pt, scratch_dir=None):
    """SVG → PNG。成功返回 True。

    width_pt/height_pt 是画布的**物理尺寸**（pt）；截图按 2 倍超采样。

    中间 SVG 写在 `scratch_dir`（默认是 PNG 所在目录），**用完即删**。
    不碰系统临时目录：本工具全程只在 outDir 内落文件，这样才能保证
    「一次转换的产物全部可追溯、可删除」，也才能过 hygiene 检查。
    """
    browsers = _browsers()
    if not browsers:
        return False, "没有找到无头浏览器（Edge / Chrome）"
    # 必须 abspath：outDir 可能是相对路径（用户 `docx2md_extract a.docx out`
    # 就这么传）。相对路径进 _to_windows_path 会原样返回，拼出来的
    # file:// URL 是坏的，浏览器打开的是错误页——而错误页照样能截成 PNG。
    tmpdir = os.path.abspath(scratch_dir or os.path.dirname(out_png))
    if not os.path.isdir(tmpdir):
        os.makedirs(tmpdir, exist_ok=True)
    svg_path = os.path.join(tmpdir, "_canvas_tmp.svg")
    with open(svg_path, "w", encoding="utf-8") as f:
        f.write(svg)
    # 截图尺寸按 2 倍超采样，正文里的图更清晰。
    # 关键：窗口尺寸是 **CSS 像素**，超采样靠 --force-device-scale-factor，
    # 不能把 2 倍直接乘进 window-size（那样 SVG 只占左上四分之一）。
    scale = 2
    w = max(1, int(round(width_pt * 96.0 / 72.0)))
    h = max(1, int(round(height_pt * 96.0 / 72.0)))
    url = _file_url(svg_path)
    target = _to_windows_path(os.path.abspath(out_png))
    ok = False
    why = "浏览器没能加载画布 SVG（截图是错误页）"
    try:
        for exe in browsers:
            cmd = [exe, "--headless=new", "--disable-gpu", "--no-sandbox",
                   "--hide-scrollbars", "--default-background-color=FFFFFFFF",
                   "--force-device-scale-factor=%d" % scale,
                   "--screenshot=%s" % target,
                   "--window-size=%d,%d" % (w, h),
                   "--dump-dom", url]
            try:
                proc = subprocess.run(cmd, capture_output=True, timeout=120)
            except Exception:                             # noqa: BLE001
                continue
            # 成功判据不能只看「PNG 存在且非空」：文件路径拼错时 Edge 会把
            # 一张「未找到文件 / ERR_FILE_NOT_FOUND」错误页截成 PNG，
            # 尺寸、字节数、退出码全都正常。必须确认它真的加载了我们的 SVG。
            dom = (proc.stdout or b"").decode("utf-8", "replace")
            if "data-docx2md" not in dom:
                continue
            if os.path.isfile(out_png) and os.path.getsize(out_png) > 0:
                ok = True
                why = ""
                break
    finally:
        try:
            os.remove(svg_path)
        except OSError:
            pass
    return ok, why


def render_canvas(group, rel_map, workdir, out_png):
    """画布 → PNG。返回 (ok, reason)。"""
    svg, cw, ch = canvas_svg(group, rel_map, workdir)
    if not svg:
        return False, "画布没有尺寸信息"
    return rasterize(svg, out_png, cw, ch, os.path.dirname(out_png))


def main():
    """自检：--probe 报告浏览器可用性。"""
    if "--probe" in sys.argv:
        bs = _browsers()
        print("浏览器: %s" % (bs[0] if bs else "未找到"))
        return 0 if bs else 3
    print(__doc__)
    return 0


if __name__ == "__main__":
    sys.exit(main())
