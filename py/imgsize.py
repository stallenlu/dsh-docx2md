#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""imgsize.py —— 读图片**自身**的像素尺寸（不依赖 docx 里的排版属性）。

为什么需要它：判断一张图「是扁条公式截图还是示意图」问的是图本身的形状，
而 docx 里记录的显示尺寸散落在三处互不相同的地方——
  * DrawingML：<wp:extent> 是 <wp:inline> 的子元素，**不是** <a:blip> 的
    祖先，沿祖先链往上永远找不到，于是尺寸算 None（实测全漏）；
  * VML <v:shape style="width:..pt;height:..pt">：正常，但只覆盖 VML；
  * VML 画布 <v:rect style="height:3742055">：裸数字没有单位。Word 在这里
    写的是 EMU（实测 3742055 与同一 group 的 294.65pt 精确相等），
    当成 px 会得出 393 万像素高的怪物尺寸。
三套口径混用，任何一套单独用都会错。图片字节里的尺寸没有歧义。

支持 PNG / JPEG / GIF / BMP / EMF / WMF（EMF、WMF 按文件头里的
物理尺寸换算出等效像素）。认不出的格式返回 None——调用方退回 XML 尺寸。
"""
import struct


def _png(b):
    if b[:8] != b"\x89PNG\r\n\x1a\n" or len(b) < 24:
        return None
    w, h = struct.unpack(">II", b[16:24])
    return (w, h) if w and h else None


def _jpeg(b):
    # 必须先认 SOI（FF D8）再扫描段标记。不认的话，任何二进制里偶然出现的
    # 0xFFC0/0xFFC2 都会被当成 SOF，把 EMF 当 JPEG 解出天文数字的尺寸
    # （实测 image163.emf 曾被读成 29679×31593）。
    if len(b) < 4 or b[:2] != b"\xff\xd8":
        return None
    i, n = 2, len(b)
    while i < n - 9:
        if b[i] != 0xFF:
            i += 1
            continue
        m = b[i + 1]
        if m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB,
                 0xCD, 0xCE, 0xCF):
            h, w = struct.unpack(">HH", b[i + 5:i + 9])
            return (w, h) if w and h else None
        if m == 0xD8 or (0xD0 <= m <= 0xD7):
            i += 2
            continue
        if i + 4 > n:
            break
        i += 2 + struct.unpack(">H", b[i + 2:i + 4])[0]
    return None


def _gif(b):
    if b[:3] != b"GIF" or len(b) < 10:
        return None
    w, h = struct.unpack("<HH", b[6:10])
    return (w, h) if w and h else None


def _bmp(b):
    if b[:2] != b"BM" or len(b) < 26:
        return None
    w, h = struct.unpack("<ii", b[18:26])
    return (abs(w), abs(h)) if w and h else None


def _emf(b):
    """ENHMETAHEADER: rclFrame（第 24 字节起，0.01mm）是设计尺寸。"""
    if len(b) < 88 or b[:4] != b"\x01\x00\x00\x00":
        return None
    l, t, r, bo = struct.unpack("<iiii", b[24:40])
    if r <= l or bo <= t:
        return None
    px = lambda v: int(round(v / 100.0 / 25.4 * 96))     # 0.01mm → px
    w, h = px(r - l), px(bo - t)
    return (w, h) if w > 0 and h > 0 else None


def _wmf(b):
    """可放置（placeable）WMF：bbox 在第 6 字节起 4 个 int16，单位 twip。"""
    if len(b) < 22 or b[:4] != b"\xd7\xcd\xc6\x9a":
        return None
    x1, y1, x2, y2 = struct.unpack("<hhhh", b[6:14])
    if x2 <= x1 or y2 <= y1:
        return None
    px = lambda v: int(round(v / 1440.0 * 96))          # twip → px
    w, h = px(x2 - x1), px(y2 - y1)
    return (w, h) if w > 0 and h > 0 else None


def read_size(blob):
    """(宽, 高) 像素；不认识的格式返回 (None, None)。"""
    if not blob:
        return (None, None)
    for fn in (_png, _jpeg, _gif, _bmp, _emf, _wmf):
        try:
            got = fn(blob)
        except Exception:
            got = None
        if got:
            return got
    return (None, None)


def load_media_sizes(z):
    """{basename: (w, h)}，把 docx 里 word/media/ 下的图全读一遍。

    一次读完比每张图再开一次 zip 便宜得多；实测 175 个文件约 40ms。
    """
    out = {}
    for name in z.namelist():
        if not name.startswith("word/media/"):
            continue
        base = name.split("/")[-1]
        if not base:
            continue
        try:
            out[base] = read_size(z.read(name))
        except Exception:
            out[base] = (None, None)
    return out
