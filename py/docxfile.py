#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docxfile.py —— .docx 打开与结构校验，供其它脚本共用。

为什么单独抽出来：坏输入（非 zip、空 zip、缺 word/document.xml、加密文档）
如果直接交给 zipfile / ElementTree，会抛出原始异常栈，对使用者没有意义。
这里统一把它们转成一句可读的错误，并以固定退出码 3 退出。
"""
import sys
import zipfile
from xml.etree import ElementTree as ET

# 坏输入的专用退出码，便于调用方区分「文档有问题」与「程序有 bug」。
EXIT_BAD_INPUT = 3

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


class BadDocument(Exception):
    """输入文件不是可处理的 .docx。"""


def force_utf8_stdio():
    """让 stdout/stderr 以 UTF-8 输出。

    中文 Windows 上 sys.stdout.encoding 默认是 gbk，非 ASCII 字符会被编成 GBK
    字节；调用方按 UTF-8 解码就得到「����」这样的乱码，而且是静默的。本插件的
    输出里有中文段落（missing_preview_context.paragraph_text），乱码会让下游判断
    失效，所以这里主动钉成 UTF-8，不依赖调用方设环境变量。

    Python 3.7+ 才有 reconfigure；更老的版本跳过（那时也只能靠环境变量）。
    """
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8")
            except (ValueError, OSError):
                pass  # 已被重定向成不可重配的对象，交给环境变量兜底


def fail(message):
    """打印一句可读错误到 stderr 并退出。"""
    force_utf8_stdio()
    sys.stderr.write("docx 输入错误: %s\n" % message)
    sys.exit(EXIT_BAD_INPUT)


def read_rels(z, part="word/_rels/document.xml.rels", include_external=False):
    """rId -> Target. Read OOXML attributes, not attribute order.

    The old regex Id="..."[^>]*Target="..." required Id before Target;
    when Word writes them in reverse, every relationship silently vanished
    (formula preview images and hyperlinks). External targets are skipped
    by default; rich.py passes include_external=True for w:hyperlink URLs.
    """
    try:
        raw = z.read(part)
    except KeyError:
        return {}
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return {}
    out = {}
    for rel in root:
        if rel.tag.split("}")[-1] != "Relationship":
            continue
        if not include_external and rel.get("TargetMode") == "External":
            continue
        rid, tgt = rel.get("Id"), rel.get("Target")
        if rid and tgt:
            out[rid] = tgt
    return out


def open_docx(path):
    """打开 .docx 并返回 (zipfile, document_root)。

    失败时抛 BadDocument，调用方用 fail() 转成可读错误。
    """
    try:
        z = zipfile.ZipFile(path)
    except FileNotFoundError:
        raise BadDocument("文件不存在: %s" % path)
    except IsADirectoryError:
        raise BadDocument("给的是目录而不是文件: %s" % path)
    except zipfile.BadZipFile:
        raise BadDocument(
            "不是有效的 .docx（内部不是 zip 结构）。"
            "旧版 .doc 请先用 Word/WPS 另存为 .docx。路径: %s" % path)
    except OSError as exc:
        raise BadDocument("无法读取 %s: %s" % (path, exc))

    names = z.namelist()
    if not names:
        z.close()
        raise BadDocument("文件是空的 zip 容器: %s" % path)

    # 解压炸弹防护：只看 ZipInfo 元数据，不实际解压，成本 O(条目数)。
    # 单文件 / 压缩比 / 全包总量三道闸，任一超限即拒并关掉句柄。
    MAX_ONE = 64 * 1024 * 1024
    MAX_RATIO = 200
    MAX_TOTAL = 256 * 1024 * 1024
    total = 0
    for info in z.infolist():
        ratio = (info.file_size / info.compress_size
                 if info.compress_size > 0 else 0)
        total += info.file_size
        if info.file_size > MAX_ONE or ratio > MAX_RATIO or total > MAX_TOTAL:
            z.close()
            raise BadDocument("压缩包疑似解压炸弹（单文件 %d 字节 / 压缩比 %d:1 / 累计 %d 字节）: %s" % (info.file_size, round(ratio), total, path))

    if "word/document.xml" not in names:
        # 加密的 .docx 会呈现这个特征（只有 EncryptionInfo 之类）
        encrypted = any(n.startswith("EncryptionInfo") or n == "EncryptedPackage"
                        for n in names)
        z.close()
        if encrypted:
            raise BadDocument(
                "文档是加密的，无法读取。请先在 Word/WPS 里去掉密码后再转换。")
        raise BadDocument(
            "压缩包内缺少 word/document.xml，不是 Word 文档"
            "（可能是 .xlsx/.pptx，或已损坏）。路径: %s" % path)

    try:
        root = ET.fromstring(z.read("word/document.xml"))
    except ET.ParseError as exc:
        z.close()
        raise BadDocument("word/document.xml 解析失败，文档可能已损坏: %s" % exc)

    body = root.find(W + "body")
    if body is None:
        z.close()
        raise BadDocument("word/document.xml 里没有 body 节点，文档结构异常。")

    return z, root

def oid_of(obj, rel_map, R):
    """OLEObject 的稳定键：embedding 文件名（extract 与 inspect 同口径）。"""
    return (rel_map.get(obj.get(R + "id")) or "?").split("/")[-1]


def pair_ole_previews(body, rel_map, W, R, V, O):
    """统一「OLE 公式 → 预览图」配对：inspect 与 extract 共用同一把尺子。

    主路径按 ShapeID 精确配对：公式的预览图是带 o:ole="t" 的 v:shape，
    其 id 等于 OLEObject 的 ShapeID，与出现顺序无关、不受同段落插图干扰。
    ShapeID 缺失或对不上时退回「同段第 k 个公式 ↔ 第 k 张 v:imagedata」，
    并把结果标为 ordinal（未经确认）。实测 145 公式文档里纯下标配对把 10 个
    公式配成了 VDDD3D 程序截图 / 位错环三维图，而且不报错、不告警。

    以前 inspect（docx_scan.py）与 extract.py 各写一份配对逻辑、口径会分叉：
    同一个文档，两个工具报出的预览图数量 / 格式可能不一致且都不报错。
    抽到这里共用，口径必然一致。

    返回 (pairs, pairing, ole_shape_id)：
      pairs[oid]        = "media/image49.wmf" 或 None
      pairing[oid]      = "shape-id" | "ordinal" | "none"
      ole_shape_id[oid] = OLEObject@ShapeID（供下游判「是否在画布内」）
    """
    shape_preview = {}
    for p in body.iter(W + "p"):
        for sh in p.iter(V + "shape"):
            if sh.get(O + "ole") != "t":
                continue
            ims = list(sh.iter(V + "imagedata"))
            if ims:
                shape_preview[sh.get("id")] = rel_map.get(ims[0].get(R + "id"))

    pairs, pairing, ole_shape_id = {}, {}, {}
    for p in body.iter(W + "p"):
        imgs = [i.get(R + "id") for i in p.iter(V + "imagedata")]
        for k, obj in enumerate(p.iter(O + "OLEObject")):
            oid = oid_of(obj, rel_map, R)
            sid = obj.get("ShapeID")
            if sid:
                ole_shape_id[oid] = sid
            target = shape_preview.get(sid)
            if target:
                pairs[oid], pairing[oid] = target, "shape-id"
            elif k < len(imgs):
                pairs[oid], pairing[oid] = rel_map.get(imgs[k]) or None, "ordinal"
            else:
                pairs[oid], pairing[oid] = None, "none"
    return pairs, pairing, ole_shape_id

