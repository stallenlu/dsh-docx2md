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
        raise BadDocument("文件是空的 zip 容器: %s" % path)

    if "word/document.xml" not in names:
        # 加密的 .docx 会呈现这个特征（只有 EncryptionInfo 之类）
        encrypted = any(n.startswith("EncryptionInfo") or n == "EncryptedPackage"
                        for n in names)
        if encrypted:
            raise BadDocument(
                "文档是加密的，无法读取。请先在 Word/WPS 里去掉密码后再转换。")
        raise BadDocument(
            "压缩包内缺少 word/document.xml，不是 Word 文档"
            "（可能是 .xlsx/.pptx，或已损坏）。路径: %s" % path)

    try:
        root = ET.fromstring(z.read("word/document.xml"))
    except ET.ParseError as exc:
        raise BadDocument("word/document.xml 解析失败，文档可能已损坏: %s" % exc)

    body = root.find(W + "body")
    if body is None:
        raise BadDocument("word/document.xml 里没有 body 节点，文档结构异常。")

    return z, root
