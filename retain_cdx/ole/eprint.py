"""Word 容器多出来的两个流：`\\x03EPRINT` 与 `\\x03ObjInfo`。

Windows Word + ChemDraw 插入的容器比 Excel 的多两个流，比 PowerPoint 的
少一个 `\\x02OlePres000`（见 `docs/word_ole_research.md`）::

    \\x01CompObj    126 B   与 Excel / PowerPoint 逐字节相同
    \\x01Ole         20 B   同上
    \\x03EPRINT    7624 B   **与 word/media/image1.emf 逐字节相同**
    \\x03ObjInfo      6 B   00 00 03 00 0d 00
    CONTENTS      1754 B   .cdx 原始字节

``\\x03EPRINT``
    Word 的渲染缓存，就是预览 EMF 原样放进来——不像 PowerPoint 的演示流
    那样要把 EMF 藏进 WMF 注释记录、算校验和。所以只有 EMF 预览才挂得上。
``\\x03ObjInfo``
    Word 自己的对象信息（MS-DOC ODT 结构）。6 字节，照抄样本。
"""

from __future__ import annotations

EPRINT_STREAM = "\x03EPRINT"
OBJINFO_STREAM = "\x03ObjInfo"

# 与 Windows Word 样本逐字节一致（有测试盯着）。
OBJINFO = bytes.fromhex("000003000d00")


def word_streams(emf: bytes | None) -> dict[str, bytes]:
    """Word 容器在 CONTENTS 之外要带的流；*emf* 为 ``None`` 时不挂缓存。"""
    streams = {OBJINFO_STREAM: OBJINFO}
    if emf is not None:
        streams[EPRINT_STREAM] = emf
    return streams


__all__ = ["EPRINT_STREAM", "OBJINFO", "OBJINFO_STREAM", "word_streams"]
