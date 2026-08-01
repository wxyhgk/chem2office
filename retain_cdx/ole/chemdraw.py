"""一个 ChemDraw OLE 对象由哪些流构成。

这里是「ChemDraw 是什么」的全部知识；容器怎么写在 `cfbf.py`，那边不认识化学。

三个流的字节与 Windows Excel + ChemDraw 插入产出的模板**逐字节一致**，
CLSID 与 ProgID 在 Excel / PowerPoint 两侧通用：

``\\x01Ole``
    20 字节固定内容，声明这是嵌入对象（不是链接）。
``\\x01CompObj``
    双击时该启动谁：用户可见类型名、剪贴板格式、ProgID。
``CONTENTS``
    就是 ``.cdx`` 文件的原始字节，一字不差。
"""

from __future__ import annotations

import struct

from .cfbf import build_compound_file

# ChemDraw.Document 的 CLSID，跨 ChemDraw 版本稳定。
# 按 MS-DTYP §2.3.4.2 混合端序存放：前三段小端，后两段大端。
CHEMDRAW_CLSID = (
    struct.pack("<IHH", 0x41BA6D21, 0xA02E, 0x11CE)
    + bytes([0x8F, 0xD9, 0x00, 0x20, 0xAF, 0xD1, 0xF2, 0x0C])
)

# "6.0" 不是应用版本号——v6 以来所有 ChemDraw 都注册这个固定名字。
PROG_ID = "ChemDraw.Document.6.0"
USER_TYPE = "CS ChemDraw Drawing"
CLIPBOARD_FORMAT = "ChemDraw Interchange Format"

CONTENTS_STREAM = "CONTENTS"

_OLE_STREAM = struct.pack("<IIIII", 0x02000001, 0, 0, 0, 0)
_UNICODE_MARKER = 0x71B239F4


def _compobj_stream() -> bytes:
    """``\\x01CompObj``：告诉 Windows 这个对象归谁。"""
    buf = bytearray()
    buf += struct.pack("<HHI", 0x0001, 0xFFFE, 0x0A03)  # 版本 / 端序 / 操作系统
    buf += b"\xff\xff\xff\xff"
    buf += CHEMDRAW_CLSID
    for text in (USER_TYPE, CLIPBOARD_FORMAT, PROG_ID):
        payload = text.encode("ascii") + b"\x00"
        buf += struct.pack("<I", len(payload)) + payload
    # 与模板一致的尾部：Unicode 段标记 + 三个空 Unicode 串
    buf += struct.pack("<I", _UNICODE_MARKER)
    buf += b"\x00" * 12
    return bytes(buf)


def build_chemdraw_ole(
    cdx_data: bytes, *, extra_streams: dict[str, bytes] | None = None
) -> bytes:
    """从零造一个内含 *cdx_data* 的 ChemDraw OLE 容器（`.bin`）。

    *extra_streams* 原样附加额外的流。PowerPoint 的原生容器还带一个
    ``\\x02OlePres000`` 演示流（缓存的 EMF 渲染）；我们不复制陈旧的那份，
    但需要时可以从这里放进去。
    """
    streams: list[tuple[str, bytes]] = [
        ("\x01Ole", _OLE_STREAM),
        ("\x01CompObj", _compobj_stream()),
        (CONTENTS_STREAM, cdx_data),
    ]
    streams.extend((extra_streams or {}).items())
    return build_compound_file(streams, root_clsid=CHEMDRAW_CLSID)


__all__ = ["CHEMDRAW_CLSID", "CONTENTS_STREAM", "PROG_ID", "build_chemdraw_ole"]
