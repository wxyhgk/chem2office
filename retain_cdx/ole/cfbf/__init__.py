"""写出一个合法的复合文件（CFBF / MS-CFB），与内容无关。

CFBF 是「一个文件里的迷你文件系统」：512 字节扇区、一张 FAT 把扇区串成链、
一棵目录树记录每个流的名字与起点。本包只管这套格式，不认识 ChemDraw——
放哪些流、根 CLSID 是什么由调用方决定（见 `ole/chemdraw.py`）。

为什么自己写：olefile 只能读不能建，而 Office 对容器合法性极其挑剔——
FAT 常量写错的后果是 Excel 打开时「修复」文件并把 OLE 对象整个删掉。

模块划分对应格式自身的结构::

    const.py      格式词汇（扇区大小、特殊扇区号）
    layout.py     规划：谁住 mini 流、各段几个扇区、起点在哪
    alloc.py      FAT / mini-FAT
    directory.py  目录树与 128 字节条目
    header.py     512 字节文件头

细节见 `ole/README.md`。
"""

from __future__ import annotations

from io import BytesIO

from .alloc import build_fat, build_minifat
from .const import (
    ENDOFCHAIN,
    FREESECT,
    MINI_SECTOR_SIZE,
    MINI_STREAM_CUTOFF,
    SECTOR_SIZE,
    ZERO_CLSID,
)
from .directory import build_directory
from .header import build_header
from .layout import Layout, plan


def build_compound_file(
    streams: list[tuple[str, bytes]], *, root_clsid: bytes = ZERO_CLSID
) -> bytes:
    """把 ``(流名, 字节)`` 列表写成一个完整的复合文件。"""
    layout = plan(streams)

    out = BytesIO()
    out.write(build_header(layout))
    out.write(build_fat(layout))
    out.write(build_directory(layout, root_clsid=root_clsid))
    out.write(build_minifat(layout))
    out.write(layout.mini_stream)
    for name, _data in layout.regular_defs:
        out.write(layout.regular_bytes[name])

    written = out.tell()
    expected = (layout.total_sectors + 1) * SECTOR_SIZE
    assert written == expected, f"sector layout mismatch: {written} != {expected}"
    return out.getvalue()


__all__ = [
    "ENDOFCHAIN",
    "FREESECT",
    "Layout",
    "MINI_SECTOR_SIZE",
    "MINI_STREAM_CUTOFF",
    "SECTOR_SIZE",
    "ZERO_CLSID",
    "build_compound_file",
]
