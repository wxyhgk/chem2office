"""512 字节文件头。"""

from __future__ import annotations

import struct

from .const import (
    ENDOFCHAIN,
    FREESECT,
    HEADER_DIFAT_SLOTS,
    MINI_STREAM_CUTOFF,
    SECTOR_SIZE,
    SIGNATURE,
)
from .layout import Layout


def build_header(layout: Layout) -> bytes:
    header = bytearray(SECTOR_SIZE)
    header[0:8] = SIGNATURE
    struct.pack_into("<HH", header, 24, 0x003E, 0x0003)  # 次版本 / 主版本 v3
    struct.pack_into("<H", header, 28, 0xFFFE)           # 小端
    struct.pack_into("<H", header, 30, 9)                # 512 字节扇区
    struct.pack_into("<H", header, 32, 6)                # 64 字节 mini 扇区
    struct.pack_into("<I", header, 40, 0)                # v3 里目录扇区数必须为 0
    struct.pack_into("<I", header, 44, layout.n_fat)
    struct.pack_into("<I", header, 48, layout.dir_start)
    struct.pack_into("<I", header, 52, 0)                # 事务签名
    struct.pack_into("<I", header, 56, MINI_STREAM_CUTOFF)
    struct.pack_into(
        "<I", header, 60, layout.minifat_start if layout.n_minifat else ENDOFCHAIN
    )
    struct.pack_into("<I", header, 64, layout.n_minifat)
    struct.pack_into("<I", header, 68, ENDOFCHAIN)       # 首个 DIFAT 扇区：无
    struct.pack_into("<I", header, 72, 0)                # DIFAT 扇区数

    # DIFAT：头部直接列出前 109 个 FAT 扇区号
    for index in range(HEADER_DIFAT_SLOTS):
        value = index if index < layout.n_fat else FREESECT
        struct.pack_into("<I", header, 76 + index * 4, value)
    return bytes(header)
