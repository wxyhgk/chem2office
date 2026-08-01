"""分配表：把规划好的扇区串成链。

FAT 管常规扇区，mini-FAT 管 mini 流内部的 64 字节小扇区。
"""

from __future__ import annotations

import struct

from .const import ENDOFCHAIN, FAT_ENTRIES_PER_SECTOR, FATSECT, FREESECT, SECTOR_SIZE
from .layout import Layout


def build_fat(layout: Layout) -> bytes:
    fat = [FREESECT] * (layout.n_fat * FAT_ENTRIES_PER_SECTOR)
    for index in range(layout.n_fat):
        fat[index] = FATSECT  # FAT 扇区自己也要标出来

    def chain(start: int, count: int) -> None:
        for index in range(count):
            sector = start + index
            fat[sector] = ENDOFCHAIN if index == count - 1 else sector + 1

    chain(layout.dir_start, layout.n_dir)
    chain(layout.minifat_start, layout.n_minifat)
    chain(layout.ministream_start, layout.n_ministream)
    for name, _data in layout.regular_defs:
        chain(layout.regular_starts[name], layout.regular_sector_count(name))
    return b"".join(struct.pack("<I", value) for value in fat)


def build_minifat(layout: Layout) -> bytes:
    """空闲位填 FREESECT——Windows 就是这么写的；填 0 会指向真实的扇区 0。"""
    minifat = bytearray()
    for value in layout.mini_chain:
        minifat += struct.pack("<I", value)
    padding = (-len(minifat)) % SECTOR_SIZE
    minifat += struct.pack("<I", FREESECT) * (padding // 4)
    return bytes(minifat)
