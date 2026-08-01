"""复合文件格式的基本词汇（MS-CFB）。"""

from __future__ import annotations

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_STREAM_CUTOFF = 0x1000  # 4096：小于此值的流住 mini 流

# 特殊扇区号
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZERO_CLSID = b"\x00" * 16

FAT_ENTRIES_PER_SECTOR = SECTOR_SIZE // 4  # 128
DIR_ENTRY_SIZE = 128
HEADER_DIFAT_SLOTS = 109  # 头部能直接放下的 FAT 扇区号个数


def pad(data: bytes, size: int) -> bytes:
    """补零到 *size* 的整数倍。"""
    remainder = len(data) % size
    return data + b"\x00" * (size - remainder) if remainder else data


def ceil_div(value: int, divisor: int) -> int:
    return -(-value // divisor)
