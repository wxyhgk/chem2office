"""写出一个合法的复合文件（CFBF / MS-CFB），与内容无关。

CFBF 是「一个文件里的迷你文件系统」：512 字节扇区、一张 FAT 分配表把扇区串成
链、一棵目录树记录每个流的名字与起点。本模块只管这套格式，不认识 ChemDraw；
「放哪些流、根节点 CLSID 是什么」由调用方决定（见 `chemdraw.py`）。

为什么不用现成库：olefile 只能读不能建，而 Office 对容器合法性极其挑剔——
FAT 常量写错的后果是 Excel 打开时「修复」文件并把 OLE 对象整个删掉。

布局（扇区从 0 开始编号，头部不算扇区）::

    header(512B)  FAT×n  目录×n  mini-FAT×n  mini 流×n  大流…

小于 4096 字节的流住在 mini 流里（由 mini-FAT 串链），大流直接占常规扇区。
FAT / mini-FAT / 目录都按载荷需要增长——早期写死单个 FAT 扇区，导致超过
63,488 字节的载荷静默产出损坏容器。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from io import BytesIO

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_STREAM_CUTOFF = 0x1000  # 4096：小于此值的流走 mini 流

# MS-CFB 特殊扇区号
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_FAT_ENTRIES_PER_SECTOR = SECTOR_SIZE // 4  # 128
_DIR_ENTRY_SIZE = 128
_HEADER_DIFAT_SLOTS = 109  # 头部能直接放下的 FAT 扇区号个数

ZERO_CLSID = b"\x00" * 16


def build_compound_file(
    streams: list[tuple[str, bytes]], *, root_clsid: bytes = ZERO_CLSID
) -> bytes:
    """把 ``(流名, 字节)`` 列表写成一个完整的复合文件。"""
    layout = _plan(streams)
    directory = _build_directory(layout, root_clsid=root_clsid)

    out = BytesIO()
    out.write(_build_header(layout))
    out.write(_build_fat(layout))
    out.write(directory)
    out.write(_build_minifat(layout))
    out.write(layout.mini_stream)
    for name, _data in layout.regular_defs:
        out.write(layout.regular_bytes[name])

    written = out.tell()
    expected = (layout.total_sectors + 1) * SECTOR_SIZE
    assert written == expected, f"sector layout mismatch: {written} != {expected}"
    return out.getvalue()


# --------------------------------------------------------------------------
# 1. 规划：每个流放哪、各区段占几个扇区
# --------------------------------------------------------------------------


@dataclass
class _Layout:
    streams: list[tuple[str, bytes]]
    mini_defs: list[tuple[str, bytes]] = field(default_factory=list)
    regular_defs: list[tuple[str, bytes]] = field(default_factory=list)
    mini_stream: bytes = b""          # 已补齐到扇区边界
    mini_stream_size: int = 0         # 未补齐的真实长度（写进根条目）
    mini_chain: list[int] = field(default_factory=list)
    mini_starts: dict[str, int] = field(default_factory=dict)
    regular_bytes: dict[str, bytes] = field(default_factory=dict)
    regular_starts: dict[str, int] = field(default_factory=dict)
    n_fat: int = 1
    dir_start: int = 0
    n_dir: int = 0
    minifat_start: int = 0
    n_minifat: int = 0
    ministream_start: int = 0
    n_ministream: int = 0
    total_sectors: int = 0


def _plan(streams: list[tuple[str, bytes]]) -> _Layout:
    layout = _Layout(streams=streams)
    layout.mini_defs = [(n, d) for n, d in streams if len(d) < MINI_STREAM_CUTOFF]
    layout.regular_defs = [(n, d) for n, d in streams if len(d) >= MINI_STREAM_CUTOFF]

    _pack_mini_stream(layout)
    layout.regular_bytes = {
        name: _pad(data, SECTOR_SIZE) for name, data in layout.regular_defs
    }
    n_regular = sum(len(b) // SECTOR_SIZE for b in layout.regular_bytes.values())

    layout.n_minifat = _ceil_div(len(layout.mini_chain) * 4, SECTOR_SIZE)
    layout.n_dir = _ceil_div((len(streams) + 1) * _DIR_ENTRY_SIZE, SECTOR_SIZE)
    layout.n_fat = _fat_sector_count(
        layout.n_dir + layout.n_minifat + layout.n_ministream + n_regular
    )

    layout.dir_start = layout.n_fat
    layout.minifat_start = layout.dir_start + layout.n_dir
    layout.ministream_start = layout.minifat_start + layout.n_minifat
    sector = layout.ministream_start + layout.n_ministream
    for name, _data in layout.regular_defs:
        layout.regular_starts[name] = sector
        sector += len(layout.regular_bytes[name]) // SECTOR_SIZE
    layout.total_sectors = sector
    return layout


def _pack_mini_stream(layout: _Layout) -> None:
    """把小流首尾相接成 mini 流，并记下每个流的起始 mini 扇区与链。"""
    buffer = bytearray()
    for name, data in layout.mini_defs:
        start = len(buffer) // MINI_SECTOR_SIZE
        layout.mini_starts[name] = start if data else ENDOFCHAIN
        padded = _pad(data, MINI_SECTOR_SIZE) if data else b""
        count = len(padded) // MINI_SECTOR_SIZE
        for index in range(count):
            last = index == count - 1
            layout.mini_chain.append(ENDOFCHAIN if last else start + index + 1)
        buffer.extend(padded)

    layout.mini_stream_size = len(buffer)
    layout.mini_stream = _pad(bytes(buffer), SECTOR_SIZE) if buffer else b""
    layout.n_ministream = len(layout.mini_stream) // SECTOR_SIZE


def _fat_sector_count(sectors_to_map: int) -> int:
    """FAT 自己也占扇区，也得被自己映射——所以要迭代到收敛。"""
    n_fat = 1
    while n_fat + sectors_to_map > n_fat * _FAT_ENTRIES_PER_SECTOR:
        n_fat += 1
    if n_fat > _HEADER_DIFAT_SLOTS:
        raise ValueError(
            f"payload too large for a header-only DIFAT ({n_fat} FAT sectors "
            f"> {_HEADER_DIFAT_SLOTS}); DIFAT sectors are not implemented."
        )
    return n_fat


# --------------------------------------------------------------------------
# 2. 分配表
# --------------------------------------------------------------------------


def _build_fat(layout: _Layout) -> bytes:
    fat = [FREESECT] * (layout.n_fat * _FAT_ENTRIES_PER_SECTOR)
    for index in range(layout.n_fat):
        fat[index] = FATSECT

    def chain(start: int, count: int) -> None:
        for index in range(count):
            sector = start + index
            fat[sector] = ENDOFCHAIN if index == count - 1 else sector + 1

    chain(layout.dir_start, layout.n_dir)
    chain(layout.minifat_start, layout.n_minifat)
    chain(layout.ministream_start, layout.n_ministream)
    for name, _data in layout.regular_defs:
        chain(layout.regular_starts[name], len(layout.regular_bytes[name]) // SECTOR_SIZE)
    return b"".join(struct.pack("<I", value) for value in fat)


def _build_minifat(layout: _Layout) -> bytes:
    """空闲位填 FREESECT——Windows 就是这么写的，填 0 会指向真实扇区 0。"""
    minifat = bytearray()
    for value in layout.mini_chain:
        minifat += struct.pack("<I", value)
    padding = (-len(minifat)) % SECTOR_SIZE
    minifat += struct.pack("<I", FREESECT) * (padding // 4)
    return bytes(minifat)


# --------------------------------------------------------------------------
# 3. 目录
# --------------------------------------------------------------------------


def _name_sort_key(name: str) -> tuple[int, str]:
    """CFB 目录序：先比名字长度，再比大写形式。"""
    return len(name), name.upper()


def _build_directory(layout: _Layout, *, root_clsid: bytes) -> bytes:
    """根条目 + 每个流一条，排成一棵按 CFB 序的平衡二叉查找树。

    颜色按深度交替，保证不会出现红节点带红子节点。
    """
    placed = [
        (
            name,
            layout.mini_starts[name]
            if name in layout.mini_starts
            else layout.regular_starts[name],
            len(data),
        )
        for name, data in layout.streams
    ]
    ordered = sorted(placed, key=lambda item: _name_sort_key(item[0]))
    sids = {name: index + 1 for index, (name, _s, _z) in enumerate(ordered)}
    links: dict[str, tuple[int, int, int]] = {}

    def build(subset: list[tuple[str, int, int]], depth: int) -> int:
        if not subset:
            return NOSTREAM
        middle = len(subset) // 2
        name = subset[middle][0]
        left = build(subset[:middle], depth + 1)
        right = build(subset[middle + 1 :], depth + 1)
        links[name] = (left, right, 1 if depth % 2 == 0 else 0)
        return sids[name]

    child = build(ordered, 0)

    directory = bytearray()
    directory += _make_dir_entry(
        "Root Entry", entry_type=5, color=0,
        left=NOSTREAM, right=NOSTREAM, child=child,
        clsid=root_clsid,
        start=layout.ministream_start if layout.n_ministream else ENDOFCHAIN,
        size=layout.mini_stream_size,
    )
    for name, start, size in ordered:
        left, right, color = links[name]
        directory += _make_dir_entry(
            name, entry_type=2, color=color,
            left=left, right=right, child=NOSTREAM,
            clsid=ZERO_CLSID, start=start, size=size,
        )
    return _pad(bytes(directory), SECTOR_SIZE)


def _encode_dir_name(name: str) -> tuple[bytes, int]:
    encoded = name.encode("utf-16-le") + b"\x00\x00"
    if len(encoded) > 64:
        raise ValueError(f"Directory name too long: {name!r}")
    return encoded.ljust(64, b"\x00"), len(encoded)


def _make_dir_entry(
    name: str,
    entry_type: int,
    color: int,
    left: int,
    right: int,
    child: int,
    clsid: bytes,
    start: int,
    size: int,
) -> bytes:
    """一条 128 字节目录项。"""
    name_bytes, name_len = _encode_dir_name(name)
    buf = BytesIO()
    buf.write(name_bytes)                    # 0-63    名字（UTF-16LE）
    buf.write(struct.pack("<H", name_len))   # 64-65   名字字节数
    buf.write(struct.pack("<B", entry_type)) # 66      2=流 5=根
    buf.write(struct.pack("<B", color))      # 67      0=红 1=黑
    buf.write(struct.pack("<I", left))       # 68-71   左兄弟
    buf.write(struct.pack("<I", right))      # 72-75   右兄弟
    buf.write(struct.pack("<I", child))      # 76-79   子树根
    buf.write(clsid)                         # 80-95
    buf.write(b"\x00" * 4)                   # 96-99   状态位
    buf.write(b"\x00" * 16)                  # 100-115 创建/修改时间
    buf.write(struct.pack("<I", start))      # 116-119 起始扇区
    buf.write(struct.pack("<I", size))       # 120-123 长度（低位）
    buf.write(struct.pack("<I", 0))          # 124-127 长度（高位）
    return buf.getvalue()


# --------------------------------------------------------------------------
# 4. 头部
# --------------------------------------------------------------------------


def _build_header(layout: _Layout) -> bytes:
    header = bytearray(SECTOR_SIZE)
    header[0:8] = _SIGNATURE
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
    for index in range(_HEADER_DIFAT_SLOTS):
        value = index if index < layout.n_fat else FREESECT
        struct.pack_into("<I", header, 76 + index * 4, value)
    return bytes(header)


# --------------------------------------------------------------------------


def _pad(data: bytes, size: int) -> bytes:
    remainder = len(data) % size
    return data + b"\x00" * (size - remainder) if remainder else data


def _ceil_div(value: int, divisor: int) -> int:
    return -(-value // divisor)


__all__ = [
    "MINI_STREAM_CUTOFF",
    "SECTOR_SIZE",
    "build_compound_file",
]
