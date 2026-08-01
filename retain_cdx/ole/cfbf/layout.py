"""规划：每个流住哪里、各区段占几个扇区、从第几号扇区开始。

这一步只算数，不产出任何字节。后面 `alloc` / `directory` / `header`
都从这里的 `Layout` 派生。

扇区布局（头部不算扇区，扇区 0 紧跟其后）::

    header(512B)  FAT×n  目录×n  mini-FAT×n  mini流×n  大流…
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .const import (
    DIR_ENTRY_SIZE,
    ENDOFCHAIN,
    FAT_ENTRIES_PER_SECTOR,
    HEADER_DIFAT_SLOTS,
    MINI_SECTOR_SIZE,
    MINI_STREAM_CUTOFF,
    SECTOR_SIZE,
    ceil_div,
    pad,
)


@dataclass
class Layout:
    """一次写入的完整规划。"""

    streams: list[tuple[str, bytes]]
    mini_defs: list[tuple[str, bytes]] = field(default_factory=list)
    regular_defs: list[tuple[str, bytes]] = field(default_factory=list)

    mini_stream: bytes = b""      # 已补齐到扇区边界
    mini_stream_size: int = 0     # 未补齐的真实长度（写进根条目）
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

    def stream_start(self, name: str) -> int:
        """流 *name* 的起始扇区（mini 流里是 mini 扇区号）。"""
        if name in self.mini_starts:
            return self.mini_starts[name]
        return self.regular_starts[name]

    def regular_sector_count(self, name: str) -> int:
        return len(self.regular_bytes[name]) // SECTOR_SIZE


def plan(streams: list[tuple[str, bytes]]) -> Layout:
    layout = Layout(streams=streams)
    layout.mini_defs = [(n, d) for n, d in streams if len(d) < MINI_STREAM_CUTOFF]
    layout.regular_defs = [(n, d) for n, d in streams if len(d) >= MINI_STREAM_CUTOFF]

    _pack_mini_stream(layout)
    layout.regular_bytes = {
        name: pad(data, SECTOR_SIZE) for name, data in layout.regular_defs
    }
    n_regular = sum(
        layout.regular_sector_count(name) for name, _d in layout.regular_defs
    )

    layout.n_minifat = ceil_div(len(layout.mini_chain) * 4, SECTOR_SIZE)
    layout.n_dir = ceil_div((len(streams) + 1) * DIR_ENTRY_SIZE, SECTOR_SIZE)
    layout.n_fat = _fat_sector_count(
        layout.n_dir + layout.n_minifat + layout.n_ministream + n_regular
    )

    layout.dir_start = layout.n_fat
    layout.minifat_start = layout.dir_start + layout.n_dir
    layout.ministream_start = layout.minifat_start + layout.n_minifat
    sector = layout.ministream_start + layout.n_ministream
    for name, _data in layout.regular_defs:
        layout.regular_starts[name] = sector
        sector += layout.regular_sector_count(name)
    layout.total_sectors = sector
    return layout


def _pack_mini_stream(layout: Layout) -> None:
    """小流首尾相接成 mini 流，同时记下各自起点与链。"""
    buffer = bytearray()
    for name, data in layout.mini_defs:
        start = len(buffer) // MINI_SECTOR_SIZE
        layout.mini_starts[name] = start if data else ENDOFCHAIN
        padded = pad(data, MINI_SECTOR_SIZE) if data else b""
        count = len(padded) // MINI_SECTOR_SIZE
        for index in range(count):
            last = index == count - 1
            layout.mini_chain.append(ENDOFCHAIN if last else start + index + 1)
        buffer.extend(padded)

    layout.mini_stream_size = len(buffer)
    layout.mini_stream = pad(bytes(buffer), SECTOR_SIZE) if buffer else b""
    layout.n_ministream = len(layout.mini_stream) // SECTOR_SIZE


def _fat_sector_count(sectors_to_map: int) -> int:
    """FAT 自己也占扇区，也必须被自己映射——所以要迭代到收敛。

    早期这里写死 1，于是载荷超过 63,488 字节就静默产出损坏容器：
    olefile 仍能打开，读流时才炸。
    """
    n_fat = 1
    while n_fat + sectors_to_map > n_fat * FAT_ENTRIES_PER_SECTOR:
        n_fat += 1
    if n_fat > HEADER_DIFAT_SLOTS:
        raise ValueError(
            f"payload too large for a header-only DIFAT ({n_fat} FAT sectors "
            f"> {HEADER_DIFAT_SLOTS}); DIFAT sectors are not implemented."
        )
    return n_fat
