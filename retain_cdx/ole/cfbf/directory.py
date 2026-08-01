"""目录：根条目 + 每个流一条，128 字节定长，排成一棵二叉查找树。

两处容易踩空：

* **排序不是字典序**——先比名字长度，再比大写形式（MS-CFB 规定）。
* 名字里的 ``\\x01`` / ``\\x02`` 前缀是内容的一部分，不是转义。
"""

from __future__ import annotations

import struct
from io import BytesIO

from .const import DIR_ENTRY_SIZE, ENDOFCHAIN, NOSTREAM, SECTOR_SIZE, ZERO_CLSID, pad
from .layout import Layout

_ENTRY_STREAM = 2
_ENTRY_ROOT = 5
_RED, _BLACK = 0, 1


def sort_key(name: str) -> tuple[int, str]:
    """MS-CFB 目录序：先长度，后大写形式。"""
    return len(name), name.upper()


def build_directory(layout: Layout, *, root_clsid: bytes) -> bytes:
    placed = [
        (name, layout.stream_start(name), len(data)) for name, data in layout.streams
    ]
    ordered = sorted(placed, key=lambda item: sort_key(item[0]))
    child, links = _build_tree(ordered)

    directory = bytearray()
    directory += make_entry(
        "Root Entry",
        entry_type=_ENTRY_ROOT,
        color=_RED,
        left=NOSTREAM,
        right=NOSTREAM,
        child=child,
        clsid=root_clsid,
        start=layout.ministream_start if layout.n_ministream else ENDOFCHAIN,
        size=layout.mini_stream_size,
    )
    for name, start, size in ordered:
        left, right, color = links[name]
        directory += make_entry(
            name,
            entry_type=_ENTRY_STREAM,
            color=color,
            left=left,
            right=right,
            child=NOSTREAM,
            clsid=ZERO_CLSID,
            start=start,
            size=size,
        )
    return pad(bytes(directory), SECTOR_SIZE)


def _build_tree(
    ordered: list[tuple[str, int, int]],
) -> tuple[int, dict[str, tuple[int, int, int]]]:
    """平衡二叉查找树；颜色按深度交替，红节点不会带红子节点。

    返回 (子树根的 sid, {名字: (左, 右, 颜色)})。sid 是条目在目录里的下标，
    根条目占 0，所以流从 1 开始。
    """
    sids = {name: index + 1 for index, (name, _s, _z) in enumerate(ordered)}
    links: dict[str, tuple[int, int, int]] = {}

    def build(subset: list[tuple[str, int, int]], depth: int) -> int:
        if not subset:
            return NOSTREAM
        middle = len(subset) // 2
        name = subset[middle][0]
        left = build(subset[:middle], depth + 1)
        right = build(subset[middle + 1 :], depth + 1)
        links[name] = (left, right, _BLACK if depth % 2 == 0 else _RED)
        return sids[name]

    return build(ordered, 0), links


def make_entry(
    name: str,
    *,
    entry_type: int,
    color: int,
    left: int,
    right: int,
    child: int,
    clsid: bytes,
    start: int,
    size: int,
) -> bytes:
    """一条 128 字节目录项（注释里的数字是字节偏移）。"""
    encoded = name.encode("utf-16-le") + b"\x00\x00"
    if len(encoded) > 64:
        raise ValueError(f"Directory name too long: {name!r}")

    buf = BytesIO()
    buf.write(encoded.ljust(64, b"\x00"))    # 0-63    名字（UTF-16LE）
    buf.write(struct.pack("<H", len(encoded)))  # 64-65 名字字节数
    buf.write(struct.pack("<B", entry_type)) # 66      2=流 5=根
    buf.write(struct.pack("<B", color))      # 67      0=红 1=黑
    buf.write(struct.pack("<I", left))       # 68-71   左兄弟
    buf.write(struct.pack("<I", right))      # 72-75   右兄弟
    buf.write(struct.pack("<I", child))      # 76-79   子树根
    buf.write(clsid)                         # 80-95
    buf.write(b"\x00" * 4)                   # 96-99   状态位
    buf.write(b"\x00" * 16)                  # 100-115 创建/修改时间
    buf.write(struct.pack("<I", start))      # 116-119 起始扇区
    buf.write(struct.pack("<I", size))       # 120-123 长度（低 32 位）
    buf.write(struct.pack("<I", 0))          # 124-127 长度（高 32 位）

    entry = buf.getvalue()
    assert len(entry) == DIR_ENTRY_SIZE
    return entry
