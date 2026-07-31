"""Build a ChemDraw OLE (CFBF) binary from raw CDX data.

This module constructs a valid Compound File Binary Format container from
scratch, eliminating the need for a Windows-generated template.  The
resulting .bin file can be embedded directly in an XLSX as an OLE object.
"""

from __future__ import annotations

import math
import struct
from io import BytesIO

# ChemDraw CLSID  41BA6D21-A02E-11CE-8FD9-0020AFD1F20C
# Stored in mixed-endian per MS-DTYP §2.3.4.2:
#   first three components little-endian, last two big-endian.
_CHEMDRAW_CLSID = (
    struct.pack("<IHH", 0x41BA6D21, 0xA02E, 0x11CE)
    + bytes([0x8F, 0xD9, 0x00, 0x20, 0xAF, 0xD1, 0xF2, 0x0C])
)

_ZERO_CLSID = b"\x00" * 16

# Fixed content of the \x01Ole stream (version 2.0, no links).
_OLE_STREAM = struct.pack("<IIIII", 0x02000001, 0, 0, 0, 0)

_SECTOR_SIZE = 512
_MINI_SECTOR_SIZE = 64
_MINI_STREAM_CUTOFF = 0x1000  # 4096 bytes – streams smaller go to mini stream

# CFBF special sector IDs per [MS-CFB].
_FREE = 0xFFFFFFFF
_ENDOFCHAIN = 0xFFFFFFFE
_FATSECT = 0xFFFFFFFD
_NOSTREAM = 0xFFFFFFFF


def _build_compobj() -> bytes:
    """Build the \\x01CompObj stream for ChemDraw."""
    buf = BytesIO()
    # Header: version, byte-order, OS, unused, CLSID
    # Match the byte layout emitted by Excel's template OLE object.
    buf.write(struct.pack("<HHI", 0x0001, 0xFFFE, 0x0A03))
    buf.write(b"\xff\xff\xff\xff")
    buf.write(_CHEMDRAW_CLSID)

    # AnsiUserType (length-prefixed ANSI string including NUL)
    user_type = b"CS ChemDraw Drawing\x00"
    buf.write(struct.pack("<I", len(user_type)))
    buf.write(user_type)

    # AnsiClipboardFormat (length-prefixed ANSI string including NUL)
    clip_fmt = b"ChemDraw Interchange Format\x00"
    buf.write(struct.pack("<I", len(clip_fmt)))
    buf.write(clip_fmt)

    # Reserved (ProgID as ANSI, length-prefixed including NUL)
    prog_id = b"ChemDraw.Document.6.0\x00"
    buf.write(struct.pack("<I", len(prog_id)))
    buf.write(prog_id)

    # UnicodeMarker + tail (matching observed template)
    buf.write(struct.pack("<I", 0x71B239F4))
    buf.write(b"\x00" * 12)

    return buf.getvalue()


def _pad(data: bytes, size: int) -> bytes:
    """Pad *data* to a multiple of *size* with zero bytes."""
    r = len(data) % size
    return data + b"\x00" * (size - r) if r else data


def _encode_dir_name(name: str) -> tuple[bytes, int]:
    """Encode a directory entry name and return (bytes[64], name_bytes)."""
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
    """Build a 128-byte CFBF directory entry."""
    name_bytes, name_len = _encode_dir_name(name)
    buf = BytesIO()
    buf.write(name_bytes)                         # 0-63:   name
    buf.write(struct.pack("<H", name_len))         # 64-65:  name size
    buf.write(struct.pack("<B", entry_type))        # 66:     type
    buf.write(struct.pack("<B", color))             # 67:     color (0=red,1=black)
    buf.write(struct.pack("<I", left))              # 68-71:  left sibling
    buf.write(struct.pack("<I", right))             # 72-75:  right sibling
    buf.write(struct.pack("<I", child))             # 76-79:  child
    buf.write(clsid)                               # 80-95:  CLSID
    buf.write(b"\x00" * 4)                         # 96-99:  state bits
    buf.write(b"\x00" * 8)                         # 100-107: creation time
    buf.write(b"\x00" * 8)                         # 108-115: modification time
    buf.write(struct.pack("<I", start))             # 116-119: start sector
    buf.write(struct.pack("<I", size))              # 120-123: size (low)
    buf.write(struct.pack("<I", 0))                 # 124-127: size (high)
    return buf.getvalue()


def _name_sort_key(name: str) -> tuple[int, str]:
    """CFB directory ordering: UTF-16 name length first, then uppercase name."""
    return len(name), name.upper()


def _build_directory(streams: list[tuple[str, int, int]], root_start: int, root_size: int) -> bytes:
    """Directory sectors for *streams* as ``(name, start_sector, size)``.

    Entries are laid out as a balanced binary search tree over the CFB name
    ordering, coloured by depth so no red node ever parents another.
    """
    ordered = sorted(streams, key=lambda item: _name_sort_key(item[0]))
    sids = {name: index + 1 for index, (name, _s, _z) in enumerate(ordered)}

    links: dict[str, tuple[int, int, int]] = {}

    def build(subset: list[tuple[str, int, int]], depth: int) -> int:
        if not subset:
            return _NOSTREAM
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
        left=_NOSTREAM, right=_NOSTREAM, child=child,
        clsid=_CHEMDRAW_CLSID, start=root_start, size=root_size,
    )
    for name, start, size in ordered:
        left, right, color = links[name]
        directory += _make_dir_entry(
            name, entry_type=2, color=color,
            left=left, right=right, child=_NOSTREAM,
            clsid=_ZERO_CLSID, start=start, size=size,
        )
    return _pad(bytes(directory), _SECTOR_SIZE)


def _ceil_div(value: int, divisor: int) -> int:
    return -(-value // divisor)


def build_chemdraw_ole(
    cdx_data: bytes, *, extra_streams: dict[str, bytes] | None = None
) -> bytes:
    """Return a complete CFBF (.bin) OLE file embedding *cdx_data*.

    Always contains the three streams Excel / ChemDraw expect:

    * ``\\x01Ole``     – OLE version marker (20 bytes, fixed)
    * ``\\x01CompObj`` – ProgID *ChemDraw.Document.6.0*
    * ``CONTENTS``     – the raw CDX data (variable length)

    *extra_streams* adds further streams verbatim (PowerPoint's containers also
    carry a ``\\x02OlePres000`` presentation stream, for instance).

    Streams below the 4096-byte cutoff live in the mini stream, larger ones in
    regular sectors. FAT, mini-FAT and directory all grow to as many sectors as
    the payload needs — a single-FAT-sector layout silently corrupted any CDX
    over ~62 KB, which is the failure mode where Excel drops the object.
    """
    streams: list[tuple[str, bytes]] = [
        ("\x01Ole", _OLE_STREAM),
        ("\x01CompObj", _build_compobj()),
        ("CONTENTS", cdx_data),
    ]
    for name, data in (extra_streams or {}).items():
        streams.append((name, data))

    entries_per_sector = _SECTOR_SIZE // 4  # 128 FAT entries per sector

    # ---- Mini stream (everything below the 4096-byte cutoff) ----
    mini_defs = [(n, d) for n, d in streams if len(d) < _MINI_STREAM_CUTOFF]
    regular_defs = [(n, d) for n, d in streams if len(d) >= _MINI_STREAM_CUTOFF]

    mini_stream = bytearray()
    mini_starts: dict[str, int] = {}
    mini_chain: list[int] = []
    for name, data in mini_defs:
        start = len(mini_stream) // _MINI_SECTOR_SIZE
        mini_starts[name] = start if data else _ENDOFCHAIN
        padded = _pad(data, _MINI_SECTOR_SIZE) if data else b""
        count = len(padded) // _MINI_SECTOR_SIZE
        for index in range(count):
            mini_chain.append(
                start + index + 1 if index < count - 1 else _ENDOFCHAIN
            )
        mini_stream.extend(padded)

    mini_stream_padded = _pad(bytes(mini_stream), _SECTOR_SIZE) if mini_stream else b""
    n_ministream = len(mini_stream_padded) // _SECTOR_SIZE
    n_minifat = _ceil_div(len(mini_chain) * 4, _SECTOR_SIZE)
    n_dir = _ceil_div((len(streams) + 1) * 128, _SECTOR_SIZE)

    regular_bytes = {name: _pad(data, _SECTOR_SIZE) for name, data in regular_defs}
    n_regular = sum(len(b) // _SECTOR_SIZE for b in regular_bytes.values())

    # ---- How many FAT sectors? Each one also occupies a sector it must map. ----
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_ministream + n_regular
        if total <= n_fat * entries_per_sector:
            break
        n_fat += 1
    if n_fat > 109:
        raise ValueError(
            f"CDX payload too large for a header-only DIFAT ({len(cdx_data)} bytes); "
            "DIFAT sectors are not implemented."
        )

    dir_start = n_fat
    minifat_start = dir_start + n_dir
    ministream_start = minifat_start + n_minifat
    next_sector = ministream_start + n_ministream
    regular_starts: dict[str, int] = {}
    for name, _data in regular_defs:
        regular_starts[name] = next_sector
        next_sector += len(regular_bytes[name]) // _SECTOR_SIZE
    total_sectors = next_sector

    # ---- FAT ----
    fat = [_FREE] * (n_fat * entries_per_sector)
    for index in range(n_fat):
        fat[index] = _FATSECT

    def chain(start: int, count: int) -> None:
        for index in range(count):
            sector = start + index
            fat[sector] = sector + 1 if index < count - 1 else _ENDOFCHAIN

    chain(dir_start, n_dir)
    chain(minifat_start, n_minifat)
    chain(ministream_start, n_ministream)
    for name, _data in regular_defs:
        chain(regular_starts[name], len(regular_bytes[name]) // _SECTOR_SIZE)

    fat_sectors = b"".join(struct.pack("<I", value) for value in fat)

    # ---- Mini FAT (free entries are FREESECT, matching Windows output) ----
    minifat = bytearray()
    for value in mini_chain:
        minifat += struct.pack("<I", value)
    padding = (-len(minifat)) % _SECTOR_SIZE
    minifat += struct.pack("<I", _FREE) * (padding // 4)

    # ---- Directory ----
    placed = [
        (
            name,
            mini_starts[name] if name in mini_starts else regular_starts[name],
            len(data),
        )
        for name, data in streams
    ]
    directory = _build_directory(
        placed,
        root_start=ministream_start if n_ministream else _ENDOFCHAIN,
        root_size=len(mini_stream),
    )

    # ---- Header ----
    header = bytearray(512)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<HH", header, 24, 0x003E, 0x0003)  # minor / major version
    struct.pack_into("<H", header, 28, 0xFFFE)           # little-endian
    struct.pack_into("<H", header, 30, 9)                # 512-byte sectors
    struct.pack_into("<H", header, 32, 6)                # 64-byte mini sectors
    struct.pack_into("<I", header, 40, 0)                # csectDir must be 0 in v3
    struct.pack_into("<I", header, 44, n_fat)
    struct.pack_into("<I", header, 48, dir_start)
    struct.pack_into("<I", header, 52, 0)                # transaction signature
    struct.pack_into("<I", header, 56, _MINI_STREAM_CUTOFF)
    struct.pack_into("<I", header, 60, minifat_start if n_minifat else _ENDOFCHAIN)
    struct.pack_into("<I", header, 64, n_minifat)
    struct.pack_into("<I", header, 68, _ENDOFCHAIN)      # first DIFAT sector
    struct.pack_into("<I", header, 72, 0)                # DIFAT sector count
    for index in range(109):
        value = index if index < n_fat else _FREE
        struct.pack_into("<I", header, 76 + index * 4, value)

    out = BytesIO()
    out.write(bytes(header))
    out.write(fat_sectors)
    out.write(directory)
    out.write(bytes(minifat))
    out.write(mini_stream_padded)
    for name, _data in regular_defs:
        out.write(regular_bytes[name])
    assert out.tell() == (total_sectors + 1) * _SECTOR_SIZE, "sector layout mismatch"
    return out.getvalue()
