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


def build_chemdraw_ole(cdx_data: bytes) -> bytes:
    """Return a complete CFBF (.bin) OLE file embedding *cdx_data*.

    The generated binary contains three streams that Excel / ChemDraw
    expect:

    * ``\\x01Ole``     – OLE version marker (20 bytes, fixed)
    * ``\\x01CompObj`` – ProgID *ChemDraw.Document.6.0*
    * ``CONTENTS``     – the raw CDX data (variable length)

    All streams use the **mini-stream** when their individual size is
    below 4096 bytes, which is the common case for typical CDX files.
    For very large CDX data (≥4096 bytes), streams are stored as
    regular sectors.
    """
    ole_data = _OLE_STREAM
    compobj_data = _build_compobj()
    contents_data = cdx_data

    # ---- Decide storage strategy ----
    # Small streams live inside the root mini-stream, while larger
    # CONTENTS payloads must use regular FAT sectors.
    mini_stream_defs: list[tuple[str, bytes]] = []
    regular_stream_defs: list[tuple[str, bytes]] = []
    for name, data in [
        ("\x01Ole", ole_data),
        ("\x01CompObj", compobj_data),
        ("CONTENTS", contents_data),
    ]:
        if len(data) < _MINI_STREAM_CUTOFF:
            mini_stream_defs.append((name, data))
        else:
            regular_stream_defs.append((name, data))

    # Build mini-stream: concatenate small stream data, each padded to
    # a multiple of mini-sector size (64 bytes).
    mini_stream = bytearray()
    mini_starts: dict[str, int] = {}  # start mini-sector for each stream
    mini_chain: list[int] = []       # mini-FAT entries

    for name, sdata in mini_stream_defs:
        start_mini_sec = len(mini_stream) // _MINI_SECTOR_SIZE
        mini_starts[name] = start_mini_sec
        padded = _pad(sdata, _MINI_SECTOR_SIZE)
        n_mini_secs = len(padded) // _MINI_SECTOR_SIZE
        for j in range(n_mini_secs):
            sec_id = start_mini_sec + j
            if j < n_mini_secs - 1:
                mini_chain.append(sec_id + 1)
            else:
                mini_chain.append(_ENDOFCHAIN)
        mini_stream.extend(padded)

    # Pad mini-stream to sector boundary.
    mini_stream_padded = _pad(bytes(mini_stream), _SECTOR_SIZE)
    n_mini_stream_sectors = len(mini_stream_padded) // _SECTOR_SIZE

    regular_stream_bytes: dict[str, bytes] = {}
    regular_stream_sectors: dict[str, int] = {}
    for name, sdata in regular_stream_defs:
        regular_stream_bytes[name] = _pad(sdata, _SECTOR_SIZE)
        regular_stream_sectors[name] = len(regular_stream_bytes[name]) // _SECTOR_SIZE

    # ---- Layout sectors ----
    # Sector 0: FAT sector
    # Sector 1: Directory sector (4 entries = 512 bytes, fits in 1 sector)
    # Sector 2: Mini-FAT sector
    # Sector 3 .. 3+N-1: Mini-stream data sectors
    # Remaining sectors: regular FAT-backed streams (e.g. large CONTENTS)
    mini_stream_start = 3
    next_sector = mini_stream_start + n_mini_stream_sectors
    regular_starts: dict[str, int] = {}
    for name, _sdata in regular_stream_defs:
        regular_starts[name] = next_sector
        next_sector += regular_stream_sectors[name]
    total_sectors = next_sector

    # ---- Build FAT ----
    fat = [0] * max(128, total_sectors)  # one FAT sector = 128 entries
    fat[0] = _FATSECT        # sector 0 is the FAT itself
    fat[1] = _ENDOFCHAIN     # sector 1 is directory (single sector)
    fat[2] = _ENDOFCHAIN     # sector 2 is mini-FAT (single sector)
    for i in range(n_mini_stream_sectors):
        s = mini_stream_start + i
        fat[s] = s + 1 if i < n_mini_stream_sectors - 1 else _ENDOFCHAIN
    for name, _sdata in regular_stream_defs:
        start = regular_starts[name]
        count = regular_stream_sectors[name]
        for i in range(count):
            s = start + i
            fat[s] = s + 1 if i < count - 1 else _ENDOFCHAIN
    # Mark rest as free.
    for i in range(total_sectors, 128):
        fat[i] = _FREE

    fat_sector = bytearray()
    for v in fat[:128]:
        fat_sector += struct.pack("<I", v)

    # ---- Build Mini-FAT ----
    mini_fat = bytearray()
    for v in mini_chain:
        mini_fat += struct.pack("<I", v)
    # Pad to sector size.
    mini_fat = _pad(bytes(mini_fat), _SECTOR_SIZE)

    # ---- Build Directory ----
    root_size = len(mini_stream)  # un-padded mini-stream size
    dir_sector = bytearray()

    # DE[0]: Root Entry – child points to CompObj (DE[2]) as subtree root
    dir_sector += _make_dir_entry(
        "Root Entry", entry_type=5, color=0,
        left=_NOSTREAM, right=_NOSTREAM, child=2,
        clsid=_CHEMDRAW_CLSID,
        start=mini_stream_start, size=root_size,
    )

    # DE[1]: \x01Ole – left sibling of CompObj
    dir_sector += _make_dir_entry(
        "\x01Ole", entry_type=2, color=0,
        left=_NOSTREAM, right=_NOSTREAM, child=_NOSTREAM,
        clsid=_ZERO_CLSID,
        start=mini_starts["\x01Ole"], size=len(ole_data),
    )

    # DE[2]: \x01CompObj – subtree root, left=Ole, right=CONTENTS
    dir_sector += _make_dir_entry(
        "\x01CompObj", entry_type=2, color=1,
        left=1, right=3, child=_NOSTREAM,
        clsid=_ZERO_CLSID,
        start=mini_starts["\x01CompObj"], size=len(compobj_data),
    )

    # DE[3]: CONTENTS
    contents_start = (
        mini_starts["CONTENTS"]
        if "CONTENTS" in mini_starts
        else regular_starts["CONTENTS"]
    )
    dir_sector += _make_dir_entry(
        "CONTENTS", entry_type=2, color=0,
        left=_NOSTREAM, right=_NOSTREAM, child=_NOSTREAM,
        clsid=_ZERO_CLSID,
        start=contents_start, size=len(contents_data),
    )

    dir_sector = _pad(bytes(dir_sector), _SECTOR_SIZE)

    # ---- Build Header ----
    header = bytearray(512)
    # Signature
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    # Minor version = 0x003E, Major = 0x0003 (v3)
    struct.pack_into("<HH", header, 24, 0x003E, 0x0003)
    # Byte order = little-endian
    struct.pack_into("<H", header, 28, 0xFFFE)
    # Sector size power = 9 (512)
    struct.pack_into("<H", header, 30, 9)
    # Mini sector size power = 6 (64)
    struct.pack_into("<H", header, 32, 6)
    # Reserved (6 bytes zero) – already zero
    # Total directory sectors = 0 (must be 0 for v3)
    struct.pack_into("<I", header, 40, 0)
    # Total FAT sectors = 1
    struct.pack_into("<I", header, 44, 1)
    # First directory sector
    struct.pack_into("<I", header, 48, 1)
    # Transaction signature = 0
    struct.pack_into("<I", header, 52, 0)
    # Mini stream cutoff
    struct.pack_into("<I", header, 56, _MINI_STREAM_CUTOFF)
    # First mini-FAT sector
    struct.pack_into("<I", header, 60, 2)
    # Number of mini-FAT sectors = 1
    struct.pack_into("<I", header, 64, 1)
    # First DIFAT sector = none
    struct.pack_into("<I", header, 68, _ENDOFCHAIN)
    # Number of DIFAT sectors = 0
    struct.pack_into("<I", header, 72, 0)
    # DIFAT array: first entry = sector 0 (FAT), rest 0xFFFFFFFE
    struct.pack_into("<I", header, 76, 0)
    for i in range(1, 109):
        struct.pack_into("<I", header, 76 + i * 4, _FREE)

    # ---- Assemble ----
    out = BytesIO()
    out.write(bytes(header))          # header (512 bytes)
    out.write(bytes(fat_sector))      # sector 0: FAT
    out.write(bytes(dir_sector))      # sector 1: directory
    out.write(bytes(mini_fat))        # sector 2: mini-FAT
    out.write(mini_stream_padded)     # sector 3+: mini-stream data
    for name, _sdata in regular_stream_defs:
        out.write(regular_stream_bytes[name])

    return out.getvalue()
