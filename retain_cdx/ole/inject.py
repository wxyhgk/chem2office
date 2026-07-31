from __future__ import annotations

from pathlib import Path

import olefile


def inject_cdx_stream(ole_path: Path, stream_name: str, cdx_data: bytes) -> None:
    ole = olefile.OleFileIO(str(ole_path), write_mode=True)
    try:
        if not ole.exists(stream_name):
            raise SystemExit(f"OLE stream not found: {stream_name}")
        old_data = ole.openstream(stream_name).read()
        if len(cdx_data) > len(old_data):
            raise SystemExit(
                f"CDX is larger than existing stream ({len(cdx_data)} > {len(old_data)})."
            )
        padded = cdx_data + (b"\x00" * (len(old_data) - len(cdx_data)))
        ole.write_stream(stream_name, padded)
    finally:
        ole.close()


def patch_direntry_size(ole_path: Path, stream_name: str, new_size: int) -> None:
    # olefile cannot change stream size; patch the directory entry size field.
    ole = olefile.OleFileIO(str(ole_path))
    try:
        dir_index = None
        for de in ole.direntries:
            if de and de.name == stream_name:
                dir_index = de.sid
                break
        if dir_index is None:
            raise SystemExit(f"Directory entry not found for stream: {stream_name}")

        size_offset = _direntry_offset(ole, dir_index) + 120  # 8-byte LE size
    finally:
        ole.close()

    with open(ole_path, "r+b") as f:
        f.seek(size_offset)
        f.write(new_size.to_bytes(8, "little"))


def _direntry_offset(ole: "olefile.OleFileIO", sid: int) -> int:
    """Byte offset of directory entry *sid*, following the directory FAT chain.

    Containers with a presentation stream (PowerPoint's ChemDraw objects carry
    ``\\x02OlePres000``) hold more than the four entries that fit in one
    512-byte sector, so the directory stream must be walked, not assumed flat.
    """
    sector_size = ole.sector_size
    per_sector = sector_size // 128
    sector = ole.first_dir_sector
    for _ in range(sid // per_sector):
        if sector < 0 or sector >= len(ole.fat):
            raise SystemExit("Directory chain ended before reaching the entry.")
        sector = ole.fat[sector]
        if sector >= olefile.MAXREGSECT:
            raise SystemExit("Directory chain ended before reaching the entry.")
    return (sector + 1) * sector_size + (sid % per_sector) * 128
