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

        dir_bytes = len(ole.direntries) * 128
        if dir_bytes > ole.sector_size:
            raise SystemExit(
                "Directory stream spans multiple sectors; size patch not supported."
            )
        first_dir_sector = ole.first_dir_sector
        sector_size = ole.sector_size
    finally:
        ole.close()

    dir_stream_offset = (first_dir_sector + 1) * sector_size
    dir_entry_offset = dir_stream_offset + (dir_index * 128)
    size_offset = dir_entry_offset + 120  # 8 bytes little-endian size
    with open(ole_path, "r+b") as f:
        f.seek(size_offset)
        f.write(new_size.to_bytes(8, "little"))
