from __future__ import annotations

from pathlib import Path

from .inject import inject_cdx_stream, patch_direntry_size
from .cfbf import build_chemdraw_ole


def write_cdx_to_ole(
    ole_path: Path,
    cdx_path: Path,
    *,
    template_ole: bytes | None = None,
    stream_name: str = "CONTENTS",
) -> None:
    """Write *cdx_path* into a ChemDraw OLE container.

    If *template_ole* is provided, it is written to *ole_path* first. If the
    CDX payload is too large for the template stream, a fresh ChemDraw OLE
    container is built instead.
    """
    if template_ole is not None:
        ole_path.write_bytes(template_ole)

    cdx_data = cdx_path.read_bytes()
    try:
        inject_cdx_stream(ole_path, stream_name, cdx_data)
        patch_direntry_size(ole_path, stream_name, len(cdx_data))
    except SystemExit as exc:
        if "CDX is larger than existing stream" not in str(exc):
            raise
        ole_path.write_bytes(build_chemdraw_ole(cdx_data))
