from __future__ import annotations

from pathlib import Path

import olefile

from .cfbf import MINI_STREAM_CUTOFF
from .chemdraw import build_chemdraw_ole
from .eprint import EPRINT_STREAM
from .presentation import PRESENTATION_STREAM
from .inject import inject_cdx_stream, patch_direntry_size

# Streams caching a rendering of the old molecule (PowerPoint containers carry
# \x02OlePres000, Word containers \x03EPRINT). We cannot regenerate them, and a
# cache that disagrees with CONTENTS is worse than none — Excel's own ChemDraw
# containers ship without.
_STALE_CACHE_PREFIXES = ("\x02OlePres", EPRINT_STREAM)


def write_cdx_to_ole(
    ole_path: Path,
    cdx_path: Path,
    *,
    template_ole: bytes | None = None,
    stream_name: str = "CONTENTS",
    presentation: bytes | None = None,
    extra_streams: dict[str, bytes] | None = None,
) -> None:
    """Write *cdx_path* into a ChemDraw OLE container.

    If *template_ole* is provided it is written to *ole_path* first. The
    template's stream is patched in place when the new payload fits the same
    storage; otherwise a fresh container is built.

    *presentation* attaches a cached rendering (`\\x02OlePres000`). Without one
    the object still activates, but copying it out of PowerPoint degrades to a
    picture — see `presentation.py`. Supplying it forces a rebuild so the cache
    cannot disagree with CONTENTS.

    *extra_streams* are written alongside CONTENTS verbatim (Word's
    ``\x03EPRINT`` / ``\x03ObjInfo``); like *presentation* they force a rebuild.
    """
    if template_ole is not None:
        ole_path.write_bytes(template_ole)

    cdx_data = cdx_path.read_bytes()
    streams = dict(extra_streams or {})
    if presentation is not None:
        streams[PRESENTATION_STREAM] = presentation
    if streams:
        ole_path.write_bytes(build_chemdraw_ole(cdx_data, extra_streams=streams))
        return
    if _needs_rebuild(ole_path, stream_name, len(cdx_data)):
        ole_path.write_bytes(build_chemdraw_ole(cdx_data))
        return

    inject_cdx_stream(ole_path, stream_name, cdx_data)
    patch_direntry_size(ole_path, stream_name, len(cdx_data))


def _needs_rebuild(ole_path: Path, stream_name: str, new_size: int) -> bool:
    """Decide between patching the template in place and rebuilding it."""
    ole = olefile.OleFileIO(str(ole_path))
    try:
        if not ole.exists(stream_name):
            raise ValueError(f"OLE stream not found: {stream_name}")
        old_size = ole.openstream(stream_name).size
        names = ["/".join(entry) for entry in ole.listdir()]
    finally:
        ole.close()

    if new_size > old_size:
        return True
    # Crossing the mini-stream cutoff moves the stream between the mini FAT and
    # the regular FAT; patching only the size field would leave the start
    # sector being resolved against the wrong allocator.
    if (old_size >= MINI_STREAM_CUTOFF) != (new_size >= MINI_STREAM_CUTOFF):
        return True
    return any(name.startswith(_STALE_CACHE_PREFIXES) for name in names)
