from __future__ import annotations

from .embed import write_cdx_to_ole
from .eprint import EPRINT_STREAM, OBJINFO_STREAM, word_streams
from .presentation import (
    MAX_EMBEDDED_EMF,
    build_enhmetafile_presentation,
    emu_to_himetric,
    wrap_emf_in_wmf,
)

__all__ = [
    "EPRINT_STREAM",
    "OBJINFO_STREAM",
    "build_enhmetafile_presentation",
    "emu_to_himetric",
    "word_streams",
    "write_cdx_to_ole",
]
