"""Generated-template OOXML details (internal to ``_source``).

This package is the implementation drawer for ``wire_surface`` / ``generated``:
how to mutate drawing, VML, oleObjects, and sheet labels when cloning the
bundled ChemDraw OLE template. Callers outside ``_source`` should not import it.
"""

from __future__ import annotations

from .namespaces import register_namespaces
from .objects import update_drawing, update_ole_objects
from .sheet import (
    extract_base_anchor,
    update_row_heights,
    update_sheet_labels,
    write_shared_strings,
)
from .vml import update_vml

__all__ = [
    "extract_base_anchor",
    "register_namespaces",
    "update_drawing",
    "update_ole_objects",
    "update_row_heights",
    "update_sheet_labels",
    "update_vml",
    "write_shared_strings",
]
