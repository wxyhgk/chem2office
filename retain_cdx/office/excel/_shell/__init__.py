"""Native ChemDraw OLE shell helpers (single-object source packages).

Substeps (aligned with generated fill/wire + merge); keep ``_`` package as internal::

  B   preview.replace_shell_preview — media art only
  C   media_rels — image relationship targets
  D   geometry — tighten box / patch cell position & size
  chrome — remove black OLE border

Orchestration of B/D/chrome order lives in ``_source.workbook`` (shell path).
Low-level: model, grid, vml, ooxml_patch.
"""

from __future__ import annotations

from .chrome import remove_shell_ole_border
from .geometry import (
    patch_shell_geometry,
    read_vml_size_px,
    tighten_shell_display_box,
)
from .grid import ExcelGridMetrics
from .preview import replace_shell_preview

__all__ = [
    "ExcelGridMetrics",
    "patch_shell_geometry",
    "read_vml_size_px",
    "remove_shell_ole_border",
    "replace_shell_preview",
    "tighten_shell_display_box",
]
