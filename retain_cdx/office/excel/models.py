"""Internal DTOs shared by excel pipeline and source builders.

**Not public API.** Prefer ``ExcelCDXWriter`` + ``options.DisplayOptions`` etc.
from the package root. These models exist so ``_insert`` and ``_source`` can
share a neutral contract without depending on each other.

Rules:
  - Must not import ``_insert`` / ``_source`` (no reverse edges).
  - Safe to change within the package as long as internal callers update.
  - Do not re-export from ``retain_cdx`` / ``retain_cdx.office.excel`` ``__all__``.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from retain_cdx.render.preview import PreviewBackend, PreviewFormat, PreviewStyle


@dataclass(frozen=True)
class PreviewSpec:
    """How to render the OLE surface preview (PNG/EMF)."""

    width: int | None = None
    height: int | None = None
    thick_lines: bool = False
    no_margin: bool = False
    mono: bool = True
    format: PreviewFormat = "png"
    backend: PreviewBackend = "auto"
    style: PreviewStyle = "acs1996"
    # Strategy B: letterbox into fixed OLE box (set by source when known).
    target_aspect: float | None = None
    fill_ratio: float = 0.90


@dataclass(frozen=True)
class LayoutSpec:
    """Where and how large the OLE display box is on the sheet."""

    col: int
    row: int
    width_px: int
    height_px: int
    row_height: float | None = None
    col_width: float | None = None
    offset_x_px: int = 0
    offset_y_px: int = 0


@dataclass(frozen=True)
class NormalizeSpec:
    """Optional CDX view-box normalization before embed."""

    enabled: bool = False
    margin: float = 18.0
    fill_ratio: float = 0.8


@dataclass(frozen=True)
class ShellSourceOptions:
    """Native ChemDraw OLE shell settings for the source workbook path."""

    path: Path
    preserve_aspect: bool = True
    allow_resize: bool = True
    # When True, pass width/height into shell geometry patch (display-driven).
    resize_display: bool = False
    fit_cdx: bool = False


@dataclass(frozen=True)
class SourceBuildRequest:
    """Everything ``_source`` needs to build a temporary source workbook.

    Pipeline constructs this from an insert plan; source must not import plan.
    """

    layout: LayoutSpec
    preview: PreviewSpec
    shell: ShellSourceOptions | None = None

    @property
    def has_shell(self) -> bool:
        return self.shell is not None
