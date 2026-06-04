from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .units import parse_px


@dataclass(frozen=True)
class Placement:
    """Top-left cell and optional in-cell offset."""

    cell: str
    offset_x: int | float | str = 0
    offset_y: int | float | str = 0

    @property
    def offset_x_px(self) -> int:
        return parse_px(self.offset_x)

    @property
    def offset_y_px(self) -> int:
        return parse_px(self.offset_y)


@dataclass(frozen=True)
class PreviewOptions:
    width: int | None = None
    height: int | None = None
    thick_lines: bool = False
    no_margin: bool = False


@dataclass(frozen=True)
class DisplayOptions:
    width: int | float | str = 145
    height: int | float | str = 100
    row_height: float | None = None
    col_width: float | None = 24
    preserve_aspect: bool = False
    normalize_cdx: bool = False
    normalize_margin: float = 18
    normalize_fill_ratio: float = 0.8
    preview: PreviewOptions = PreviewOptions()

    @property
    def width_px(self) -> int:
        return parse_px(self.width)

    @property
    def height_px(self) -> int:
        return parse_px(self.height)


@dataclass(frozen=True)
class ShellTemplate:
    """Native Excel/ChemDraw OLE shell to reuse for stable OLE display."""

    xlsx: str | Path
    preserve_aspect: bool = True
    allow_resize: bool = False
    fit_cdx: bool = False
