from __future__ import annotations

from dataclasses import dataclass

from retain_cdx.office.excel._utils import parse_emu
from retain_cdx.office.excel.options import PreviewOptions

Length = int | float | str


@dataclass(frozen=True)
class SlidePlacement:
    """Absolute box on a slide.

    Lengths accept ``"6cm"`` / ``"2in"`` / ``"200pt"`` / ``"96px"`` or a bare
    number (already EMU). Leave *width* or *height* unset to derive it from the
    molecule's own aspect ratio — set neither and *size* falls back to
    ``PowerPointCDXWriter`` defaults.
    """

    x: Length = 0
    y: Length = 0
    width: Length | None = None
    height: Length | None = None

    @property
    def x_emu(self) -> int:
        return parse_emu(self.x)

    @property
    def y_emu(self) -> int:
        return parse_emu(self.y)

    @property
    def width_emu(self) -> int | None:
        return None if self.width is None else parse_emu(self.width)

    @property
    def height_emu(self) -> int | None:
        return None if self.height is None else parse_emu(self.height)


@dataclass(frozen=True)
class SlideGrid:
    """Lay several structures out on one slide, row by row."""

    columns: int = 3
    x: Length = "1cm"
    y: Length = "1cm"
    cell_width: Length = "5cm"
    cell_height: Length = "4cm"
    gap_x: Length = "0.5cm"
    gap_y: Length = "0.5cm"

    def box(self, index: int) -> tuple[int, int, int, int]:
        """(x, y, width, height) in EMU for the *index*-th structure."""
        if self.columns < 1:
            raise ValueError("SlideGrid.columns must be >= 1")
        column, row = index % self.columns, index // self.columns
        width, height = parse_emu(self.cell_width), parse_emu(self.cell_height)
        x = parse_emu(self.x) + column * (width + parse_emu(self.gap_x))
        y = parse_emu(self.y) + row * (height + parse_emu(self.gap_y))
        return x, y, width, height


__all__ = ["PreviewOptions", "SlideGrid", "SlidePlacement"]
