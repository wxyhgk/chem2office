from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from retain_cdx.office.options import PreviewOptions

from ._utils import parse_px


@dataclass(frozen=True)
class ExcelPlacement:
    """Anchor placement: top-left cell + offset, plus optional second anchor.

    单锚点（默认）：``cell`` (+ ``offset_x``/``offset_y``) 只定左上角，
    显示框大小由 ``width_px``/``height_px`` 或 ``ExcelDisplayOptions`` 决定。

    两点锚（two-point anchor）：设置 ``to_cell`` 后，显示框从
    ``cell + offset`` 一直延伸到 ``to_cell + to_offset``。``to_cell`` 指该
    单元格的**左上角**（与 OOXML twoCellAnchor 的 ``to`` 语义一致），
    跨度按目标 sheet 的真实列宽/行高解析（未显式设置的列/行用 Excel
    默认 64px/19px）。

    Size precedence / 尺寸优先级::

        to_cell（两点锚，网格解析）
          > ExcelDisplayOptions.width/height
          > width_px/height_px 参数

    两点锚模式下不会再调整目标表的行高/列宽（框由现有网格本身定义）。
    Offsets accept numbers (px) or strings with px/pt/in/mm units,
    parsed via :func:`parse_px`.
    """

    cell: str
    offset_x: int | float | str = 0
    offset_y: int | float | str = 0
    # Optional second anchor (bottom-right corner of the display box).
    to_cell: str | None = None
    to_offset_x: int | float | str = 0
    to_offset_y: int | float | str = 0

    @property
    def offset_x_px(self) -> int:
        return parse_px(self.offset_x)

    @property
    def offset_y_px(self) -> int:
        return parse_px(self.offset_y)

    @property
    def to_offset_x_px(self) -> int:
        return parse_px(self.to_offset_x)

    @property
    def to_offset_y_px(self) -> int:
        return parse_px(self.to_offset_y)


@dataclass(frozen=True)
class ExcelDisplayOptions:
    width: int | float | str = 145
    height: int | float | str = 100
    row_height: float | None = None
    col_width: float | None = 24
    preserve_aspect: bool = True
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
class ExcelShellTemplate:
    """Native Excel/ChemDraw OLE shell to reuse for stable OLE display."""

    xlsx: str | Path
    preserve_aspect: bool = True
    # When True (default), honor insert cell/offset so the OLE is not stuck at the
    # shell template's original anchor (e.g. shell_large mid-sheet placement).
    allow_resize: bool = True
    fit_cdx: bool = False
