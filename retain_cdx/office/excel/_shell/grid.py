from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import xml.etree.ElementTree as ET

from retain_cdx.office.excel._utils import POINTS_PER_PIXEL

from .model import AnchorCorner, ShellSize, TwoCellAnchor


# SpreadsheetML namespace (local sheet XML after parse).
_NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def excel_col_width_to_px(width: float) -> int:
    """Approximate Excel column width (character units) → pixels."""
    # Same convention used elsewhere in this repo (_source/template/vml.py).
    return max(1, int(width * 7 + 5))


def excel_row_height_pt_to_px(height_pt: float) -> int:
    """Excel row height is stored in points."""
    return max(1, round(float(height_pt) / POINTS_PER_PIXEL))


@dataclass
class ExcelGridMetrics:
    """Column/row metrics used to convert pixel boxes to twoCell anchors.

    Always prefer loading from the target worksheet so anchors match the real
    grid Excel will use (custom col widths / row heights). Using hardcoded
    defaults after ``copy_sheet_layout`` is what stretched OLE previews.
    """

    default_col_width_px: int = 64
    default_row_height_px: int = 19
    col_width_by_index: dict[int, int] = field(default_factory=dict)
    row_height_by_index: dict[int, int] = field(default_factory=dict)

    @classmethod
    def from_worksheet(cls, sheet_path: Path) -> "ExcelGridMetrics":
        metrics = cls()
        if not sheet_path.exists():
            return metrics

        root = ET.parse(sheet_path).getroot()
        fmt = root.find("s:sheetFormatPr", _NS)
        if fmt is not None:
            if fmt.get("defaultColWidth") is not None:
                metrics.default_col_width_px = excel_col_width_to_px(
                    float(fmt.get("defaultColWidth"))
                )
            if fmt.get("defaultRowHeight") is not None:
                metrics.default_row_height_px = excel_row_height_pt_to_px(
                    float(fmt.get("defaultRowHeight"))
                )
            # baseColWidth is in characters when defaultColWidth absent
            if fmt.get("defaultColWidth") is None and fmt.get("baseColWidth") is not None:
                metrics.default_col_width_px = excel_col_width_to_px(
                    float(fmt.get("baseColWidth"))
                )

        cols = root.find("s:cols", _NS)
        if cols is not None:
            for col in cols.findall("s:col", _NS):
                if col.get("width") is None:
                    continue
                width_px = excel_col_width_to_px(float(col.get("width")))
                start = int(col.get("min", "1")) - 1
                end = int(col.get("max", col.get("min", "1"))) - 1
                for idx in range(start, end + 1):
                    metrics.col_width_by_index[idx] = width_px

        sheet_data = root.find("s:sheetData", _NS)
        if sheet_data is not None:
            for row in sheet_data.findall("s:row", _NS):
                if row.get("ht") is None:
                    continue
                r_idx = int(row.get("r", "1")) - 1
                metrics.row_height_by_index[r_idx] = excel_row_height_pt_to_px(
                    float(row.get("ht"))
                )
        return metrics

    def anchor_for_box(
        self,
        *,
        start_col: int,
        start_row: int,
        offset_x_px: int,
        offset_y_px: int,
        size: ShellSize,
    ) -> TwoCellAnchor:
        from_col, from_col_off_px = self.split_col(start_col - 1, offset_x_px)
        from_row, from_row_off_px = self.split_row(start_row - 1, offset_y_px)
        to_col, to_col_off_px = self.advance_cols(
            from_col, from_col_off_px + size.width_px
        )
        to_row, to_row_off_px = self.advance_rows(
            from_row, from_row_off_px + size.height_px
        )
        return TwoCellAnchor(
            from_=AnchorCorner(from_col, from_col_off_px, from_row, from_row_off_px),
            to=AnchorCorner(to_col, to_col_off_px, to_row, to_row_off_px),
        )

    def cell_origin_x_px(self, col: int) -> int:
        """Absolute x (px) of the left edge of 0-based column *col*."""
        return sum(self.col_width_px(idx) for idx in range(col))

    def cell_origin_y_px(self, row: int) -> int:
        """Absolute y (px) of the top edge of 0-based row *row*."""
        return sum(self.row_height_px(idx) for idx in range(row))

    def absolute_x_px(self, corner: AnchorCorner) -> int:
        return self.cell_origin_x_px(corner.col) + corner.col_off_px

    def absolute_y_px(self, corner: AnchorCorner) -> int:
        return self.cell_origin_y_px(corner.row) + corner.row_off_px

    def split_col(self, col: int, offset_px: int) -> tuple[int, int]:
        # Guard pathological infinite loops.
        for _ in range(10000):
            width = self.col_width_px(col)
            if offset_px < width:
                return col, offset_px
            offset_px -= width
            col += 1
        return col, offset_px

    def split_row(self, row: int, offset_px: int) -> tuple[int, int]:
        for _ in range(10000):
            height = self.row_height_px(row)
            if offset_px < height:
                return row, offset_px
            offset_px -= height
            row += 1
        return row, offset_px

    def advance_cols(self, col: int, width_px: int) -> tuple[int, int]:
        return self.split_col(col, width_px)

    def advance_rows(self, row: int, height_px: int) -> tuple[int, int]:
        return self.split_row(row, height_px)

    def col_width_px(self, col: int) -> int:
        return self.col_width_by_index.get(col, self.default_col_width_px)

    def row_height_px(self, row: int) -> int:
        return self.row_height_by_index.get(row, self.default_row_height_px)
