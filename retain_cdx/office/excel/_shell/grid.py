from __future__ import annotations

from dataclasses import dataclass

from .model import AnchorCorner, ShellSize, TwoCellAnchor


@dataclass(frozen=True)
class ExcelGridMetrics:
    default_col_width_px: int = 64
    default_row_height_px: int = 19

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
        to_col, to_col_off_px = self.advance_cols(from_col, from_col_off_px + size.width_px)
        to_row, to_row_off_px = self.advance_rows(from_row, from_row_off_px + size.height_px)
        return TwoCellAnchor(
            from_=AnchorCorner(from_col, from_col_off_px, from_row, from_row_off_px),
            to=AnchorCorner(to_col, to_col_off_px, to_row, to_row_off_px),
        )

    def absolute_x_px(self, corner: AnchorCorner) -> int:
        return sum(self.col_width_px(idx) for idx in range(corner.col)) + corner.col_off_px

    def absolute_y_px(self, corner: AnchorCorner) -> int:
        return sum(self.row_height_px(idx) for idx in range(corner.row)) + corner.row_off_px

    def split_col(self, col: int, offset_px: int) -> tuple[int, int]:
        while offset_px >= self.col_width_px(col):
            offset_px -= self.col_width_px(col)
            col += 1
        return col, offset_px

    def split_row(self, row: int, offset_px: int) -> tuple[int, int]:
        while offset_px >= self.row_height_px(row):
            offset_px -= self.row_height_px(row)
            row += 1
        return row, offset_px

    def advance_cols(self, col: int, width_px: int) -> tuple[int, int]:
        return self.split_col(col, width_px)

    def advance_rows(self, row: int, height_px: int) -> tuple[int, int]:
        return self.split_row(row, height_px)

    def col_width_px(self, _col: int) -> int:
        return self.default_col_width_px

    def row_height_px(self, _row: int) -> int:
        return self.default_row_height_px
