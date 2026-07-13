from __future__ import annotations

from pathlib import Path

from .grid import ExcelGridMetrics
from .ooxml_patch import patch_first_drawing_anchor
from .vml import patch_vml_shape, read_vml_shell_size


def read_vml_size_px(vml_text: str) -> tuple[int, int]:
    from .vml import read_vml_size_px as _read_vml_size_px

    return _read_vml_size_px(vml_text)


def patch_shell_geometry(
    package_dir: Path,
    *,
    start_col: int,
    start_row: int,
    offset_x_px: int,
    offset_y_px: int,
    width_px: int | None,
    height_px: int | None,
    preserve_aspect: bool,
    center_in_cell: bool = False,
) -> None:
    sheet_path = package_dir / "xl" / "worksheets" / "sheet1.xml"
    drawing_path = package_dir / "xl" / "drawings" / "drawing1.xml"
    vml_path = package_dir / "xl" / "drawings" / "vmlDrawing1.vml"

    vml_text = vml_path.read_text(encoding="utf-8")
    shell_size = read_vml_shell_size(vml_text)
    target_size = shell_size.fit(
        width_px=width_px,
        height_px=height_px,
        preserve_aspect=preserve_aspect,
    )

    # Use the worksheet's real col/row sizes (after layout copy), not defaults.
    grid = ExcelGridMetrics.from_worksheet(sheet_path)

    # Center the OLE inside the target cell when no explicit in-cell offset.
    # shell_large makes A very wide / row 1 very tall, so (0,0) looks "stuck"
    # in the top-left corner of the cell.
    if center_in_cell and offset_x_px == 0 and offset_y_px == 0:
        cell_w = grid.col_width_px(start_col - 1)
        cell_h = grid.row_height_px(start_row - 1)
        offset_x_px = max(0, (cell_w - target_size.width_px) // 2)
        offset_y_px = max(0, (cell_h - target_size.height_px) // 2)

    anchor = grid.anchor_for_box(
        start_col=start_col,
        start_row=start_row,
        offset_x_px=offset_x_px,
        offset_y_px=offset_y_px,
        size=target_size,
    )

    for path in [sheet_path, drawing_path]:
        patch_first_drawing_anchor(path, anchor)

    patch_vml_shape(
        vml_path,
        vml_text=vml_text,
        anchor=anchor,
        left_px=grid.absolute_x_px(anchor.from_),
        top_px=grid.absolute_y_px(anchor.from_),
        size=target_size,
    )
