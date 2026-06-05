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

    grid = ExcelGridMetrics()
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
