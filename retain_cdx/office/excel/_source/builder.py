from __future__ import annotations

from pathlib import Path

from retain_cdx.office.excel._insert.plan import InsertPlan
from retain_cdx.office.excel._utils import index_to_col

from .workbook import build_source_ole_workbook, build_source_workbook_from_shell


def build_source_workbook(
    cdx_paths: list[Path],
    output_xlsx: Path,
    plan: InsertPlan,
) -> None:
    if plan.shell_path is not None:
        build_source_workbook_from_shell(
            cdx_paths[0],
            plan.shell_path,
            output_xlsx,
            start_col=plan.col if plan.shell_allow_resize else None,
            start_row=plan.row if plan.shell_allow_resize else None,
            offset_x_px=plan.placement.offset_x_px,
            offset_y_px=plan.placement.offset_y_px,
            width_px=plan.width_px if plan.shell_resize_display else None,
            height_px=plan.height_px if plan.shell_resize_display else None,
            preserve_aspect=plan.shell_preserve_aspect,
            fit_cdx=plan.shell_fit_cdx,
            preview_width=plan.preview_width,
            preview_height=plan.preview_height,
            preview_thick_lines=plan.preview_thick_lines,
            preview_no_margin=plan.preview_no_margin,
        )
        return

    build_source_ole_workbook(
        cdx_paths,
        output_xlsx,
        col_letter=index_to_col(plan.col),
        start_row=plan.row,
        offset_x_px=plan.placement.offset_x_px,
        offset_y_px=plan.placement.offset_y_px,
        width_px=plan.width_px,
        height_px=plan.height_px,
        row_height=plan.row_height,
        preview_width=plan.preview_width,
        preview_height=plan.preview_height,
        preview_thick_lines=plan.preview_thick_lines,
        preview_no_margin=plan.preview_no_margin,
    )
