"""Insert pipeline: plan is already built; assemble source then place into target.

Steps (current architecture)::

  1. resolve target sheet in user package_dir
     (two-point placement resolves width/height against its real grid here)
  2. write temporary CDX files (optional normalize)
  3. _source: 在 staging 目录里装齐零件（bin + preview + surface XML）
  4. _parts:  merge source parts into package_dir
  5. touch row/col layout when not using shell
  6. ensure content types for drawing / media
"""

from __future__ import annotations

import tempfile
from dataclasses import dataclass, replace
from pathlib import Path

from retain_cdx.structure import StructureSource

from retain_cdx.office.excel._ooxml import (
    ensure_ole_drawing_content_types,
    parse_workbook_sheets,
    sheet_index_from_path,
    touch_sheet_layout,
)
from retain_cdx.office.excel._parts import merge_source_workbook
from retain_cdx.office.excel._parts.anchor_overflow import normalize_anchor_overflow
from retain_cdx.office.excel._grid import ExcelGridMetrics
from retain_cdx.office.excel._source import build_source_workbook
from retain_cdx.office.excel._utils import parse_cell, px_to_points

from .cdx import write_cdx_sources
from .plan import InsertPlan


@dataclass(frozen=True)
class TargetSheet:
    name: str
    path: Path
    index: int
    prefix: str


def execute_insert(
    package_dir: Path,
    sources: list[StructureSource],
    *,
    sheet: str,
    plan: InsertPlan,
) -> None:
    # 1. target sheet in the already-unzipped user workbook
    target_sheet = resolve_target_sheet(package_dir, sheet)

    # 1b. two-point placement → concrete width/height against the real grid
    plan = resolve_two_point_layout(plan, target_sheet.path)

    with tempfile.TemporaryDirectory(prefix="retain_cdx_source_") as tmp:
        tmpdir = Path(tmp)

        # 2. temporary CDX files (chem payload on disk)
        cdx_paths = write_cdx_sources(sources, tmpdir, plan)

        # 3. assemble the OLE parts in a staging directory (never zipped —
        #    packing then immediately re-extracting was pure overhead)
        source_dir = tmpdir / "source"
        build_source_workbook(cdx_paths, source_dir, plan.to_source_request())

        # 4. place: copy bin/media/surface from source → user package_dir
        drawing_name = merge_source_workbook(
            package_dir,
            source_dir,
            sheet_path=target_sheet.path,
            sheet_idx=target_sheet.index,
            prefix=target_sheet.prefix,
            preserve_object_pr=plan.has_shell,
            copy_layout=plan.has_shell,
        )

    # 5. layout (generated path only — shell may copy layout in merge).
    # Two-point placements never reshape the grid: the box was resolved
    # against the existing rows/cols and must keep matching them.
    layout = plan.layout
    if (
        not plan.has_shell
        and plan.placement.to_cell is None
        and (layout.row_height is not None or layout.col_width is not None)
    ):
        touch_sheet_layout(
            target_sheet.path,
            col=layout.col,
            start_row=layout.row,
            count=len(sources),
            row_height=layout.row_height,
            col_width=layout.col_width,
        )

    # 6. content types for OLE drawing / media parts
    ensure_ole_drawing_content_types(package_dir, drawing_name)

    # 7. walk any overflowing anchor offset into the cell it really lands in
    #    (no-op unless the box spans cells, e.g. a two-point placement)
    drawings = package_dir / "xl" / "drawings"
    normalize_anchor_overflow(
        target_sheet.path,
        drawings / drawing_name,
        next(iter(sorted(drawings.glob(f"{target_sheet.prefix}*.vml"))), None),
    )


def resolve_two_point_layout(plan: InsertPlan, sheet_path: Path) -> InsertPlan:
    """Resolve a two-point ``ExcelPlacement`` into concrete ``width_px``/``height_px``.

    两点锚：对照目标 sheet 的真实列宽/行高（未声明的列/行用 Excel 默认
    64px/19px），把 ``cell + offset → to_cell + to_offset`` 的跨度换算成
    像素框，并替换 ``LayoutSpec``。单锚点 placement 原样返回。

    The resolved size overrides ``width_px``/``height_px``/``ExcelDisplayOptions``
    and, for multi-object inserts, applies per row like ``width_px`` does.
    """
    placement = plan.placement
    if placement.to_cell is None:
        return plan

    metrics = ExcelGridMetrics.from_worksheet(sheet_path)
    from_col, from_row = parse_cell(placement.cell)
    to_col, to_row = parse_cell(placement.to_cell)

    x0 = metrics.cell_origin_x_px(from_col - 1) + placement.offset_x_px
    y0 = metrics.cell_origin_y_px(from_row - 1) + placement.offset_y_px
    x1 = metrics.cell_origin_x_px(to_col - 1) + placement.to_offset_x_px
    y1 = metrics.cell_origin_y_px(to_row - 1) + placement.to_offset_y_px

    width_px = x1 - x0
    height_px = y1 - y0
    if width_px <= 0 or height_px <= 0:
        raise ValueError(
            "Two-point placement resolves to an empty box "
            f"({width_px}x{height_px}px): to_cell={placement.to_cell!r} must be "
            f"below and to the right of cell={placement.cell!r} "
            "(to_cell marks the top-left corner of that cell)."
        )

    layout = replace(
        plan.layout,
        width_px=width_px,
        height_px=height_px,
        # Generated source builder needs a row height for its temporary
        # sheet; the target grid itself stays untouched (see execute_insert).
        row_height=(
            plan.layout.row_height
            if plan.has_shell
            else px_to_points(height_px)
        ),
    )
    return replace(plan, layout=layout)


def resolve_target_sheet(package_dir: Path, sheet: str) -> TargetSheet:
    sheets = parse_workbook_sheets(package_dir)
    if sheet not in sheets:
        raise ValueError(f"Sheet not found: {sheet}")

    sheet_path = package_dir / sheets[sheet]
    sheet_index = sheet_index_from_path(sheets[sheet])
    return TargetSheet(
        name=sheet,
        path=sheet_path,
        index=sheet_index,
        prefix=f"sheet{sheet_index}_retain",
    )
