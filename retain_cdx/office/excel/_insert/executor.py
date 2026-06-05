from __future__ import annotations

import tempfile
from dataclasses import dataclass
from pathlib import Path

from retain_cdx.structure import StructureSource

from retain_cdx.office.excel._ooxml import (
    ensure_ole_drawing_content_types,
    extract_xlsx,
    parse_workbook_sheets,
    sheet_index_from_path,
    touch_sheet_layout,
)
from retain_cdx.office.excel._parts import merge_source_workbook
from retain_cdx.office.excel._source import build_source_workbook

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
    target_sheet = resolve_target_sheet(package_dir, sheet)

    with tempfile.TemporaryDirectory(prefix="retain_cdx_source_") as tmp:
        tmpdir = Path(tmp)
        cdx_paths = write_cdx_sources(sources, tmpdir, plan)
        source_xlsx = tmpdir / "source_ole.xlsx"
        build_source_workbook(cdx_paths, source_xlsx, plan)

        source_dir = tmpdir / "source"
        source_dir.mkdir()
        extract_xlsx(source_xlsx, source_dir)
        drawing_name = merge_source_workbook(
            package_dir,
            source_dir,
            sheet_path=target_sheet.path,
            sheet_idx=target_sheet.index,
            prefix=target_sheet.prefix,
            preserve_object_pr=plan.has_shell,
            copy_layout=plan.has_shell,
        )

    if not plan.has_shell and (plan.row_height is not None or plan.col_width is not None):
        touch_sheet_layout(
            target_sheet.path,
            col=plan.col,
            start_row=plan.row,
            count=len(sources),
            row_height=plan.row_height,
            col_width=plan.col_width,
        )
    ensure_ole_drawing_content_types(package_dir, drawing_name)


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
