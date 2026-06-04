from __future__ import annotations

import tempfile
from pathlib import Path

from retain_cdx.structure import StructureSource
from retain_cdx.office.excel._insert import build_insert_plan, write_cdx_sources
from retain_cdx.office.excel._ooxml import (
    ensure_ole_drawing_content_types,
    extract_xlsx,
    pack_xlsx,
    parse_workbook_sheets,
    sheet_index_from_path,
    touch_sheet_layout,
)
from retain_cdx.office.excel._parts import merge_source_workbook
from retain_cdx.office.excel.options import DisplayOptions, Placement, ShellTemplate
from retain_cdx.office.excel._source import build_source_workbook


class ExcelCDXWriter:
    """Insert editable ChemDraw CDX OLE objects into an XLSX workbook."""

    def __init__(self, input_xlsx: str | Path):
        self.input_xlsx = Path(input_xlsx)
        self._tmp = tempfile.TemporaryDirectory(prefix="retain_cdx_excel_")
        self.package_dir = Path(self._tmp.name) / "workbook"
        self.package_dir.mkdir()
        extract_xlsx(self.input_xlsx, self.package_dir)

    def close(self) -> None:
        self._tmp.cleanup()

    def __enter__(self) -> "ExcelCDXWriter":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def save(self, output_xlsx: str | Path) -> None:
        pack_xlsx(self.package_dir, Path(output_xlsx))

    def insert_cdx(
        self,
        source: StructureSource | str | Path,
        *,
        sheet: str,
        cell: str | None = None,
        placement: Placement | None = None,
        display: DisplayOptions | None = None,
        shell: ShellTemplate | str | Path | None = None,
        width_px: int = 145,
        height_px: int = 100,
        row_height: float | None = 110,
        col_width: float | None = 24,
        preview_width: int | None = None,
        preview_height: int | None = None,
        preview_thick_lines: bool = False,
        preview_no_margin: bool = False,
        normalize_cdx: bool = False,
        normalize_margin: float = 18,
        normalize_fill_ratio: float = 0.8,
        preserve_aspect: bool = False,
    ) -> None:
        """Insert one CDX object at *sheet*!*cell*."""
        if placement is None:
            if cell is None:
                raise ValueError("Either cell or placement is required.")
            placement = Placement(cell)
        cdx_source = (
            source
            if isinstance(source, StructureSource)
            else StructureSource.from_cdx(source)
        )
        self.insert_cdx_many(
            [cdx_source],
            sheet=sheet,
            start_cell=placement.cell,
            placement=placement,
            display=display,
            shell=shell,
            width_px=width_px,
            height_px=height_px,
            row_height=row_height,
            col_width=col_width,
            preview_width=preview_width,
            preview_height=preview_height,
            preview_thick_lines=preview_thick_lines,
            preview_no_margin=preview_no_margin,
            normalize_cdx=normalize_cdx,
            normalize_margin=normalize_margin,
            normalize_fill_ratio=normalize_fill_ratio,
            preserve_aspect=preserve_aspect,
        )

    def insert_cdx_many(
        self,
        sources: list[StructureSource],
        *,
        sheet: str,
        start_cell: str,
        placement: Placement | None = None,
        display: DisplayOptions | None = None,
        shell: ShellTemplate | str | Path | None = None,
        width_px: int = 145,
        height_px: int = 100,
        row_height: float | None = 110,
        col_width: float | None = 24,
        preview_width: int | None = None,
        preview_height: int | None = None,
        preview_thick_lines: bool = False,
        preview_no_margin: bool = False,
        normalize_cdx: bool = False,
        normalize_margin: float = 18,
        normalize_fill_ratio: float = 0.8,
        preserve_aspect: bool = False,
    ) -> None:
        """Insert CDX objects vertically starting at *start_cell*."""
        if not sources:
            return

        plan = build_insert_plan(
            sources,
            start_cell=start_cell,
            placement=placement,
            display=display,
            shell=shell,
            width_px=width_px,
            height_px=height_px,
            row_height=row_height,
            col_width=col_width,
            preview_width=preview_width,
            preview_height=preview_height,
            preview_thick_lines=preview_thick_lines,
            preview_no_margin=preview_no_margin,
            normalize_cdx=normalize_cdx,
            normalize_margin=normalize_margin,
            normalize_fill_ratio=normalize_fill_ratio,
            preserve_aspect=preserve_aspect,
        )
        sheets = parse_workbook_sheets(self.package_dir)
        if sheet not in sheets:
            raise ValueError(f"Sheet not found: {sheet}")

        sheet_path = self.package_dir / sheets[sheet]
        sheet_idx = sheet_index_from_path(sheets[sheet])
        prefix = f"sheet{sheet_idx}_retain"

        with tempfile.TemporaryDirectory(prefix="retain_cdx_source_") as tmp:
            tmpdir = Path(tmp)
            cdx_paths = write_cdx_sources(sources, tmpdir, plan)
            source_xlsx = tmpdir / "source_ole.xlsx"
            build_source_workbook(cdx_paths, source_xlsx, plan)

            src_dir = tmpdir / "source"
            src_dir.mkdir()
            extract_xlsx(source_xlsx, src_dir)
            drawing_name = merge_source_workbook(
                self.package_dir,
                src_dir,
                sheet_path=sheet_path,
                sheet_idx=sheet_idx,
                prefix=prefix,
                preserve_object_pr=plan.has_shell,
                copy_layout=plan.has_shell,
            )

        if not plan.has_shell and (plan.row_height is not None or plan.col_width is not None):
            touch_sheet_layout(
                sheet_path,
                col=plan.col,
                start_row=plan.row,
                count=len(sources),
                row_height=plan.row_height,
                col_width=plan.col_width,
            )
        ensure_ole_drawing_content_types(self.package_dir, drawing_name)
