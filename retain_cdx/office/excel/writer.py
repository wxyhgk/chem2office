from __future__ import annotations

import tempfile
from collections.abc import Iterable, Mapping
from pathlib import Path

from retain_cdx.render.preview import PreviewBackend, PreviewFormat, PreviewStyle
from retain_cdx.structure import StructureSource
from retain_cdx.office.excel._insert import build_insert_plan, execute_insert
from retain_cdx.office.excel._insert.cdx import write_cdx_sources
from retain_cdx.office.excel._insert.executor import (
    resolve_target_sheet,
    resolve_two_point_layout,
)
from retain_cdx.office.excel._picture import insert_structure_pictures
from retain_cdx.office.excel._ooxml import (
    extract_xlsx,
    pack_xlsx,
    parse_workbook_sheets,
    set_inline_cell_text,
    touch_sheet_layout,
)
from retain_cdx.office.excel._utils import col_to_index
from retain_cdx.office.excel.options import DisplayOptions, Placement, ShellTemplate


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

    def insert_pictures_many(
        self,
        sources: Iterable[StructureSource | str | Path],
        *,
        sheet: str,
        start_cell: str,
        display: DisplayOptions | None = None,
        placement: Placement | None = None,
    ) -> None:
        """Insert structures as plain pictures, one per row, **not** as OLE.

        Use this when a sheet needs hundreds or thousands of structures: an OLE
        object carries a CFBF container, a VML shape and activation machinery,
        and Excel slows to a crawl in the thousands, while pictures stay cheap.
        The trade-off is that pictures cannot be opened in ChemDraw by
        double-click — use :meth:`insert_cdx_many` for rows that must stay
        editable.
        """
        cdx_sources = [
            source
            if isinstance(source, StructureSource)
            else StructureSource.from_cdx(source)
            for source in sources
        ]
        if not cdx_sources:
            return

        options = display or DisplayOptions()
        plan = build_insert_plan(
            cdx_sources,
            start_cell=start_cell,
            placement=placement,
            display=options,
        )
        target_sheet = resolve_target_sheet(self.package_dir, sheet)
        plan = resolve_two_point_layout(plan, target_sheet.path)

        if plan.layout.row_height is not None or plan.layout.col_width is not None:
            if plan.placement.to_cell is None:
                touch_sheet_layout(
                    target_sheet.path,
                    col=plan.layout.col,
                    start_row=plan.layout.row,
                    count=len(cdx_sources),
                    row_height=plan.layout.row_height,
                    col_width=plan.layout.col_width,
                )

        with tempfile.TemporaryDirectory(prefix="retain_cdx_pictures_") as tmp:
            cdx_paths = write_cdx_sources(cdx_sources, Path(tmp), plan)
            insert_structure_pictures(
                self.package_dir,
                cdx_paths,
                sheet_path=target_sheet.path,
                prefix=target_sheet.prefix,
                layout=plan.layout,
                preview=plan.preview,
            )

    def set_cell_text(self, sheet: str, cell: str, text: str) -> None:
        """Write a plain text value (inline string) into *sheet*!*cell*."""
        self.set_cell_texts(sheet, {cell: text})

    def set_cell_texts(self, sheet: str, cells: Mapping[str, str]) -> None:
        """Write plain text values, e.g. ``{"A1": "编号", "A2": "D0001"}``.

        Use this (not openpyxl, which drops OLE parts) to put IDs or headers
        next to inserted structures.
        """
        set_inline_cell_text(self._sheet_path(sheet), cells)

    def set_column_width(self, sheet: str, col: str | int, width: float) -> None:
        """Set a column width (Excel character units); *col* is ``"A"`` or 1."""
        col_index = col_to_index(col) if isinstance(col, str) else int(col)
        touch_sheet_layout(
            self._sheet_path(sheet),
            col=col_index,
            start_row=1,
            count=0,
            row_height=None,
            col_width=width,
        )

    def set_row_height(
        self, sheet: str, row: int, height: float, *, count: int = 1
    ) -> None:
        """Set custom heights (points) for *count* rows starting at *row*.

        Safe on workbooks that already carry OLE parts (unlike an openpyxl
        round-trip, which drops them).
        """
        touch_sheet_layout(
            self._sheet_path(sheet),
            col=1,  # unused by the row-height branch
            start_row=row,
            count=count,
            row_height=height,
            col_width=None,
        )

    def _sheet_path(self, sheet: str) -> Path:
        sheets = parse_workbook_sheets(self.package_dir)
        if sheet not in sheets:
            raise ValueError(f"Sheet not found: {sheet!r}")
        return self.package_dir / sheets[sheet]

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
        preview_format: PreviewFormat = "png",
        preview_mono: bool = True,
        preview_backend: PreviewBackend = "auto",
        preview_style: PreviewStyle = "acs1996",
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
            preview_format=preview_format,
            preview_mono=preview_mono,
            preview_backend=preview_backend,
            preview_style=preview_style,
            normalize_cdx=normalize_cdx,
            normalize_margin=normalize_margin,
            normalize_fill_ratio=normalize_fill_ratio,
            preserve_aspect=preserve_aspect,
        )

    def insert_cdx_many(
        self,
        sources: Iterable[StructureSource | str | Path],
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
        preview_format: PreviewFormat = "png",
        preview_mono: bool = True,
        preview_backend: PreviewBackend = "auto",
        preview_style: PreviewStyle = "acs1996",
        normalize_cdx: bool = False,
        normalize_margin: float = 18,
        normalize_fill_ratio: float = 0.8,
        preserve_aspect: bool = False,
    ) -> None:
        """Insert CDX objects vertically starting at *start_cell*."""
        cdx_sources = [
            source
            if isinstance(source, StructureSource)
            else StructureSource.from_cdx(source)
            for source in sources
        ]
        if not cdx_sources:
            return

        plan = build_insert_plan(
            cdx_sources,
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
            preview_format=preview_format,
            preview_mono=preview_mono,
            preview_backend=preview_backend,
            preview_style=preview_style,
            normalize_cdx=normalize_cdx,
            normalize_margin=normalize_margin,
            normalize_fill_ratio=normalize_fill_ratio,
            preserve_aspect=preserve_aspect,
        )
        execute_insert(self.package_dir, cdx_sources, sheet=sheet, plan=plan)
