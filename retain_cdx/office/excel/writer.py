from __future__ import annotations

import tempfile
from collections.abc import Iterable
from pathlib import Path

from retain_cdx.render.preview import PreviewBackend, PreviewFormat, PreviewStyle
from retain_cdx.structure import StructureSource
from retain_cdx.office.excel._insert import build_insert_plan, execute_insert
from retain_cdx.office.excel._ooxml import extract_xlsx, pack_xlsx
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
