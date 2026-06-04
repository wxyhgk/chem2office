from __future__ import annotations

import tempfile
from pathlib import Path

from retain_cdx.structure import StructureSource
from retain_cdx.office.excel.address import index_to_col, parse_cell
from retain_cdx.office.excel.content_types import ensure_ole_drawing_content_types
from retain_cdx.office.excel.ole_parts import merge_source_workbook
from retain_cdx.office.excel.options import DisplayOptions, Placement, ShellTemplate
from retain_cdx.office.excel.package import extract_xlsx, pack_xlsx
from retain_cdx.office.excel.source_workbook import (
    build_source_ole_workbook,
    build_source_workbook_from_shell,
)
from retain_cdx.office.excel.units import px_to_points
from retain_cdx.office.excel.worksheet import touch_sheet_layout
from retain_cdx.office.excel.workbook import parse_workbook_sheets, sheet_index_from_path


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

        if placement is None:
            placement = Placement(start_cell)
        shell_path = _shell_path(shell)
        if shell_path is not None and len(sources) != 1:
            raise ValueError("Shell templates currently support one CDX object per insert.")
        if display is not None:
            width_px = display.width_px
            height_px = display.height_px
            row_height = display.row_height
            col_width = display.col_width
            preview_width = display.preview.width
            preview_height = display.preview.height
            preview_thick_lines = display.preview.thick_lines
            preview_no_margin = display.preview.no_margin
            normalize_cdx = display.normalize_cdx
            normalize_margin = display.normalize_margin
            normalize_fill_ratio = display.normalize_fill_ratio
            preserve_aspect = display.preserve_aspect
        if shell_path is not None and display is None:
            row_height = None
            col_width = None
        if row_height is None and shell_path is None:
            row_height = px_to_points(height_px)

        col, row = parse_cell(placement.cell)
        sheets = parse_workbook_sheets(self.package_dir)
        if sheet not in sheets:
            raise ValueError(f"Sheet not found: {sheet}")

        sheet_path = self.package_dir / sheets[sheet]
        sheet_idx = sheet_index_from_path(sheets[sheet])
        prefix = f"sheet{sheet_idx}_retain"
        if preserve_aspect:
            aspect = sources[0].content_aspect
            if aspect and aspect > 0:
                height_px = round(width_px / aspect)
                if row_height is not None:
                    row_height = height_px * 0.75

        with tempfile.TemporaryDirectory(prefix="retain_cdx_source_") as tmp:
            tmpdir = Path(tmp)
            cdx_paths = []
            for idx, source in enumerate(sources, start=1):
                if normalize_cdx:
                    source = source.normalized(
                        margin=normalize_margin,
                        target_aspect=width_px / height_px,
                        target_width=width_px * 0.75,
                        target_height=height_px * 0.75,
                        fill_ratio=normalize_fill_ratio,
                    )
                path = tmpdir / f"source_{idx}.cdx"
                source.write_cdx(path)
                cdx_paths.append(path)

            source_xlsx = tmpdir / "source_ole.xlsx"
            if shell_path is not None:
                build_source_workbook_from_shell(
                    cdx_paths[0],
                    shell_path,
                    source_xlsx,
                    start_col=col if _shell_allow_resize(shell) else None,
                    start_row=row if _shell_allow_resize(shell) else None,
                    offset_x_px=placement.offset_x_px,
                    offset_y_px=placement.offset_y_px,
                    width_px=width_px if _shell_allow_resize(shell) and display is not None else None,
                    height_px=height_px if _shell_allow_resize(shell) and display is not None else None,
                    preserve_aspect=_shell_preserve_aspect(shell),
                    fit_cdx=_shell_fit_cdx(shell),
                )
            else:
                build_source_ole_workbook(
                    cdx_paths,
                    source_xlsx,
                    col_letter=index_to_col(col),
                    start_row=row,
                    offset_x_px=placement.offset_x_px,
                    offset_y_px=placement.offset_y_px,
                    width_px=width_px,
                    height_px=height_px,
                    row_height=row_height,
                    preview_width=preview_width,
                    preview_height=preview_height,
                    preview_thick_lines=preview_thick_lines,
                    preview_no_margin=preview_no_margin,
                )

            src_dir = tmpdir / "source"
            src_dir.mkdir()
            extract_xlsx(source_xlsx, src_dir)
            drawing_name = merge_source_workbook(
                self.package_dir,
                src_dir,
                sheet_path=sheet_path,
                sheet_idx=sheet_idx,
                prefix=prefix,
                preserve_object_pr=shell_path is not None,
                copy_layout=shell_path is not None,
            )

        if shell_path is None and (row_height is not None or col_width is not None):
            touch_sheet_layout(
                sheet_path,
                col=col,
                start_row=row,
                count=len(sources),
                row_height=row_height,
                col_width=col_width,
            )
        ensure_ole_drawing_content_types(self.package_dir, drawing_name)


def _shell_path(shell: ShellTemplate | str | Path | None) -> Path | None:
    if shell is None:
        return None
    if isinstance(shell, ShellTemplate):
        return Path(shell.xlsx)
    return Path(shell)


def _shell_preserve_aspect(shell: ShellTemplate | str | Path | None) -> bool:
    if isinstance(shell, ShellTemplate):
        return shell.preserve_aspect
    return True


def _shell_allow_resize(shell: ShellTemplate | str | Path | None) -> bool:
    if isinstance(shell, ShellTemplate):
        return shell.allow_resize
    return False


def _shell_fit_cdx(shell: ShellTemplate | str | Path | None) -> bool:
    if isinstance(shell, ShellTemplate):
        return shell.fit_cdx
    return False
