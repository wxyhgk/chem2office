from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from retain_cdx.render.preview import PreviewBackend, PreviewFormat, PreviewStyle
from retain_cdx.structure import StructureSource

from retain_cdx.office.excel._utils import parse_cell, px_to_points
from retain_cdx.office.excel.options import DisplayOptions, Placement, ShellTemplate


@dataclass(frozen=True)
class InsertPlan:
    placement: Placement
    col: int
    row: int
    width_px: int
    height_px: int
    row_height: float | None
    col_width: float | None
    preview_width: int | None
    preview_height: int | None
    preview_thick_lines: bool
    preview_no_margin: bool
    preview_format: PreviewFormat
    preview_mono: bool
    preview_backend: PreviewBackend
    preview_style: PreviewStyle
    normalize_cdx: bool
    normalize_margin: float
    normalize_fill_ratio: float
    shell_path: Path | None
    shell_preserve_aspect: bool
    shell_allow_resize: bool
    shell_resize_display: bool
    shell_fit_cdx: bool

    @property
    def has_shell(self) -> bool:
        return self.shell_path is not None


def build_insert_plan(
    sources: list[StructureSource],
    *,
    start_cell: str,
    placement: Placement | None,
    display: DisplayOptions | None,
    shell: ShellTemplate | str | Path | None,
    width_px: int,
    height_px: int,
    row_height: float | None,
    col_width: float | None,
    preview_width: int | None,
    preview_height: int | None,
    preview_thick_lines: bool,
    preview_no_margin: bool,
    normalize_cdx: bool,
    normalize_margin: float,
    normalize_fill_ratio: float,
    preserve_aspect: bool,
    preview_format: PreviewFormat = "png",
    preview_mono: bool = True,
    preview_backend: PreviewBackend = "auto",
    preview_style: PreviewStyle = "acs1996",
) -> InsertPlan:
    if placement is None:
        placement = Placement(start_cell)

    shell_path = shell_template_path(shell)
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
        preview_format = display.preview.format
        preview_mono = display.preview.mono
        preview_backend = display.preview.backend
        preview_style = display.preview.style
        normalize_cdx = display.normalize_cdx
        normalize_margin = display.normalize_margin
        normalize_fill_ratio = display.normalize_fill_ratio
        preserve_aspect = display.preserve_aspect

    if shell_path is not None and display is None:
        row_height = None
        col_width = None
    if row_height is None and shell_path is None:
        row_height = px_to_points(height_px)

    if preserve_aspect:
        aspect = sources[0].content_aspect
        if aspect and aspect > 0:
            height_px = round(width_px / aspect)
            if row_height is not None:
                row_height = height_px * 0.75

    col, row = parse_cell(placement.cell)
    return InsertPlan(
        placement=placement,
        col=col,
        row=row,
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
        shell_path=shell_path,
        shell_preserve_aspect=shell_preserve_aspect(shell),
        shell_allow_resize=shell_allow_resize(shell),
        shell_resize_display=shell_allow_resize(shell) and display is not None,
        shell_fit_cdx=shell_fit_cdx(shell),
    )


def shell_template_path(shell: ShellTemplate | str | Path | None) -> Path | None:
    if shell is None:
        return None
    if isinstance(shell, ShellTemplate):
        return Path(shell.xlsx)
    return Path(shell)


def shell_preserve_aspect(shell: ShellTemplate | str | Path | None) -> bool:
    if isinstance(shell, ShellTemplate):
        return shell.preserve_aspect
    return True


def shell_allow_resize(shell: ShellTemplate | str | Path | None) -> bool:
    if isinstance(shell, ShellTemplate):
        return shell.allow_resize
    return False


def shell_fit_cdx(shell: ShellTemplate | str | Path | None) -> bool:
    if isinstance(shell, ShellTemplate):
        return shell.fit_cdx
    return False
