from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from retain_cdx.render.preview import PreviewBackend, PreviewFormat, PreviewStyle
from retain_cdx.structure import StructureSource

from retain_cdx.office.excel._utils import parse_cell, px_to_points
from retain_cdx.office.excel.models import (
    LayoutSpec,
    NormalizeSpec,
    PreviewSpec,
    ShellSourceOptions,
    SourceBuildRequest,
)
from retain_cdx.office.excel.options import ExcelDisplayOptions, ExcelPlacement, ExcelShellTemplate


@dataclass(frozen=True)
class InsertPlan:
    """Resolved insert inputs for the pipeline (plan → CDX files → source → merge)."""

    placement: ExcelPlacement
    layout: LayoutSpec
    preview: PreviewSpec
    normalize: NormalizeSpec
    shell: ShellSourceOptions | None

    @property
    def has_shell(self) -> bool:
        return self.shell is not None

    # --- thin accessors (pipeline / tests) ---
    @property
    def col(self) -> int:
        return self.layout.col

    @property
    def row(self) -> int:
        return self.layout.row

    @property
    def width_px(self) -> int:
        return self.layout.width_px

    @property
    def height_px(self) -> int:
        return self.layout.height_px

    @property
    def row_height(self) -> float | None:
        return self.layout.row_height

    @property
    def col_width(self) -> float | None:
        return self.layout.col_width

    def to_source_request(self) -> SourceBuildRequest:
        return SourceBuildRequest(
            layout=self.layout,
            preview=self.preview,
            shell=self.shell,
        )


def build_insert_plan(
    sources: list[StructureSource],
    *,
    start_cell: str,
    placement: ExcelPlacement | None,
    display: ExcelDisplayOptions | None,
    shell: ExcelShellTemplate | str | Path | None = None,
    # Ignored when *display* is given (it overrides all of them); the defaults
    # mirror ``ExcelCDXWriter.insert_cdx_many``.
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
    preview_format: PreviewFormat = "png",
    preview_mono: bool = True,
    preview_backend: PreviewBackend = "auto",
    preview_style: PreviewStyle = "acs1996",
) -> InsertPlan:
    if placement is None:
        placement = ExcelPlacement(start_cell)

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

    # Two-point anchor: the display box spans the target sheet's real grid.
    # Width/height are resolved later by the executor (it can read the target
    # worksheet); do not reshape the target rows/cols here — the grid itself
    # defines the box. 两点锚：框由现有网格定义，不再动目标行高/列宽。
    two_point = placement.to_cell is not None
    if two_point:
        parse_cell(placement.to_cell)  # fail fast on malformed refs
        row_height = None
        col_width = None

    if shell_path is not None and display is None:
        row_height = None
        col_width = None
    if row_height is None and shell_path is None and not two_point:
        row_height = px_to_points(height_px)

    if preserve_aspect and not two_point:
        aspect = sources[0].content_aspect
        if aspect and aspect > 0:
            height_px = round(width_px / aspect)
            if row_height is not None:
                row_height = px_to_points(height_px)

    col, row = parse_cell(placement.cell)
    layout = LayoutSpec(
        col=col,
        row=row,
        width_px=width_px,
        height_px=height_px,
        row_height=row_height,
        col_width=col_width,
        offset_x_px=placement.offset_x_px,
        offset_y_px=placement.offset_y_px,
    )
    preview = PreviewSpec(
        width=preview_width,
        height=preview_height,
        thick_lines=preview_thick_lines,
        no_margin=preview_no_margin,
        mono=preview_mono,
        format=preview_format,
        backend=preview_backend,
        style=preview_style,
    )
    normalize = NormalizeSpec(
        enabled=normalize_cdx,
        margin=normalize_margin,
        fill_ratio=normalize_fill_ratio,
    )
    shell_opts: ShellSourceOptions | None = None
    if shell_path is not None:
        shell_opts = ShellSourceOptions(
            path=shell_path,
            preserve_aspect=shell_preserve_aspect(shell),
            allow_resize=shell_allow_resize(shell),
            resize_display=shell_allow_resize(shell)
            and (display is not None or two_point),
            fit_cdx=shell_fit_cdx(shell),
        )

    return InsertPlan(
        placement=placement,
        layout=layout,
        preview=preview,
        normalize=normalize,
        shell=shell_opts,
    )


def shell_template_path(shell: ExcelShellTemplate | str | Path | None) -> Path | None:
    if shell is None:
        return None
    if isinstance(shell, ExcelShellTemplate):
        return Path(shell.xlsx)
    return Path(shell)


def shell_preserve_aspect(shell: ExcelShellTemplate | str | Path | None) -> bool:
    if isinstance(shell, ExcelShellTemplate):
        return shell.preserve_aspect
    return True


def shell_allow_resize(shell: ExcelShellTemplate | str | Path | None) -> bool:
    if isinstance(shell, ExcelShellTemplate):
        return shell.allow_resize
    return False


def shell_fit_cdx(shell: ExcelShellTemplate | str | Path | None) -> bool:
    if isinstance(shell, ExcelShellTemplate):
        return shell.fit_cdx
    return False
