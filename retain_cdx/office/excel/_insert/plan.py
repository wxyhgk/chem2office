from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from retain_cdx.render.preview import (
    PreviewBackend,
    PreviewFormat,
    PreviewStyle,
    resolve_preview_format,
)
from retain_cdx.structure import StructureSource

from retain_cdx.office.excel._utils import parse_cell, px_to_points
from retain_cdx.office.excel.models import (
    LayoutSpec,
    NormalizeSpec,
    PreviewSpec,
    ShellSourceOptions,
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
    preserve_aspect: bool = True,
    presentation_cache: bool | None = None,
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
        if display.preview.presentation_cache is not None:
            presentation_cache = display.preview.presentation_cache

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

    if presentation_cache is None:
        # 缺了渲染缓存，把结构复制到 PowerPoint 会唤起 ChemDraw 现场出图，
        # 软件会自己弹出来——所以矢量预览时默认挂上。位图预览没有可用的载荷
        # （CF_ENHMETAFILE 只吃 EMF），单为缓存再跑一遍矢量管线不值当，故不挂。
        presentation_cache = resolve_preview_format(preview_format) == "emf"

    boxes: tuple[tuple[int, int], ...] = ()
    if preserve_aspect and not two_point:
        # 逐个分子算——整批共用一个比例的话，除第一个以外全会被拉伸。
        # 高度为基准（表格行高齐整），宽度随分子变；宽到超出 width_px 才反过来受限。
        boxes = tuple(_fit_box(s.content_aspect, width_px, height_px) for s in sources)
        # 行高不动：调用方给什么就是什么。框只在 width_px×height_px 里缩放，
        # 不会长出去，所以没有把行撑高的理由；反过来按框改行高会把默认的
        # 110pt 悄悄压成 75pt，等于替调用方改了排版。

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
        boxes=boxes,
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
        presentation_cache=presentation_cache,
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


def _fit_box(aspect: float | None, width_px: int, height_px: int) -> tuple[int, int]:
    """把比例为 *aspect* 的分子塞进 (width_px, height_px)，不变形。

    显示框的比例必须等于 CDX 内容的比例，否则对象一旦被别的程序按 CDX 重新
    渲染（比如从 Excel 复制到 PowerPoint），就会被拉伸——Excel 里看不出来，
    因为预览图已经 letterbox 成了框的比例。
    """
    if not aspect or aspect <= 0:
        return width_px, height_px
    width = round(height_px * aspect)
    if width <= width_px:
        return max(1, width), height_px
    return width_px, max(1, round(width_px / aspect))


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
