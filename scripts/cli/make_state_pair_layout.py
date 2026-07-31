#!/usr/bin/env python3
"""Build a two-row state layout: OLE on the left (A1:A2), labels on the right.

Layout::

    OLE | 基态
    OLE | 激发态 s1

The OLE display box preserves the molecule's aspect ratio and is centered in
the square-ish A1+A2 area, filling ``--fill-ratio`` of the limiting side.

Built on the retain_cdx public API: openpyxl prepares the blank two-row grid
(row heights / column widths live in plain worksheet XML), then
``ExcelCDXWriter`` inserts the ChemDraw OLE object.

Usage::

    python3 -m scripts.cli.make_state_pair_layout --cdx examples/san.cdx -o state_pair.xlsx
"""

from __future__ import annotations

import argparse
from pathlib import Path

from retain_cdx import ExcelCDXWriter, Placement, ShellTemplate, StructureSource

DEFAULT_SHEET = "Sheet1"
POINTS_PER_PIXEL = 0.75
DEFAULT_A_COLUMN_PIXELS = 180
DEFAULT_ROW_PIXELS = 90
DEFAULT_B_COLUMN_WIDTH = 21.375
DEFAULT_FILL_RATIO = 0.9
DEFAULT_ASPECT = 1.0  # fallback when the CDX exposes no bounding box


def _excel_width_to_pixels(width: float) -> int:
    return int(((256 * width + int(128 / 7)) / 256) * 7)


def _pixels_to_excel_width(pixels: int) -> float:
    for raw in range(0, 256 * 100):
        width = raw / 256
        if _excel_width_to_pixels(width) >= pixels:
            return width
    raise ValueError(f"Unable to represent pixel width: {pixels}")


def centered_box(
    aspect: float,
    *,
    area_width_px: int,
    area_height_px: int,
    fill_ratio: float,
) -> tuple[int, int, int, int]:
    """Return (width, height, offset_x, offset_y) in px, centered in the area."""
    width = area_width_px * fill_ratio
    height = width / aspect
    if height > area_height_px * fill_ratio:
        height = area_height_px * fill_ratio
        width = height * aspect
    width_px = max(1, round(width))
    height_px = max(1, round(height))
    offset_x = max(0, round((area_width_px - width_px) / 2))
    offset_y = max(0, round((area_height_px - height_px) / 2))
    return width_px, height_px, offset_x, offset_y


def new_base_workbook(
    path: Path,
    *,
    a_column_pixels: int,
    row_pixels: int,
    b_column_width: float = DEFAULT_B_COLUMN_WIDTH,
    sheet: str = DEFAULT_SHEET,
) -> Path:
    try:
        from openpyxl import Workbook
    except ImportError as exc:  # pragma: no cover
        raise SystemExit("openpyxl is required to create the output workbook.") from exc
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet
    worksheet.column_dimensions["A"].width = _pixels_to_excel_width(a_column_pixels)
    worksheet.column_dimensions["B"].width = b_column_width
    for row in (1, 2):
        worksheet.row_dimensions[row].height = row_pixels * POINTS_PER_PIXEL
    workbook.save(path)
    return path


def build_state_pair_layout(
    cdx_path: Path,
    output: Path,
    *,
    top_label: str,
    bottom_label: str,
    template: Path | None,
    a_column_pixels: int,
    row_pixels: int,
    fill_ratio: float,
) -> None:
    source = StructureSource.from_cdx(cdx_path)
    aspect = source.content_aspect or source.page_aspect or DEFAULT_ASPECT

    width_px, height_px, offset_x, offset_y = centered_box(
        aspect,
        area_width_px=a_column_pixels,
        area_height_px=row_pixels * 2,
        fill_ratio=fill_ratio,
    )

    if output.parent != Path(""):
        output.parent.mkdir(parents=True, exist_ok=True)
    base = output.with_name(f".{output.stem}_base.tmp.xlsx")
    new_base_workbook(
        base, a_column_pixels=a_column_pixels, row_pixels=row_pixels
    )
    try:
        with ExcelCDXWriter(base) as writer:
            writer.insert_cdx(
                source,
                sheet=DEFAULT_SHEET,
                placement=Placement("A1", offset_x=offset_x, offset_y=offset_y),
                shell=ShellTemplate(xlsx=template) if template else None,
                width_px=width_px,
                height_px=height_px,
                # Same value openpyxl already wrote; keeps row 1's height
                # stable while the generated path pins its layout.
                row_height=row_pixels * POINTS_PER_PIXEL,
                col_width=None,
                preview_no_margin=True,
            )
            writer.set_cell_texts(
                DEFAULT_SHEET, {"B1": top_label, "B2": bottom_label}
            )
            writer.save(output)
    finally:
        base.unlink(missing_ok=True)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build a two-row state layout: OLE on the left, labels on the right.",
    )
    parser.add_argument("--cdx", required=True, help="Input CDX file")
    parser.add_argument("-o", "--output", default="state_pair.xlsx", help="Output XLSX")
    parser.add_argument("--top-label", default="基态", help="Top-right label")
    parser.add_argument("--bottom-label", default="激发态 s1", help="Bottom-right label")
    parser.add_argument(
        "--a-column-pixels",
        type=int,
        default=DEFAULT_A_COLUMN_PIXELS,
        help="Width of column A in pixels (default: 180)",
    )
    parser.add_argument(
        "--row-pixels",
        type=int,
        default=DEFAULT_ROW_PIXELS,
        help="Height of each of rows 1 and 2 in pixels (default: 90)",
    )
    parser.add_argument(
        "--template",
        default=None,
        help="Native Excel/ChemDraw OLE shell workbook to reuse (optional)",
    )
    parser.add_argument(
        "--fill-ratio",
        type=float,
        default=DEFAULT_FILL_RATIO,
        help="Fraction of the A1+A2 box to fill on the limiting side while "
             "preserving aspect ratio (default: 0.9)",
    )
    args = parser.parse_args()

    cdx_path = Path(args.cdx)
    if not cdx_path.exists():
        raise SystemExit(f"CDX not found: {cdx_path}")
    template = Path(args.template) if args.template else None
    if template is not None and not template.exists():
        raise SystemExit(f"Template not found: {template}")

    build_state_pair_layout(
        cdx_path,
        Path(args.output),
        top_label=args.top_label,
        bottom_label=args.bottom_label,
        template=template,
        a_column_pixels=args.a_column_pixels,
        row_pixels=args.row_pixels,
        fill_ratio=args.fill_ratio,
    )
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
