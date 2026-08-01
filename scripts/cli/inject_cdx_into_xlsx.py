#!/usr/bin/env python3
"""Inject CDX files into an Excel workbook as editable ChemDraw OLE objects.

One OLE object per row: labels (file stems) in column A, structures in the
chosen column (default B). Built on the retain_cdx public API.

Usage::

    python3 -m scripts.cli.inject_cdx_into_xlsx --cdx tests/fixtures/san.cdx -o output.xlsx
    python3 -m scripts.cli.inject_cdx_into_xlsx --cdx mol.cdx --col C --start-row 2 -o output.xlsx
    python3 -m scripts.cli.inject_cdx_into_xlsx --cdx mol.cdx --from A1+5mm,3mm -o output.xlsx
    python3 -m scripts.cli.inject_cdx_into_xlsx --cdx mol.cdx --from A1 --to C4+2mm,10pt -o output.xlsx

Notes on the public-API port:

- ``--to`` defines the display box with a second anchor (``ExcelPlacement.to_cell``)
  and overrides ``--width``/``--height``; the span is resolved against the
  sheet's real column widths / row heights.
- ``--stream`` was removed: the writer always uses the ChemDraw ``CONTENTS``
  stream.
- ``--template`` now reuses the given workbook as a native ChemDraw OLE
  *shell* (``ExcelShellTemplate``); shell inserts support one CDX at a time.
- Anchor offsets accept ``mm``/``pt``/``in``/``px`` (``emu`` is gone).
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from retain_cdx import (
    ExcelDisplayOptions,
    ExcelCDXWriter,
    ExcelPlacement,
    PreviewOptions,
    ExcelShellTemplate,
    StructureSource,
)

DEFAULT_SHEET = "Sheet1"

_ANCHOR_RE = re.compile(
    r"^([A-Za-z]+\d+)"
    r"(?:\+([0-9]*\.?[0-9]+(?:mm|pt|in|px)),([0-9]*\.?[0-9]+(?:mm|pt|in|px)))?$",
    re.I,
)


def parse_anchor_spec(spec: str, *, flag: str = "--from") -> tuple[str, str, str]:
    """Parse ``A1`` or ``A1+5mm,3mm`` into (cell, offset_x, offset_y)."""
    match = _ANCHOR_RE.match(spec.strip())
    if not match:
        raise SystemExit(
            f"Invalid {flag} spec: {spec!r}. Expected 'A1' or 'A1+5mm,3mm' "
            "(units: mm/pt/in/px)."
        )
    return match.group(1).upper(), match.group(2) or "0mm", match.group(3) or "0mm"


def new_blank_workbook(path: Path, sheet: str = DEFAULT_SHEET) -> Path:
    try:
        from openpyxl import Workbook
    except ImportError as exc:  # pragma: no cover
        raise SystemExit("openpyxl is required to create the output workbook.") from exc
    workbook = Workbook()
    workbook.active.title = sheet
    workbook.save(path)
    return path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Inject CDX into Excel with ChemDraw OLE objects.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
examples:
  # Default (recommended):
  %(prog)s --cdx mol1.cdx mol2.cdx -o output.xlsx
  %(prog)s --cdx mol.cdx --col C --start-row 2 -o output.xlsx
  %(prog)s --cdx mol.cdx --from A1+0mm,40pt -o output.xlsx
  %(prog)s --cdx mol.cdx --from A1 --to C4 -o output.xlsx

  # Native shell template (stable OLE display; one CDX per run):
  %(prog)s --template template.xlsx --cdx mol.cdx -o output.xlsx
""",
    )
    parser.add_argument("--template", default=None,
                        help="Native Excel/ChemDraw OLE shell workbook to reuse")
    parser.add_argument("--cdx", required=True, nargs="+",
                        help="CDX file(s) to inject")
    parser.add_argument("-o", "--output", default="output.xlsx",
                        help="Output xlsx (default: output.xlsx)")

    pos = parser.add_argument_group("position control")
    pos.add_argument("--col", default="B",
                     help="Column for OLE objects (default: B)")
    pos.add_argument("--start-row", type=int, default=1,
                     help="Starting row, 1-indexed (default: 1)")
    pos.add_argument("--margin-left", default="12mm",
                     help="Left offset within cell (default: 12mm)")
    pos.add_argument("--margin-top", default="5mm",
                     help="Top offset within cell (default: 5mm)")
    pos.add_argument("--width", default="22mm",
                     help="OLE object width (default: 22mm)")
    pos.add_argument("--height", default="25mm",
                     help="OLE object height (default: 25mm)")
    pos.add_argument("--from", dest="from_spec", default=None,
                     help="Precise from-anchor, e.g. 'A1+5mm,3mm' "
                          "(overrides --col/--start-row/--margin-*)")
    pos.add_argument("--to", dest="to_spec", default=None,
                     help="Second anchor defining the box extent, e.g. 'C4' or "
                          "'C4+2mm,10pt' (top-left corner of that cell); "
                          "overrides --width/--height by spanning the real "
                          "sheet grid")
    pos.add_argument("--row-height", type=float, default=97.5,
                     help="Row height in points (default: 97.5)")

    args = parser.parse_args()

    cdx_paths = [Path(p) for p in args.cdx]
    for cdx_path in cdx_paths:
        if not cdx_path.exists():
            raise SystemExit(f"CDX not found: {cdx_path}")
    sources = [StructureSource.from_cdx(p) for p in cdx_paths]

    shell: ExcelShellTemplate | None = None
    if args.template:
        template = Path(args.template)
        if not template.exists():
            raise SystemExit(f"Template not found: {template}")
        if len(sources) != 1:
            raise SystemExit("--template (shell) mode supports exactly one --cdx file.")
        shell = ExcelShellTemplate(xlsx=template)

    if args.from_spec:
        cell, offset_x, offset_y = parse_anchor_spec(args.from_spec, flag="--from")
    else:
        cell = f"{args.col.upper()}{args.start_row}"
        offset_x, offset_y = args.margin_left, args.margin_top

    to_cell: str | None = None
    to_offset_x: str = "0mm"
    to_offset_y: str = "0mm"
    if args.to_spec:
        to_cell, to_offset_x, to_offset_y = parse_anchor_spec(
            args.to_spec, flag="--to"
        )

    display = ExcelDisplayOptions(
        width=args.width,
        height=args.height,
        row_height=args.row_height,
        col_width=None,
        preview=PreviewOptions(),
    )

    output = Path(args.output)
    if output.parent != Path(""):
        output.parent.mkdir(parents=True, exist_ok=True)
    blank = output.with_name(f".{output.stem}_blank.tmp.xlsx")
    new_blank_workbook(blank)
    try:
        with ExcelCDXWriter(blank) as writer:
            writer.insert_cdx_many(
                sources,
                sheet=DEFAULT_SHEET,
                start_cell=cell,
                placement=ExcelPlacement(
                    cell,
                    offset_x=offset_x,
                    offset_y=offset_y,
                    to_cell=to_cell,
                    to_offset_x=to_offset_x,
                    to_offset_y=to_offset_y,
                ),
                display=display,
                shell=shell,
            )
            start_row = int(re.sub(r"[A-Za-z]+", "", cell))
            writer.set_cell_texts(
                DEFAULT_SHEET,
                {f"A{start_row + idx}": src.name for idx, src in enumerate(sources)},
            )
            writer.save(output)
    finally:
        blank.unlink(missing_ok=True)

    print(f"Wrote {output}")


if __name__ == "__main__":
    main()
