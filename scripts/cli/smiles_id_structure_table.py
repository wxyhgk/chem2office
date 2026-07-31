#!/usr/bin/env python3
"""Build Excel tables: column A = ID, column B = ChemDraw OLE structure.

Input line format (no separator between SMILES and ID)::

    c1ccccc1D0000

Usage::

    python3 -m scripts.cli.smiles_id_structure_table \\
        --input /path/to/a.smiles /path/to/b.smiles \\
        --output-dir output/smiles_tables
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path

from retain_cdx import ExcelCDXWriter, Placement, StructureSource

# SMILES body + trailing ID like D0000 / D0123
_LINE_RE = re.compile(r"^(?P<smiles>.+?)(?P<id>D\d+)\s*$")


def parse_smiles_line(line: str, *, line_no: int, path: Path) -> tuple[str, str]:
    text = line.strip()
    if not text or text.startswith("#"):
        raise ValueError("empty")
    match = _LINE_RE.match(text)
    if not match:
        raise ValueError(
            f"{path}:{line_no}: cannot parse SMILES+ID (expected ...D0000): {text[:80]!r}"
        )
    return match.group("smiles"), match.group("id")


def load_smiles_file(path: Path) -> list[tuple[str, str]]:
    """Return list of (smiles, id) in file order."""
    rows: list[tuple[str, str]] = []
    for line_no, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        try:
            rows.append(parse_smiles_line(line, line_no=line_no, path=path))
        except ValueError as exc:
            if str(exc) == "empty":
                continue
            raise
    if not rows:
        raise SystemExit(f"No molecules in {path}")
    return rows


def _excel_col_width_to_px(width: float) -> int:
    return max(1, int(float(width) * 7 + 5))


def _row_height_pt_to_px(height_pt: float) -> int:
    return max(1, round(float(height_pt) / 0.75))


def center_offsets_px(
    *,
    object_width_px: int,
    object_height_px: int,
    col_width_chars: float,
    row_height_pt: float,
) -> tuple[int, int]:
    """In-cell offsets so the OLE box is centered in the structure column cell."""
    cell_w = _excel_col_width_to_px(col_width_chars)
    cell_h = _row_height_pt_to_px(row_height_pt)
    ox = max(0, (cell_w - object_width_px) // 2)
    oy = max(0, (cell_h - object_height_px) // 2)
    return ox, oy


def build_table(
    smiles_path: Path,
    output_xlsx: Path,
    *,
    blank_xlsx: Path,
    sheet: str = "Sheet1",
    limit: int | None = None,
    preview_format: str = "png",
    preview_backend: str = "auto",
    # Every row uses the same OLE box; molecules are letterboxed
    # (preserve aspect, may have side padding) — never stretched to fill.
    width_px: int = 130,
    height_px: int = 100,
    row_height: float = 90,
    col_width: float = 20,
    id_header: str = "编号",
    structure_header: str = "结构",
) -> Path:
    rows = load_smiles_file(smiles_path)
    if limit is not None:
        rows = rows[:limit]

    print(f"[{smiles_path.name}] molecules={len(rows)} → {output_xlsx}")
    sources: list[StructureSource] = []
    ids: list[str] = []
    errors: list[str] = []

    t0 = time.time()
    for i, (smiles, mol_id) in enumerate(rows, start=1):
        try:
            src = StructureSource.from_smiles(smiles, name=mol_id)
            sources.append(src)
            ids.append(mol_id)
        except Exception as exc:  # noqa: BLE001 — collect and continue
            errors.append(f"{mol_id}: {exc}")
            print(f"  ! skip {mol_id}: {exc}", file=sys.stderr)
        if i % 50 == 0 or i == len(rows):
            print(f"  SMILES→CDX {i}/{len(rows)} ({time.time() - t0:.1f}s)")

    if not sources:
        raise SystemExit(f"No valid molecules from {smiles_path}")

    print(f"  inserting {len(sources)} OLE objects (preview={preview_format})…")
    t1 = time.time()
    offset_x_px, offset_y_px = center_offsets_px(
        object_width_px=width_px,
        object_height_px=height_px,
        col_width_chars=col_width,
        row_height_pt=row_height,
    )
    with ExcelCDXWriter(blank_xlsx) as writer:
        # Structures in column B starting row 2; centered in each B cell.
        writer.insert_cdx_many(
            sources,
            sheet=sheet,
            start_cell="B2",
            placement=Placement("B2", offset_x=offset_x_px, offset_y=offset_y_px),
            width_px=width_px,
            height_px=height_px,
            row_height=row_height,
            col_width=col_width,
            preview_format=preview_format,  # type: ignore[arg-type]
            preview_backend=preview_backend,  # type: ignore[arg-type]
            preview_style="acs1996",
            preview_mono=True,
            preview_no_margin=True,
        )
        # Header row 1; IDs in column A next to their structures (row 2+).
        cells = {"A1": id_header, "B1": structure_header}
        cells.update({f"A{idx + 2}": mol_id for idx, mol_id in enumerate(ids)})
        writer.set_cell_texts(sheet, cells)
        # Widen ID column; structure column width already set by insert.
        writer.set_column_width(sheet, "A", 12)
        writer.save(output_xlsx)

    print(
        f"  done → {output_xlsx} "
        f"(insert {time.time() - t1:.1f}s, total {time.time() - t0:.1f}s, "
        f"ok={len(sources)}, skipped={len(errors)})"
    )
    if errors:
        err_path = output_xlsx.with_suffix(".errors.txt")
        err_path.write_text("\n".join(errors) + "\n", encoding="utf-8")
        print(f"  errors logged → {err_path}")
    return output_xlsx


def main(argv: list[str] | None = None) -> None:
    repo = Path(__file__).resolve().parents[2]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        "-i",
        nargs="+",
        type=Path,
        required=True,
        help="One or more .smiles files (SMILES+ID per line)",
    )
    parser.add_argument(
        "--output-dir",
        "-o",
        type=Path,
        default=repo / "output" / "smiles_tables",
        help="Directory for output xlsx files",
    )
    parser.add_argument(
        "--blank",
        type=Path,
        default=repo / "tests" / "fixtures" / "blank.xlsx",
        help="Blank workbook template",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Only first N molecules per file (for smoke tests)",
    )
    parser.add_argument(
        "--preview-format",
        choices=("png", "emf"),
        default="png",
        help="Preview format (png is much faster for hundreds of molecules)",
    )
    parser.add_argument(
        "--preview-backend",
        choices=("auto", "rdkit", "obabel"),
        default="auto",
    )
    parser.add_argument(
        "--id-header",
        default="编号",
        help="Header text for the ID column (A1)",
    )
    parser.add_argument(
        "--structure-header",
        default="结构",
        help="Header text for the structure column (B1)",
    )
    parser.add_argument(
        "--width-px",
        type=int,
        default=130,
        help="Fixed OLE box width in pixels (all rows share this box)",
    )
    parser.add_argument(
        "--height-px",
        type=int,
        default=100,
        help="Fixed OLE box height in pixels (all rows share this box)",
    )
    args = parser.parse_args(argv)

    if not args.blank.exists():
        raise SystemExit(f"Blank workbook not found: {args.blank}")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    for smiles_path in args.input:
        if not smiles_path.exists():
            raise SystemExit(f"Input not found: {smiles_path}")
        out = args.output_dir / f"{smiles_path.stem}_id_structure.xlsx"
        build_table(
            smiles_path,
            out,
            blank_xlsx=args.blank,
            limit=args.limit,
            preview_format=args.preview_format,
            preview_backend=args.preview_backend,
            width_px=args.width_px,
            height_px=args.height_px,
            id_header=args.id_header,
            structure_header=args.structure_header,
        )


if __name__ == "__main__":
    main()
