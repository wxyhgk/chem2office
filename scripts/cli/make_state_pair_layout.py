#!/usr/bin/env python3
from __future__ import annotations

import argparse
import copy
import tempfile
from pathlib import Path
import xml.etree.ElementTree as ET

from scripts.ole.embed import write_cdx_to_ole
from scripts.render.preview import render_png
from scripts.xlsx.content_types import ensure_png_content_type
from scripts.xlsx.namespaces import NS, NS_MAIN, NS_X14AC, ensure_sheet_namespaces, register_namespaces
from scripts.xlsx.package import extract_xlsx, pack_xlsx
from scripts.xlsx.relationships import build_sheet_relationships, build_vml_relationships
from scripts.xlsx.sheet import (
    extract_base_anchor,
    extract_base_ole_anchor,
    update_drawing,
    update_ole_objects,
    update_vml,
    write_shared_strings,
)

EMU_PER_PIXEL = 9525
POINTS_PER_PIXEL = 0.75
DEFAULT_A_COLUMN_PIXELS = 180
DEFAULT_ROW_PIXELS = 90
DEFAULT_B_COLUMN_WIDTH = 21.375
DEFAULT_FILL_RATIO = 0.9


def _excel_width_to_pixels(width: float) -> int:
    return int(((256 * width + int(128 / 7)) / 256) * 7)


def _pixels_to_excel_width(pixels: int) -> float:
    for raw in range(0, 256 * 100):
        width = raw / 256
        if _excel_width_to_pixels(width) >= pixels:
            return width
    raise ValueError(f"Unable to represent pixel width: {pixels}")


def _add_vertical_offset(row: int, row_off: int, delta_emu: int, row_heights_emu: list[int]) -> tuple[int, int]:
    total = row_off + delta_emu
    cur_row = row
    while cur_row < len(row_heights_emu) and total >= row_heights_emu[cur_row]:
        total -= row_heights_emu[cur_row]
        cur_row += 1
    return cur_row, total


def _base_anchor_size_emu(base_anchor, *, a_column_pixels: int, row_pixels: int) -> tuple[int, int]:
    width_emu = (base_anchor.to_pt.col - base_anchor.from_pt.col) * a_column_pixels * EMU_PER_PIXEL
    width_emu += base_anchor.to_pt.col_off - base_anchor.from_pt.col_off
    height_emu = (base_anchor.to_pt.row - base_anchor.from_pt.row) * row_pixels * EMU_PER_PIXEL
    height_emu += base_anchor.to_pt.row_off - base_anchor.from_pt.row_off
    return width_emu, height_emu


def _center_anchor_in_a1_a2(
    base_anchor,
    *,
    a_column_pixels: int,
    row_pixels: int,
    fill_ratio: float,
):
    width_emu, height_emu = _base_anchor_size_emu(
        base_anchor,
        a_column_pixels=a_column_pixels,
        row_pixels=row_pixels,
    )

    area_width_emu = a_column_pixels * EMU_PER_PIXEL
    area_height_emu = (row_pixels * 2) * EMU_PER_PIXEL

    target_width_emu = area_width_emu * fill_ratio
    target_height_emu = area_height_emu * fill_ratio
    scale = min(target_width_emu / width_emu, target_height_emu / height_emu)
    scaled_width_emu = max(1, round(width_emu * scale))
    scaled_height_emu = max(1, round(height_emu * scale))

    from_col_off = max(0, round((area_width_emu - scaled_width_emu) / 2))
    from_row_off = max(0, round((area_height_emu - scaled_height_emu) / 2))

    row_heights_emu = [row_pixels * EMU_PER_PIXEL, row_pixels * EMU_PER_PIXEL]
    to_row, to_row_off = _add_vertical_offset(0, from_row_off, scaled_height_emu, row_heights_emu)

    return [{
        "from_col": 0,
        "from_col_off": from_col_off,
        "from_row": 0,
        "from_row_off": from_row_off,
        "to_col": 0,
        "to_col_off": from_col_off + scaled_width_emu,
        "to_row": to_row,
        "to_row_off": to_row_off,
    }]


def _dict_anchor_to_model(anchor_dict):
    from scripts.xlsx.position import AnchorPoint, OleAnchor

    return OleAnchor(
        from_pt=AnchorPoint(
            col=anchor_dict["from_col"],
            col_off=anchor_dict["from_col_off"],
            row=anchor_dict["from_row"],
            row_off=anchor_dict["from_row_off"],
        ),
        to_pt=AnchorPoint(
            col=anchor_dict["to_col"],
            col_off=anchor_dict["to_col_off"],
            row=anchor_dict["to_row"],
            row_off=anchor_dict["to_row_off"],
        ),
    )


def _write_two_row_labels(
    sheet_path: Path,
    top_label: str,
    bottom_label: str,
    *,
    a_column_pixels: int,
    row_pixels: int,
    b_column_width: float = DEFAULT_B_COLUMN_WIDTH,
) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    ensure_sheet_namespaces(root)

    dim = root.find("s:dimension", NS)
    if dim is not None:
        dim.set("ref", "A1:B2")

    cols = root.find("s:cols", NS)
    if cols is None:
        cols = ET.SubElement(root, f"{{{NS_MAIN}}}cols")
    for col in list(cols):
        cols.remove(col)
    ET.SubElement(
        cols,
        f"{{{NS_MAIN}}}col",
        min="1",
        max="1",
        width=str(_pixels_to_excel_width(a_column_pixels)),
        customWidth="1",
    )
    ET.SubElement(
        cols,
        f"{{{NS_MAIN}}}col",
        min="2",
        max="2",
        width=str(b_column_width),
        customWidth="1",
    )

    sheet_data = root.find("s:sheetData", NS)
    if sheet_data is None:
        raise SystemExit("sheetData not found")

    template_row = sheet_data.find("s:row", NS)
    if template_row is None:
        raise SystemExit("template row not found")

    for row in list(sheet_data):
        sheet_data.remove(row)

    for idx, _label in enumerate([top_label, bottom_label], start=1):
        row = copy.deepcopy(template_row)
        row.set("r", str(idx))
        row.set("ht", str(row_pixels * POINTS_PER_PIXEL))
        row.set("customHeight", "1")
        if f"{{{NS_X14AC}}}dyDescent" in row.attrib:
            row.set(f"{{{NS_X14AC}}}dyDescent", "0.2")

        for cell in list(row.findall("s:c", NS)):
            row.remove(cell)

        cell = ET.SubElement(row, f"{{{NS_MAIN}}}c", r=f"B{idx}", t="s")
        v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
        v.text = str(idx - 1)
        sheet_data.append(row)

    tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)


def build_state_pair_layout(
    cdx_path: Path,
    output: Path,
    *,
    top_label: str,
    bottom_label: str,
    template: Path,
    a_column_pixels: int,
    row_pixels: int,
    fill_ratio: float,
) -> None:
    register_namespaces()

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        extract_xlsx(template, tmpdir)

        ole_dir = tmpdir / "xl" / "embeddings"
        ole_dir.mkdir(parents=True, exist_ok=True)
        ole_path = ole_dir / "oleObject1.bin"
        if not ole_path.exists():
            raise SystemExit(f"OLE object not found: {ole_path}")

        sheet_path = tmpdir / "xl" / "worksheets" / "sheet1.xml"
        drawing_path = tmpdir / "xl" / "drawings" / "drawing1.xml"
        vml_path = tmpdir / "xl" / "drawings" / "vmlDrawing1.vml"
        sheet_rels_path = tmpdir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels"
        vml_rels_path = tmpdir / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels"
        shared_strings_path = tmpdir / "xl" / "sharedStrings.xml"
        content_types_path = tmpdir / "[Content_Types].xml"

        base_shape_id, from_row, to_row = extract_base_anchor(sheet_path)
        base_anchor = extract_base_ole_anchor(sheet_path)
        anchors = [
            _dict_anchor_to_model(anchor_dict)
            for anchor_dict in _center_anchor_in_a1_a2(
                base_anchor,
                a_column_pixels=a_column_pixels,
                row_pixels=row_pixels,
                fill_ratio=fill_ratio,
            )
        ]

        _write_two_row_labels(
            sheet_path,
            top_label,
            bottom_label,
            a_column_pixels=a_column_pixels,
            row_pixels=row_pixels,
        )
        write_shared_strings(shared_strings_path, [top_label, bottom_label])

        write_cdx_to_ole(ole_path, cdx_path)

        image_path = tmpdir / "xl" / "media" / "image1.png"
        image_path.parent.mkdir(parents=True, exist_ok=True)
        render_png(cdx_path, image_path)

        emf_path = tmpdir / "xl" / "media" / "image1.emf"
        if emf_path.exists():
            emf_path.unlink()

        rel_pairs = build_sheet_relationships(
            sheet_rels_path,
            ["../embeddings/oleObject1.bin"],
            ["../media/image1.png"],
        )
        vml_rids = build_vml_relationships(vml_rels_path, ["../media/image1.png"])

        update_ole_objects(
            sheet_path, base_shape_id, from_row, to_row, rel_pairs, anchors=anchors
        )
        update_drawing(
            drawing_path, base_shape_id, from_row, to_row, 1, anchors=anchors
        )
        update_vml(vml_path, base_shape_id, from_row, to_row, vml_rids, anchors=anchors)
        ensure_png_content_type(content_types_path)

        pack_xlsx(tmpdir, output)


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
        default=str(Path(__file__).resolve().parents[2] / "template.xlsx"),
        help="Template XLSX with a ChemDraw OLE object",
    )
    parser.add_argument(
        "--fill-ratio",
        type=float,
        default=DEFAULT_FILL_RATIO,
        help="Fraction of the A1+A2 box to fill on the limiting side while preserving aspect ratio (default: 0.9)",
    )
    args = parser.parse_args()

    cdx_path = Path(args.cdx)
    if not cdx_path.exists():
        raise SystemExit(f"CDX not found: {cdx_path}")

    build_state_pair_layout(
        cdx_path,
        Path(args.output),
        top_label=args.top_label,
        bottom_label=args.bottom_label,
        template=Path(args.template),
        a_column_pixels=args.a_column_pixels,
        row_pixels=args.row_pixels,
        fill_ratio=args.fill_ratio,
    )
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
