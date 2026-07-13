#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from rdkit import Chem
from rdkit.Chem import AllChem
from pycdxml.cdxml_converter.rdkit_chemdraw import mol_to_document
from pycdxml.cdxml_converter.chemdraw_io import write_cdx_file

from scripts.cli.inject_cdx_into_xlsx import _build_generate_mode
from scripts.xlsx.content_types import ensure_png_content_type
from scripts.xlsx.namespaces import (
    NS,
    NS_MAIN,
    NS_MC,
    NS_R,
    NS_X14,
    NS_XR,
    NS_XR2,
    NS_XR3,
)
from scripts.xlsx.package import extract_xlsx, pack_xlsx
from scripts.xlsx.relationships import build_rels


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
DEFAULT_SHEETS = ("P型", "N型")


def col_to_index(col: str) -> int:
    value = 0
    for ch in col.upper():
        value = value * 26 + ord(ch) - ord("A") + 1
    return value


def index_to_col(index: int) -> str:
    out = ""
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


def shift_cell_ref(ref: str, *, start_col: int, delta: int) -> str:
    match = re.match(r"^([A-Za-z]+)(\d+)$", ref)
    if not match:
        return ref
    col = col_to_index(match.group(1))
    if col >= start_col:
        col += delta
    return f"{index_to_col(col)}{match.group(2)}"


def sheet_index_from_path(path: str) -> int:
    match = re.search(r"sheet(\d+)\.xml$", path)
    if not match:
        raise ValueError(f"Cannot determine sheet index from {path}")
    return int(match.group(1))


def parse_workbook_sheets(tmpdir: Path) -> dict[str, str]:
    wb_root = ET.parse(tmpdir / "xl" / "workbook.xml").getroot()
    rel_root = ET.parse(tmpdir / "xl" / "_rels" / "workbook.xml.rels").getroot()
    rels = {
        rel.get("Id"): rel.get("Target")
        for rel in rel_root.findall(f"{{{REL_NS}}}Relationship")
    }
    out: dict[str, str] = {}
    for sheet in wb_root.findall("s:sheets/s:sheet", NS):
        rid = sheet.get(f"{{{NS_R}}}id")
        target = rels.get(rid or "")
        if target:
            out[sheet.get("name", "")] = "xl/" + target.lstrip("/")
    return out


def shared_strings(tmpdir: Path) -> list[str]:
    root = ET.parse(tmpdir / "xl" / "sharedStrings.xml").getroot()
    return [
        "".join(t.text or "" for t in si.findall(".//s:t", NS))
        for si in root.findall("s:si", NS)
    ]


def append_shared_string(tmpdir: Path, text: str) -> int:
    path = tmpdir / "xl" / "sharedStrings.xml"
    tree = ET.parse(path)
    root = tree.getroot()
    values = [
        "".join(t.text or "" for t in si.findall(".//s:t", NS))
        for si in root.findall("s:si", NS)
    ]
    if text in values:
        return values.index(text)
    si = ET.SubElement(root, f"{{{NS_MAIN}}}si")
    t = ET.SubElement(si, f"{{{NS_MAIN}}}t")
    t.text = text
    root.set("count", str(int(root.get("count", str(len(values)))) + 1))
    root.set("uniqueCount", str(len(values) + 1))
    tree.write(path, encoding="UTF-8", xml_declaration=True)
    return len(values)


def cell_text(sheet_path: Path, row_num: int, col_num: int, tmpdir: Path) -> str:
    root = ET.parse(sheet_path).getroot()
    ref = f"{index_to_col(col_num)}{row_num}"
    cell = root.find(f".//s:c[@r='{ref}']", NS)
    if cell is None:
        return ""
    value = cell.find("s:v", NS)
    if value is None or value.text is None:
        return ""
    if cell.get("t") == "s":
        return shared_strings(tmpdir)[int(value.text)]
    return value.text


def smiles_to_cdx(
    smiles: str,
    output: Path,
    *,
    bond_length: float,
    line_width: float,
    bold_width: float,
    label_size: float,
) -> None:
    mol = Chem.MolFromSmiles(smiles)
    if mol is None:
        raise ValueError(f"Invalid SMILES: {smiles}")
    AllChem.Compute2DCoords(mol)
    doc = mol_to_document(
        mol,
        chemdraw_style={
            "BondLength": bond_length,
            "LineWidth": line_width,
            "BoldWidth": bold_width,
            "LabelSize": label_size,
            "CaptionSize": label_size,
        },
    )
    write_cdx_file(doc, output)


def ensure_content_types(tmpdir: Path, drawing_name: str) -> None:
    ensure_png_content_type(tmpdir / "[Content_Types].xml")
    path = tmpdir / "[Content_Types].xml"
    tree = ET.parse(path)
    root = tree.getroot()
    defaults = {el.get("Extension") for el in root.findall(f"{{{CT_NS}}}Default")}
    for ext, ct in [
        ("bin", "application/vnd.openxmlformats-officedocument.oleObject"),
        ("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"),
    ]:
        if ext not in defaults:
            ET.SubElement(root, f"{{{CT_NS}}}Default", Extension=ext, ContentType=ct)
    part_name = f"/xl/drawings/{drawing_name}"
    overrides = {el.get("PartName") for el in root.findall(f"{{{CT_NS}}}Override")}
    if part_name not in overrides:
        ET.SubElement(
            root,
            f"{{{CT_NS}}}Override",
            PartName=part_name,
            ContentType="application/vnd.openxmlformats-officedocument.drawing+xml",
        )
    xml = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    xml = xml.replace(b"ns0:", b"").replace(b"xmlns:ns0", b"xmlns")
    path.write_bytes(xml)


def prepare_sheet_xml(
    sheet_path: Path,
    *,
    struct_shared_index: int,
    limit: int,
    row_height: float,
    struct_col_width: float,
    drawing_rid: str,
    vml_rid: str,
    source_ole_objects: ET.Element,
) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    root.set("xmlns:x14", NS_X14)
    root.set("xmlns:xr", NS_XR)
    root.set("xmlns:xr2", NS_XR2)
    root.set("xmlns:xr3", NS_XR3)
    root.set(f"{{{NS_MC}}}Ignorable", "x14ac xr xr2 xr3")

    dimension = root.find("s:dimension", NS)
    if dimension is not None and dimension.get("ref"):
        ref = dimension.get("ref", "")
        if ":" in ref:
            start, end = ref.split(":", 1)
            dimension.set("ref", f"{start}:{shift_cell_ref(end, start_col=2, delta=1)}")

    old_cols = root.find("s:cols", NS)
    if old_cols is not None:
        root.remove(old_cols)
    sheet_data = root.find("s:sheetData", NS)
    if sheet_data is None:
        raise SystemExit(f"sheetData not found in {sheet_path}")

    cols = ET.Element(f"{{{NS_MAIN}}}cols")
    ET.SubElement(cols, f"{{{NS_MAIN}}}col", min="1", max="1", width="55", customWidth="1")
    ET.SubElement(cols, f"{{{NS_MAIN}}}col", min="2", max="2", width=str(struct_col_width), customWidth="1")
    ET.SubElement(cols, f"{{{NS_MAIN}}}col", min="3", max="3", width="18", customWidth="1")
    root.insert(list(root).index(sheet_data), cols)

    for row in sheet_data.findall("s:row", NS):
        row_num = int(row.get("r", "0"))
        row.set("spans", "1:3")
        if 2 <= row_num <= limit + 1:
            row.set("ht", str(row_height))
            row.set("customHeight", "1")
        for cell in row.findall("s:c", NS):
            ref = cell.get("r")
            if ref:
                cell.set("r", shift_cell_ref(ref, start_col=2, delta=1))

        struct_cell = ET.Element(f"{{{NS_MAIN}}}c", r=f"B{row_num}")
        if row_num == 1:
            struct_cell.set("t", "s")
            ET.SubElement(struct_cell, f"{{{NS_MAIN}}}v").text = str(struct_shared_index)
        row.insert(1, struct_cell)

    for tag in ["s:drawing", "s:legacyDrawing", "s:oleObjects"]:
        old = root.find(tag, NS)
        if old is not None:
            root.remove(old)

    page_margins = root.find("s:pageMargins", NS)
    insert_after = page_margins if page_margins is not None else sheet_data
    pos = list(root).index(insert_after) + 1
    root.insert(pos, ET.Element(f"{{{NS_MAIN}}}drawing", {f"{{{NS_R}}}id": drawing_rid}))
    root.insert(pos + 1, ET.Element(f"{{{NS_MAIN}}}legacyDrawing", {f"{{{NS_R}}}id": vml_rid}))
    root.insert(pos + 2, source_ole_objects)

    tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)


def build_source_ole_workbook(
    cdx_paths: list[Path],
    output: Path,
    *,
    preview_width: int,
    preview_height: int,
    preview_thick_lines: bool,
    preview_no_margin: bool,
) -> None:
    _build_generate_mode(
        cdx_paths,
        output,
        col="B",
        start_row=2,
        margin_left="2mm",
        margin_top="2mm",
        width="36mm",
        height="26mm",
        from_spec=None,
        to_spec=None,
        row_height=110,
        preview_width=preview_width,
        preview_height=preview_height,
        preview_thick_lines=preview_thick_lines,
        preview_no_margin=preview_no_margin,
    )


def merge_one_sheet(
    input_xlsx: Path,
    output_xlsx: Path,
    *,
    sheet_name: str,
    limit: int,
    row_height: float,
    struct_col_width: float,
    bond_length: float,
    line_width: float,
    bold_width: float,
    label_size: float,
    preview_width: int,
    preview_height: int,
    preview_thick_lines: bool,
    preview_no_margin: bool,
) -> None:
    with tempfile.TemporaryDirectory(prefix="smiles_ole_merge_") as tmp:
        work = Path(tmp) / "work"
        src = Path(tmp) / "src"
        work.mkdir()
        src.mkdir()
        extract_xlsx(input_xlsx, work)

        sheets = parse_workbook_sheets(work)
        if sheet_name not in sheets:
            raise SystemExit(f"Sheet not found: {sheet_name}")
        target_sheet_path = work / sheets[sheet_name]
        sheet_idx = sheet_index_from_path(sheets[sheet_name])

        data_rows = min(limit, max(0, len(ET.parse(target_sheet_path).getroot().findall("s:sheetData/s:row", NS)) - 1))
        cdx_paths: list[Path] = []
        for row in range(2, data_rows + 2):
            smiles = cell_text(target_sheet_path, row, 1, work)
            cdx_path = Path(tmp) / f"{sheet_idx}_{row}.cdx"
            smiles_to_cdx(
                smiles,
                cdx_path,
                bond_length=bond_length,
                line_width=line_width,
                bold_width=bold_width,
                label_size=label_size,
            )
            cdx_paths.append(cdx_path)

        source_xlsx = Path(tmp) / "source_ole.xlsx"
        build_source_ole_workbook(
            cdx_paths,
            source_xlsx,
            preview_width=preview_width,
            preview_height=preview_height,
            preview_thick_lines=preview_thick_lines,
            preview_no_margin=preview_no_margin,
        )
        extract_xlsx(source_xlsx, src)

        prefix = f"sheet{sheet_idx}"
        drawing_name = f"{prefix}_drawing.xml"
        vml_name = f"{prefix}_vmlDrawing.vml"
        drawing_rid = "rId1"
        vml_rid = "rId2"

        (work / "xl" / "drawings").mkdir(exist_ok=True)
        (work / "xl" / "drawings" / "_rels").mkdir(exist_ok=True)
        (work / "xl" / "embeddings").mkdir(exist_ok=True)
        (work / "xl" / "media").mkdir(exist_ok=True)

        (work / "xl" / "drawings" / drawing_name).write_bytes((src / "xl" / "drawings" / "drawing1.xml").read_bytes())
        (work / "xl" / "drawings" / vml_name).write_bytes((src / "xl" / "drawings" / "vmlDrawing1.vml").read_bytes())

        rels: list[tuple[str, str, str]] = [
            (drawing_rid, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", f"../drawings/{drawing_name}"),
            (vml_rid, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing", f"../drawings/{vml_name}"),
        ]
        vml_rels: list[tuple[str, str, str]] = []

        source_rels = ET.parse(src / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels").getroot()
        for rel in source_rels.findall(f"{{{REL_NS}}}Relationship"):
            rid = rel.get("Id", "")
            rtype = rel.get("Type", "")
            target = rel.get("Target", "")
            if rtype.endswith("/oleObject"):
                name = Path(target).name
                new_name = f"{prefix}_{name}"
                (work / "xl" / "embeddings" / new_name).write_bytes((src / "xl" / "embeddings" / name).read_bytes())
                rels.append((rid, rtype, f"../embeddings/{new_name}"))
            elif rtype.endswith("/image"):
                name = Path(target).name
                new_name = f"{prefix}_{name}"
                (work / "xl" / "media" / new_name).write_bytes((src / "xl" / "media" / name).read_bytes())
                rels.append((rid, rtype, f"../media/{new_name}"))

        source_vml_rels = ET.parse(src / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels").getroot()
        for rel in source_vml_rels.findall(f"{{{REL_NS}}}Relationship"):
            target = rel.get("Target", "")
            name = Path(target).name
            vml_rels.append((rel.get("Id", ""), rel.get("Type", ""), f"../media/{prefix}_{name}"))

        target_rels = work / "xl" / "worksheets" / "_rels" / f"sheet{sheet_idx}.xml.rels"
        target_rels.parent.mkdir(exist_ok=True)
        build_rels(target_rels, rels)
        build_rels(work / "xl" / "drawings" / "_rels" / f"{vml_name}.rels", vml_rels)

        source_sheet_root = ET.parse(src / "xl" / "worksheets" / "sheet1.xml").getroot()
        source_ole_objects = source_sheet_root.find("s:oleObjects", NS)
        if source_ole_objects is None:
            raise SystemExit("source OLE workbook has no oleObjects")

        struct_idx = append_shared_string(work, "Struct")
        prepare_sheet_xml(
            target_sheet_path,
            struct_shared_index=struct_idx,
            limit=data_rows,
            row_height=row_height,
            struct_col_width=struct_col_width,
            drawing_rid=drawing_rid,
            vml_rid=vml_rid,
            source_ole_objects=source_ole_objects,
        )
        ensure_content_types(work, drawing_name)
        pack_xlsx(work, output_xlsx)


def main() -> None:
    parser = argparse.ArgumentParser(description="Merge verified OLE workbooks into data1.xlsx sheet by sheet.")
    parser.add_argument("input", nargs="?", default="Data/data1.xlsx")
    parser.add_argument("-o", "--output", default="output/data1_ole_merge.xlsx")
    parser.add_argument("--sheets", nargs="+", default=list(DEFAULT_SHEETS))
    parser.add_argument("--limit", type=int, default=3)
    parser.add_argument(
        "--all",
        action="store_true",
        help="Process all data rows in each selected sheet.",
    )
    parser.add_argument("--row-height", type=float, default=110)
    parser.add_argument("--struct-col-width", type=float, default=24)
    parser.add_argument("--bond-length", type=float, default=24)
    parser.add_argument("--line-width", type=float, default=1.2)
    parser.add_argument("--bold-width", type=float, default=4)
    parser.add_argument("--label-size", type=float, default=16)
    parser.add_argument("--preview-width", type=int, default=900)
    parser.add_argument("--preview-height", type=int, default=650)
    parser.add_argument("--preview-thin-lines", action="store_true")
    parser.add_argument("--preview-keep-margin", action="store_true")
    args = parser.parse_args()

    with tempfile.TemporaryDirectory(prefix="smiles_ole_merge_seq_") as tmp:
        current = Path(args.input)
        limit = 10**9 if args.all else args.limit
        for idx, sheet in enumerate(args.sheets, start=1):
            next_path = Path(args.output) if idx == len(args.sheets) else Path(tmp) / f"pass_{idx}.xlsx"
            merge_one_sheet(
                current,
                next_path,
                sheet_name=sheet,
                limit=limit,
                row_height=args.row_height,
                struct_col_width=args.struct_col_width,
                bond_length=args.bond_length,
                line_width=args.line_width,
                bold_width=args.bold_width,
                label_size=args.label_size,
                preview_width=args.preview_width,
                preview_height=args.preview_height,
                preview_thick_lines=not args.preview_thin_lines,
                preview_no_margin=not args.preview_keep_margin,
            )
            current = next_path
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
