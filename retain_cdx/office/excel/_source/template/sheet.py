from __future__ import annotations

import copy
from pathlib import Path
import xml.etree.ElementTree as ET

from .namespaces import (
    NS,
    NS_MAIN,
    NS_X14AC,
    ensure_sheet_namespaces,
)


def extract_base_anchor(sheet_path: Path) -> tuple[int, int, int]:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    base_alt = root.find("s:oleObjects/mc:AlternateContent", NS)
    if base_alt is None:
        raise SystemExit("base OLE object not found in sheet1.xml")
    base_choice = base_alt.find("mc:Choice", NS)
    if base_choice is None:
        raise SystemExit("base OLE choice not found in sheet1.xml")
    base_ole = base_choice.find("s:oleObject", NS)
    if base_ole is None:
        raise SystemExit("base OLE element not found in sheet1.xml")
    base_shape_id = int(base_ole.get("shapeId", "1025"))
    obj_pr = base_ole.find("s:objectPr/s:anchor", NS)
    if obj_pr is None:
        raise SystemExit("objectPr anchor not found in sheet1.xml")
    from_row = int(obj_pr.find("s:from/xdr:row", NS).text)  # type: ignore[union-attr]
    to_row = int(obj_pr.find("s:to/xdr:row", NS).text)  # type: ignore[union-attr]
    return base_shape_id, from_row, to_row


def update_sheet_labels(
    sheet_path: Path,
    labels: list[str],
    total_rows: int | None = None,
) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    ensure_sheet_namespaces(root)

    dim = root.find("s:dimension", NS)
    if dim is not None:
        max_rows = max(total_rows or len(labels), len(labels))
        dim.set("ref", f"A1:B{max_rows}")

    sheet_data = root.find("s:sheetData", NS)
    if sheet_data is None:
        raise SystemExit("sheetData not found")

    template_row = sheet_data.find("s:row", NS)
    if template_row is None:
        raise SystemExit("template row not found")

    for row in list(sheet_data):
        sheet_data.remove(row)

    max_rows = max(total_rows or len(labels), len(labels))
    for idx in range(1, max_rows + 1):
        row = copy.deepcopy(template_row)
        row.set("r", str(idx))
        if f"{{{NS_X14AC}}}dyDescent" in row.attrib:
            row.set(f"{{{NS_X14AC}}}dyDescent", "0.2")

        cell = row.find("s:c", NS)
        if cell is None:
            raise SystemExit("cell not found in template row")
        if idx <= len(labels):
            cell.set("r", f"A{idx}")
            cell.set("t", "s")
            v = cell.find("s:v", NS)
            if v is None:
                v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
            v.text = str(idx - 1)
        else:
            row.remove(cell)
        sheet_data.append(row)

    tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)


def update_row_heights(sheet_path: Path, row_height: float, total_rows: int) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    ensure_sheet_namespaces(root)

    sheet_format = root.find("s:sheetFormatPr", NS)
    if sheet_format is not None:
        sheet_format.set("defaultRowHeight", str(row_height))
        sheet_format.set("customHeight", "1")

    sheet_data = root.find("s:sheetData", NS)
    if sheet_data is None:
        raise SystemExit("sheetData not found")
    for idx in range(1, total_rows + 1):
        row = sheet_data.find(f"s:row[@r='{idx}']", NS)
        if row is not None:
            row.set("ht", str(row_height))
            row.set("customHeight", "1")

    tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)


def write_shared_strings(path: Path, labels: list[str]) -> None:
    root = ET.Element(
        f"{{{NS_MAIN}}}sst", count=str(len(labels)), uniqueCount=str(len(labels))
    )
    for label in labels:
        si = ET.SubElement(root, f"{{{NS_MAIN}}}si")
        t = ET.SubElement(si, f"{{{NS_MAIN}}}t")
        t.text = label
    ET.ElementTree(root).write(path, encoding="UTF-8", xml_declaration=True)
