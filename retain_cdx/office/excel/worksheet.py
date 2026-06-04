from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
NS_XR2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
NS_XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"

NS = {"s": NS_MAIN}


def _register_namespaces() -> None:
    ET.register_namespace("", NS_MAIN)
    ET.register_namespace("mc", NS_MC)
    ET.register_namespace("r", NS_R)
    ET.register_namespace("x14", NS_X14)
    ET.register_namespace("x14ac", NS_X14AC)
    ET.register_namespace("xr", NS_XR)
    ET.register_namespace("xr2", NS_XR2)
    ET.register_namespace("xr3", NS_XR3)
    ET.register_namespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")


def attach_ole_objects(
    sheet_path: Path,
    *,
    drawing_rid: str,
    vml_rid: str,
    ole_objects: ET.Element,
) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    _ensure_sheet_compat_namespaces(root)

    for tag in ["s:drawing", "s:legacyDrawing", "s:oleObjects"]:
        old = root.find(tag, NS)
        if old is not None:
            root.remove(old)

    sheet_data = root.find("s:sheetData", NS)
    if sheet_data is None:
        raise ValueError(f"sheetData not found: {sheet_path}")
    page_margins = root.find("s:pageMargins", NS)
    insert_after = page_margins if page_margins is not None else sheet_data
    pos = list(root).index(insert_after) + 1
    root.insert(pos, ET.Element(f"{{{NS_MAIN}}}drawing", {f"{{{NS_R}}}id": drawing_rid}))
    root.insert(pos + 1, ET.Element(f"{{{NS_MAIN}}}legacyDrawing", {f"{{{NS_R}}}id": vml_rid}))
    root.insert(pos + 2, ole_objects)
    write_sheet_xml(tree, sheet_path)


def touch_sheet_layout(
    sheet_path: Path,
    *,
    col: int,
    start_row: int,
    count: int,
    row_height: float | None,
    col_width: float | None,
) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    _ensure_sheet_compat_namespaces(root)

    if col_width is not None:
        cols = root.find("s:cols", NS)
        sheet_data = root.find("s:sheetData", NS)
        if sheet_data is None:
            raise ValueError(f"sheetData not found: {sheet_path}")
        if cols is None:
            cols = ET.Element(f"{{{NS_MAIN}}}cols")
            root.insert(list(root).index(sheet_data), cols)
        ET.SubElement(
            cols,
            f"{{{NS_MAIN}}}col",
            min=str(col),
            max=str(col),
            width=str(col_width),
            customWidth="1",
        )

    if row_height is not None:
        sheet_data = root.find("s:sheetData", NS)
        if sheet_data is None:
            raise ValueError(f"sheetData not found: {sheet_path}")
        rows_by_number = {
            int(row.get("r", "0")): row
            for row in sheet_data.findall("s:row", NS)
        }
        for row_num in range(start_row, start_row + count):
            row = rows_by_number.get(row_num)
            if row is None:
                row = ET.Element(f"{{{NS_MAIN}}}row", r=str(row_num))
                _insert_row_sorted(sheet_data, row)
            if start_row <= row_num < start_row + count:
                row.set("ht", str(row_height))
                row.set("customHeight", "1")
    write_sheet_xml(tree, sheet_path)


def write_sheet_xml(tree: ET.ElementTree, sheet_path: Path) -> None:
    _register_namespaces()
    tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)
    text = sheet_path.read_text(encoding="UTF-8")
    root_start = text.find("<worksheet")
    if root_start == -1:
        return
    root_end = text.find(">", root_start)
    if root_end == -1:
        return

    root_tag = text[root_start:root_end]
    declarations = {
        "xmlns:x14": NS_X14,
        "xmlns:x14ac": NS_X14AC,
        "xmlns:xr": NS_XR,
        "xmlns:xr2": NS_XR2,
        "xmlns:xr3": NS_XR3,
    }
    for name, uri in declarations.items():
        if f"{name}=" not in root_tag:
            root_tag += f' {name}="{uri}"'

    text = text[:root_start] + root_tag + text[root_end:]
    sheet_path.write_text(text, encoding="UTF-8")


def _ensure_sheet_compat_namespaces(root: ET.Element) -> None:
    root.set("xmlns:x14", NS_X14)
    root.set("xmlns:x14ac", NS_X14AC)
    root.set("xmlns:xr", NS_XR)
    root.set("xmlns:xr2", NS_XR2)
    root.set("xmlns:xr3", NS_XR3)
    root.set(f"{{{NS_MC}}}Ignorable", "x14ac xr xr2 xr3")


def _insert_row_sorted(sheet_data: ET.Element, row: ET.Element) -> None:
    row_num = int(row.get("r", "0"))
    for idx, existing in enumerate(list(sheet_data)):
        try:
            existing_num = int(existing.get("r", "0"))
        except ValueError:
            continue
        if existing_num > row_num:
            sheet_data.insert(idx, row)
            return
    sheet_data.append(row)
