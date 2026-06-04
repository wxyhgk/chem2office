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


def copy_sheet_layout(source_sheet_path: Path, target_sheet_path: Path) -> None:
    source_tree = ET.parse(source_sheet_path)
    source_root = source_tree.getroot()
    target_tree = ET.parse(target_sheet_path)
    target_root = target_tree.getroot()
    _ensure_sheet_compat_namespaces(target_root)

    _replace_or_remove_child(target_root, source_root, "sheetFormatPr")
    _replace_or_remove_child(target_root, source_root, "cols")
    _replace_sheet_rows(source_root, target_root)
    write_sheet_xml(target_tree, target_sheet_path)


def _replace_or_remove_child(target_root: ET.Element, source_root: ET.Element, tag: str) -> None:
    qtag = f"{{{NS_MAIN}}}{tag}"
    source_child = source_root.find(f"s:{tag}", NS)
    target_child = target_root.find(f"s:{tag}", NS)
    if target_child is not None:
        target_root.remove(target_child)
    if source_child is None:
        return

    insert_before = target_root.find("s:sheetData", NS)
    if tag == "sheetFormatPr":
        sheet_views = target_root.find("s:sheetViews", NS)
        insert_pos = list(target_root).index(sheet_views) + 1 if sheet_views is not None else 0
    elif insert_before is not None:
        insert_pos = list(target_root).index(insert_before)
    else:
        insert_pos = len(target_root)
    target_root.insert(insert_pos, ET.fromstring(ET.tostring(source_child)))


def _replace_sheet_rows(source_root: ET.Element, target_root: ET.Element) -> None:
    source_sheet_data = source_root.find("s:sheetData", NS)
    target_sheet_data = target_root.find("s:sheetData", NS)
    if source_sheet_data is None or target_sheet_data is None:
        return

    target_rows_by_number = {
        int(row.get("r", "0")): row
        for row in target_sheet_data.findall("s:row", NS)
    }
    for source_row in source_sheet_data.findall("s:row", NS):
        try:
            row_num = int(source_row.get("r", "0"))
        except ValueError:
            continue
        target_row = target_rows_by_number.get(row_num)
        if target_row is None:
            target_row = ET.Element(f"{{{NS_MAIN}}}row", r=str(row_num))
            _insert_row_sorted(target_sheet_data, target_row)
            target_rows_by_number[row_num] = target_row
        _copy_row_layout(source_row, target_row)


def _copy_row_layout(source_row: ET.Element, target_row: ET.Element) -> None:
    layout_attrs = [
        "ht",
        "customHeight",
        "hidden",
        "outlineLevel",
        "collapsed",
        "thickTop",
        "thickBot",
        f"{{{NS_X14AC}}}dyDescent",
    ]
    for attr in layout_attrs:
        if attr in source_row.attrib:
            target_row.set(attr, source_row.attrib[attr])
        else:
            target_row.attrib.pop(attr, None)


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
