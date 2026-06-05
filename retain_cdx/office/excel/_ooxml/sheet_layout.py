from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

from .xml_io import (
    NS,
    NS_MAIN,
    NS_X14AC,
    ensure_sheet_compat_namespaces,
    write_sheet_xml,
)


def copy_sheet_layout(source_sheet_path: Path, target_sheet_path: Path) -> None:
    source_tree = ET.parse(source_sheet_path)
    source_root = source_tree.getroot()
    target_tree = ET.parse(target_sheet_path)
    target_root = target_tree.getroot()
    ensure_sheet_compat_namespaces(target_root)

    _replace_or_remove_child(target_root, source_root, "sheetFormatPr")
    _replace_or_remove_child(target_root, source_root, "cols")
    _replace_sheet_rows(source_root, target_root)
    write_sheet_xml(target_tree, target_sheet_path)


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
    ensure_sheet_compat_namespaces(root)

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
            row.set("ht", str(row_height))
            row.set("customHeight", "1")
    write_sheet_xml(tree, sheet_path)


def _replace_or_remove_child(target_root: ET.Element, source_root: ET.Element, tag: str) -> None:
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
