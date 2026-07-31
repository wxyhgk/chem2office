"""Plain cell values (inline strings) for worksheets carrying OLE parts.

openpyxl cannot post-process these workbooks — it rewrites the package and
drops OLE/VML parts — so tables around the structures are written here.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from collections.abc import Mapping
from pathlib import Path

from retain_cdx.office.excel._utils import index_to_col, parse_cell

from .xml_io import (
    NS,
    NS_MAIN,
    ensure_sheet_compat_namespaces,
    insert_worksheet_child,
    write_sheet_xml,
)


def set_inline_cell_text(sheet_path: Path, cells: Mapping[str, str]) -> None:
    """Write inline-string values like ``{"A1": "编号", "A2": "D0001"}``.

    Creates missing rows, preserves existing row attributes (heights set by
    the OLE insert), and keeps rows/cells sorted so Excel does not repair.
    """
    if not cells:
        return
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    ensure_sheet_compat_namespaces(root)

    sheet_data = root.find("s:sheetData", NS)
    if sheet_data is None:
        sheet_data = ET.SubElement(root, f"{{{NS_MAIN}}}sheetData")

    rows_by_r: dict[int, ET.Element] = {}
    for row in sheet_data.findall("s:row", NS):
        rows_by_r[int(row.get("r", "0") or 0)] = row

    replaced_formula = False
    for ref, text in cells.items():
        _, row_num = parse_cell(ref)
        row = rows_by_r.get(row_num)
        if row is None:
            row = ET.Element(f"{{{NS_MAIN}}}row", r=str(row_num))
            rows_by_r[row_num] = row
        replaced_formula |= _set_inline(row, ref.strip().upper(), text)
        _sort_row_cells(row)

    # Excel requires ascending row order; re-append everything sorted.
    for row in list(sheet_data.findall("s:row", NS)):
        sheet_data.remove(row)
    for row_num in sorted(rows_by_r):
        sheet_data.append(rows_by_r[row_num])

    _expand_dimension(root, cells)
    write_sheet_xml(tree, sheet_path)

    if replaced_formula:
        # sheet_path is <package>/xl/worksheets/sheetN.xml
        drop_calc_chain(sheet_path.parent.parent.parent)


def _set_inline(row: ET.Element, ref: str, text: str) -> bool:
    """Replace the cell at *ref*. Returns True if it used to hold a formula."""
    style = None
    had_formula = False
    for cell in list(row.findall("s:c", NS)):
        if cell.get("r") != ref:
            continue
        # Keep the cell's format — replacing text should not strip bold/fill.
        style = cell.get("s")
        had_formula = cell.find("s:f", NS) is not None
        row.remove(cell)
    cell = ET.SubElement(row, f"{{{NS_MAIN}}}c", r=ref, t="inlineStr")
    if style is not None:
        cell.set("s", style)
    is_el = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
    t_el = ET.SubElement(is_el, f"{{{NS_MAIN}}}t")
    t_el.text = text
    return had_formula


def drop_calc_chain(package_dir: Path) -> None:
    """Remove xl/calcChain.xml and its wiring.

    A chain entry pointing at a cell that no longer holds a formula makes Excel
    report "Removed Records: Formula". The part is a cache, so dropping it is
    safe — Excel rebuilds it on the next calculation.
    """
    chain = package_dir / "xl" / "calcChain.xml"
    if not chain.exists():
        return
    chain.unlink()

    content_types = package_dir / "[Content_Types].xml"
    if content_types.exists():
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        tree = ET.parse(content_types)
        root = tree.getroot()
        for override in root.findall(f"{{{ct_ns}}}Override"):
            if override.get("PartName") == "/xl/calcChain.xml":
                root.remove(override)
        xml = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
        content_types.write_bytes(xml.replace(b"ns0:", b"").replace(b"xmlns:ns0", b"xmlns"))

    rels_path = package_dir / "xl" / "_rels" / "workbook.xml.rels"
    if rels_path.exists():
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        tree = ET.parse(rels_path)
        root = tree.getroot()
        for rel in root.findall(f"{{{rels_ns}}}Relationship"):
            if (rel.get("Target") or "").endswith("calcChain.xml"):
                root.remove(rel)
        xml = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
        rels_path.write_bytes(xml.replace(b"ns0:", b"").replace(b"xmlns:ns0", b"xmlns"))


def _sort_row_cells(row: ET.Element) -> None:
    cells = row.findall("s:c", NS)
    ordered = sorted(cells, key=lambda c: _cell_sort_key(c.get("r", "")))
    if cells == ordered:
        return
    for cell in cells:
        row.remove(cell)
    for cell in ordered:
        row.append(cell)


def _cell_sort_key(ref: str) -> int:
    try:
        col, _ = parse_cell(ref)
    except ValueError:
        return 0
    return col


def _expand_dimension(root: ET.Element, cells: Mapping[str, str]) -> None:
    """Grow <dimension> to cover the written cells (create it if missing)."""
    refs = [parse_cell(ref) for ref in cells]
    min_col = min(c for c, _ in refs)
    max_col = max(c for c, _ in refs)
    min_row = min(r for _, r in refs)
    max_row = max(r for _, r in refs)

    dim = root.find("s:dimension", NS)
    if dim is None:
        dim = ET.Element(f"{{{NS_MAIN}}}dimension")
        insert_worksheet_child(root, dim)
    else:
        current = dim.get("ref", "")
        if ":" in current:
            start, end = current.split(":", 1)
            try:
                c1, r1 = parse_cell(start)
                c2, r2 = parse_cell(end)
                min_col, max_col = min(min_col, c1), max(max_col, c2)
                min_row, max_row = min(min_row, r1), max(max_row, r2)
            except ValueError:
                pass
    dim.set("ref", f"{_ref(min_col, min_row)}:{_ref(max_col, max_row)}")


def _ref(col: int, row: int) -> str:
    return f"{index_to_col(col)}{row}"


__all__ = ["set_inline_cell_text"]
