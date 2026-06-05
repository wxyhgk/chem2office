from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

from .xml_io import (
    NS,
    NS_MAIN,
    NS_R,
    ensure_sheet_compat_namespaces,
    write_sheet_xml,
)


def attach_ole_objects(
    sheet_path: Path,
    *,
    drawing_rid: str,
    vml_rid: str,
    ole_objects: ET.Element,
) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    ensure_sheet_compat_namespaces(root)

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
