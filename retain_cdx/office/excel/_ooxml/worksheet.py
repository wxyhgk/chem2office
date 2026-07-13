from __future__ import annotations

import copy
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
    rid_map: dict[str, str] | None = None,
) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    ensure_sheet_compat_namespaces(root)

    sheet_data = root.find("s:sheetData", NS)
    if sheet_data is None:
        raise ValueError(f"sheetData not found: {sheet_path}")

    drawing = root.find("s:drawing", NS)
    legacy_drawing = root.find("s:legacyDrawing", NS)
    target_ole_objects = root.find("s:oleObjects", NS)
    if drawing is None or legacy_drawing is None or target_ole_objects is None:
        page_margins = root.find("s:pageMargins", NS)
        insert_after = page_margins if page_margins is not None else sheet_data
        pos = list(root).index(insert_after) + 1
        if drawing is None:
            drawing = ET.Element(f"{{{NS_MAIN}}}drawing", {f"{{{NS_R}}}id": drawing_rid})
            root.insert(pos, drawing)
            pos += 1
        if legacy_drawing is None:
            legacy_drawing = ET.Element(
                f"{{{NS_MAIN}}}legacyDrawing",
                {f"{{{NS_R}}}id": vml_rid},
            )
            root.insert(pos, legacy_drawing)
            pos += 1
        if target_ole_objects is None:
            target_ole_objects = ET.Element(f"{{{NS_MAIN}}}oleObjects")
            root.insert(pos, target_ole_objects)

    if drawing.get(f"{{{NS_R}}}id") is None:
        drawing.set(f"{{{NS_R}}}id", drawing_rid)
    if legacy_drawing.get(f"{{{NS_R}}}id") is None:
        legacy_drawing.set(f"{{{NS_R}}}id", vml_rid)

    rid_map = rid_map or {}
    for child in list(ole_objects):
        imported = copy.deepcopy(child)
        _rewrite_relationship_ids(imported, rid_map)
        target_ole_objects.append(imported)
    write_sheet_xml(tree, sheet_path)


def _rewrite_relationship_ids(element: ET.Element, rid_map: dict[str, str]) -> None:
    rid_attr = f"{{{NS_R}}}id"
    for node in element.iter():
        rid = node.get(rid_attr)
        if rid in rid_map:
            node.set(rid_attr, rid_map[rid])
