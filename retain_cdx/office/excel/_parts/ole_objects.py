"""Worksheet <oleObjects> helpers (挂 OLE 声明).

装进 Excel · 子步骤：
  E. 在 sheet 上声明 oleObject，并与 shapeId 对齐
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

from retain_cdx.office.excel._ooxml import NS_MAIN

from .util import parse_int


def rewrite_ole_shape_ids(ole_objects: ET.Element, shape_id_map: dict[int, int]) -> None:
    for ole_object in ole_objects.findall(f".//{{{NS_MAIN}}}oleObject"):
        source_id = parse_int(ole_object.get("shapeId"))
        if source_id in shape_id_map:
            ole_object.set("shapeId", str(shape_id_map[source_id]))


def normalize_object_pr(ole_objects: ET.Element) -> None:
    """Pin the display box to our anchor instead of the preview image.

    ``autoPict`` defaults to *true* (ECMA-376 §18.3.1.56) = size the object from
    the referenced picture. Leaving it unset made Excel measure every preview at
    layout time — costly with thousands of objects, and wrong for us since the
    previews are deliberately supersampled. Excel's own files carry
    ``autoPict="0"``; write it explicitly rather than dropping the attribute.
    """
    for object_pr in ole_objects.findall(".//s:objectPr", {"s": NS_MAIN}):
        object_pr.set("autoPict", "0")
        object_pr.set("defaultSize", "0")
