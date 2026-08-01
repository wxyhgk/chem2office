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
