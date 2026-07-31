"""Merge drawing.xml + vmlDrawing.vml (钉框 / 表面图形).

装进 Excel · 子步骤：
  D. 钉位置 — two-cell anchors、shape 几何
  E 一部分 — shapeId 在 drawing/VML 侧对齐
"""

from __future__ import annotations

import copy
import re
import xml.etree.ElementTree as ET

from retain_cdx.office.excel._ooxml import NS_O, NS_V, NS_XDR

from .context import OlePartMergeContext
from .util import parse_int, write_xml


def merge_drawing_and_vml(
    context: OlePartMergeContext,
    vml_rid_map: dict[str, str],
) -> dict[int, int]:
    """Merge source drawing/VML into target; return source→target shapeId map."""
    drawing_path = context.target_dir / "xl" / "drawings" / context.drawing_name
    vml_path = context.target_dir / "xl" / "drawings" / context.vml_name
    target_drawing = ET.parse(drawing_path)
    source_drawing = ET.parse(context.source_dir / "xl" / "drawings" / "drawing1.xml")
    target_vml = ET.parse(vml_path)
    source_vml = ET.parse(context.source_dir / "xl" / "drawings" / "vmlDrawing1.vml")

    if context.created_drawing and context.created_vml:
        # Nothing pre-existing to collide with; only fix VML image rIds.
        rewrite_vml_relationship_ids(target_vml.getroot(), vml_rid_map)
        write_xml(target_vml, vml_path)
        return {}

    # Exactly one side was freshly copied (a chart supplies the drawing, or
    # comments supply the VML). The copy still carries the source's shape ids,
    # so it has to be remapped too — otherwise the oleObject binds to a shape
    # id that exists on only one of the two layers.
    shape_id_map = allocate_shape_ids(
        target_drawing.getroot(),
        target_vml.getroot(),
        source_drawing.getroot(),
    )

    if context.created_drawing:
        for child in target_drawing.getroot():
            rewrite_drawing_shape_ids(child, shape_id_map)
        write_xml(target_drawing, drawing_path)
    elif has_existing_drawing_content(target_drawing.getroot()):
        append_drawing_content(
            target_drawing.getroot(),
            source_drawing.getroot(),
            shape_id_map,
        )
        write_xml(target_drawing, drawing_path)

    if context.created_vml:
        for shape in target_vml.getroot().findall(f"{{{NS_V}}}shape"):
            rewrite_vml_shape_ids(shape, shape_id_map)
        rewrite_vml_relationship_ids(target_vml.getroot(), vml_rid_map)
        write_xml(target_vml, vml_path)
    elif has_existing_vml_shapes(target_vml.getroot()):
        append_vml_shapes(
            target_vml.getroot(),
            source_vml.getroot(),
            shape_id_map,
            vml_rid_map,
        )
        write_xml(target_vml, vml_path)
    else:
        rewrite_vml_relationship_ids(target_vml.getroot(), vml_rid_map)
        write_xml(target_vml, vml_path)
    return shape_id_map


def has_existing_drawing_content(root: ET.Element) -> bool:
    return bool(list(root))


def has_existing_vml_shapes(root: ET.Element) -> bool:
    return bool(root.findall(f"{{{NS_V}}}shape"))


def append_drawing_content(
    target_root: ET.Element,
    source_root: ET.Element,
    shape_id_map: dict[int, int],
) -> None:
    for child in list(source_root):
        imported = copy.deepcopy(child)
        rewrite_drawing_shape_ids(imported, shape_id_map)
        target_root.append(imported)


def append_vml_shapes(
    target_root: ET.Element,
    source_root: ET.Element,
    shape_id_map: dict[int, int],
    rid_map: dict[str, str],
) -> None:
    _copy_missing_shapetypes(target_root, source_root)
    for shape in source_root.findall(f"{{{NS_V}}}shape"):
        imported = copy.deepcopy(shape)
        rewrite_vml_shape_ids(imported, shape_id_map)
        rewrite_vml_relationship_ids(imported, rid_map)
        target_root.append(imported)


def _copy_missing_shapetypes(target_root: ET.Element, source_root: ET.Element) -> None:
    """Carry over ``v:shapetype`` definitions the appended shapes rely on.

    A comment's VML only defines ``_x0000_t202``; dropping our picture shapes
    into it leaves them pointing at an undefined ``#_x0000_t75``.
    """
    have = {
        element.get("id")
        for element in target_root.findall(f"{{{NS_V}}}shapetype")
    }
    for index, shapetype in enumerate(source_root.findall(f"{{{NS_V}}}shapetype")):
        if shapetype.get("id") in have:
            continue
        # Definitions must precede the shapes that reference them.
        target_root.insert(index, copy.deepcopy(shapetype))


def allocate_shape_ids(
    target_drawing_root: ET.Element,
    target_vml_root: ET.Element,
    source_drawing_root: ET.Element,
) -> dict[int, int]:
    max_shape_id = 1024
    for value in iter_shape_ids(target_drawing_root, target_vml_root):
        max_shape_id = max(max_shape_id, value)

    shape_id_map: dict[int, int] = {}
    for c_nv_pr in source_drawing_root.findall(f".//{{{NS_XDR}}}cNvPr"):
        source_id = parse_int(c_nv_pr.get("id"))
        if source_id is None or source_id in shape_id_map:
            continue
        max_shape_id += 1
        shape_id_map[source_id] = max_shape_id
    return shape_id_map


def iter_shape_ids(*roots: ET.Element):
    for root in roots:
        for c_nv_pr in root.findall(f".//{{{NS_XDR}}}cNvPr"):
            value = parse_int(c_nv_pr.get("id"))
            if value is not None:
                yield value
        for shape in root.findall(f".//{{{NS_V}}}shape"):
            value = shape_id_from_vml(shape.get("id", ""))
            if value is not None:
                yield value


def rewrite_drawing_shape_ids(element: ET.Element, shape_id_map: dict[int, int]) -> None:
    for c_nv_pr in element.findall(f".//{{{NS_XDR}}}cNvPr"):
        source_id = parse_int(c_nv_pr.get("id"))
        if source_id not in shape_id_map:
            continue
        target_id = shape_id_map[source_id]
        c_nv_pr.set("id", str(target_id))
        c_nv_pr.set("name", f"Object {target_id - 1024}")
        for compat in c_nv_pr.findall(
            ".//{http://schemas.microsoft.com/office/drawing/2010/main}compatExt"
        ):
            compat.set("spid", f"_x0000_s{target_id}")


def rewrite_vml_shape_ids(element: ET.Element, shape_id_map: dict[int, int]) -> None:
    source_id = shape_id_from_vml(element.get("id", ""))
    if source_id in shape_id_map:
        element.set("id", f"_x0000_s{shape_id_map[source_id]}")


def rewrite_vml_relationship_ids(element: ET.Element, rid_map: dict[str, str]) -> None:
    relid_attr = f"{{{NS_O}}}relid"
    for node in element.iter():
        relid = node.get(relid_attr)
        if relid in rid_map:
            node.set(relid_attr, rid_map[relid])


def shape_id_from_vml(value: str) -> int | None:
    match = re.search(r"_x0000_s(\d+)$", value)
    if not match:
        return None
    return int(match.group(1))
