from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field
from pathlib import Path
import xml.etree.ElementTree as ET

from retain_cdx.office.excel._ooxml import (
    NS_MAIN,
    attach_ole_objects,
    build_rels,
    copy_sheet_layout,
)


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_V = "urn:schemas-microsoft-com:vml"

REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
REL_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"


def merge_source_workbook(
    target_dir: Path,
    source_dir: Path,
    *,
    sheet_path: Path,
    sheet_idx: int,
    prefix: str,
    preserve_object_pr: bool = True,
    copy_layout: bool = False,
) -> str:
    """Merge OLE/drawing parts from a source OLE workbook into target sheet.

    Returns the drawing part name used by the target sheet.
    """
    context = _OlePartMergeContext(
        source_dir=source_dir,
        target_dir=target_dir,
        prefix=prefix,
        sheet_rels_path=target_dir
        / "xl"
        / "worksheets"
        / "_rels"
        / f"sheet{sheet_idx}.xml.rels",
    )

    context.ensure_part_dirs()
    context.load_sheet_rels()
    context.ensure_drawing_parts()
    rid_map = context.merge_sheet_relationships()
    vml_rid_map = context.merge_vml_relationships()
    shape_id_map = context.merge_drawing_and_vml(vml_rid_map)

    source_sheet_root = ET.parse(source_dir / "xl" / "worksheets" / "sheet1.xml").getroot()
    source_ole_objects = copy.deepcopy(
        source_sheet_root.find("s:oleObjects", {"s": NS_MAIN})
    )
    if source_ole_objects is None:
        raise ValueError("Source OLE workbook has no oleObjects")
    if not preserve_object_pr:
        _normalize_object_pr(source_ole_objects)
    _rewrite_ole_shape_ids(source_ole_objects, shape_id_map)

    if copy_layout:
        copy_sheet_layout(source_dir / "xl" / "worksheets" / "sheet1.xml", sheet_path)

    attach_ole_objects(
        sheet_path,
        drawing_rid=context.drawing_rid,
        vml_rid=context.vml_rid,
        ole_objects=source_ole_objects,
        rid_map=rid_map,
    )
    return context.drawing_name


@dataclass
class _OlePartMergeContext:
    source_dir: Path
    target_dir: Path
    prefix: str
    sheet_rels_path: Path
    sheet_rels: list[tuple[str, str, str]] = field(default_factory=list)
    drawing_rid: str = ""
    vml_rid: str = ""
    drawing_name: str = ""
    vml_name: str = ""
    created_drawing: bool = False
    created_vml: bool = False
    copied_media: dict[str, str] = field(default_factory=dict)

    def ensure_part_dirs(self) -> None:
        for rel in ["xl/drawings", "xl/drawings/_rels", "xl/embeddings", "xl/media"]:
            (self.target_dir / rel).mkdir(parents=True, exist_ok=True)
        self.sheet_rels_path.parent.mkdir(parents=True, exist_ok=True)

    def load_sheet_rels(self) -> None:
        self.sheet_rels = _read_rels(self.sheet_rels_path)
        for rid, rtype, target in self.sheet_rels:
            if rtype.endswith("/drawing") and not self.drawing_rid:
                self.drawing_rid = rid
                self.drawing_name = Path(target).name
            elif rtype.endswith("/vmlDrawing") and not self.vml_rid:
                self.vml_rid = rid
                self.vml_name = Path(target).name

    def ensure_drawing_parts(self) -> None:
        if not self.drawing_name:
            self.created_drawing = True
            self.drawing_name = _unique_part_name(
                self.target_dir / "xl" / "drawings",
                f"{self.prefix}_drawing.xml",
            )
            self.drawing_rid = self.next_sheet_rid()
            self.sheet_rels.append((self.drawing_rid, REL_DRAWING, f"../drawings/{self.drawing_name}"))
            (self.target_dir / "xl" / "drawings" / self.drawing_name).write_bytes(
                (self.source_dir / "xl" / "drawings" / "drawing1.xml").read_bytes()
            )

        if not self.vml_name:
            self.created_vml = True
            self.vml_name = _unique_part_name(
                self.target_dir / "xl" / "drawings",
                f"{self.prefix}_vmlDrawing.vml",
            )
            self.vml_rid = self.next_sheet_rid()
            self.sheet_rels.append((self.vml_rid, REL_VML, f"../drawings/{self.vml_name}"))
            (self.target_dir / "xl" / "drawings" / self.vml_name).write_bytes(
                (self.source_dir / "xl" / "drawings" / "vmlDrawing1.vml").read_bytes()
            )

    def merge_sheet_relationships(self) -> dict[str, str]:
        rid_map: dict[str, str] = {}
        source_rels = _read_rels(
            self.source_dir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels"
        )
        for source_rid, rtype, target in source_rels:
            if rtype.endswith("/oleObject"):
                new_rid = self.next_sheet_rid()
                copied_target = self.copy_ole_target(target)
                self.sheet_rels.append((new_rid, rtype, copied_target))
                rid_map[source_rid] = new_rid
            elif rtype.endswith("/image"):
                new_rid = self.next_sheet_rid()
                copied_target = self.copy_media_target(target)
                self.sheet_rels.append((new_rid, rtype, copied_target))
                rid_map[source_rid] = new_rid
        build_rels(self.sheet_rels_path, self.sheet_rels)
        return rid_map

    def merge_vml_relationships(self) -> dict[str, str]:
        vml_rels_path = self.target_dir / "xl" / "drawings" / "_rels" / f"{self.vml_name}.rels"
        vml_rels = _read_rels(vml_rels_path)
        rid_map: dict[str, str] = {}
        source_vml_rels = _read_rels(
            self.source_dir / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels"
        )
        for source_rid, rtype, target in source_vml_rels:
            new_rid = _next_rid(vml_rels)
            copied_target = self.copy_media_target(target)
            vml_rels.append((new_rid, rtype, copied_target))
            rid_map[source_rid] = new_rid
        build_rels(vml_rels_path, vml_rels)
        return rid_map

    def merge_drawing_and_vml(self, vml_rid_map: dict[str, str]) -> dict[int, int]:
        drawing_path = self.target_dir / "xl" / "drawings" / self.drawing_name
        vml_path = self.target_dir / "xl" / "drawings" / self.vml_name
        target_drawing = ET.parse(drawing_path)
        source_drawing = ET.parse(self.source_dir / "xl" / "drawings" / "drawing1.xml")
        target_vml = ET.parse(vml_path)
        source_vml = ET.parse(self.source_dir / "xl" / "drawings" / "vmlDrawing1.vml")

        if self.created_drawing and self.created_vml:
            _rewrite_vml_relationship_ids(target_vml.getroot(), vml_rid_map)
            _write_xml(target_vml, vml_path)
            return {}

        shape_id_map = _allocate_shape_ids(
            target_drawing.getroot(),
            target_vml.getroot(),
            source_drawing.getroot(),
        )
        if not self.created_drawing and _has_existing_drawing_content(target_drawing.getroot()):
            _append_drawing_content(
                target_drawing.getroot(),
                source_drawing.getroot(),
                shape_id_map,
            )
            _write_xml(target_drawing, drawing_path)

        if not self.created_vml and _has_existing_vml_shapes(target_vml.getroot()):
            _append_vml_shapes(
                target_vml.getroot(),
                source_vml.getroot(),
                shape_id_map,
                vml_rid_map,
            )
            _write_xml(target_vml, vml_path)
        else:
            _rewrite_vml_relationship_ids(target_vml.getroot(), vml_rid_map)
            _write_xml(target_vml, vml_path)
        return shape_id_map

    def copy_ole_target(self, target: str) -> str:
        name = Path(target).name
        new_name = _unique_part_name(self.target_dir / "xl" / "embeddings", f"{self.prefix}_{name}")
        (self.target_dir / "xl" / "embeddings" / new_name).write_bytes(
            (self.source_dir / "xl" / "embeddings" / name).read_bytes()
        )
        return f"../embeddings/{new_name}"

    def copy_media_target(self, target: str) -> str:
        name = Path(target).name
        if name in self.copied_media:
            return self.copied_media[name]
        new_name = _unique_part_name(self.target_dir / "xl" / "media", f"{self.prefix}_{name}")
        copied_target = f"../media/{new_name}"
        (self.target_dir / "xl" / "media" / new_name).write_bytes(
            (self.source_dir / "xl" / "media" / name).read_bytes()
        )
        self.copied_media[name] = copied_target
        return copied_target

    def next_sheet_rid(self) -> str:
        return _next_rid(self.sheet_rels)


def _read_rels(rels_path: Path) -> list[tuple[str, str, str]]:
    if not rels_path.exists():
        return []
    root = ET.parse(rels_path).getroot()
    return [
        (rel.get("Id", ""), rel.get("Type", ""), rel.get("Target", ""))
        for rel in root.findall(f"{{{REL_NS}}}Relationship")
    ]


def _next_rid(rels: list[tuple[str, str, str]]) -> str:
    max_id = 0
    for rid, _, _ in rels:
        match = re.fullmatch(r"rId(\d+)", rid)
        if match:
            max_id = max(max_id, int(match.group(1)))
    return f"rId{max_id + 1}"


def _unique_part_name(directory: Path, desired_name: str) -> str:
    candidate = desired_name
    path = directory / candidate
    if not path.exists():
        return candidate
    stem = path.stem
    suffix = path.suffix
    idx = 2
    while True:
        candidate = f"{stem}_{idx}{suffix}"
        if not (directory / candidate).exists():
            return candidate
        idx += 1


def _has_existing_drawing_content(root: ET.Element) -> bool:
    return bool(list(root))


def _has_existing_vml_shapes(root: ET.Element) -> bool:
    return bool(root.findall(f"{{{NS_V}}}shape"))


def _append_drawing_content(
    target_root: ET.Element,
    source_root: ET.Element,
    shape_id_map: dict[int, int],
) -> None:
    for child in list(source_root):
        imported = copy.deepcopy(child)
        _rewrite_drawing_shape_ids(imported, shape_id_map)
        target_root.append(imported)


def _append_vml_shapes(
    target_root: ET.Element,
    source_root: ET.Element,
    shape_id_map: dict[int, int],
    rid_map: dict[str, str],
) -> None:
    for shape in source_root.findall(f"{{{NS_V}}}shape"):
        imported = copy.deepcopy(shape)
        _rewrite_vml_shape_ids(imported, shape_id_map)
        _rewrite_vml_relationship_ids(imported, rid_map)
        target_root.append(imported)


def _allocate_shape_ids(
    target_drawing_root: ET.Element,
    target_vml_root: ET.Element,
    source_drawing_root: ET.Element,
) -> dict[int, int]:
    max_shape_id = 1024
    for value in _iter_shape_ids(target_drawing_root, target_vml_root):
        max_shape_id = max(max_shape_id, value)

    shape_id_map: dict[int, int] = {}
    for c_nv_pr in source_drawing_root.findall(f".//{{{NS_XDR}}}cNvPr"):
        source_id = _parse_int(c_nv_pr.get("id"))
        if source_id is None or source_id in shape_id_map:
            continue
        max_shape_id += 1
        shape_id_map[source_id] = max_shape_id
    return shape_id_map


def _iter_shape_ids(*roots: ET.Element):
    for root in roots:
        for c_nv_pr in root.findall(f".//{{{NS_XDR}}}cNvPr"):
            value = _parse_int(c_nv_pr.get("id"))
            if value is not None:
                yield value
        for shape in root.findall(f".//{{{NS_V}}}shape"):
            value = _shape_id_from_vml(shape.get("id", ""))
            if value is not None:
                yield value


def _rewrite_drawing_shape_ids(element: ET.Element, shape_id_map: dict[int, int]) -> None:
    for c_nv_pr in element.findall(f".//{{{NS_XDR}}}cNvPr"):
        source_id = _parse_int(c_nv_pr.get("id"))
        if source_id not in shape_id_map:
            continue
        target_id = shape_id_map[source_id]
        c_nv_pr.set("id", str(target_id))
        c_nv_pr.set("name", f"Object {target_id - 1024}")
        for compat in c_nv_pr.findall(".//{http://schemas.microsoft.com/office/drawing/2010/main}compatExt"):
            compat.set("spid", f"_x0000_s{target_id}")


def _rewrite_vml_shape_ids(element: ET.Element, shape_id_map: dict[int, int]) -> None:
    source_id = _shape_id_from_vml(element.get("id", ""))
    if source_id in shape_id_map:
        element.set("id", f"_x0000_s{shape_id_map[source_id]}")


def _rewrite_vml_relationship_ids(element: ET.Element, rid_map: dict[str, str]) -> None:
    relid_attr = f"{{{NS_O}}}relid"
    for node in element.iter():
        relid = node.get(relid_attr)
        if relid in rid_map:
            node.set(relid_attr, rid_map[relid])


def _rewrite_ole_shape_ids(ole_objects: ET.Element, shape_id_map: dict[int, int]) -> None:
    for ole_object in ole_objects.findall(f".//{{{NS_MAIN}}}oleObject"):
        source_id = _parse_int(ole_object.get("shapeId"))
        if source_id in shape_id_map:
            ole_object.set("shapeId", str(shape_id_map[source_id]))


def _shape_id_from_vml(value: str) -> int | None:
    match = re.search(r"_x0000_s(\d+)$", value)
    if not match:
        return None
    return int(match.group(1))


def _parse_int(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _write_xml(tree: ET.ElementTree, path: Path) -> None:
    tree.write(path, encoding="UTF-8", xml_declaration=True)


def _normalize_object_pr(ole_objects: ET.Element) -> None:
    for object_pr in ole_objects.findall(".//s:objectPr", {"s": NS_MAIN}):
        object_pr.attrib.pop("autoPict", None)
