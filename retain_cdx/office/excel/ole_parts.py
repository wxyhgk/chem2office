from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

from .relationships import build_rels
from .worksheet import NS_MAIN, attach_ole_objects


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def merge_source_workbook(
    target_dir: Path,
    source_dir: Path,
    *,
    sheet_path: Path,
    sheet_idx: int,
    prefix: str,
    preserve_object_pr: bool = True,
) -> str:
    """Merge OLE/drawing parts from a source OLE workbook into target sheet.

    Returns the drawing part name added to ``xl/drawings``.
    """
    drawing_name = f"{prefix}_drawing.xml"
    vml_name = f"{prefix}_vmlDrawing.vml"

    for rel in ["xl/drawings", "xl/drawings/_rels", "xl/embeddings", "xl/media"]:
        (target_dir / rel).mkdir(parents=True, exist_ok=True)

    (target_dir / "xl" / "drawings" / drawing_name).write_bytes(
        (source_dir / "xl" / "drawings" / "drawing1.xml").read_bytes()
    )
    (target_dir / "xl" / "drawings" / vml_name).write_bytes(
        (source_dir / "xl" / "drawings" / "vmlDrawing1.vml").read_bytes()
    )

    rels = [
        (
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            f"../drawings/{drawing_name}",
        ),
        (
            "rId2",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
            f"../drawings/{vml_name}",
        ),
    ]
    vml_rels: list[tuple[str, str, str]] = []

    source_rels = ET.parse(
        source_dir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels"
    ).getroot()
    for rel in source_rels.findall(f"{{{REL_NS}}}Relationship"):
        rid = rel.get("Id", "")
        rtype = rel.get("Type", "")
        target = rel.get("Target", "")
        if rtype.endswith("/oleObject"):
            name = Path(target).name
            new_name = f"{prefix}_{name}"
            (target_dir / "xl" / "embeddings" / new_name).write_bytes(
                (source_dir / "xl" / "embeddings" / name).read_bytes()
            )
            rels.append((rid, rtype, f"../embeddings/{new_name}"))
        elif rtype.endswith("/image"):
            name = Path(target).name
            new_name = f"{prefix}_{name}"
            (target_dir / "xl" / "media" / new_name).write_bytes(
                (source_dir / "xl" / "media" / name).read_bytes()
            )
            rels.append((rid, rtype, f"../media/{new_name}"))

    source_vml_rels = ET.parse(
        source_dir / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels"
    ).getroot()
    for rel in source_vml_rels.findall(f"{{{REL_NS}}}Relationship"):
        target = rel.get("Target", "")
        name = Path(target).name
        vml_rels.append(
            (rel.get("Id", ""), rel.get("Type", ""), f"../media/{prefix}_{name}")
        )

    sheet_rels_path = target_dir / "xl" / "worksheets" / "_rels" / f"sheet{sheet_idx}.xml.rels"
    sheet_rels_path.parent.mkdir(parents=True, exist_ok=True)
    build_rels(sheet_rels_path, rels)
    build_rels(
        target_dir / "xl" / "drawings" / "_rels" / f"{vml_name}.rels",
        vml_rels,
    )

    source_sheet_root = ET.parse(source_dir / "xl" / "worksheets" / "sheet1.xml").getroot()
    source_ole_objects = source_sheet_root.find("s:oleObjects", {"s": NS_MAIN})
    if source_ole_objects is None:
        raise ValueError("Source OLE workbook has no oleObjects")
    if not preserve_object_pr:
        _normalize_object_pr(source_ole_objects)

    attach_ole_objects(
        sheet_path,
        drawing_rid="rId1",
        vml_rid="rId2",
        ole_objects=source_ole_objects,
    )
    return drawing_name


def _normalize_object_pr(ole_objects: ET.Element) -> None:
    for object_pr in ole_objects.findall(".//s:objectPr", {"s": NS_MAIN}):
        object_pr.attrib.pop("autoPict", None)
