from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET


def build_rels(rels_path: Path, rels: list[tuple[str, str, str]]) -> None:
    root = ET.Element(
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships"
    )
    for rid, rtype, target in rels:
        ET.SubElement(
            root,
            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
            Id=rid,
            Type=rtype,
            Target=target,
        )
    ET.ElementTree(root).write(rels_path, encoding="UTF-8", xml_declaration=True)


def build_sheet_relationships(
    rels_path: Path, ole_targets: list[str], image_targets: list[str]
) -> list[tuple[str, str]]:
    rels_xml = ET.parse(rels_path)
    rels_root = rels_xml.getroot()

    keep = []
    max_id = 0
    for rel in rels_root.findall(
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    ):
        rid = rel.get("Id", "")
        if rid.startswith("rId"):
            try:
                max_id = max(max_id, int(rid[3:]))
            except ValueError:
                pass
        rtype = rel.get("Type", "")
        if rtype.endswith("/drawing") or rtype.endswith("/vmlDrawing"):
            keep.append((rid, rtype, rel.get("Target", "")))

    rel_pairs: list[tuple[str, str]] = []
    new_rels = keep[:]
    next_id = max_id + 1
    for ole_t, img_t in zip(ole_targets, image_targets):
        ole_rid = f"rId{next_id}"
        next_id += 1
        img_rid = f"rId{next_id}"
        next_id += 1
        rel_pairs.append((ole_rid, img_rid))
        new_rels.append(
            (
                ole_rid,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                ole_t,
            )
        )
        new_rels.append(
            (
                img_rid,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                img_t,
            )
        )

    build_rels(rels_path, new_rels)
    return rel_pairs


def build_vml_relationships(rels_path: Path, image_targets: list[str]) -> list[str]:
    rels = []
    rids = []
    for idx, target in enumerate(image_targets, start=1):
        rid = f"rId{idx}"
        rids.append(rid)
        rels.append(
            (
                rid,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                target,
            )
        )
    build_rels(rels_path, rels)
    return rids
