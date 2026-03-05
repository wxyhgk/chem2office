from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def build_rels(rels_path: Path, rels: list[tuple[str, str, str]]) -> None:
    root = ET.Element(f"{{{_RELS_NS}}}Relationships")
    for rid, rtype, target in rels:
        ET.SubElement(
            root,
            f"{{{_RELS_NS}}}Relationship",
            Id=rid,
            Type=rtype,
            Target=target,
        )
    xml = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    xml = xml.replace(b"ns0:", b"").replace(b"xmlns:ns0", b"xmlns")
    rels_path.write_bytes(xml)


def build_sheet_relationships(
    rels_path: Path, ole_targets: list[str], image_targets: list[str]
) -> list[tuple[str, str]]:
    rels_xml = ET.parse(rels_path)
    rels_root = rels_xml.getroot()

    keep = []
    existing_ole_ids: list[str] = []
    existing_image_ids: list[str] = []
    max_id = 0
    for rel in rels_root.findall(
        f"{{{_RELS_NS}}}Relationship"
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
        elif rtype.endswith("/oleObject"):
            existing_ole_ids.append(rid)
        elif rtype.endswith("/image"):
            existing_image_ids.append(rid)

    rel_pairs: list[tuple[str, str]] = []
    new_rels = keep[:]
    next_id = max_id + 1
    for idx, (ole_t, img_t) in enumerate(zip(ole_targets, image_targets)):
        if idx < len(existing_ole_ids):
            ole_rid = existing_ole_ids[idx]
        else:
            ole_rid = f"rId{next_id}"
            next_id += 1
        if idx < len(existing_image_ids):
            img_rid = existing_image_ids[idx]
        else:
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
