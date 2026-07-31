from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

# Canonical relationship namespace + type URIs (single source of truth).
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
REL_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
REL_OLE_OBJECT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def read_rels(rels_path: Path) -> list[tuple[str, str, str]]:
    """Return ``(Id, Type, Target)`` for every relationship (missing file → [])."""
    if not rels_path.exists():
        return []
    root = ET.parse(rels_path).getroot()
    return [
        (rel.get("Id", ""), rel.get("Type", ""), rel.get("Target", ""))
        for rel in root.findall(f"{{{RELS_NS}}}Relationship")
    ]


def rewrite_rel_targets(
    rels_path: Path,
    target: str,
    *,
    rel_id: str | None = None,
    rel_type: str | None = None,
) -> None:
    """Point matching relationships at *target* in place (other attributes kept)."""
    tree = ET.parse(rels_path)
    for rel in tree.getroot():
        if rel_id is not None and rel.get("Id") != rel_id:
            continue
        if rel_type is not None and rel.get("Type", "") != rel_type:
            continue
        rel.set("Target", target)
    tree.write(rels_path, encoding="UTF-8", xml_declaration=True)


def build_rels(rels_path: Path, rels: list[tuple[str, str, str]]) -> None:
    root = ET.Element(f"{{{RELS_NS}}}Relationships")
    for rid, rtype, target in rels:
        ET.SubElement(
            root,
            f"{{{RELS_NS}}}Relationship",
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
    keep = []
    existing_ole_ids: list[str] = []
    existing_image_ids: list[str] = []
    max_id = 0
    for rid, rtype, target in read_rels(rels_path):
        if rid.startswith("rId"):
            try:
                max_id = max(max_id, int(rid[3:]))
            except ValueError:
                pass
        if rtype.endswith("/drawing") or rtype.endswith("/vmlDrawing"):
            keep.append((rid, rtype, target))
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
        new_rels.append((ole_rid, REL_OLE_OBJECT, ole_t))
        new_rels.append((img_rid, REL_IMAGE, img_t))

    build_rels(rels_path, new_rels)
    return rel_pairs


def build_vml_relationships(rels_path: Path, image_targets: list[str]) -> list[str]:
    rels = []
    rids = []
    for idx, target in enumerate(image_targets, start=1):
        rid = f"rId{idx}"
        rids.append(rid)
        rels.append((rid, REL_IMAGE, target))
    build_rels(rels_path, rels)
    return rids
