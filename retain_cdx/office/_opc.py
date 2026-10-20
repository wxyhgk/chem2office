"""Open Packaging Conventions: relationships, content types, part names.

The package-level bookkeeping every OOXML format shares. PowerPoint and Word
both attach an OLE object as two parts (``.bin`` + preview image) reached
through two relationships, so this lives above the per-format packages — the
formats must not reach into each other's private modules for it.
"""

from __future__ import annotations

import re
from pathlib import Path
import xml.etree.ElementTree as ET

NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

REL_OFFICE_DOCUMENT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
REL_OLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

CONTENT_TYPE_OLE = "application/vnd.openxmlformats-officedocument.oleObject"
MEDIA_CONTENT_TYPES = {"emf": "image/x-emf", "png": "image/png"}

Relationship = tuple[str, str, str]  # (Id, Type, Target)


def rels_path_for(part_path: Path) -> Path:
    """``word/document.xml`` → ``word/_rels/document.xml.rels``."""
    return part_path.parent / "_rels" / f"{part_path.name}.rels"


def read_rels(rels_path: Path) -> list[Relationship]:
    if not rels_path.exists():
        return []
    root = ET.parse(rels_path).getroot()
    return [
        (el.get("Id"), el.get("Type"), el.get("Target"))
        for el in root.findall(f"{{{NS_RELS}}}Relationship")
    ]


def write_rels(rels_path: Path, rels: list[Relationship]) -> None:
    root = ET.Element(f"{{{NS_RELS}}}Relationships")
    for rid, rel_type, target in rels:
        ET.SubElement(root, f"{{{NS_RELS}}}Relationship", Id=rid, Type=rel_type, Target=target)
    rels_path.parent.mkdir(parents=True, exist_ok=True)
    rels_path.write_bytes(_default_namespace(ET.tostring(root, encoding="UTF-8", xml_declaration=True)))


def next_rid(rels: list[Relationship]) -> str:
    used = {
        int(match.group(1))
        for rid, _, _ in rels
        if (match := re.fullmatch(r"rId(\d+)", rid))
    }
    return f"rId{max(used, default=0) + 1}"


def unique_part_name(directory: Path, desired: str) -> str:
    candidate = desired
    index = 2
    while (directory / candidate).exists():
        stem, suffix = Path(desired).stem, Path(desired).suffix
        candidate = f"{stem}_{index}{suffix}"
        index += 1
    return candidate


def main_part_path(package_dir: Path) -> Path:
    """The package's main part (``word/document.xml``, ``ppt/presentation.xml``…).

    Resolved through ``_rels/.rels`` rather than assumed: the part name is a
    convention, the relationship is the contract.
    """
    for _, rel_type, target in read_rels(package_dir / "_rels" / ".rels"):
        if rel_type == REL_OFFICE_DOCUMENT:
            return package_dir / target.lstrip("/")
    raise ValueError(f"No officeDocument relationship in {package_dir}")


def ensure_ole_content_types(package_dir: Path, preview_ext: str) -> None:
    """Declare ``.bin`` (OLE) and the preview's extension as Default types."""
    path = package_dir / "[Content_Types].xml"
    if not path.exists():
        return
    tree = ET.parse(path)
    root = tree.getroot()
    existing = {
        (element.get("Extension") or "").lower()
        for element in root.findall(f"{{{NS_CT}}}Default")
    }
    wanted = {
        "bin": CONTENT_TYPE_OLE,
        preview_ext: MEDIA_CONTENT_TYPES.get(preview_ext, "image/x-emf"),
    }
    changed = False
    for extension, content_type in wanted.items():
        if extension not in existing:
            ET.SubElement(
                root,
                f"{{{NS_CT}}}Default",
                Extension=extension,
                ContentType=content_type,
            )
            changed = True
    if changed:
        path.write_bytes(_default_namespace(ET.tostring(root, encoding="UTF-8", xml_declaration=True)))


def _default_namespace(xml: bytes) -> bytes:
    """ElementTree names the single namespace ``ns0``; OPC parts use the default."""
    return xml.replace(b"ns0:", b"").replace(b"xmlns:ns0", b"xmlns")


__all__ = [
    "CONTENT_TYPE_OLE",
    "MEDIA_CONTENT_TYPES",
    "NS_RELS",
    "REL_IMAGE",
    "REL_OLE",
    "REL_OFFICE_DOCUMENT",
    "Relationship",
    "ensure_ole_content_types",
    "main_part_path",
    "next_rid",
    "read_rels",
    "rels_path_for",
    "unique_part_name",
    "write_rels",
]
