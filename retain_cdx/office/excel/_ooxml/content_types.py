from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def ensure_png_content_type(ct_path: Path) -> None:
    ensure_content_type_defaults(ct_path, {"png": "image/png"})


def ensure_content_type_defaults(ct_path: Path, defaults: dict[str, str]) -> None:
    if not ct_path.exists():
        return
    ct_xml = ET.parse(ct_path)
    ct_root = ct_xml.getroot()
    existing = {
        el.get("Extension")
        for el in ct_root.findall(f"{{{_CT_NS}}}Default")
    }
    for extension, content_type in defaults.items():
        if extension not in existing:
            ET.SubElement(
                ct_root,
                f"{{{_CT_NS}}}Default",
                Extension=extension,
                ContentType=content_type,
            )
    xml = ET.tostring(ct_root, encoding="UTF-8", xml_declaration=True)
    xml = xml.replace(b"ns0:", b"").replace(b"xmlns:ns0", b"xmlns")
    ct_path.write_bytes(xml)


def ensure_ole_drawing_content_types(package_dir: Path, drawing_name: str) -> None:
    path = package_dir / "[Content_Types].xml"
    media_defaults: dict[str, str] = {}
    media_dir = package_dir / "xl" / "media"
    if media_dir.exists():
        for media in media_dir.iterdir():
            ext = media.suffix.lower().lstrip(".")
            if ext == "png":
                media_defaults["png"] = "image/png"
            elif ext == "emf":
                media_defaults["emf"] = "image/x-emf"
    ensure_content_type_defaults(path, media_defaults)

    tree = ET.parse(path)
    root = tree.getroot()

    defaults = {el.get("Extension") for el in root.findall(f"{{{_CT_NS}}}Default")}
    for ext, content_type in [
        ("bin", "application/vnd.openxmlformats-officedocument.oleObject"),
        ("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"),
    ]:
        if ext not in defaults:
            ET.SubElement(
                root,
                f"{{{_CT_NS}}}Default",
                Extension=ext,
                ContentType=content_type,
            )

    part_name = f"/xl/drawings/{drawing_name}"
    overrides = {el.get("PartName") for el in root.findall(f"{{{_CT_NS}}}Override")}
    if part_name not in overrides:
        ET.SubElement(
            root,
            f"{{{_CT_NS}}}Override",
            PartName=part_name,
            ContentType="application/vnd.openxmlformats-officedocument.drawing+xml",
        )

    xml = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    xml = xml.replace(b"ns0:", b"").replace(b"xmlns:ns0", b"xmlns")
    path.write_bytes(xml)
