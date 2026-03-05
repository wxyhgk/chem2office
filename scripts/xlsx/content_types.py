from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET


def ensure_png_content_type(ct_path: Path) -> None:
    if not ct_path.exists():
        return
    ct_xml = ET.parse(ct_path)
    ct_root = ct_xml.getroot()
    has_png = any(
        el.tag == "{http://schemas.openxmlformats.org/package/2006/content-types}Default"
        and el.get("Extension") == "png"
        for el in ct_root
    )
    if not has_png:
        ET.SubElement(
            ct_root,
            "{http://schemas.openxmlformats.org/package/2006/content-types}Default",
            Extension="png",
            ContentType="image/png",
        )
    ct_xml.write(ct_path, encoding="UTF-8", xml_declaration=True)
