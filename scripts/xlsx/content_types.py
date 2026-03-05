from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def ensure_png_content_type(ct_path: Path) -> None:
    if not ct_path.exists():
        return
    ct_xml = ET.parse(ct_path)
    ct_root = ct_xml.getroot()
    has_png = any(
        el.tag == f"{{{_CT_NS}}}Default"
        and el.get("Extension") == "png"
        for el in ct_root
    )
    if not has_png:
        ET.SubElement(
            ct_root,
            f"{{{_CT_NS}}}Default",
            Extension="png",
            ContentType="image/png",
        )
    xml = ET.tostring(ct_root, encoding="UTF-8", xml_declaration=True)
    xml = xml.replace(b"ns0:", b"").replace(b"xmlns:ns0", b"xmlns")
    ct_path.write_bytes(xml)
