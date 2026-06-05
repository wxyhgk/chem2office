from __future__ import annotations

import xml.etree.ElementTree as ET

from .template_namespaces import NS


def set_anchor_corner(
    corner: ET.Element | None,
    *,
    col: int,
    col_off: int,
    row: int,
    row_off: int,
) -> None:
    if corner is None:
        return
    mapping = {
        "xdr:col": str(col),
        "xdr:colOff": str(col_off),
        "xdr:row": str(row),
        "xdr:rowOff": str(row_off),
    }
    for tag, value in mapping.items():
        el = corner.find(tag, NS)
        if el is not None:
            el.text = value
