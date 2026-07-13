from __future__ import annotations

from pathlib import Path
import re

from .model import TwoCellAnchor


def patch_first_drawing_anchor(path: Path, anchor: TwoCellAnchor) -> None:
    """Patch the first two-cell anchor corner pair in drawing or worksheet XML.

    DrawingML uses ``<xdr:from>`` / ``<xdr:to>``.
    Worksheet ``oleObject`` uses ``<from>`` / ``<to>`` with ``xdr:*`` children.
    Both must stay in sync or Excel will size the OLE by the sheet anchor and
    stretch the preview (classic distortion bug).
    """
    text = path.read_text(encoding="utf-8")
    text = _replace_first_corner(text, "from", anchor.from_.to_xdr_xml())
    text = _replace_first_corner(text, "to", anchor.to.to_xdr_xml())
    path.write_text(text, encoding="utf-8")


def _replace_first_corner(text: str, corner: str, inner_xml: str) -> str:
    patterns = [
        # drawing1.xml / worksheet drawing parts
        (
            rf"<xdr:{corner}>.*?</xdr:{corner}>",
            f"<xdr:{corner}>{inner_xml}</xdr:{corner}>",
        ),
        # worksheet oleObject objectPr/anchor
        (
            rf"<{corner}>.*?</{corner}>",
            f"<{corner}>{inner_xml}</{corner}>",
        ),
    ]
    for pattern, replacement in patterns:
        new_text, count = re.subn(
            pattern,
            replacement,
            text,
            count=1,
            flags=re.S,
        )
        if count:
            return new_text
    return text
