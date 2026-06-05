from __future__ import annotations

from pathlib import Path
import re

from .model import TwoCellAnchor


def patch_first_drawing_anchor(path: Path, anchor: TwoCellAnchor) -> None:
    text = path.read_text(encoding="utf-8")
    text = _replace_first_corner(text, "from", anchor.from_.to_xdr_xml())
    text = _replace_first_corner(text, "to", anchor.to.to_xdr_xml())
    path.write_text(text, encoding="utf-8")


def _replace_first_corner(text: str, corner: str, replacement: str) -> str:
    return re.sub(
        rf"<xdr:{corner}>.*?</xdr:{corner}>",
        f"<xdr:{corner}>{replacement}</xdr:{corner}>",
        text,
        count=1,
        flags=re.S,
    )
