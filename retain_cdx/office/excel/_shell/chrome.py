"""Shell visual chrome only (not geometry, not preview art).

装进 Excel · 子步骤 chrome：去掉 OLE/预览外框黑边。
"""

from __future__ import annotations

from pathlib import Path
import re


def remove_shell_ole_border(package_dir: Path) -> None:
    """Remove black outline around the OLE/preview shape (VML + DrawingML)."""
    for vml_path in package_dir.glob("xl/drawings/*.vml"):
        text = vml_path.read_text(encoding="utf-8")
        text = re.sub(r'\bstroked="t"', 'stroked="f"', text)
        # Drop explicit stroke color so Excel does not repaint a border.
        text = re.sub(r'\s+strokecolor="[^"]*"', "", text)
        # If stroked is absent on the shape, force it off.
        text = re.sub(
            r"(<v:shape\b(?![^>]*\bstroked=)[^>]*?)(/?>)",
            r'\1 stroked="f"\2',
            text,
            count=1,
            flags=re.S,
        )
        vml_path.write_text(text, encoding="utf-8")

    for drawing_path in package_dir.glob("xl/drawings/*.xml"):
        text = drawing_path.read_text(encoding="utf-8")
        if "<a:ln" not in text:
            continue
        text = re.sub(
            r"<a:ln\b[^>]*>.*?</a:ln>",
            "<a:ln><a:noFill/></a:ln>",
            text,
            flags=re.S,
        )
        drawing_path.write_text(text, encoding="utf-8")
