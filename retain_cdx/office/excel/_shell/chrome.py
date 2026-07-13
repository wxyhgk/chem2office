from __future__ import annotations

from pathlib import Path
import re

from .grid import ExcelGridMetrics
from .model import AnchorCorner, ShellSize, TwoCellAnchor
from .ooxml_patch import patch_first_drawing_anchor
from .vml import patch_vml_shape, read_vml_shell_size


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


def fit_size_to_content_aspect(
    shell: ShellSize,
    content_aspect: float,
) -> ShellSize:
    """Shrink the shell box so its aspect matches the molecule (tight frame)."""
    if content_aspect <= 0 or shell.width_px <= 0 or shell.height_px <= 0:
        return shell
    shell_aspect = shell.width_px / shell.height_px
    if content_aspect >= shell_aspect:
        # Molecule wider than shell → limited by shell width.
        width = shell.width_px
        height = max(1, round(width / content_aspect))
    else:
        height = shell.height_px
        width = max(1, round(height * content_aspect))
    return ShellSize(width, height)


def tighten_shell_display_box(
    package_dir: Path,
    *,
    content_aspect: float,
) -> ShellSize:
    """Resize OLE display box to molecule aspect, keeping current top-left."""
    vml_path = package_dir / "xl" / "drawings" / "vmlDrawing1.vml"
    drawing_path = package_dir / "xl" / "drawings" / "drawing1.xml"
    sheet_path = package_dir / "xl" / "worksheets" / "sheet1.xml"

    vml_text = vml_path.read_text(encoding="utf-8")
    shell = read_vml_shell_size(vml_text)
    target = fit_size_to_content_aspect(shell, content_aspect)
    if target.width_px == shell.width_px and target.height_px == shell.height_px:
        return target

    from_corner = _read_first_from_corner(drawing_path, vml_text)
    grid = ExcelGridMetrics.from_worksheet(sheet_path)
    to_col, to_col_off = grid.advance_cols(
        from_corner.col, from_corner.col_off_px + target.width_px
    )
    to_row, to_row_off = grid.advance_rows(
        from_corner.row, from_corner.row_off_px + target.height_px
    )
    anchor = TwoCellAnchor(
        from_=from_corner,
        to=AnchorCorner(to_col, to_col_off, to_row, to_row_off),
    )

    for path in (sheet_path, drawing_path):
        if path.exists():
            patch_first_drawing_anchor(path, anchor)

    patch_vml_shape(
        vml_path,
        vml_text=vml_text,
        anchor=anchor,
        left_px=grid.absolute_x_px(anchor.from_),
        top_px=grid.absolute_y_px(anchor.from_),
        size=target,
    )
    # patch_vml_shape rewrites style but may leave stroked=t from template text;
    # border cleanup runs after this.
    return target


def _read_first_from_corner(drawing_path: Path, vml_text: str) -> AnchorCorner:
    if drawing_path.exists():
        text = drawing_path.read_text(encoding="utf-8")
        match = re.search(
            r"<xdr:from>\s*"
            r"<xdr:col>(\d+)</xdr:col>\s*"
            r"<xdr:colOff>(\d+)</xdr:colOff>\s*"
            r"<xdr:row>(\d+)</xdr:row>\s*"
            r"<xdr:rowOff>(\d+)</xdr:rowOff>\s*"
            r"</xdr:from>",
            text,
            flags=re.S,
        )
        if match:
            from retain_cdx.office.excel._utils import EMU_PER_PIXEL

            return AnchorCorner(
                col=int(match.group(1)),
                col_off_px=round(int(match.group(2)) / EMU_PER_PIXEL),
                row=int(match.group(3)),
                row_off_px=round(int(match.group(4)) / EMU_PER_PIXEL),
            )

    anchor_match = re.search(r"<x:Anchor>\s*([^<]+)</x:Anchor>", vml_text)
    if anchor_match:
        parts = [part.strip() for part in anchor_match.group(1).split(",")]
        if len(parts) >= 4:
            return AnchorCorner(
                col=int(parts[0]),
                col_off_px=int(parts[1]),
                row=int(parts[2]),
                row_off_px=int(parts[3]),
            )
    return AnchorCorner(0, 0, 0, 0)
