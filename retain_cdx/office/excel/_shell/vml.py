from __future__ import annotations

from pathlib import Path
import re

from retain_cdx.office.excel._utils import POINTS_PER_PIXEL

from .model import ShellSize, TwoCellAnchor


def read_vml_size_px(vml_text: str) -> tuple[int, int]:
    match = re.search(r"width:([0-9.]+)pt;.*?height:([0-9.]+)pt", vml_text, re.S)
    if not match:
        raise ValueError("Shell VML shape has no width/height style.")
    return round(float(match.group(1)) / POINTS_PER_PIXEL), round(
        float(match.group(2)) / POINTS_PER_PIXEL
    )


def read_vml_shell_size(vml_text: str) -> ShellSize:
    width_px, height_px = read_vml_size_px(vml_text)
    return ShellSize(width_px, height_px)


def patch_vml_shape(
    vml_path: Path,
    *,
    vml_text: str,
    anchor: TwoCellAnchor,
    left_px: int,
    top_px: int,
    size: ShellSize,
) -> None:
    style = (
        "position:absolute;"
        f"margin-left:{_px_to_pt(left_px):g}pt;"
        f"margin-top:{_px_to_pt(top_px):g}pt;"
        f"width:{_px_to_pt(size.width_px):g}pt;"
        f"height:{_px_to_pt(size.height_px):g}pt;"
        "z-index:1"
    )
    if "mso-wrap-style:tight" in vml_text:
        style += ";mso-wrap-style:tight"
    vml_text = re.sub(
        r"(<v:shape\b[^>]*\sstyle=)(['\"])(.*?)\2",
        lambda match: f"{match.group(1)}{match.group(2)}{style}{match.group(2)}",
        vml_text,
        count=1,
        flags=re.S,
    )
    vml_text = re.sub(
        r"<x:Anchor>.*?</x:Anchor>",
        f"<x:Anchor>\n    {anchor.to_vml_anchor()}</x:Anchor>",
        vml_text,
        count=1,
        flags=re.S,
    )
    vml_path.write_text(vml_text, encoding="utf-8")


def _px_to_pt(px: int) -> float:
    return px * POINTS_PER_PIXEL
