from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

from retain_cdx.render.preview import render_png

from .geometry import read_vml_size_px


REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def replace_shell_preview(
    cdx_path: Path,
    package_dir: Path,
    *,
    width: int | None,
    height: int | None,
    thick_lines: bool,
    no_margin: bool,
) -> None:
    vml_rels_path = package_dir / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels"
    if not vml_rels_path.exists():
        return

    image_rel = _find_image_relationship(vml_rels_path)
    if image_rel is None:
        return

    rel_id, old_target = image_rel
    media_dir = package_dir / "xl" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)
    old_name = Path(old_target).name or "image1.emf"
    new_name = f"{Path(old_name).stem}.png"
    png_path = media_dir / new_name
    shell_width_px, shell_height_px = read_vml_size_px(
        (package_dir / "xl" / "drawings" / "vmlDrawing1.vml").read_text(encoding="utf-8")
    )
    render_png(
        cdx_path,
        png_path,
        width=width,
        height=height,
        thick_lines=thick_lines,
        no_margin=no_margin,
        crop=True,
        crop_padding=2,
        target_aspect=shell_width_px / shell_height_px,
    )
    new_target = f"../media/{new_name}"
    _set_image_relationship(vml_rels_path, rel_id, new_target)
    _set_sheet_image_relationships(
        package_dir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels",
        new_target,
    )


def _find_image_relationship(rels_path: Path) -> tuple[str, str] | None:
    root = ET.parse(rels_path).getroot()
    for rel in root:
        if rel.get("Type", "") == REL_IMAGE:
            return rel.get("Id", ""), rel.get("Target", "")
    return None


def _set_image_relationship(rels_path: Path, rel_id: str, target: str) -> None:
    tree = ET.parse(rels_path)
    root = tree.getroot()
    for rel in root:
        if rel.get("Id") == rel_id:
            rel.set("Target", target)
            break
    tree.write(rels_path, encoding="UTF-8", xml_declaration=True)


def _set_sheet_image_relationships(rels_path: Path, target: str) -> None:
    if not rels_path.exists():
        return
    tree = ET.parse(rels_path)
    root = tree.getroot()
    for rel in root:
        if rel.get("Type", "") == REL_IMAGE:
            rel.set("Target", target)
    tree.write(rels_path, encoding="UTF-8", xml_declaration=True)
