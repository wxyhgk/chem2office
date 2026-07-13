from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

from retain_cdx.render.preview import (
    PreviewBackend,
    PreviewFormat,
    PreviewStyle,
    preview_extension,
    render_preview,
)
from retain_cdx.structure.geometry import cdx_content_aspect

from .chrome import remove_shell_ole_border, tighten_shell_display_box
from .geometry import read_vml_size_px


REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

# Structure should nearly fill the (already tightened) display box.
_SHELL_FILL_RATIO = 0.98


def replace_shell_preview(
    cdx_path: Path,
    package_dir: Path,
    *,
    width: int | None,
    height: int | None,
    thick_lines: bool,
    no_margin: bool,
    format: PreviewFormat = "png",
    mono: bool = True,
    backend: PreviewBackend = "auto",
    style: PreviewStyle = "acs1996",
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
    stem = Path(old_name).stem or "image1"
    ext = preview_extension(format)
    new_name = f"{stem}.{ext}"
    preview_path = media_dir / new_name

    # 1) Shrink OLE frame to molecule aspect so the outer box hugs the structure.
    content_aspect = cdx_content_aspect(Path(cdx_path).read_bytes())
    if content_aspect and content_aspect > 0:
        tighten_shell_display_box(package_dir, content_aspect=content_aspect)

    # 2) Drop black VML/DrawingML outline (shell_large has a visible border).
    remove_shell_ole_border(package_dir)

    shell_width_px, shell_height_px = read_vml_size_px(
        (package_dir / "xl" / "drawings" / "vmlDrawing1.vml").read_text(encoding="utf-8")
    )
    target_aspect = None
    if shell_width_px > 0 and shell_height_px > 0:
        target_aspect = shell_width_px / shell_height_px
    elif content_aspect and content_aspect > 0:
        target_aspect = content_aspect

    # 3) Preview art matches the (tight) display box aspect.
    #    PNG: crop + letterbox; EMF: SVG viewBox letterbox → vector EMF.
    render_preview(
        cdx_path,
        preview_path,
        format=format,
        width=width,
        height=height,
        thick_lines=thick_lines,
        no_margin=no_margin,
        mono=mono,
        backend=backend,
        style=style,
        crop=(ext == "png"),
        crop_padding=2,
        target_aspect=target_aspect,
        fill_ratio=_SHELL_FILL_RATIO,
    )

    # Drop the previous media file when the extension changes (e.g. emf -> png).
    old_path = media_dir / old_name
    if old_path != preview_path and old_path.exists():
        old_path.unlink()

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
