"""Replace shell preview image (media art + image rels only).

装进 Excel · shell 路径 · B + C（预览侧）：
  B  渲染新预览图写入 xl/media
  C  更新 vml/sheet image relationships

Geometry (tighten) and chrome (border) belong to the caller
(``_shell.staging``), so this module stays a single responsibility.
"""

from __future__ import annotations

from pathlib import Path

from retain_cdx.render.preview import preview_extension, render_preview
from retain_cdx.structure.geometry import cdx_content_aspect

from retain_cdx.office.excel.models import PreviewSpec

from .geometry import read_vml_size_px
from .media_rels import (
    find_image_relationship,
    set_all_sheet_image_targets,
    set_image_relationship,
)

# Structure should nearly fill the display box (caller may have tightened it).
_SHELL_FILL_RATIO = 0.98


def replace_shell_preview(
    cdx_path: Path,
    package_dir: Path,
    *,
    preview: PreviewSpec,
    fill_ratio: float = _SHELL_FILL_RATIO,
) -> None:
    """Write preview media and rewire image relationships.

    If ``preview.target_aspect`` is unset, uses the current VML display box
    aspect. *fill_ratio* stays a separate shell knob: the display box is
    already tightened here, so ``preview.fill_ratio`` (a generated-path
    letterbox default) does not apply.
    """
    width = preview.width
    height = preview.height
    vml_rels_path = package_dir / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels"
    image_rel = find_image_relationship(vml_rels_path)
    if image_rel is None:
        return

    rel_id, old_target = image_rel
    media_dir = package_dir / "xl" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)
    old_name = Path(old_target).name or "image1.emf"
    stem = Path(old_name).stem or "image1"
    ext = preview_extension(preview.format)
    new_name = f"{stem}.{ext}"
    preview_path = media_dir / new_name

    vml_path = package_dir / "xl" / "drawings" / "vmlDrawing1.vml"
    shell_w = shell_h = 0
    if vml_path.exists():
        try:
            shell_w, shell_h = read_vml_size_px(vml_path.read_text(encoding="utf-8"))
        except ValueError:
            pass

    aspect = preview.target_aspect
    if (aspect is None or aspect <= 0) and shell_w > 0 and shell_h > 0:
        aspect = shell_w / shell_h
    if aspect is None or aspect <= 0:
        # VML box unparseable: fall back to the molecule's own aspect.
        content_aspect = cdx_content_aspect(Path(cdx_path).read_bytes())
        if content_aspect and content_aspect > 0:
            aspect = content_aspect

    # Raster default: supersample the shell box; small molecules otherwise
    # flexi-render at ~25 px and Excel upscales them blurrily.
    if width is None and height is None and ext == "png" and shell_w > 0 and shell_h > 0:
        width = round(shell_w * 4)
        height = round(shell_h * 4)

    # B: preview art (letterbox into display aspect — never stretch molecule).
    render_preview(
        cdx_path,
        preview_path,
        format=preview.format,
        width=width,
        height=height,
        thick_lines=preview.thick_lines,
        no_margin=preview.no_margin,
        mono=preview.mono,
        backend=preview.backend,
        style=preview.style,
        crop=(ext == "png"),
        crop_padding=2,
        target_aspect=aspect,
        fill_ratio=fill_ratio,
    )

    # Drop the previous media file when the extension changes (e.g. emf -> png).
    old_path = media_dir / old_name
    if old_path != preview_path and old_path.exists():
        old_path.unlink()

    # C: rewire image relationships
    new_target = f"../media/{new_name}"
    set_image_relationship(vml_rels_path, rel_id, new_target)
    set_all_sheet_image_targets(
        package_dir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels",
        new_target,
    )
