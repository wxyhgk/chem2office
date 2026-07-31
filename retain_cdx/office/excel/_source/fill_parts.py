"""在 source package 里填充 .bin 与预览图（装进 Excel 之前的「造零件落盘」）.

对应子步骤 B：
  · xl/embeddings/oleObjectN.bin  ← CDX 写入 OLE 壳
  · xl/media/imageN.png|emf       ← 预览图
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from retain_cdx.ole.embed import write_cdx_to_ole
from retain_cdx.render.preview import (
    preview_extension,
    render_emf_many,
    render_preview,
    resolve_preview_format,
)

from retain_cdx.office.excel.models import PreviewSpec


@dataclass(frozen=True)
class FilledParts:
    """Relative targets suitable for sheet/vml .rels."""

    ole_targets: list[str]
    image_targets: list[str]
    media_ext: str


def fill_ole_bins_and_previews(
    package_dir: Path,
    cdx_paths: list[Path],
    *,
    preview: PreviewSpec,
) -> FilledParts:
    """Clone template OLE shells, inject CDX, render previews into *package_dir*."""
    ole_dir = package_dir / "xl" / "embeddings"
    ole_dir.mkdir(parents=True, exist_ok=True)
    base_ole_path = ole_dir / "oleObject1.bin"
    if not base_ole_path.exists():
        raise ValueError(f"OLE object not found: {base_ole_path}")
    base_ole_bytes = base_ole_path.read_bytes()

    media_dir = package_dir / "xl" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)

    ext = preview_extension(preview.format)
    resolved_format = resolve_preview_format(preview.format)

    ole_targets: list[str] = []
    image_targets: list[str] = []
    emf_jobs: list[tuple[Path, Path]] = []

    for idx, cdx_path in enumerate(cdx_paths, start=1):
        ole_path = ole_dir / f"oleObject{idx}.bin"
        if idx == 1:
            if not ole_path.exists():
                raise ValueError(f"OLE object not found: {ole_path}")
        else:
            ole_path.write_bytes(base_ole_bytes)

        write_cdx_to_ole(ole_path, cdx_path, stream_name="CONTENTS")

        image_name = f"image{idx}.{ext}"
        image_path = media_dir / image_name
        # Strategy B: fixed OLE box; preview letterboxed — never stretch molecule.
        if resolved_format == "emf":
            emf_jobs.append((cdx_path, image_path))
        else:
            render_preview(
                cdx_path,
                image_path,
                format=preview.format,
                width=preview.width,
                height=preview.height,
                thick_lines=preview.thick_lines,
                no_margin=preview.no_margin,
                mono=preview.mono,
                backend=preview.backend,
                style=preview.style,
                crop=True,
                target_aspect=preview.target_aspect,
                fill_ratio=preview.fill_ratio,
            )

        ole_targets.append(f"../embeddings/oleObject{idx}.bin")
        image_targets.append(f"../media/{image_name}")

    if emf_jobs:
        render_emf_many(
            emf_jobs,
            width=preview.width,
            height=preview.height,
            thick_lines=preview.thick_lines,
            no_margin=preview.no_margin,
            mono=preview.mono,
            backend=preview.backend,
            style=preview.style,
            target_aspect=preview.target_aspect,
            fill_ratio=preview.fill_ratio,
        )

    # Drop leftover template media with other extensions.
    keep = {Path(t).name for t in image_targets}
    for leftover in media_dir.glob("image*"):
        if leftover.name not in keep:
            leftover.unlink()

    return FilledParts(
        ole_targets=ole_targets,
        image_targets=image_targets,
        media_ext=ext,
    )
