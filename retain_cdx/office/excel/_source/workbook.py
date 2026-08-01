"""Source workbook entry points: generated multi-OLE vs native shell.

Both paths produce a temporary source.xlsx that already contains a complete
OLE assembly; ``_parts.merge`` then places it into the user workbook.
"""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from retain_cdx.ole.embed import write_cdx_to_ole
from retain_cdx.render.preview import resolve_preview_format
from retain_cdx.structure.geometry import cdx_content_aspect

from retain_cdx.office._package import extract_ooxml, pack_ooxml
from retain_cdx.office.excel._shell import (
    patch_shell_geometry,
    read_vml_size_px,
    remove_shell_ole_border,
    replace_shell_preview,
    tighten_shell_display_box,
)
from retain_cdx.office.excel._utils import POINTS_PER_PIXEL, index_to_col, px_to_points
from retain_cdx.office.excel.models import LayoutSpec, PreviewSpec, ShellSourceOptions

from .generated import build_generated_source_workbook


def build_source_ole_workbook(
    cdx_paths: list[Path],
    output: Path,
    *,
    layout: LayoutSpec,
    preview: PreviewSpec,
) -> None:
    if layout.row_height is None:
        raise ValueError("LayoutSpec.row_height is required for the generated path.")
    # XLSX anchors use 96-dpi pixels. The existing generator accepts metric
    # lengths, so convert pixels to points for a stable display box.
    offset_x_pt = px_to_points(layout.offset_x_px)
    offset_y_pt = px_to_points(layout.offset_y_px)
    width_pt = px_to_points(layout.width_px)
    height_pt = px_to_points(layout.height_px)
    # Strategy B: fixed OLE box; preview letterboxed to box aspect (no stretch).
    box_aspect = (layout.width_px / layout.height_px) if layout.height_px else None
    preview_width = preview.width
    preview_height = preview.height
    # Raster default: supersample the display box; small molecules otherwise
    # flexi-render at ~25 px and Excel upscales them blurrily.
    if (
        preview_width is None
        and preview_height is None
        and resolve_preview_format(preview.format) == "png"
    ):
        preview_width = layout.width_px * 4
        preview_height = layout.height_px * 4
    # The spec DTO carries the computed letterbox target downstream
    # (fill_ratio keeps its PreviewSpec value; no re-declared kwargs).
    preview = replace(
        preview,
        width=preview_width,
        height=preview_height,
        target_aspect=box_aspect,
    )
    build_generated_source_workbook(
        cdx_paths,
        output,
        col=index_to_col(layout.col),
        start_row=layout.row,
        margin_left=f"{offset_x_pt}pt",
        margin_top=f"{offset_y_pt}pt",
        width=f"{width_pt}pt",
        height=f"{height_pt}pt",
        row_height=layout.row_height,
        preview=preview,
    )


def build_source_workbook_from_shell(
    cdx_path: Path,
    output: Path,
    *,
    layout: LayoutSpec,
    shell: ShellSourceOptions,
    preview: PreviewSpec,
) -> None:
    """Build a one-object source workbook by reusing a native OLE shell.

    Explicit substeps (same language as generated fill/wire + merge)::

      B   inject CDX into existing .bin
      D'  tighten display box to molecule aspect
      chrome  remove black border
      B+C render preview + wire image rels
      D   optional move/resize into target cell
      chrome  border again (geometry rewrite may restore stroke)
      pack
    """
    import tempfile

    # Shell options gate how much of the layout box is applied.
    start_col = layout.col if shell.allow_resize else None
    start_row = layout.row if shell.allow_resize else None
    width_px = layout.width_px if shell.resize_display else None
    height_px = layout.height_px if shell.resize_display else None

    with tempfile.TemporaryDirectory(prefix="retain_cdx_shell_") as tmp:
        package_dir = Path(tmp) / "package"
        package_dir.mkdir()
        extract_ooxml(shell.path, package_dir)

        embeddings = sorted((package_dir / "xl" / "embeddings").glob("*.bin"))
        if not embeddings:
            raise ValueError(f"Shell template has no OLE embedding: {shell.path}")

        # B: OLE payload
        if shell.fit_cdx:
            cdx_path = _fit_cdx_to_shell(cdx_path, package_dir)
        write_cdx_to_ole(embeddings[0], cdx_path, stream_name="CONTENTS")

        # D': hug molecule aspect before drawing the preview
        content_aspect = cdx_content_aspect(Path(cdx_path).read_bytes())
        if content_aspect and content_aspect > 0:
            tighten_shell_display_box(package_dir, content_aspect=content_aspect)

        # chrome before preview so art matches a clean frame
        remove_shell_ole_border(package_dir)

        # B+C: preview media + image relationships (no geometry/chrome inside)
        replace_shell_preview(cdx_path, package_dir, preview=preview)

        # D: pin position / size when caller allows resize
        if start_col is not None and start_row is not None:
            patch_shell_geometry(
                package_dir,
                start_col=start_col,
                start_row=start_row,
                offset_x_px=layout.offset_x_px,
                offset_y_px=layout.offset_y_px,
                width_px=width_px,
                height_px=height_px,
                preserve_aspect=shell.preserve_aspect,
                # Auto-center when caller did not supply an explicit in-cell offset.
                center_in_cell=(layout.offset_x_px == 0 and layout.offset_y_px == 0),
            )

        # Geometry rewrite can reintroduce template stroke attrs; clear again.
        remove_shell_ole_border(package_dir)
        pack_ooxml(package_dir, output)


def _fit_cdx_to_shell(cdx_path: Path, package_dir: Path) -> Path:
    from retain_cdx.structure import StructureSource

    vml_path = package_dir / "xl" / "drawings" / "vmlDrawing1.vml"
    vml_text = vml_path.read_text(encoding="utf-8")
    shell_width_px, shell_height_px = read_vml_size_px(vml_text)
    if shell_width_px <= 0 or shell_height_px <= 0:
        return cdx_path
    fitted = StructureSource.from_cdx(cdx_path).normalized(
        target_aspect=shell_width_px / shell_height_px,
        target_width=shell_width_px * POINTS_PER_PIXEL,
        target_height=shell_height_px * POINTS_PER_PIXEL,
        fill_ratio=0.8,
    )
    output = cdx_path.with_name(f"{cdx_path.stem}_fit_shell.cdx")
    fitted.write_cdx(output)
    return output
