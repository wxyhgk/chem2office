"""Source 组装入口（现仅 shell 路径）。

Shell 路径：把一份原生 Excel+ChemDraw shell 包解包、改造，留在 source 目录里，
``_parts.merge`` 再搬进用户表。

generated 路径已改为直接写目标包（`_insert_ole`），不再经过这里。
"""

from __future__ import annotations

from pathlib import Path

from retain_cdx.ole.embed import write_cdx_to_ole
from retain_cdx.structure.geometry import cdx_content_aspect

from retain_cdx.office._package import extract_ooxml
from retain_cdx.office.excel._shell import (
    patch_shell_geometry,
    read_vml_size_px,
    remove_shell_ole_border,
    replace_shell_preview,
    tighten_shell_display_box,
)
from retain_cdx.office.excel._utils import POINTS_PER_PIXEL
from retain_cdx.office.excel.models import LayoutSpec, PreviewSpec, ShellSourceOptions



def build_source_workbook_from_shell(
    cdx_path: Path,
    source_dir: Path,
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
    """
    # Shell options gate how much of the layout box is applied.
    start_col = layout.col if shell.allow_resize else None
    start_row = layout.row if shell.allow_resize else None
    width_px = layout.width_px if shell.resize_display else None
    height_px = layout.height_px if shell.resize_display else None

    package_dir = source_dir
    package_dir.mkdir(parents=True, exist_ok=True)
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
