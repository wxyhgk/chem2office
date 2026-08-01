"""把一份原生 shell 包改造成可 merge 的 staging 目录。

解包 → 注入 CDX → 贴合分子 → 去边框 → 换预览 → 挪格 → 去边框，
留在 staging 目录里由 ``_parts.merge`` 搬进用户表。

generated 与 picture 路径都是直写目标包，只有 shell 需要这一步——它改的是
一整份现成的包，而不是往目标里拼零件。
"""

from __future__ import annotations

from pathlib import Path

from retain_cdx.ole.embed import write_cdx_to_ole
from retain_cdx.structure.geometry import cdx_content_aspect

from retain_cdx.office._package import extract_ooxml
from .chrome import remove_shell_ole_border
from .geometry import patch_shell_geometry, tighten_shell_display_box
from .preview import replace_shell_preview
from .vml import read_vml_size_px
from retain_cdx.office.excel._utils import POINTS_PER_PIXEL
from retain_cdx.office.excel.models import LayoutSpec, PreviewSpec, ShellSourceOptions



def build_shell_staging(
    cdx_path: Path,
    staging_dir: Path,
    *,
    layout: LayoutSpec,
    shell: ShellSourceOptions,
    preview: PreviewSpec,
) -> None:
    """Turn a native OLE shell package into a staging directory.

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

    package_dir = staging_dir
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
