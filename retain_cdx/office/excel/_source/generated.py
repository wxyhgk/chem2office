"""Generated（非 shell）路径：在临时小表里组装完整 OLE 零件包.

编排顺序（与「装进 Excel」子步骤对齐，只是目标是 source 而非用户表）：

  1. 解包 bundled/template
  2. 算 anchors（钉框几何）
  3. B  fill_parts — .bin + 预览图
  4. C–F wire_surface — rels + drawing/vml + oleObjects + content types
  5. 打成 source.xlsx（供 _parts.merge 搬进目标表）
"""

from __future__ import annotations

import tempfile
from pathlib import Path

from retain_cdx.resources import bundled_template_xlsx
from retain_cdx.office._package import extract_ooxml, pack_ooxml
from retain_cdx.office.excel.models import PreviewSpec

from .anchors import OleAnchor, compute_anchors
from .fill_parts import fill_ole_bins_and_previews
from .template import (
    extract_base_anchor,
    register_namespaces,
    update_row_heights,
    update_sheet_labels,
    write_shared_strings,
)
from .wire_surface import wire_source_surface


def build_generated_source_workbook(
    cdx_paths: list[Path],
    output: Path,
    *,
    col: str,
    start_row: int,
    margin_left: str,
    margin_top: str,
    width: str,
    height: str,
    row_height: float,
    preview: PreviewSpec,
    template: Path | None = None,
) -> None:
    """Assemble a generated source workbook (bundled template by default).

    组装 generated source：不传 *template* 用打包内置模板；传了就用显式模板。
    """
    anchors = compute_anchors(
        len(cdx_paths),
        col=col,
        start_row=start_row,
        margin_left=margin_left,
        margin_top=margin_top,
        width=width,
        height=height,
    )
    if template is not None:
        _assemble_source_package(
            template,
            cdx_paths,
            output,
            anchors=anchors,
            row_height=row_height,
            preview=preview,
        )
        return
    with bundled_template_xlsx() as bundled_template:
        _assemble_source_package(
            bundled_template,
            cdx_paths,
            output,
            anchors=anchors,
            row_height=row_height,
            preview=preview,
        )


def _assemble_source_package(
    template: Path,
    cdx_paths: list[Path],
    output: Path,
    *,
    anchors: list[OleAnchor],
    row_height: float,
    preview: PreviewSpec,
) -> None:
    register_namespaces()
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        extract_ooxml(template, tmpdir)

        sheet_path = tmpdir / "xl" / "worksheets" / "sheet1.xml"
        shared_strings_path = tmpdir / "xl" / "sharedStrings.xml"

        base_shape_id, from_row, to_row = extract_base_anchor(sheet_path)
        labels = [path.stem for path in cdx_paths]
        total_rows = max(anchor.to_pt.row + 1 for anchor in anchors)
        update_sheet_labels(sheet_path, labels, total_rows=total_rows)
        update_row_heights(sheet_path, row_height, total_rows)
        write_shared_strings(shared_strings_path, labels)

        # B: bin + preview images
        filled = fill_ole_bins_and_previews(tmpdir, cdx_paths, preview=preview)

        # C–F: rels + drawing/vml + oleObjects + content types
        wire_source_surface(
            tmpdir,
            filled=filled,
            anchors=anchors,
            base_shape_id=base_shape_id,
            from_row=from_row,
            to_row=to_row,
            row_height=row_height,
            object_count=len(cdx_paths),
        )

        pack_ooxml(tmpdir, output)
