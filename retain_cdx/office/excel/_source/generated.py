"""Generated（非 shell）路径：在临时小表里组装完整 OLE 零件包.

编排顺序（与「装进 Excel」子步骤对齐，只是目标是 source 而非用户表）：

  1. 解包 bundled/template 到 source 目录
  2. 算 anchors（钉框几何）
  3. B  fill_parts — .bin + 预览图
  4. C–F wire_surface — rels + drawing/vml + oleObjects + content types

产出是**解包状态的目录**，直接交给 `_parts.merge`；早先会先打包成
source.xlsx 再立刻解开，那一轮 zip 往返没有任何人读得到。
"""

from __future__ import annotations

from pathlib import Path

from retain_cdx.resources import bundled_template_xlsx
from retain_cdx.office._package import extract_ooxml
from retain_cdx.office.excel.models import PreviewSpec

from .anchors import OleAnchor, compute_anchors
from .fill_parts import fill_ole_bins_and_previews
from .template import extract_base_anchor, register_namespaces
from .wire_surface import wire_source_surface


def build_generated_source_workbook(
    cdx_paths: list[Path],
    source_dir: Path,
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
            source_dir,
            anchors=anchors,
            row_height=row_height,
            preview=preview,
        )
        return
    with bundled_template_xlsx() as bundled_template:
        _assemble_source_package(
            bundled_template,
            cdx_paths,
            source_dir,
            anchors=anchors,
            row_height=row_height,
            preview=preview,
        )


def _assemble_source_package(
    template: Path,
    cdx_paths: list[Path],
    source_dir: Path,
    *,
    anchors: list[OleAnchor],
    row_height: float,
    preview: PreviewSpec,
) -> None:
    register_namespaces()
    source_dir.mkdir(parents=True, exist_ok=True)
    extract_ooxml(template, source_dir)

    sheet_path = source_dir / "xl" / "worksheets" / "sheet1.xml"
    # 模板的 shapeId 与锚点行做基准；这张临时表的单元格内容/行高不会被
    # 搬进目标（merge 只取 bin、media、drawing、vml、<oleObjects>），
    # 所以不写标签、共享字符串和行高。
    base_shape_id, from_row, to_row = extract_base_anchor(sheet_path)

    # B: bin + preview images
    filled = fill_ole_bins_and_previews(source_dir, cdx_paths, preview=preview)

    # C–F: rels + drawing/vml + oleObjects + content types
    wire_source_surface(
        source_dir,
        filled=filled,
        anchors=anchors,
        base_shape_id=base_shape_id,
        from_row=from_row,
        to_row=to_row,
        row_height=row_height,
        object_count=len(cdx_paths),
    )
