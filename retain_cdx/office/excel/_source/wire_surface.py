"""在 source package 里接线 + 钉框 + 挂 OLE（小表内完成「装进」的 XML 侧）.

对应子步骤：
  C. sheet/vml relationships
  D. drawing.xml / vmlDrawing.vml 锚点
  E. worksheet <oleObjects>
  F. [Content_Types] 媒体类型
"""

from __future__ import annotations

from pathlib import Path

from retain_cdx.office.excel._ooxml import (
    build_sheet_relationships,
    build_vml_relationships,
    ensure_content_type_defaults,
)

from .anchors import OleAnchor
from .fill_parts import FilledParts
from .template import update_drawing, update_ole_objects, update_vml


def wire_source_surface(
    package_dir: Path,
    *,
    filled: FilledParts,
    anchors: list[OleAnchor],
    base_shape_id: int,
    from_row: int,
    to_row: int,
    row_height: float,
    object_count: int,
) -> None:
    """Wire rels and update drawing/VML/oleObjects for already-filled parts."""
    sheet_path = package_dir / "xl" / "worksheets" / "sheet1.xml"
    drawing_path = package_dir / "xl" / "drawings" / "drawing1.xml"
    vml_path = package_dir / "xl" / "drawings" / "vmlDrawing1.vml"
    sheet_rels_path = package_dir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels"
    vml_rels_path = package_dir / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels"
    content_types_path = package_dir / "[Content_Types].xml"

    # C: relationships
    rel_pairs = build_sheet_relationships(
        sheet_rels_path, filled.ole_targets, filled.image_targets
    )
    vml_rids = build_vml_relationships(vml_rels_path, filled.image_targets)

    # E + D: declare OLE and pin geometry
    update_ole_objects(
        sheet_path,
        base_shape_id,
        from_row,
        to_row,
        rel_pairs,
        anchors=anchors,
    )
    update_drawing(
        drawing_path,
        base_shape_id,
        from_row,
        to_row,
        object_count,
        anchors=anchors,
    )
    update_vml(
        vml_path,
        base_shape_id,
        from_row,
        to_row,
        vml_rids,
        anchors=anchors,
        row_height=row_height,
    )

    # F: content types for preview media
    media_defaults = (
        {"emf": "image/x-emf"} if filled.media_ext == "emf" else {"png": "image/png"}
    )
    ensure_content_type_defaults(content_types_path, media_defaults)
