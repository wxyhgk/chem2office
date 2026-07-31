"""装进 Excel（当前实现）：把 source 工作簿里的 OLE 零件合并进目标 package.

流程（与用户心智「bin+图 → 放进 Excel」对应，只是零件先在 source 里装好）：

  1. 准备目标 package 目录与已有 drawing/VML 槽位
  2. 落盘 .bin + 预览图，并写 sheet/vml relationships
  3. 合并 drawing / VML（钉框、shapeId）
  4. 挂 worksheet <oleObjects>，可选复制 layout
"""

from __future__ import annotations

import copy
from pathlib import Path
import xml.etree.ElementTree as ET

from retain_cdx.office.excel._ooxml import (
    NS_MAIN,
    attach_ole_objects,
    copy_sheet_layout,
)

from .context import OlePartMergeContext
from .ole_objects import normalize_object_pr, rewrite_ole_shape_ids
from .surface import merge_drawing_and_vml


def merge_source_workbook(
    target_dir: Path,
    source_dir: Path,
    *,
    sheet_path: Path,
    sheet_idx: int,
    prefix: str,
    preserve_object_pr: bool = True,
    copy_layout: bool = False,
) -> str:
    """Merge OLE/drawing parts from a source OLE workbook into target sheet.

    Returns the drawing part name used by the target sheet.
    """
    context = OlePartMergeContext(
        source_dir=source_dir,
        target_dir=target_dir,
        prefix=prefix,
        sheet_rels_path=target_dir
        / "xl"
        / "worksheets"
        / "_rels"
        / f"sheet{sheet_idx}.xml.rels",
    )

    # B: dirs + C: discover existing drawing/vml rIds
    context.ensure_part_dirs()
    context.load_sheet_rels()
    context.ensure_drawing_parts()

    # B+C: copy bin/media and rewire sheet + VML relationships
    rid_map = context.merge_sheet_relationships()
    vml_rid_map = context.merge_vml_relationships()

    # D+E surface: drawing/VML geometry and shape ids
    shape_id_map = merge_drawing_and_vml(context, vml_rid_map)

    # E: oleObjects declaration on the worksheet
    source_sheet_root = ET.parse(source_dir / "xl" / "worksheets" / "sheet1.xml").getroot()
    source_ole_objects = copy.deepcopy(
        source_sheet_root.find("s:oleObjects", {"s": NS_MAIN})
    )
    if source_ole_objects is None:
        raise ValueError("Source OLE workbook has no oleObjects")
    if not preserve_object_pr:
        normalize_object_pr(source_ole_objects)
    rewrite_ole_shape_ids(source_ole_objects, shape_id_map)

    if copy_layout:
        copy_sheet_layout(source_dir / "xl" / "worksheets" / "sheet1.xml", sheet_path)

    attach_ole_objects(
        sheet_path,
        drawing_rid=context.drawing_rid,
        vml_rid=context.vml_rid,
        ole_objects=source_ole_objects,
        rid_map=rid_map,
    )
    return context.drawing_name
