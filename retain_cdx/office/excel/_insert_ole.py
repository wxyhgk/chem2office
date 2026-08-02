"""把 OLE 对象直接写进目标 package（不经过 staging 工作簿）。

与 `_picture.py` 同一形态，只是多了 OLE 的三层：容器 `.bin`、VML 形状、
worksheet 上的 `<oleObject>` 声明。零件来自 `ExcelOleTemplate`——克隆
Windows Excel 真实产出的片段，而不是自己拼 XML。

一个结构 = 4 个写入 + 3 个关系：

    xl/embeddings/<prefix>_oleObjectN.bin   容器（注入 CDX）
    xl/media/<prefix>_imageN.png|emf        预览图
    drawing 里一个 xdr:twoCellAnchor        现代绘图层（隐藏占位）
    VML 里一个 v:shape                      Excel 实际渲染的预览
    sheet rels: oleObject + image           worksheet 引用
    vml rels:   image                       VML 引用
"""

from __future__ import annotations

import tempfile
import xml.etree.ElementTree as ET
from contextlib import ExitStack
from dataclasses import dataclass
from pathlib import Path

from retain_cdx.ole import (
    MAX_EMBEDDED_EMF,
    build_enhmetafile_presentation,
    emu_to_himetric,
)
from retain_cdx.ole.embed import write_cdx_to_ole
from retain_cdx.render.preview import (
    preview_extension,
    render_emf_many,
    render_preview,
    resolve_preview_format,
)

from ._grid import ExcelGridMetrics
from ._model import ShellSize
from ._ole_template import NS_V, ExcelOleTemplate
from ._ooxml import (
    REL_DRAWING,
    REL_IMAGE,
    REL_OLE_OBJECT,
    REL_VML,
    attach_ole_objects,
    build_rels,
    ensure_ole_drawing_content_types,
    insert_worksheet_child,
    read_rels,
    register_namespaces,
    write_sheet_xml,
)
from ._ooxml.xml_io import NS, NS_MAIN, NS_O, NS_R, NS_XDR
from ._parts.util import next_rid, unique_part_name, write_xml
from .models import LayoutSpec, PreviewSpec

_RASTER_SUPERSAMPLE = 4
_BASE_SHAPE_ID = 1024  # Excel 自己的 OLE shapeId 从 1025 起


@dataclass
class _Slot:
    """一个结构落地后拿到的名字与关系 id。"""

    bin_name: str
    image_name: str
    ole_rid: str
    image_rid: str
    vml_rid: str
    shape_id: int


def insert_ole_objects(
    package_dir: Path,
    cdx_paths: list[Path],
    *,
    sheet_path: Path,
    prefix: str,
    template: ExcelOleTemplate,
    layout: LayoutSpec,
    preview: PreviewSpec,
) -> str:
    """把结构逐行写进 *sheet_path*；返回目标使用的 drawing 部件名。"""
    if not cdx_paths:
        raise ValueError("no structures to insert")

    register_namespaces()
    for relative in ("xl/embeddings", "xl/media", "xl/drawings", "xl/drawings/_rels"):
        (package_dir / relative).mkdir(parents=True, exist_ok=True)

    drawing_path, drawing_rid, vml_path, vml_rid = _open_surfaces(
        package_dir, sheet_path=sheet_path, prefix=prefix, template=template
    )
    slots = _write_parts(
        package_dir,
        cdx_paths,
        sheet_path=sheet_path,
        prefix=prefix,
        template=template,
        layout=layout,
        preview=preview,
        drawing_path=drawing_path,
        vml_path=vml_path,
    )

    metrics = ExcelGridMetrics.from_worksheet(sheet_path)
    anchors = _anchors(metrics, layout, len(cdx_paths))
    _append_drawing(drawing_path, template, slots, anchors)
    _append_vml(vml_path, template, slots, anchors, layout, metrics)
    _attach_declarations(sheet_path, template, slots, anchors, drawing_rid, vml_rid)

    ensure_ole_drawing_content_types(package_dir, drawing_path.name)
    return drawing_path.name


# --- 目标侧的 drawing / VML 槽位 -------------------------------------------


def _open_surfaces(
    package_dir: Path, *, sheet_path: Path, prefix: str, template: ExcelOleTemplate
) -> tuple[Path, str, Path, str]:
    """复用 sheet 已有的 drawing/legacyDrawing，没有就新建。"""
    drawings = package_dir / "xl" / "drawings"
    rels_path = sheet_path.parent / "_rels" / f"{sheet_path.name}.rels"
    rels_path.parent.mkdir(parents=True, exist_ok=True)
    rels = read_rels(rels_path)
    root = ET.parse(sheet_path).getroot()

    drawing_path, drawing_rid = _slot(
        root, rels, "s:drawing", REL_DRAWING, drawings, f"{prefix}_drawing.xml",
        sheet_path, lambda path: _write_shell(template.new_drawing_shell(), path),
    )
    vml_path, vml_rid = _slot(
        root, rels, "s:legacyDrawing", REL_VML, drawings, f"{prefix}_vmlDrawing.vml",
        sheet_path, lambda path: _write_shell(template.new_vml_shell(), path),
    )
    build_rels(rels_path, rels)
    return drawing_path, drawing_rid, vml_path, vml_rid


def _slot(root, rels, tag, rel_type, directory, desired, sheet_path, create):
    """找到已有部件，或建一个并登记关系。"""
    element = root.find(tag, NS)
    if element is not None:
        rid = element.get(f"{{{NS_R}}}id")
        for existing_rid, _type, target in rels:
            if existing_rid == rid:
                return (sheet_path.parent / target).resolve(), rid

    name = unique_part_name(directory, desired)
    path = directory / name
    create(path)
    rid = next_rid(rels)
    rels.append((rid, rel_type, f"../drawings/{name}"))
    return path, rid


def _write_shell(root: ET.Element, path: Path) -> None:
    """把模板外壳写成一个空部件。"""
    write_xml(ET.ElementTree(root), path)


# --- 零件落盘 ---------------------------------------------------------------


def _write_parts(
    package_dir: Path,
    cdx_paths: list[Path],
    *,
    sheet_path: Path,
    prefix: str,
    template: ExcelOleTemplate,
    layout: LayoutSpec,
    preview: PreviewSpec,
    drawing_path: Path,
    vml_path: Path,
) -> list[_Slot]:
    embeddings = package_dir / "xl" / "embeddings"
    media = package_dir / "xl" / "media"
    extension = preview_extension(preview.format)
    resolved = resolve_preview_format(preview.format)

    sheet_rels_path = sheet_path.parent / "_rels" / f"{sheet_path.name}.rels"
    sheet_rels = read_rels(sheet_rels_path)
    vml_rels_path = vml_path.parent / "_rels" / f"{vml_path.name}.rels"
    vml_rels = read_rels(vml_rels_path)

    shape_id = _next_shape_id(drawing_path, vml_path)

    if preview.presentation_cache and resolved != "emf":
        raise ValueError(
            "presentation_cache needs a vector preview; pass preview_format='emf'"
        )

    slots: list[_Slot] = []
    emf_jobs: list[tuple[Path, Path]] = []
    emf_aspects: list[float | None] = []
    cache_jobs: list[tuple[Path, Path, Path, int]] = []
    for index, cdx_path in enumerate(cdx_paths, start=1):
        width, height = _preview_size(layout, preview, extension, index - 1)
        target_aspect = _target_aspect(layout, preview, index - 1)
        bin_name = unique_part_name(embeddings, f"{prefix}_oleObject{index}.bin")
        write_cdx_to_ole(
            embeddings / bin_name, cdx_path, template_ole=template.container
        )

        image_name = unique_part_name(media, f"{prefix}_image{index}.{extension}")
        image_path = media / image_name
        if resolved == "emf":
            emf_jobs.append((cdx_path, image_path))
            emf_aspects.append(target_aspect)
            if preview.presentation_cache:
                cache_jobs.append(
                    (cdx_path, embeddings / bin_name, image_path, index - 1)
                )
        else:
            render_preview(
                cdx_path,
                image_path,
                format=preview.format,
                width=width,
                height=height,
                thick_lines=preview.thick_lines,
                no_margin=preview.no_margin,
                mono=preview.mono,
                backend=preview.backend,
                style=preview.style,
                crop=True,
                target_aspect=target_aspect,
                fill_ratio=preview.fill_ratio,
            )

        ole_rid = next_rid(sheet_rels)
        sheet_rels.append((ole_rid, REL_OLE_OBJECT, f"../embeddings/{bin_name}"))
        image_rid = next_rid(sheet_rels)
        sheet_rels.append((image_rid, REL_IMAGE, f"../media/{image_name}"))
        vml_image_rid = next_rid(vml_rels)
        vml_rels.append((vml_image_rid, REL_IMAGE, f"../media/{image_name}"))

        slots.append(
            _Slot(bin_name, image_name, ole_rid, image_rid, vml_image_rid, shape_id)
        )
        shape_id += 1

    if emf_jobs:
        with ExitStack() as stack:
            scratch = (
                Path(stack.enter_context(tempfile.TemporaryDirectory(prefix="retain_cdx_wmf_")))
                if cache_jobs
                else None
            )
            render_emf_many(
                emf_jobs,
                width=preview.width,
                height=preview.height,
                thick_lines=preview.thick_lines,
                no_margin=preview.no_margin,
                mono=preview.mono,
                backend=preview.backend,
                style=preview.style,
                # 每个分子的框比例不同，letterbox 也要逐个来。
                target_aspects=emf_aspects,
                fill_ratio=preview.fill_ratio,
                # 伴生 WMF 只是缓存的原料，绝不能落在 xl/media 里——那会变成
                # 一个没有内容类型声明的部件，Excel 打开即报修复。
                companion_wmf_dir=scratch,
            )
            for cdx_path, bin_path, image_path, index in cache_jobs:
                _attach_presentation(
                    cdx_path, bin_path, image_path, scratch, layout.box(index), template
                )

    build_rels(sheet_rels_path, sheet_rels)
    vml_rels_path.parent.mkdir(parents=True, exist_ok=True)
    build_rels(vml_rels_path, vml_rels)
    return slots


def _attach_presentation(
    cdx_path: Path,
    bin_path: Path,
    image_path: Path,
    scratch: Path,
    box: tuple[int, int],
    template: ExcelOleTemplate,
) -> None:
    """把渲染缓存（`\x02OlePres000`）写进容器。

    没有它，把结构从 Excel 复制到 PowerPoint 时，目标程序拿不到现成的图，
    只好唤起 OLE 服务器现场渲染——ChemDraw 会被自动拉起来。

    EMF 必须整份放进一条 WMFC 注释记录，超过 `MAX_EMBEDDED_EMF` 的就跳过：
    复制时多启动一次 ChemDraw，好过产出会让 PowerPoint 崩的文件。
    """
    from retain_cdx.office.excel._utils import EMU_PER_PIXEL

    emf = image_path.read_bytes()
    wmf_path = scratch / f"{image_path.stem}.wmf"
    if len(emf) > MAX_EMBEDDED_EMF or not wmf_path.exists():
        return
    width_px, height_px = box
    presentation = build_enhmetafile_presentation(
        emf,
        wmf_path.read_bytes(),
        width_himetric=emu_to_himetric(width_px * EMU_PER_PIXEL),
        height_himetric=emu_to_himetric(height_px * EMU_PER_PIXEL),
    )
    write_cdx_to_ole(
        bin_path, cdx_path, template_ole=template.container, presentation=presentation
    )


def _preview_size(
    layout: LayoutSpec, preview: PreviewSpec, extension: str, index: int = 0
) -> tuple[int | None, int | None]:
    if preview.width is not None or preview.height is not None:
        return preview.width, preview.height
    if extension != "png":
        return None, None
    # 位图默认按显示框超采样：小分子 flexi 渲染只有 ~25px，Excel 放大后发糊。
    box_w, box_h = layout.box(index)
    return box_w * _RASTER_SUPERSAMPLE, box_h * _RASTER_SUPERSAMPLE


def _target_aspect(
    layout: LayoutSpec, preview: PreviewSpec, index: int = 0
) -> float | None:
    if preview.target_aspect and preview.target_aspect > 0:
        return preview.target_aspect
    box_w, box_h = layout.box(index)
    return (box_w / box_h) if box_h else None


def _next_shape_id(drawing_path: Path, vml_path: Path) -> int:
    largest = _BASE_SHAPE_ID
    if drawing_path.exists():
        for element in ET.parse(drawing_path).getroot().iter(f"{{{NS_XDR}}}cNvPr"):
            if (element.get("id") or "").isdigit():
                largest = max(largest, int(element.get("id")))
    if vml_path.exists():
        text = vml_path.read_text(encoding="utf-8")
        for match in ET.fromstring(text).iter(f"{{{NS_V}}}shape"):
            spid = (match.get("id") or "").removeprefix("_x0000_s")
            if spid.isdigit():
                largest = max(largest, int(spid))
    return largest + 1


# --- 表面 XML ---------------------------------------------------------------


def _anchors(metrics: ExcelGridMetrics, layout: LayoutSpec, count: int) -> list:
    return [
        metrics.anchor_for_box(
            start_col=layout.col,
            start_row=layout.row + index,
            offset_x_px=layout.offset_x_px,
            offset_y_px=layout.offset_y_px,
            size=ShellSize(*layout.box(index)),
        )
        for index in range(count)
    ]


def _set_corner(parent: ET.Element, tag: str, corner, namespace: str) -> None:
    node = parent.find(f"{{{namespace}}}{tag}")
    if node is None:
        return
    from retain_cdx.office.excel._utils import EMU_PER_PIXEL

    for child_tag, value in (
        ("col", corner.col),
        ("colOff", corner.col_off_px * EMU_PER_PIXEL),
        ("row", corner.row),
        ("rowOff", corner.row_off_px * EMU_PER_PIXEL),
    ):
        child = node.find(f"{{{NS_XDR}}}{child_tag}")
        if child is not None:
            child.text = str(value)


def _append_drawing(
    drawing_path: Path, template: ExcelOleTemplate, slots: list[_Slot], anchors: list
) -> None:
    tree = ET.parse(drawing_path)
    root = tree.getroot()
    # 模板把锚点放在 mc:Choice Requires="a14" 里（锚点用了 a14 扩展）；
    # 直接挂到根上会丢掉这层包装。
    container = root.find("mc:AlternateContent/mc:Choice", NS)
    if container is None:
        container = root
    for slot, anchor in zip(slots, anchors):
        element = template.new_drawing_anchor()
        _set_corner(element, "from", anchor.from_, NS_XDR)
        _set_corner(element, "to", anchor.to, NS_XDR)
        for c_nv_pr in element.iter(f"{{{NS_XDR}}}cNvPr"):
            c_nv_pr.set("id", str(slot.shape_id))
            c_nv_pr.set("name", f"Object {slot.shape_id - _BASE_SHAPE_ID}")
        for ext in element.iter():
            if ext.get("spid"):
                ext.set("spid", f"_x0000_s{slot.shape_id}")
        container.append(element)
    write_xml(tree, drawing_path)


def _append_vml(
    vml_path: Path,
    template: ExcelOleTemplate,
    slots: list[_Slot],
    anchors: list,
    layout: LayoutSpec,
    metrics: ExcelGridMetrics,
) -> None:
    from retain_cdx.office.excel._utils import POINTS_PER_PIXEL

    root = ET.fromstring(vml_path.read_text(encoding="utf-8"))
    for index, (slot, anchor) in enumerate(zip(slots, anchors)):
        box_w, box_h = layout.box(index)
        shape = template.new_vml_shape()
        shape.set("id", f"_x0000_s{slot.shape_id}")
        # style 的绝对位置按目标表真实网格算。Excel 用 x:Anchor 定位，
        # 但按 style 摆放的第三方消费者需要它是对的。
        left_pt = metrics.absolute_x_px(anchor.from_) * POINTS_PER_PIXEL
        top_pt = metrics.absolute_y_px(anchor.from_) * POINTS_PER_PIXEL
        shape.set(
            "style",
            f"position:absolute;margin-left:{left_pt:g}pt;margin-top:{top_pt:g}pt;"
            f"width:{box_w * POINTS_PER_PIXEL:g}pt;"
            f"height:{box_h * POINTS_PER_PIXEL:g}pt;z-index:{slot.shape_id}",
        )
        for imagedata in shape.iter(f"{{{NS_V}}}imagedata"):
            imagedata.set(f"{{{NS_O}}}relid", slot.vml_rid)
        client = shape.find("{urn:schemas-microsoft-com:office:excel}ClientData")
        if client is not None:
            node = client.find("{urn:schemas-microsoft-com:office:excel}Anchor")
            if node is not None:
                node.text = anchor.to_vml_anchor()
        root.append(shape)
    write_xml(ET.ElementTree(root), vml_path)


def _attach_declarations(
    sheet_path: Path,
    template: ExcelOleTemplate,
    slots: list[_Slot],
    anchors: list,
    drawing_rid: str,
    vml_rid: str,
) -> None:
    container = ET.Element(f"{{{NS_MAIN}}}oleObjects")
    for slot, anchor in zip(slots, anchors):
        element = template.new_ole_object()
        for ole in element.iter(f"{{{NS_MAIN}}}oleObject"):
            ole.set("shapeId", str(slot.shape_id))
            ole.set(f"{{{NS_R}}}id", slot.ole_rid)
            object_pr = ole.find("s:objectPr", NS)
            if object_pr is None:
                continue
            object_pr.set(f"{{{NS_R}}}id", slot.image_rid)
            object_pr.set("autoPict", "0")
            object_pr.set("defaultSize", "0")
            box = object_pr.find("s:anchor", NS)
            if box is not None:
                _set_corner(box, "from", anchor.from_, NS_MAIN)
                _set_corner(box, "to", anchor.to, NS_MAIN)
        container.append(element)

    attach_ole_objects(
        sheet_path, drawing_rid=drawing_rid, vml_rid=vml_rid, ole_objects=container
    )


__all__ = ["insert_ole_objects"]
