"""把结构渲染成普通 Excel 图片（不放 OLE 对象）.

图片模式 · 与 `_source` + `_parts` 的 OLE 路径并列的另一种落地方式：

  B  渲染预览写入 xl/media
  C  drawing 的 image relationships
  D  drawing.xml 里按目标网格写 `xdr:pic` 锚点

一个 OLE 对象要带 CFBF 容器、VML 形状和激活机制，几千个会把 Excel 压垮；
一张图片只有 media + 一个 drawing 锚点，所以单表可以流畅装下几千个结构。
代价：图片不能双击进 ChemDraw 编辑（需要时用 OLE 路径单独生成那一行）。
"""

from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

from retain_cdx.render.preview import preview_extension, render_emf_many, render_preview

from retain_cdx.office.excel._ooxml import (
    NS_A,
    NS_R,
    NS_XDR,
    REL_DRAWING,
    REL_IMAGE,
    build_rels,
    ensure_content_type_defaults,
    insert_worksheet_child,
    read_rels,
    register_namespaces,
    write_sheet_xml,
)
from retain_cdx.office.excel._ooxml.xml_io import NS, NS_MAIN
from retain_cdx.office.excel._parts.util import next_rid, unique_part_name
from retain_cdx.office.excel._shell import ExcelGridMetrics
from retain_cdx.office.excel._shell.model import ShellSize
from retain_cdx.office.excel._utils import EMU_PER_PIXEL
from retain_cdx.office.excel.models import LayoutSpec, PreviewSpec

_CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
_MEDIA_CONTENT_TYPES = {"png": "image/png", "emf": "image/x-emf"}
# Pictures are the final visual (no OLE server to redraw them), so rasters are
# supersampled against the display box the same way the OLE previews are.
_RASTER_SUPERSAMPLE = 4


def insert_structure_pictures(
    package_dir: Path,
    cdx_paths: list[Path],
    *,
    sheet_path: Path,
    prefix: str,
    layout: LayoutSpec,
    preview: PreviewSpec,
) -> None:
    """Render *cdx_paths* into *sheet_path* as pictures, one per row."""
    if not cdx_paths:
        return

    media_dir = package_dir / "xl" / "media"
    drawings_dir = package_dir / "xl" / "drawings"
    for directory in (media_dir, drawings_dir, drawings_dir / "_rels"):
        directory.mkdir(parents=True, exist_ok=True)

    ext = preview_extension(preview.format)
    image_names = _render_previews(
        cdx_paths, media_dir, prefix=prefix, ext=ext, layout=layout, preview=preview
    )

    metrics = ExcelGridMetrics.from_worksheet(sheet_path)
    size = ShellSize(layout.width_px, layout.height_px)
    anchors = [
        metrics.anchor_for_box(
            start_col=layout.col,
            start_row=layout.row + index,
            offset_x_px=layout.offset_x_px,
            offset_y_px=layout.offset_y_px,
            size=size,
        )
        for index in range(len(cdx_paths))
    ]

    drawing_path, drawing_rid, created = _open_drawing(
        package_dir, sheet_path=sheet_path, prefix=prefix
    )
    _append_pictures(drawing_path, anchors=anchors, image_names=image_names, size=size)
    if created:
        _attach_drawing_element(sheet_path, drawing_rid)

    ensure_content_type_defaults(
        package_dir / "[Content_Types].xml",
        {ext: _MEDIA_CONTENT_TYPES.get(ext, "image/png")},
    )
    _ensure_drawing_override(package_dir, drawing_path.name)


def _render_previews(
    cdx_paths: list[Path],
    media_dir: Path,
    *,
    prefix: str,
    ext: str,
    layout: LayoutSpec,
    preview: PreviewSpec,
) -> list[str]:
    """Write one preview per structure; returns the media part names."""
    width, height = preview.width, preview.height
    if width is None and height is None and ext == "png":
        width = layout.width_px * _RASTER_SUPERSAMPLE
        height = layout.height_px * _RASTER_SUPERSAMPLE
    # Letterbox to the display box so the molecule is never stretched.
    target_aspect = preview.target_aspect
    if not target_aspect or target_aspect <= 0:
        target_aspect = layout.width_px / layout.height_px if layout.height_px else None

    names: list[str] = []
    emf_jobs: list[tuple[Path, Path]] = []
    for index, cdx_path in enumerate(cdx_paths, start=1):
        name = unique_part_name(media_dir, f"{prefix}_picture{index}.{ext}")
        names.append(name)
        image_path = media_dir / name
        if ext == "emf":
            emf_jobs.append((cdx_path, image_path))
            continue
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

    if emf_jobs:
        render_emf_many(
            emf_jobs,
            width=width,
            height=height,
            thick_lines=preview.thick_lines,
            no_margin=preview.no_margin,
            mono=preview.mono,
            backend=preview.backend,
            style=preview.style,
            target_aspect=target_aspect,
            fill_ratio=preview.fill_ratio,
        )
    return names


def _open_drawing(
    package_dir: Path, *, sheet_path: Path, prefix: str
) -> tuple[Path, str, bool]:
    """Return (drawing_path, sheet_rid, created). Reuses the sheet's drawing."""
    drawings_dir = package_dir / "xl" / "drawings"
    sheet_rels_path = sheet_path.parent / "_rels" / f"{sheet_path.name}.rels"
    sheet_rels = read_rels(sheet_rels_path)

    existing_rid = _existing_drawing_rid(sheet_path)
    if existing_rid is not None:
        for rid, _rel_type, target in sheet_rels:
            if rid == existing_rid:
                return (sheet_path.parent / target).resolve(), rid, False

    name = unique_part_name(drawings_dir, f"{prefix}_pictures.xml")
    drawing_path = drawings_dir / name
    register_namespaces()
    root = ET.Element(f"{{{NS_XDR}}}wsDr")
    ET.ElementTree(root).write(drawing_path, encoding="UTF-8", xml_declaration=True)

    rid = next_rid(sheet_rels)
    sheet_rels.append((rid, REL_DRAWING, f"../drawings/{name}"))
    sheet_rels_path.parent.mkdir(parents=True, exist_ok=True)
    build_rels(sheet_rels_path, sheet_rels)
    return drawing_path, rid, True


def _existing_drawing_rid(sheet_path: Path) -> str | None:
    root = ET.parse(sheet_path).getroot()
    drawing = root.find("s:drawing", NS)
    return None if drawing is None else drawing.get(f"{{{NS_R}}}id")


def _append_pictures(
    drawing_path: Path,
    *,
    anchors: list,
    image_names: list[str],
    size: ShellSize,
) -> None:
    register_namespaces()
    tree = ET.parse(drawing_path)
    root = tree.getroot()

    rels_path = drawing_path.parent / "_rels" / f"{drawing_path.name}.rels"
    rels = read_rels(rels_path)
    shape_id = _max_shape_id(root) + 1

    for anchor, image_name in zip(anchors, image_names):
        rid = next_rid(rels)
        rels.append((rid, REL_IMAGE, f"../media/{image_name}"))
        root.append(
            _picture_anchor(
                anchor, rid=rid, shape_id=shape_id, name=Path(image_name).stem, size=size
            )
        )
        shape_id += 1

    rels_path.parent.mkdir(parents=True, exist_ok=True)
    build_rels(rels_path, rels)
    tree.write(drawing_path, encoding="UTF-8", xml_declaration=True)


def _max_shape_id(root: ET.Element) -> int:
    ids = [
        int(element.get("id"))
        for element in root.iter(f"{{{NS_XDR}}}cNvPr")
        if (element.get("id") or "").isdigit()
    ]
    return max(ids, default=0)


def _picture_anchor(
    anchor, *, rid: str, shape_id: int, name: str, size: ShellSize
) -> ET.Element:
    # editAs="oneCell": move with the cell, keep the box size (same choice the
    # OLE path inherits from Excel-authored templates).
    element = ET.Element(f"{{{NS_XDR}}}twoCellAnchor", {"editAs": "oneCell"})
    _corner(element, "from", anchor.from_)
    _corner(element, "to", anchor.to)

    pic = ET.SubElement(element, f"{{{NS_XDR}}}pic")
    nv_pic_pr = ET.SubElement(pic, f"{{{NS_XDR}}}nvPicPr")
    ET.SubElement(
        nv_pic_pr, f"{{{NS_XDR}}}cNvPr", {"id": str(shape_id), "name": name}
    )
    cnv_pic_pr = ET.SubElement(nv_pic_pr, f"{{{NS_XDR}}}cNvPicPr")
    ET.SubElement(cnv_pic_pr, f"{{{NS_A}}}picLocks", {"noChangeAspect": "1"})

    blip_fill = ET.SubElement(pic, f"{{{NS_XDR}}}blipFill")
    ET.SubElement(blip_fill, f"{{{NS_A}}}blip", {f"{{{NS_R}}}embed": rid})
    stretch = ET.SubElement(blip_fill, f"{{{NS_A}}}stretch")
    ET.SubElement(stretch, f"{{{NS_A}}}fillRect")

    sp_pr = ET.SubElement(pic, f"{{{NS_XDR}}}spPr")
    xfrm = ET.SubElement(sp_pr, f"{{{NS_A}}}xfrm")
    ET.SubElement(xfrm, f"{{{NS_A}}}off", {"x": "0", "y": "0"})
    ET.SubElement(
        xfrm,
        f"{{{NS_A}}}ext",
        {
            "cx": str(size.width_px * EMU_PER_PIXEL),
            "cy": str(size.height_px * EMU_PER_PIXEL),
        },
    )
    geom = ET.SubElement(sp_pr, f"{{{NS_A}}}prstGeom", {"prst": "rect"})
    ET.SubElement(geom, f"{{{NS_A}}}avLst")

    ET.SubElement(element, f"{{{NS_XDR}}}clientData")
    return element


def _corner(parent: ET.Element, tag: str, corner) -> None:
    node = ET.SubElement(parent, f"{{{NS_XDR}}}{tag}")
    ET.SubElement(node, f"{{{NS_XDR}}}col").text = str(corner.col)
    ET.SubElement(node, f"{{{NS_XDR}}}colOff").text = str(
        corner.col_off_px * EMU_PER_PIXEL
    )
    ET.SubElement(node, f"{{{NS_XDR}}}row").text = str(corner.row)
    ET.SubElement(node, f"{{{NS_XDR}}}rowOff").text = str(
        corner.row_off_px * EMU_PER_PIXEL
    )


def _attach_drawing_element(sheet_path: Path, drawing_rid: str) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    if root.find("s:drawing", NS) is not None:
        return
    element = ET.Element(f"{{{NS_MAIN}}}drawing", {f"{{{NS_R}}}id": drawing_rid})
    insert_worksheet_child(root, element)
    write_sheet_xml(tree, sheet_path)


def _ensure_drawing_override(package_dir: Path, drawing_name: str) -> None:
    ct_path = package_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    tree = ET.parse(ct_path)
    root = tree.getroot()
    part_name = f"/xl/drawings/{drawing_name}"
    existing = {
        element.get("PartName") for element in root.findall(f"{{{ct_ns}}}Override")
    }
    if part_name in existing:
        return
    ET.SubElement(
        root, f"{{{ct_ns}}}Override", PartName=part_name, ContentType=_CT_DRAWING
    )
    xml = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    xml = xml.replace(b"ns0:", b"").replace(b"xmlns:ns0", b"xmlns")
    ct_path.write_bytes(xml)


__all__ = ["insert_structure_pictures"]
