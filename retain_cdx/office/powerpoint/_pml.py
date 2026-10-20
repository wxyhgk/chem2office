"""PresentationML surgery: 往 slide 上挂 ChemDraw OLE graphicFrame.

PPT 侧比 Excel 简单得多（详见 `docs/powerpoint_ole_research.md`）：

  * 绝对 EMU 坐标，没有行列网格，也就没有字体/列宽那一整类几何问题
  * 没有 VML 层
  * `p:spTree` 里追加即可，没有 CT_Worksheet 那种固定子元素顺序

一个对象 = `p:graphicFrame`（内含 `p:oleObj` + `mc:Fallback` 里的预览图）
+ slide rels 两条（oleObject / image）+ 两个部件（.bin / .emf|.png）。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import xml.etree.ElementTree as ET

from retain_cdx.office._opc import (
    REL_IMAGE,
    REL_OLE,
    ensure_ole_content_types,
    next_rid,
    read_rels,
    rels_path_for,
    unique_part_name,
    write_rels,
)

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_V = "urn:schemas-microsoft-com:vml"

NS = {"p": NS_P, "a": NS_A, "r": NS_R, "mc": NS_MC}


def register_namespaces() -> None:
    for prefix, uri in (
        ("p", NS_P),
        ("a", NS_A),
        ("r", NS_R),
        ("mc", NS_MC),
    ):
        ET.register_namespace(prefix, uri)


def write_slide_xml(tree: ET.ElementTree, slide_path: Path) -> None:
    """Serialize a slide, keeping namespace declarations MCE needs.

    The ChemDraw frame carries ``<mc:Choice Requires="v">`` but no element in
    that subtree uses the ``v`` prefix, so ElementTree drops the ``xmlns:v``
    declaration it inherited — leaving ``Requires`` pointing at an undeclared
    prefix. PowerPoint rejects that and offers to repair the file, so the
    declaration is re-attached to the root element.
    """
    register_namespaces()
    tree.write(slide_path, encoding="UTF-8", xml_declaration=True)
    text = slide_path.read_text(encoding="utf-8")
    start = text.find("<p:sld")
    if start == -1:
        return
    end = text.find(">", start)
    if end == -1:
        return
    root_tag = text[start:end]
    if "xmlns:v=" not in root_tag:
        root_tag += f' xmlns:v="{NS_V}"'
        slide_path.write_text(text[:start] + root_tag + text[end:], encoding="utf-8")


@dataclass(frozen=True)
class OleFrameTemplate:
    """A native PowerPoint+ChemDraw ``p:graphicFrame`` plus its OLE container."""

    frame_xml: str
    container: bytes

    @classmethod
    def from_package(cls, package_dir: Path) -> "OleFrameTemplate":
        slide_path = package_dir / "ppt" / "slides" / "slide1.xml"
        root = ET.parse(slide_path).getroot()
        frame = _find_ole_frame(root)
        if frame is None:
            raise ValueError(f"No ChemDraw OLE graphicFrame in {slide_path}")

        rels = read_rels(rels_path_for(slide_path))
        ole_target = next(
            (target for _, rel_type, target in rels if rel_type == REL_OLE), None
        )
        if ole_target is None:
            raise ValueError("Template slide has no oleObject relationship")
        container = (slide_path.parent / ole_target).resolve().read_bytes()
        register_namespaces()
        return cls(frame_xml=ET.tostring(frame, encoding="unicode"), container=container)


def _find_ole_frame(root: ET.Element) -> ET.Element | None:
    for frame in root.iter(f"{{{NS_P}}}graphicFrame"):
        if frame.find(f".//{{{NS_P}}}oleObj") is not None:
            return frame
    return None


@dataclass(frozen=True)
class FramePlacement:
    """Absolute box on the slide, in EMU."""

    x: int
    y: int
    width: int
    height: int


def append_ole_frame(
    package_dir: Path,
    *,
    slide_path: Path,
    template: OleFrameTemplate,
    container: bytes,
    preview_bytes: bytes,
    preview_ext: str,
    placement: FramePlacement,
    name: str,
) -> None:
    """Clone the template frame onto *slide_path* with its own parts and rels."""
    register_namespaces()
    embeddings_dir = package_dir / "ppt" / "embeddings"
    media_dir = package_dir / "ppt" / "media"
    for directory in (embeddings_dir, media_dir):
        directory.mkdir(parents=True, exist_ok=True)

    bin_name = unique_part_name(embeddings_dir, "oleObject1.bin")
    (embeddings_dir / bin_name).write_bytes(container)
    image_name = unique_part_name(media_dir, f"image1.{preview_ext}")
    (media_dir / image_name).write_bytes(preview_bytes)

    rels_path = rels_path_for(slide_path)
    rels = read_rels(rels_path)
    ole_rid = next_rid(rels)
    rels.append((ole_rid, REL_OLE, f"../embeddings/{bin_name}"))
    image_rid = next_rid(rels)
    rels.append((image_rid, REL_IMAGE, f"../media/{image_name}"))
    write_rels(rels_path, rels)

    tree = ET.parse(slide_path)
    root = tree.getroot()
    tree_element = root.find("p:cSld/p:spTree", NS)
    if tree_element is None:
        raise ValueError(f"Slide has no spTree: {slide_path}")

    frame = ET.fromstring(template.frame_xml)
    _apply_ids(frame, shape_id=_next_shape_id(root), name=name)
    _apply_rels(frame, ole_rid=ole_rid, image_rid=image_rid)
    _apply_placement(frame, placement)
    tree_element.append(frame)
    write_slide_xml(tree, slide_path)

    ensure_ole_content_types(package_dir, preview_ext)


def _next_shape_id(root: ET.Element) -> int:
    ids = [
        int(element.get("id"))
        for element in root.iter(f"{{{NS_P}}}cNvPr")
        if (element.get("id") or "").isdigit()
    ]
    return max(ids, default=1) + 1


def _apply_ids(frame: ET.Element, *, shape_id: int, name: str) -> int:
    """Give every cNvPr in the clone a fresh id. Returns the next free id.

    The template's fallback picture ships as ``id="0"``; cloning it verbatim
    would repeat that id once per structure, and shape ids must be unique
    within the slide's shape tree.
    """
    outer = frame.find("p:nvGraphicFramePr/p:cNvPr", NS)
    next_id = shape_id
    for cnv_pr in frame.iter(f"{{{NS_P}}}cNvPr"):
        cnv_pr.set("id", str(next_id))
        next_id += 1
        if cnv_pr is outer:
            cnv_pr.set("name", name)
            # creationId/modId are per-object identity hints; drop cloned ones.
            for ext_lst in cnv_pr.findall(f"{{{NS_A}}}extLst"):
                cnv_pr.remove(ext_lst)
    nv_pr = frame.find("p:nvGraphicFramePr/p:nvPr", NS)
    if nv_pr is not None:
        for ext_lst in nv_pr.findall(f"{{{NS_P}}}extLst"):
            nv_pr.remove(ext_lst)
    return next_id


def _apply_rels(frame: ET.Element, *, ole_rid: str, image_rid: str) -> None:
    for ole_obj in frame.iter(f"{{{NS_P}}}oleObj"):
        ole_obj.set(f"{{{NS_R}}}id", ole_rid)
    for blip in frame.iter(f"{{{NS_A}}}blip"):
        blip.set(f"{{{NS_R}}}embed", image_rid)


def _apply_placement(frame: ET.Element, placement: FramePlacement) -> None:
    # The frame's own p:xfrm and the fallback picture's a:xfrm must agree.
    for xfrm in list(frame.iter(f"{{{NS_P}}}xfrm")) + list(frame.iter(f"{{{NS_A}}}xfrm")):
        off = xfrm.find(f"{{{NS_A}}}off")
        ext = xfrm.find(f"{{{NS_A}}}ext")
        if off is not None:
            off.set("x", str(placement.x))
            off.set("y", str(placement.y))
        if ext is not None:
            ext.set("cx", str(placement.width))
            ext.set("cy", str(placement.height))
    # imgW/imgH is the object's intrinsic size; keep it consistent with the box.
    for ole_obj in frame.iter(f"{{{NS_P}}}oleObj"):
        ole_obj.set("imgW", str(placement.width))
        ole_obj.set("imgH", str(placement.height))


def slide_size(package_dir: Path) -> tuple[int, int]:
    """Slide width/height in EMU."""
    root = ET.parse(package_dir / "ppt" / "presentation.xml").getroot()
    size = root.find("p:sldSz", NS)
    if size is None:
        return 12192000, 6858000
    return int(size.get("cx")), int(size.get("cy"))


def slide_paths(package_dir: Path) -> list[Path]:
    """Slides in presentation order (by rels from presentation.xml)."""
    presentation = package_dir / "ppt" / "presentation.xml"
    rels = {
        rid: target
        for rid, _type, target in read_rels(rels_path_for(presentation))
    }
    root = ET.parse(presentation).getroot()
    paths: list[Path] = []
    for sld_id in root.findall("p:sldIdLst/p:sldId", NS):
        target = rels.get(sld_id.get(f"{{{NS_R}}}id"))
        if target:
            paths.append((presentation.parent / target).resolve())
    return paths


__all__ = [
    "FramePlacement",
    "OleFrameTemplate",
    "append_ole_frame",
    "slide_paths",
    "slide_size",
]
