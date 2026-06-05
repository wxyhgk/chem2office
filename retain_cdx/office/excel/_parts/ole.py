from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import xml.etree.ElementTree as ET

from retain_cdx.office.excel._ooxml import (
    NS_MAIN,
    attach_ole_objects,
    build_rels,
    copy_sheet_layout,
)


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


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

    Returns the drawing part name added to ``xl/drawings``.
    """
    drawing_name = f"{prefix}_drawing.xml"
    vml_name = f"{prefix}_vmlDrawing.vml"
    context = _OlePartMergeContext(
        source_dir=source_dir,
        target_dir=target_dir,
        prefix=prefix,
        drawing_name=drawing_name,
        vml_name=vml_name,
    )

    context.ensure_part_dirs()
    context.copy_static_drawing_parts()
    rels = context.build_sheet_rels()
    vml_rels = context.build_vml_rels()

    sheet_rels_path = target_dir / "xl" / "worksheets" / "_rels" / f"sheet{sheet_idx}.xml.rels"
    sheet_rels_path.parent.mkdir(parents=True, exist_ok=True)
    build_rels(sheet_rels_path, rels)
    build_rels(
        target_dir / "xl" / "drawings" / "_rels" / f"{vml_name}.rels",
        vml_rels,
    )

    source_sheet_root = ET.parse(source_dir / "xl" / "worksheets" / "sheet1.xml").getroot()
    source_ole_objects = source_sheet_root.find("s:oleObjects", {"s": NS_MAIN})
    if source_ole_objects is None:
        raise ValueError("Source OLE workbook has no oleObjects")
    if not preserve_object_pr:
        _normalize_object_pr(source_ole_objects)

    if copy_layout:
        copy_sheet_layout(source_dir / "xl" / "worksheets" / "sheet1.xml", sheet_path)

    attach_ole_objects(
        sheet_path,
        drawing_rid="rId1",
        vml_rid="rId2",
        ole_objects=source_ole_objects,
    )
    return drawing_name


@dataclass
class _OlePartMergeContext:
    source_dir: Path
    target_dir: Path
    prefix: str
    drawing_name: str
    vml_name: str
    copied_media: set[str] = field(default_factory=set)

    def ensure_part_dirs(self) -> None:
        for rel in ["xl/drawings", "xl/drawings/_rels", "xl/embeddings", "xl/media"]:
            (self.target_dir / rel).mkdir(parents=True, exist_ok=True)

    def copy_static_drawing_parts(self) -> None:
        (self.target_dir / "xl" / "drawings" / self.drawing_name).write_bytes(
            (self.source_dir / "xl" / "drawings" / "drawing1.xml").read_bytes()
        )
        (self.target_dir / "xl" / "drawings" / self.vml_name).write_bytes(
            (self.source_dir / "xl" / "drawings" / "vmlDrawing1.vml").read_bytes()
        )

    def build_sheet_rels(self) -> list[tuple[str, str, str]]:
        rels = [
            (
                "rId1",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                f"../drawings/{self.drawing_name}",
            ),
            (
                "rId2",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
                f"../drawings/{self.vml_name}",
            ),
        ]
        source_rels = ET.parse(
            self.source_dir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels"
        ).getroot()
        for rel in source_rels.findall(f"{{{REL_NS}}}Relationship"):
            rid = rel.get("Id", "")
            rtype = rel.get("Type", "")
            target = rel.get("Target", "")
            if rtype.endswith("/oleObject"):
                rels.append(self.copy_ole_relationship(rid, rtype, target))
            elif rtype.endswith("/image"):
                rels.append(self.copy_media_relationship(rid, rtype, target))
        return rels

    def build_vml_rels(self) -> list[tuple[str, str, str]]:
        vml_rels: list[tuple[str, str, str]] = []
        source_vml_rels = ET.parse(
            self.source_dir / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels"
        ).getroot()
        for rel in source_vml_rels.findall(f"{{{REL_NS}}}Relationship"):
            vml_rels.append(
                self.copy_media_relationship(
                    rel.get("Id", ""),
                    rel.get("Type", ""),
                    rel.get("Target", ""),
                )
            )
        return vml_rels

    def copy_ole_relationship(
        self,
        rid: str,
        rtype: str,
        target: str,
    ) -> tuple[str, str, str]:
        name = Path(target).name
        new_name = f"{self.prefix}_{name}"
        (self.target_dir / "xl" / "embeddings" / new_name).write_bytes(
            (self.source_dir / "xl" / "embeddings" / name).read_bytes()
        )
        return rid, rtype, f"../embeddings/{new_name}"

    def copy_media_relationship(
        self,
        rid: str,
        rtype: str,
        target: str,
    ) -> tuple[str, str, str]:
        name = Path(target).name
        new_name = f"{self.prefix}_{name}"
        self.copy_media(name, new_name)
        return rid, rtype, f"../media/{new_name}"

    def copy_media(self, source_name: str, target_name: str) -> None:
        if not source_name or target_name in self.copied_media:
            return
        (self.target_dir / "xl" / "media" / target_name).write_bytes(
            (self.source_dir / "xl" / "media" / source_name).read_bytes()
        )
        self.copied_media.add(target_name)


def _normalize_object_pr(ole_objects: ET.Element) -> None:
    for object_pr in ole_objects.findall(".//s:objectPr", {"s": NS_MAIN}):
        object_pr.attrib.pop("autoPict", None)
