"""Merge context: open target package slots and copy binary parts (bin + media).

装进 Excel · 子步骤：
  B. 落盘零件 — embeddings/*.bin、media/*
  C. 接线 — sheet/vml relationships（rId）
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from retain_cdx.office.excel._ooxml import REL_DRAWING, REL_VML, build_rels, read_rels

from .util import next_rid, unique_part_name


@dataclass
class OlePartMergeContext:
    """State while merging one source OLE workbook into a target sheet package."""

    source_dir: Path
    target_dir: Path
    prefix: str
    sheet_rels_path: Path
    sheet_rels: list[tuple[str, str, str]] = field(default_factory=list)
    drawing_rid: str = ""
    vml_rid: str = ""
    drawing_name: str = ""
    vml_name: str = ""
    created_drawing: bool = False
    created_vml: bool = False
    copied_media: dict[str, str] = field(default_factory=dict)

    # --- B: directories + binary parts ---

    def ensure_part_dirs(self) -> None:
        for rel in ["xl/drawings", "xl/drawings/_rels", "xl/embeddings", "xl/media"]:
            (self.target_dir / rel).mkdir(parents=True, exist_ok=True)
        self.sheet_rels_path.parent.mkdir(parents=True, exist_ok=True)

    def copy_ole_target(self, target: str) -> str:
        name = Path(target).name
        new_name = unique_part_name(
            self.target_dir / "xl" / "embeddings", f"{self.prefix}_{name}"
        )
        (self.target_dir / "xl" / "embeddings" / new_name).write_bytes(
            (self.source_dir / "xl" / "embeddings" / name).read_bytes()
        )
        return f"../embeddings/{new_name}"

    def copy_media_target(self, target: str) -> str:
        name = Path(target).name
        if name in self.copied_media:
            return self.copied_media[name]
        new_name = unique_part_name(
            self.target_dir / "xl" / "media", f"{self.prefix}_{name}"
        )
        copied_target = f"../media/{new_name}"
        (self.target_dir / "xl" / "media" / new_name).write_bytes(
            (self.source_dir / "xl" / "media" / name).read_bytes()
        )
        self.copied_media[name] = copied_target
        return copied_target

    # --- C: relationships ---

    def load_sheet_rels(self) -> None:
        self.sheet_rels = read_rels(self.sheet_rels_path)
        for rid, rtype, target in self.sheet_rels:
            if rtype.endswith("/drawing") and not self.drawing_rid:
                self.drawing_rid = rid
                self.drawing_name = Path(target).name
            elif rtype.endswith("/vmlDrawing") and not self.vml_rid:
                self.vml_rid = rid
                self.vml_name = Path(target).name

    def ensure_drawing_parts(self) -> None:
        """Ensure target sheet has drawing + VML parts (create from source if missing)."""
        if not self.drawing_name:
            self.created_drawing = True
            self.drawing_name = unique_part_name(
                self.target_dir / "xl" / "drawings",
                f"{self.prefix}_drawing.xml",
            )
            self.drawing_rid = self.next_sheet_rid()
            self.sheet_rels.append(
                (self.drawing_rid, REL_DRAWING, f"../drawings/{self.drawing_name}")
            )
            (self.target_dir / "xl" / "drawings" / self.drawing_name).write_bytes(
                (self.source_dir / "xl" / "drawings" / "drawing1.xml").read_bytes()
            )

        if not self.vml_name:
            self.created_vml = True
            self.vml_name = unique_part_name(
                self.target_dir / "xl" / "drawings",
                f"{self.prefix}_vmlDrawing.vml",
            )
            self.vml_rid = self.next_sheet_rid()
            self.sheet_rels.append(
                (self.vml_rid, REL_VML, f"../drawings/{self.vml_name}")
            )
            (self.target_dir / "xl" / "drawings" / self.vml_name).write_bytes(
                (self.source_dir / "xl" / "drawings" / "vmlDrawing1.vml").read_bytes()
            )

    def merge_sheet_relationships(self) -> dict[str, str]:
        """Copy OLE/image rels from source sheet into target; return source→target rId map."""
        rid_map: dict[str, str] = {}
        source_rels = read_rels(
            self.source_dir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels"
        )
        for source_rid, rtype, target in source_rels:
            if rtype.endswith("/oleObject"):
                new_rid = self.next_sheet_rid()
                copied_target = self.copy_ole_target(target)
                self.sheet_rels.append((new_rid, rtype, copied_target))
                rid_map[source_rid] = new_rid
            elif rtype.endswith("/image"):
                new_rid = self.next_sheet_rid()
                copied_target = self.copy_media_target(target)
                self.sheet_rels.append((new_rid, rtype, copied_target))
                rid_map[source_rid] = new_rid
        build_rels(self.sheet_rels_path, self.sheet_rels)
        return rid_map

    def merge_vml_relationships(self) -> dict[str, str]:
        """Copy image rels used by VML into target VML .rels."""
        vml_rels_path = (
            self.target_dir / "xl" / "drawings" / "_rels" / f"{self.vml_name}.rels"
        )
        vml_rels = read_rels(vml_rels_path)
        rid_map: dict[str, str] = {}
        source_vml_rels = read_rels(
            self.source_dir / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels"
        )
        for source_rid, rtype, target in source_vml_rels:
            new_rid = next_rid(vml_rels)
            copied_target = self.copy_media_target(target)
            vml_rels.append((new_rid, rtype, copied_target))
            rid_map[source_rid] = new_rid
        build_rels(vml_rels_path, vml_rels)
        return rid_map

    def next_sheet_rid(self) -> str:
        return next_rid(self.sheet_rels)
