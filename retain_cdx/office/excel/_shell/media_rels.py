"""Image relationship helpers for shell packages (wiring preview media).

装进 Excel · 子步骤 C（shell 路径）：更新 vml/sheet .rels 指向新预览图。
"""

from __future__ import annotations

from pathlib import Path

from retain_cdx.office.excel._ooxml import REL_IMAGE, read_rels, rewrite_rel_targets


def find_image_relationship(rels_path: Path) -> tuple[str, str] | None:
    """Return ``(rId, Target)`` for the first image relationship, if any."""
    for rid, rtype, target in read_rels(rels_path):
        if rtype == REL_IMAGE:
            return rid, target
    return None


def set_image_relationship(rels_path: Path, rel_id: str, target: str) -> None:
    rewrite_rel_targets(rels_path, target, rel_id=rel_id)


def set_all_sheet_image_targets(rels_path: Path, target: str) -> None:
    """Point every sheet image relationship at *target* (single-OLE shell)."""
    if not rels_path.exists():
        return
    rewrite_rel_targets(rels_path, target, rel_type=REL_IMAGE)
