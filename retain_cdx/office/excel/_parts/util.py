"""Small OOXML helpers used only by the merge/place step."""

from __future__ import annotations

import re
from pathlib import Path
import xml.etree.ElementTree as ET


def next_rid(rels: list[tuple[str, str, str]]) -> str:
    max_id = 0
    for rid, _, _ in rels:
        match = re.fullmatch(r"rId(\d+)", rid)
        if match:
            max_id = max(max_id, int(match.group(1)))
    return f"rId{max_id + 1}"


def unique_part_name(directory: Path, desired_name: str) -> str:
    candidate = desired_name
    path = directory / candidate
    if not path.exists():
        return candidate
    stem = path.stem
    suffix = path.suffix
    idx = 2
    while True:
        candidate = f"{stem}_{idx}{suffix}"
        if not (directory / candidate).exists():
            return candidate
        idx += 1


def parse_int(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def write_xml(tree: ET.ElementTree, path: Path) -> None:
    tree.write(path, encoding="UTF-8", xml_declaration=True)
