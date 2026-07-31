from __future__ import annotations

from pathlib import Path
import posixpath
import re
import xml.etree.ElementTree as ET

from .relationships import read_rels
from .xml_io import NS, NS_R


def parse_workbook_sheets(package_dir: Path) -> dict[str, str]:
    wb_root = ET.parse(package_dir / "xl" / "workbook.xml").getroot()
    rels = {
        rid: target
        for rid, _, target in read_rels(
            package_dir / "xl" / "_rels" / "workbook.xml.rels"
        )
    }
    sheets: dict[str, str] = {}
    for sheet in wb_root.findall("s:sheets/s:sheet", NS):
        rid = sheet.get(f"{{{NS_R}}}id")
        target = rels.get(rid or "")
        if target:
            if target.startswith("/"):
                path = target.lstrip("/")
            else:
                path = posixpath.normpath(posixpath.join("xl", target))
            sheets[sheet.get("name", "")] = path
    return sheets


def sheet_index_from_path(path: str) -> int:
    match = re.search(r"sheet(\d+)\.xml$", path)
    if not match:
        raise ValueError(f"Cannot determine sheet index from {path}")
    return int(match.group(1))
