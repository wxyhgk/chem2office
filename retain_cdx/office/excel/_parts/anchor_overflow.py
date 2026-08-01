"""把越界的锚点偏移归一化到真实网格（子步骤 D 的收尾）.

shell 路径沿用模板的锚点，偏移可能远超该格的宽高。Excel 按 origin+offset
解析并在保存时自行归一化，但 Excel 自己产出的文件从不这么写，第三方消费者
可能直接 clamp。

（generated 路径改为直写后，锚点由 `ExcelGridMetrics` 按真实网格算出，本来
就落在正确的格子里，这里对它是 no-op。）

三处表示（worksheet ``<oleObject>`` / drawing ``xdr:twoCellAnchor`` /
VML ``x:Anchor``）必须一起改，否则会重新引入它们之间的不一致。
偏移本来就在格内时全程 no-op。
"""

from __future__ import annotations

from pathlib import Path
import re
import xml.etree.ElementTree as ET

from retain_cdx.office.excel._ooxml import NS_XDR, write_sheet_xml
from retain_cdx.office.excel._ooxml.xml_io import NS
from retain_cdx.office.excel._grid import ExcelGridMetrics
from retain_cdx.office.excel._utils import EMU_PER_PIXEL

from .util import write_xml

_NS_V = "urn:schemas-microsoft-com:vml"
_NS_X = "urn:schemas-microsoft-com:office:excel"


def normalize_anchor_overflow(
    sheet_path: Path, drawing_path: Path | None, vml_path: Path | None
) -> None:
    """Walk overflowing offsets into the cells they actually land in."""
    metrics = ExcelGridMetrics.from_worksheet(sheet_path)
    _normalize_worksheet(sheet_path, metrics)
    if drawing_path is not None and drawing_path.exists():
        _normalize_drawing(drawing_path, metrics)
    if vml_path is not None and vml_path.exists():
        _normalize_vml(vml_path, metrics)


def _walk(metrics: ExcelGridMetrics, col: int, col_off: int, row: int, row_off: int):
    col, col_off = metrics.split_col(col, col_off)
    row, row_off = metrics.split_row(row, row_off)
    return col, col_off, row, row_off


def _normalize_emu_corner(corner: ET.Element, metrics: ExcelGridMetrics) -> bool:
    parts = {
        tag: corner.find(f"{{{NS_XDR}}}{tag}")
        for tag in ("col", "colOff", "row", "rowOff")
    }
    if any(node is None for node in parts.values()):
        return False
    col, col_off, row, row_off = _walk(
        metrics,
        int(parts["col"].text),
        int(parts["colOff"].text) // EMU_PER_PIXEL,
        int(parts["row"].text),
        int(parts["rowOff"].text) // EMU_PER_PIXEL,
    )
    changed = parts["col"].text != str(col) or parts["row"].text != str(row)
    parts["col"].text = str(col)
    parts["colOff"].text = str(col_off * EMU_PER_PIXEL)
    parts["row"].text = str(row)
    parts["rowOff"].text = str(row_off * EMU_PER_PIXEL)
    return changed


def _normalize_worksheet(sheet_path: Path, metrics: ExcelGridMetrics) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    changed = False
    for anchor in root.iter():
        if anchor.tag.rsplit("}", 1)[-1] != "anchor":
            continue
        for tag in ("from", "to"):
            # Excel writes these children unprefixed inside <anchor>.
            corner = anchor.find(tag) or anchor.find(f"{{{NS_XDR}}}{tag}")
            if corner is not None:
                changed |= _normalize_emu_corner(corner, metrics)
    if changed:
        write_sheet_xml(tree, sheet_path)


def _normalize_drawing(drawing_path: Path, metrics: ExcelGridMetrics) -> None:
    tree = ET.parse(drawing_path)
    changed = False
    for anchor in tree.getroot().iter(f"{{{NS_XDR}}}twoCellAnchor"):
        for tag in ("from", "to"):
            corner = anchor.find(f"{{{NS_XDR}}}{tag}")
            if corner is not None:
                changed |= _normalize_emu_corner(corner, metrics)
    if changed:
        write_xml(tree, drawing_path)


def _normalize_vml(vml_path: Path, metrics: ExcelGridMetrics) -> None:
    text = vml_path.read_text(encoding="utf-8")

    def replace(match: re.Match[str]) -> str:
        values = [int(part) for part in match.group(1).split(",")]
        if len(values) != 8:
            return match.group(0)
        from_corner = _walk(metrics, *values[:4])
        to_corner = _walk(metrics, *values[4:])
        joined = ", ".join(str(value) for value in (*from_corner, *to_corner))
        return f"<x:Anchor>{joined}</x:Anchor>"

    updated = re.sub(r"<x:Anchor>([^<]+)</x:Anchor>", replace, text)
    if updated != text:
        vml_path.write_text(updated, encoding="utf-8")


__all__ = ["normalize_anchor_overflow"]
