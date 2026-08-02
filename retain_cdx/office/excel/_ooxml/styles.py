"""按需在 `xl/styles.xml` 里登记单元格对齐样式。

Excel 的对齐不能写在单元格上，得指向 `cellXfs` 里的一条 `<xf>`；单元格通过
`s="索引"` 引用它。所以「把这格垂直居中」实际上是「找到或新建一条带该对齐的
xf，再把索引写进单元格」。

同一种对齐只登记一条并复用——否则每写一格就长一条 xf，几千个结构就是几千条。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

from .xml_io import NS, NS_MAIN

_ALIGN_ATTRS = ("horizontal", "vertical", "wrapText")


@dataclass(frozen=True)
class ExcelCellAlignment:
    """单元格对齐。字段为 None 表示不设，沿用默认。"""

    horizontal: str | None = None  # left / center / right
    vertical: str | None = None    # top / center / bottom
    wrap_text: bool = False

    def as_attrs(self) -> dict[str, str]:
        attrs: dict[str, str] = {}
        if self.horizontal:
            attrs["horizontal"] = self.horizontal
        if self.vertical:
            attrs["vertical"] = self.vertical
        if self.wrap_text:
            attrs["wrapText"] = "1"
        return attrs

    def __bool__(self) -> bool:
        return bool(self.as_attrs())


def ensure_alignment_style(package_dir: Path, alignment: ExcelCellAlignment) -> int:
    """返回可用于单元格 ``s=`` 的 xf 索引；相同对齐只登记一次。"""
    attrs = alignment.as_attrs()
    if not attrs:
        return 0

    styles_path = package_dir / "xl" / "styles.xml"
    if not styles_path.exists():
        return 0

    tree = ET.parse(styles_path)
    root = tree.getroot()
    cell_xfs = root.find("s:cellXfs", NS)
    if cell_xfs is None:
        return 0

    for index, xf in enumerate(cell_xfs.findall("s:xf", NS)):
        existing = xf.find("s:alignment", NS)
        if existing is not None and _alignment_attrs(existing) == attrs:
            return index

    base = cell_xfs.find("s:xf", NS)
    xf = ET.SubElement(cell_xfs, f"{{{NS_MAIN}}}xf")
    for key in ("numFmtId", "fontId", "fillId", "borderId", "xfId"):
        xf.set(key, base.get(key, "0") if base is not None else "0")
    xf.set("applyAlignment", "1")
    ET.SubElement(xf, f"{{{NS_MAIN}}}alignment", attrs)

    index = len(cell_xfs.findall("s:xf", NS)) - 1
    cell_xfs.set("count", str(index + 1))
    _write(tree, styles_path)
    return index


def _alignment_attrs(element: ET.Element) -> dict[str, str]:
    return {key: value for key, value in element.attrib.items() if key in _ALIGN_ATTRS}


def _write(tree: ET.ElementTree, path: Path) -> None:
    ET.register_namespace("", NS_MAIN)
    xml = ET.tostring(tree.getroot(), encoding="UTF-8", xml_declaration=True)
    path.write_bytes(xml)


__all__ = ["ExcelCellAlignment", "ensure_alignment_style"]
