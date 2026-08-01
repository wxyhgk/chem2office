"""从内置模板里取出「一个 OLE 对象」所需的全部零件。

为什么是克隆而不是从零拼：早期版本自己拼 worksheet/drawing/VML XML，
Excel 打开后会「修复」文件并把 OLE 对象整个删掉（见 `docs/developer_log.md`
2026-03-06）。模板是 Windows Excel + ChemDraw 真实插入产出的，克隆它的片段
才能保证 Excel 认。

一个对象需要四份东西，都从模板取：

===================  =====================================================
容器 `.bin`          CFBF 壳，注入新 CDX 后即为该分子的 OLE
`<oleObject>` 片段   worksheet 上的声明（含 `mc:AlternateContent` 包装）
drawing 锚点片段     `xdr:twoCellAnchor`，含隐藏占位 `sp` 与 a14 扩展
VML shape 片段       `v:shape`，Excel 真正拿来渲染预览的那层
===================  =====================================================

与 PowerPoint 的 `_pml.OleFrameTemplate` 是同一个思路：模板是**零件供体**，
不是工作目录。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

from retain_cdx.office._package import extract_ooxml
from retain_cdx.resources import bundled_template_xlsx

from ._ooxml.xml_io import NS, register_namespaces

_TEMPLATE_SHEET = "xl/worksheets/sheet1.xml"
_TEMPLATE_DRAWING = "xl/drawings/drawing1.xml"
_TEMPLATE_VML = "xl/drawings/vmlDrawing1.vml"
_TEMPLATE_EMBEDDING = "xl/embeddings/oleObject1.bin"

NS_V = "urn:schemas-microsoft-com:vml"


@dataclass(frozen=True)
class ExcelOleTemplate:
    """一个 Excel+ChemDraw OLE 对象的可克隆零件集。

    XML 片段以字符串保存（而不是 Element），这样每次克隆都是独立解析，
    不会有共享可变状态；PowerPoint 侧同理。
    """

    container: bytes
    ole_object_xml: str      # <mc:AlternateContent> 包着的 <oleObject>
    drawing_anchor_xml: str  # <xdr:twoCellAnchor>
    vml_shape_xml: str       # <v:shape>
    base_shape_id: int
    from_row: int
    to_row: int

    @classmethod
    def from_package(cls, package_dir: Path) -> "ExcelOleTemplate":
        # 不先注册前缀的话，tostring 会写成 ns0:/ns1:，克隆出来的片段
        # 带着一堆无意义的前缀声明。
        register_namespaces()
        sheet_root = ET.parse(package_dir / _TEMPLATE_SHEET).getroot()
        ole_alt = sheet_root.find("s:oleObjects/mc:AlternateContent", NS)
        if ole_alt is None:
            raise ValueError(f"模板 {_TEMPLATE_SHEET} 里没有 oleObjects")

        base_ole = ole_alt.find("mc:Choice/s:oleObject", NS)
        if base_ole is None:
            raise ValueError("模板的 oleObject 缺 mc:Choice 分支")
        anchor = base_ole.find("s:objectPr/s:anchor", NS)
        if anchor is None:
            raise ValueError("模板的 objectPr 没有 anchor")

        drawing_root = ET.parse(package_dir / _TEMPLATE_DRAWING).getroot()
        drawing_anchor = drawing_root.find(
            "mc:AlternateContent/mc:Choice/xdr:twoCellAnchor", NS
        )
        if drawing_anchor is None:
            raise ValueError(f"模板 {_TEMPLATE_DRAWING} 里没有 twoCellAnchor")

        vml_root = ET.parse(package_dir / _TEMPLATE_VML).getroot()
        vml_shape = vml_root.find(f"{{{NS_V}}}shape")
        if vml_shape is None:
            raise ValueError(f"模板 {_TEMPLATE_VML} 里没有 v:shape")

        return cls(
            container=(package_dir / _TEMPLATE_EMBEDDING).read_bytes(),
            ole_object_xml=ET.tostring(ole_alt, encoding="unicode"),
            drawing_anchor_xml=ET.tostring(drawing_anchor, encoding="unicode"),
            vml_shape_xml=ET.tostring(vml_shape, encoding="unicode"),
            base_shape_id=int(base_ole.get("shapeId", "1025")),
            from_row=_row(anchor, "from"),
            to_row=_row(anchor, "to"),
        )

    @classmethod
    def bundled(cls) -> "ExcelOleTemplate":
        """内置模板（`retain_cdx/data/template.xlsx`）。"""
        import tempfile

        with tempfile.TemporaryDirectory(prefix="retain_cdx_tpl_") as tmp:
            package = Path(tmp)
            with bundled_template_xlsx() as template:
                extract_ooxml(template, package)
            return cls.from_package(package)

    # --- 克隆 ---

    def new_ole_object(self) -> ET.Element:
        return ET.fromstring(self.ole_object_xml)

    def new_drawing_anchor(self) -> ET.Element:
        return ET.fromstring(self.drawing_anchor_xml)

    def new_vml_shape(self) -> ET.Element:
        return ET.fromstring(self.vml_shape_xml)


def _row(anchor: ET.Element, corner: str) -> int:
    node = anchor.find(f"s:{corner}/xdr:row", NS)
    if node is None or node.text is None:
        raise ValueError(f"模板 anchor 的 <{corner}> 缺 row")
    return int(node.text)


__all__ = ["ExcelOleTemplate"]
