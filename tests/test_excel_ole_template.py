"""ExcelOleTemplate：从内置模板抽出可克隆的 OLE 零件。

模板是 Windows Excel + ChemDraw 真实产出的，克隆它的片段是 generated 路径
正确性的根基——早期自己拼 XML 会让 Excel 修复文件并删掉对象。
"""

from __future__ import annotations

import io
import xml.etree.ElementTree as ET
import zipfile

import olefile
import pytest

from retain_cdx.office.excel._ole_template import NS_V, ExcelOleTemplate
from retain_cdx.office.excel._ooxml.xml_io import NS, NS_MAIN, NS_R, NS_XDR


@pytest.fixture(scope="module")
def template() -> ExcelOleTemplate:
    return ExcelOleTemplate.bundled()


def test_container_is_a_chemdraw_ole(template):
    ole = olefile.OleFileIO(io.BytesIO(template.container))
    assert set(name[0] for name in ole.listdir()) == {
        "\x01Ole",
        "\x01CompObj",
        "CONTENTS",
    }
    assert ole.openstream("CONTENTS").read()[:8] == b"VjCD0100"


def test_fragments_keep_the_wrappers_excel_wrote(template):
    """包装层不能丢：oleObject 的 mc:AlternateContent、drawing 的隐藏占位。"""
    ole_object = template.new_ole_object()
    assert ole_object.tag.endswith("AlternateContent")
    assert ole_object.find("mc:Choice/s:oleObject", NS) is not None
    assert ole_object.find("mc:Fallback/s:oleObject", NS) is not None

    anchor = template.new_drawing_anchor()
    assert anchor.tag == f"{{{NS_XDR}}}twoCellAnchor"
    assert anchor.find(f"{{{NS_XDR}}}from") is not None
    assert anchor.find(f".//{{{NS_XDR}}}cNvPr") is not None

    shape = template.new_vml_shape()
    assert shape.tag == f"{{{NS_V}}}shape"


def test_fragments_have_no_generated_prefixes(template):
    """未注册命名空间就 tostring 会产出 ns0:/ns1:，污染克隆结果。"""
    for xml in (
        template.ole_object_xml,
        template.drawing_anchor_xml,
        template.vml_shape_xml,
    ):
        assert "ns0:" not in xml and "ns1:" not in xml


def test_clones_are_independent(template):
    first, second = template.new_drawing_anchor(), template.new_drawing_anchor()
    first.set("editAs", "changed")
    assert second.get("editAs") != "changed"


def test_base_anchor_matches_the_legacy_reader(template, tmp_path):
    """与现有 generated 路径读到的基准值一致。"""
    from retain_cdx.office._package import extract_ooxml
    from retain_cdx.office.excel._source.template import extract_base_anchor
    from retain_cdx.resources import bundled_template_xlsx

    package = tmp_path / "pkg"
    package.mkdir()
    with bundled_template_xlsx() as bundled:
        extract_ooxml(bundled, package)
    legacy = extract_base_anchor(package / "xl" / "worksheets" / "sheet1.xml")
    assert legacy == (template.base_shape_id, template.from_row, template.to_row)


def test_ole_object_carries_the_relationship_slots(template):
    """克隆后要改的两个 r:id 必须存在（oleObject 与 objectPr）。"""
    ole = template.new_ole_object().find("mc:Choice/s:oleObject", NS)
    assert ole.get(f"{{{NS_R}}}id"), "oleObject 缺 r:id"
    object_pr = ole.find("s:objectPr", NS)
    assert object_pr is not None and object_pr.get(f"{{{NS_R}}}id"), "objectPr 缺 r:id"
    assert object_pr.find("s:anchor", NS) is not None


def test_missing_pieces_raise_clear_errors(tmp_path):
    from retain_cdx.office._package import extract_ooxml
    from retain_cdx.resources import bundled_template_xlsx

    package = tmp_path / "broken"
    package.mkdir()
    with bundled_template_xlsx() as bundled:
        extract_ooxml(bundled, package)
    vml = package / "xl" / "drawings" / "vmlDrawing1.vml"
    vml.write_text("<xml xmlns:v='urn:schemas-microsoft-com:vml'/>", encoding="utf-8")

    with pytest.raises(ValueError, match="v:shape"):
        ExcelOleTemplate.from_package(package)
