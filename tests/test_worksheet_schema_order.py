from __future__ import annotations

from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

import pytest

from retain_cdx import ExcelCDXWriter
from retain_cdx.office._package import extract_ooxml, pack_ooxml
from retain_cdx.office.excel._ooxml.xml_io import (
    WORKSHEET_CHILD_ORDER,
    local_name,
)

from test_excel_shell_layout import patch_preview_renderers

_RANK = {name: idx for idx, name in enumerate(WORKSHEET_CHILD_ORDER)}


def _order_violations(sheet_xml: bytes) -> list[tuple[str, str]]:
    kids = [local_name(child.tag) for child in ET.fromstring(sheet_xml)]
    ranks = [_RANK.get(k, -1) for k in kids]
    return [
        (kids[i], kids[i + 1])
        for i in range(len(kids) - 1)
        if ranks[i] > ranks[i + 1] >= 0
    ]


def _sheet_children(xlsx_path: Path) -> list[str]:
    with zipfile.ZipFile(xlsx_path) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    return [local_name(child.tag) for child in root]


def _add_trailing_page_parts(source: Path, target: Path, tmp_path: Path) -> None:
    """Give a sheet the pageSetup/headerFooter that real Excel files carry."""
    package = tmp_path / "unpacked"
    package.mkdir()
    extract_ooxml(source, package)
    sheet_path = package / "xl" / "worksheets" / "sheet1.xml"
    text = sheet_path.read_text(encoding="utf-8")
    assert "<pageSetup" not in text
    extra = (
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75"'
        ' header="0.3" footer="0.3"/>'
        '<pageSetup paperSize="9" orientation="portrait"/>'
        "<headerFooter><oddHeader>&amp;A</oddHeader></headerFooter>"
    )
    text = text.replace("</worksheet>", f"{extra}</worksheet>")
    sheet_path.write_text(text, encoding="utf-8")
    pack_ooxml(package, target)


def test_insert_keeps_page_setup_before_drawing(tmp_path, monkeypatch):
    """Regression: drawing/oleObjects were inserted right after pageMargins,
    pushing pageSetup/headerFooter behind them — Excel repairs such a file and
    drops the OLE objects."""
    patch_preview_renderers(monkeypatch)
    target = tmp_path / "with_page_setup.xlsx"
    _add_trailing_page_parts(Path("tests/fixtures/blank.xlsx"), target, tmp_path)

    out = tmp_path / "inserted.xlsx"
    with ExcelCDXWriter(target) as writer:
        writer.insert_cdx("tests/fixtures/test.cdx", sheet="Sheet1", cell="A1")
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        sheet_xml = archive.read("xl/worksheets/sheet1.xml")
    assert _order_violations(sheet_xml) == []

    kids = _sheet_children(out)
    for earlier, later in (
        ("pageMargins", "pageSetup"),
        ("pageSetup", "headerFooter"),
        ("headerFooter", "drawing"),
        ("drawing", "legacyDrawing"),
        ("legacyDrawing", "oleObjects"),
    ):
        assert kids.index(earlier) < kids.index(later), f"{earlier} must precede {later}"


def test_plain_sheet_insert_order_is_legal(tmp_path, monkeypatch):
    patch_preview_renderers(monkeypatch)
    out = tmp_path / "plain.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx("tests/fixtures/test.cdx", sheet="Sheet1", cell="A1")
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        assert _order_violations(archive.read("xl/worksheets/sheet1.xml")) == []


def test_set_cell_texts_dimension_lands_after_sheet_pr(tmp_path):
    package = tmp_path / "pkg"
    package.mkdir()
    extract_ooxml(Path("tests/fixtures/blank.xlsx"), package)
    sheet_path = package / "xl" / "worksheets" / "sheet1.xml"
    text = sheet_path.read_text(encoding="utf-8")
    # Strip dimension so the writer has to create one; the fixture already
    # carries a <sheetPr>, which is the only legal predecessor.
    text, count = re.subn(r"<dimension[^>]*/>", "", text, count=1)
    assert count == 1
    assert "<sheetPr>" in text
    sheet_path.write_text(text, encoding="utf-8")
    stripped = tmp_path / "stripped.xlsx"
    pack_ooxml(package, stripped)

    out = tmp_path / "cells.xlsx"
    with ExcelCDXWriter(stripped) as writer:
        writer.set_cell_text("Sheet1", "B2", "x")
        writer.save(out)

    kids = _sheet_children(out)
    assert kids[0] == "sheetPr"
    assert kids[1] == "dimension"
    with zipfile.ZipFile(out) as archive:
        assert _order_violations(archive.read("xl/worksheets/sheet1.xml")) == []


@pytest.mark.parametrize(
    "element,expected_before",
    [("drawing", "oleObjects"), ("pageSetup", "drawing")],
)
def test_worksheet_order_table_is_self_consistent(element, expected_before):
    assert _RANK[element] < _RANK[expected_before]
