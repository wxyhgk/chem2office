from __future__ import annotations

import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from ooxml_checks import check_package
from retain_cdx import ExcelCDXWriter
from retain_cdx.office.excel._ooxml import extract_xlsx, pack_xlsx
from retain_cdx.office.excel._shell.grid import (
    ExcelGridMetrics,
    base_col_width_to_px,
    excel_col_width_to_px,
    workbook_max_digit_width,
)

from test_excel_shell_layout import patch_preview_renderers

NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _unpacked(source: str, tmp_path: Path) -> Path:
    package = tmp_path / "pkg"
    package.mkdir(parents=True)
    extract_xlsx(Path(source), package)
    return package


def test_base_col_width_rounds_up_to_multiple_of_eight():
    """Regression: 8 characters of Calibri 11 is 64px, not 8*7+5 = 61."""
    assert base_col_width_to_px(8, 7) == 64
    assert base_col_width_to_px(8, 8) == 72  # a CJK-default workbook
    # defaultColWidth / <col width> already include the padding, no rounding.
    assert excel_col_width_to_px(8, 7) == 61


def test_max_digit_width_follows_the_workbook_font(tmp_path):
    calibri = _unpacked("tests/fixtures/blank.xlsx", tmp_path)
    assert workbook_max_digit_width(calibri / "xl" / "worksheets" / "sheet1.xml") == 7

    cjk = _unpacked("retain_cdx/data/template.xlsx", tmp_path / "cjk")
    assert workbook_max_digit_width(cjk / "xl" / "worksheets" / "sheet1.xml") == 8


def test_grid_metrics_use_the_workbook_font(tmp_path):
    """The bundled template is a 等线-11 workbook: default column is 72px."""
    package = _unpacked("retain_cdx/data/template.xlsx", tmp_path)
    metrics = ExcelGridMetrics.from_worksheet(
        package / "xl" / "worksheets" / "sheet1.xml"
    )
    assert metrics.max_digit_width == 8
    assert metrics.default_col_width_px == 72


def test_grid_metrics_default_to_calibri(tmp_path):
    package = _unpacked("tests/fixtures/blank.xlsx", tmp_path)
    metrics = ExcelGridMetrics.from_worksheet(
        package / "xl" / "worksheets" / "sheet1.xml"
    )
    assert metrics.max_digit_width == 7
    assert metrics.default_col_width_px == 64


def test_unknown_font_falls_back_to_calibri_metrics(tmp_path):
    package = _unpacked("tests/fixtures/blank.xlsx", tmp_path)
    styles = package / "xl" / "styles.xml"
    styles.write_text(
        styles.read_text(encoding="utf-8").replace("Calibri", "Nonesuch Sans"),
        encoding="utf-8",
    )
    assert workbook_max_digit_width(package / "xl" / "worksheets" / "sheet1.xml") == 7


def _sheet_with_hyperlink(tmp_path: Path) -> Path:
    package = _unpacked("tests/fixtures/blank.xlsx", tmp_path)
    rels = package / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels"
    rels.parent.mkdir(parents=True, exist_ok=True)
    rels.write_text(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{NS_RELS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/hyperlink" '
        'Target="https://example.com/x" TargetMode="External"/>'
        "</Relationships>",
        encoding="utf-8",
    )
    source = tmp_path / "hyperlink.xlsx"
    pack_xlsx(package, source)
    return source


def test_insert_preserves_external_target_mode(tmp_path, monkeypatch):
    """Regression: rebuilding the .rels dropped TargetMode, breaking hyperlinks."""
    patch_preview_renderers(monkeypatch)
    source = _sheet_with_hyperlink(tmp_path)
    out = tmp_path / "out.xlsx"
    with ExcelCDXWriter(source) as writer:
        writer.insert_cdx("tests/fixtures/test.cdx", sheet="Sheet1", cell="C3")
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/_rels/sheet1.xml.rels"))
    hyperlinks = [
        rel
        for rel in root.findall(f"{{{NS_RELS}}}Relationship")
        if rel.get("Type", "").endswith("/hyperlink")
    ]
    assert len(hyperlinks) == 1
    assert hyperlinks[0].get("TargetMode") == "External"
    assert hyperlinks[0].get("Target") == "https://example.com/x"
    assert check_package(out) == []


def test_repeated_insert_keeps_the_hyperlink(tmp_path, monkeypatch):
    patch_preview_renderers(monkeypatch)
    source = _sheet_with_hyperlink(tmp_path)
    out = tmp_path / "twice.xlsx"
    with ExcelCDXWriter(source) as writer:
        writer.insert_cdx("tests/fixtures/test.cdx", sheet="Sheet1", cell="C3")
        writer.insert_cdx("tests/fixtures/test.cdx", sheet="Sheet1", cell="C8")
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        rels = archive.read("xl/worksheets/_rels/sheet1.xml.rels").decode()
    assert 'TargetMode="External"' in rels
    assert check_package(out) == []
