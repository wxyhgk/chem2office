"""Insertion must not damage what the target workbook already contains."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from ooxml_checks import check_package
from retain_cdx import ExcelCDXWriter
from retain_cdx.office._package import extract_ooxml, pack_ooxml
from retain_cdx.office.excel._grid import ExcelGridMetrics

from test_excel_shell_layout import patch_preview_renderers

CDX = "tests/fixtures/test.cdx"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"s": NS_MAIN}
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_V = "urn:schemas-microsoft-com:vml"


def _repack(tmp_path: Path, mutate) -> Path:
    package = tmp_path / "pkg"
    package.mkdir(parents=True, exist_ok=True)
    extract_ooxml(Path("tests/fixtures/blank.xlsx"), package)
    mutate(package)
    target = tmp_path / "target.xlsx"
    pack_ooxml(package, target)
    return target


# --- C8: cell style and calcChain ---------------------------------------


def test_overwriting_a_cell_keeps_its_style(tmp_path):
    def mutate(package: Path) -> None:
        sheet = package / "xl" / "worksheets" / "sheet1.xml"
        text = sheet.read_text(encoding="utf-8")
        text = text.replace(
            "<sheetData></sheetData>",
            '<sheetData><row r="1"><c r="A1" s="3" t="inlineStr">'
            "<is><t>old</t></is></c></row></sheetData>",
        )
        sheet.write_text(text, encoding="utf-8")

    target = _repack(tmp_path, mutate)
    out = tmp_path / "out.xlsx"
    with ExcelCDXWriter(target) as writer:
        writer.set_cell_text("Sheet1", "A1", "new")
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    cell = root.find('.//s:c[@r="A1"]', NS)
    assert cell is not None
    assert cell.get("s") == "3", "cell format must survive a text overwrite"
    assert cell.find("s:is/s:t", NS).text == "new"


def test_overwriting_a_formula_cell_drops_calc_chain(tmp_path):
    """A chain entry for a cell that no longer has a formula triggers repair."""

    def mutate(package: Path) -> None:
        sheet = package / "xl" / "worksheets" / "sheet1.xml"
        text = sheet.read_text(encoding="utf-8")
        text = text.replace(
            "<sheetData></sheetData>",
            '<sheetData><row r="2"><c r="B2"><f>1+2</f><v>3</v></c></row></sheetData>',
        )
        sheet.write_text(text, encoding="utf-8")
        (package / "xl" / "calcChain.xml").write_text(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<calcChain xmlns="{NS_MAIN}"><c r="B2" i="1"/></calcChain>',
            encoding="utf-8",
        )

    target = _repack(tmp_path, mutate)
    assert "xl/calcChain.xml" in zipfile.ZipFile(target).namelist()

    out = tmp_path / "out.xlsx"
    with ExcelCDXWriter(target) as writer:
        writer.set_cell_text("Sheet1", "B2", "text now")
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        assert "xl/calcChain.xml" not in archive.namelist()
        assert "calcChain" not in archive.read("[Content_Types].xml").decode()
    assert check_package(out) == []


def test_cells_untouched_by_us_keep_their_formula(tmp_path):
    def mutate(package: Path) -> None:
        sheet = package / "xl" / "worksheets" / "sheet1.xml"
        text = sheet.read_text(encoding="utf-8")
        sheet.write_text(
            text.replace(
                "<sheetData></sheetData>",
                '<sheetData><row r="2"><c r="B2"><f>1+2</f><v>3</v></c></row></sheetData>',
            ),
            encoding="utf-8",
        )

    target = _repack(tmp_path, mutate)
    out = tmp_path / "out.xlsx"
    with ExcelCDXWriter(target) as writer:
        writer.set_cell_text("Sheet1", "A5", "elsewhere")
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    assert "<f>1+2</f>" in sheet


# --- C9: .xlsm content types --------------------------------------------


def test_embeddings_get_an_override_when_bin_means_vba(tmp_path, monkeypatch):
    patch_preview_renderers(monkeypatch)

    def mutate(package: Path) -> None:
        content_types = package / "[Content_Types].xml"
        text = content_types.read_text(encoding="utf-8")
        text = text.replace(
            "<Types ",
            "<Types ",
        ).replace(
            "</Types>",
            '<Default Extension="bin" '
            'ContentType="application/vnd.ms-office.vbaProject"/></Types>',
        )
        content_types.write_text(text, encoding="utf-8")

    target = _repack(tmp_path, mutate)
    out = tmp_path / "out.xlsm"
    with ExcelCDXWriter(target) as writer:
        writer.insert_cdx(CDX, sheet="Sheet1", cell="A1")
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        content_types = archive.read("[Content_Types].xml").decode()
        embeddings = [n for n in archive.namelist() if "/embeddings/" in n]
    assert embeddings
    for name in embeddings:
        assert f'PartName="/{name}"' in content_types, (
            "embedded OLE must be typed explicitly, not inherit the VBA default"
        )
    assert 'ContentType="application/vnd.ms-office.vbaProject"' in content_types


# --- F3: mc:Ignorable ----------------------------------------------------


def test_existing_ignorable_prefixes_are_kept(tmp_path, monkeypatch):
    patch_preview_renderers(monkeypatch)

    def mutate(package: Path) -> None:
        sheet = package / "xl" / "worksheets" / "sheet1.xml"
        text = sheet.read_text(encoding="utf-8")
        text = text.replace(
            "<worksheet ",
            '<worksheet xmlns:mc="http://schemas.openxmlformats.org/'
            'markup-compatibility/2006" xmlns:xr6="http://schemas.microsoft.com/'
            'office/spreadsheetml/2016/revision6" mc:Ignorable="xr6" ',
            1,
        )
        sheet.write_text(text, encoding="utf-8")

    target = _repack(tmp_path, mutate)
    out = tmp_path / "out.xlsx"
    with ExcelCDXWriter(target) as writer:
        writer.insert_cdx(CDX, sheet="Sheet1", cell="A1")
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    ignorable = re.search(r'mc:Ignorable="([^"]+)"', sheet).group(1).split()
    assert "xr6" in ignorable, "target's own ignorable prefixes were dropped"
    assert "x14ac" in ignorable
    assert check_package(out) == []


# --- F4: grouped column ranges -------------------------------------------


def test_setting_a_width_splits_a_grouped_col_range(tmp_path):
    def mutate(package: Path) -> None:
        sheet = package / "xl" / "worksheets" / "sheet1.xml"
        text = sheet.read_text(encoding="utf-8")
        sheet.write_text(
            text.replace(
                "<sheetData>",
                '<cols><col min="1" max="10" width="20" customWidth="1"/></cols>'
                "<sheetData>",
            ),
            encoding="utf-8",
        )

    target = _repack(tmp_path, mutate)
    out = tmp_path / "out.xlsx"
    with ExcelCDXWriter(target) as writer:
        writer.set_column_width("Sheet1", "B", 33)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    spans = [
        (int(col.get("min")), int(col.get("max")), col.get("width"))
        for col in root.findall("s:cols/s:col", NS)
    ]
    covering = [(lo, hi) for lo, hi, _w in spans if lo <= 2 <= hi]
    assert covering == [(2, 2)], f"overlapping definitions for column B: {spans}"
    assert dict(((lo, hi), w) for lo, hi, w in spans)[(2, 2)] == "33"
    assert (1, 1, "20") in spans and (3, 10, "20") in spans


# --- F2: anchor overflow --------------------------------------------------


def test_two_point_anchor_lands_in_the_right_cell(tmp_path, monkeypatch):
    """Regression: the to-corner used to stay on the from-cell with a huge
    offset — Excel copes, other consumers clamp."""
    patch_preview_renderers(monkeypatch)
    from retain_cdx import ExcelPlacement

    out = tmp_path / "out.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx(
            CDX, sheet="Sheet1", placement=ExcelPlacement("B2", to_cell="E8")
        )
        writer.save(out)

    package = tmp_path / "unpacked"
    package.mkdir()
    extract_ooxml(out, package)
    metrics = ExcelGridMetrics.from_worksheet(
        package / "xl" / "worksheets" / "sheet1.xml"
    )
    drawing = ET.parse(package / "xl" / "drawings" / "sheet1_retain_drawing.xml")
    anchor = next(drawing.getroot().iter(f"{{{NS_XDR}}}twoCellAnchor"))
    corner = anchor.find(f"{{{NS_XDR}}}to")
    col = int(corner.find(f"{{{NS_XDR}}}col").text)
    col_off = int(corner.find(f"{{{NS_XDR}}}colOff").text) // 9525
    row = int(corner.find(f"{{{NS_XDR}}}row").text)
    row_off = int(corner.find(f"{{{NS_XDR}}}rowOff").text) // 9525

    assert col_off < metrics.col_width_px(col), "colOff still overflows its column"
    assert row_off < metrics.row_height_px(row), "rowOff still overflows its row"
    # B2 → E8 means the box ends at E8's top-left corner.
    assert (col, col_off, row, row_off) == (4, 0, 7, 0)


def test_vml_anchor_matches_the_drawing_after_normalisation(tmp_path, monkeypatch):
    patch_preview_renderers(monkeypatch)
    from retain_cdx import ExcelPlacement

    out = tmp_path / "out.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx(
            CDX, sheet="Sheet1", placement=ExcelPlacement("B2", to_cell="E8")
        )
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        vml = archive.read("xl/drawings/sheet1_retain_vmlDrawing.vml").decode()
    values = [
        int(part)
        for part in re.search(r"<x:Anchor>([^<]+)</x:Anchor>", vml).group(1).split(",")
    ]
    assert values[4:] == [4, 0, 7, 0], f"VML anchor disagrees with the drawing: {values}"


def test_ordinary_insert_anchor_is_unchanged(tmp_path, monkeypatch):
    """Normalisation must be a no-op when the box fits its cell."""
    patch_preview_renderers(monkeypatch)
    out = tmp_path / "out.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx(CDX, sheet="Sheet1", cell="B2")
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        drawing = archive.read("xl/drawings/sheet1_retain_drawing.xml").decode()
    to_block = re.search(r"<xdr:to>[\s\S]*?</xdr:to>", drawing).group(0)
    assert "<xdr:col>1</xdr:col>" in to_block
    assert "<xdr:colOff>1381125</xdr:colOff>" in to_block
    assert "<xdr:row>1</xdr:row>" in to_block
