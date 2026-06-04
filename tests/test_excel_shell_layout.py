from __future__ import annotations

from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

from retain_cdx import ExcelCDXWriter, ShellTemplate
from retain_cdx.office.excel.package import extract_xlsx, pack_xlsx
from PIL import Image


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"s": NS_MAIN}


def test_shell_insert_preserves_existing_sheet_cells(tmp_path: Path) -> None:
    source = Path("tests/fixtures/blank.xlsx")
    target = tmp_path / "target.xlsx"
    package_dir = tmp_path / "target_package"
    package_dir.mkdir()
    extract_xlsx(source, package_dir)

    sheet_path = package_dir / "xl" / "worksheets" / "sheet1.xml"
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    sheet_data = root.find("s:sheetData", NS)
    assert sheet_data is not None
    row = ET.SubElement(sheet_data, f"{{{NS_MAIN}}}row", r="1")
    cell = ET.SubElement(row, f"{{{NS_MAIN}}}c", r="A1", t="inlineStr")
    inline = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
    text = ET.SubElement(inline, f"{{{NS_MAIN}}}t")
    text.text = "keep me"
    tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)
    pack_xlsx(package_dir, target)

    output = tmp_path / "output.xlsx"
    with ExcelCDXWriter(target) as writer:
        writer.insert_cdx(
            Path("tests/fixtures/test.cdx"),
            sheet="Sheet1",
            cell="A1",
            shell=ShellTemplate(Path("tests/fixtures/shell_large.xlsx")),
        )
        writer.save(output)

    output_dir = tmp_path / "output_package"
    output_dir.mkdir()
    extract_xlsx(output, output_dir)
    output_sheet = ET.parse(output_dir / "xl" / "worksheets" / "sheet1.xml").getroot()

    kept_text = output_sheet.find(".//s:c[@r='A1']/s:is/s:t", NS)
    assert kept_text is not None
    assert kept_text.text == "keep me"

    output_row = output_sheet.find("s:sheetData/s:row[@r='1']", NS)
    assert output_row is not None
    assert output_row.get("ht") == "95.25"
    assert output_row.get("customHeight") == "1"


def test_shell_template_does_not_fit_cdx_by_default() -> None:
    assert ShellTemplate(Path("tests/fixtures/shell_large.xlsx")).fit_cdx is False


def test_shell_preview_matches_shell_display_aspect(tmp_path: Path) -> None:
    output = tmp_path / "output.xlsx"
    shell = Path("tests/fixtures/shell_large.xlsx")
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx(
            Path("tests/fixtures/test.cdx"),
            sheet="Sheet1",
            cell="A1",
            shell=ShellTemplate(shell),
            preview_width=700,
            preview_height=700,
            preview_no_margin=True,
        )
        writer.save(output)

    with zipfile.ZipFile(shell) as z:
        vml = z.read("xl/drawings/vmlDrawing1.vml").decode("utf-8")
    match = re.search(r"width:([0-9.]+)pt;.*?height:([0-9.]+)pt", vml, re.S)
    assert match is not None
    shell_aspect = float(match.group(1)) / float(match.group(2))

    with zipfile.ZipFile(output) as z:
        preview = tmp_path / "preview.png"
        preview.write_bytes(z.read("xl/media/sheet1_retain_image1.png"))
    with Image.open(preview) as image:
        preview_aspect = image.width / image.height

    assert abs(preview_aspect - shell_aspect) < 0.01
