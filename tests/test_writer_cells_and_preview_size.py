from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

import pytest
from PIL import Image

from retain_cdx import ExcelCDXWriter

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"s": NS_MAIN}


def _sheet_root(xlsx_path: Path) -> ET.Element:
    with zipfile.ZipFile(xlsx_path) as z:
        return ET.fromstring(z.read("xl/worksheets/sheet1.xml"))


def test_set_cell_texts_writes_inline_strings(tmp_path):
    out = tmp_path / "cells.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.set_cell_texts("Sheet1", {"B3": "structure", "A1": "编号"})
        writer.set_cell_text("Sheet1", "A2", "D0001")
        writer.set_column_width("Sheet1", "A", 12)
        writer.save(out)

    root = _sheet_root(out)
    values = {}
    for row in root.findall("s:sheetData/s:row", NS):
        for cell in row.findall("s:c", NS):
            assert cell.get("t") == "inlineStr"
            values[cell.get("r")] = cell.find("s:is/s:t", NS).text
    assert values == {"A1": "编号", "A2": "D0001", "B3": "structure"}

    # Rows must stay in ascending order or Excel repairs the sheet.
    row_numbers = [int(r.get("r")) for r in root.findall("s:sheetData/s:row", NS)]
    assert row_numbers == sorted(row_numbers)

    dim = root.find("s:dimension", NS)
    assert dim is not None and dim.get("ref") == "A1:B3"

    col = root.find("s:cols/s:col", NS)
    assert col is not None
    assert (col.get("min"), col.get("max"), col.get("width")) == ("1", "1", "12")


def test_set_row_height_marks_custom_height(tmp_path):
    out = tmp_path / "rows.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.set_row_height("Sheet1", 2, 67.5, count=2)
        writer.save(out)

    root = _sheet_root(out)
    heights = {
        int(r.get("r")): (r.get("ht"), r.get("customHeight"))
        for r in root.findall("s:sheetData/s:row", NS)
    }
    assert heights[2] == ("67.5", "1")
    assert heights[3] == ("67.5", "1")


def test_set_cell_texts_unknown_sheet_raises():
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        with pytest.raises(ValueError, match="Sheet not found"):
            writer.set_cell_text("NoSuchSheet", "A1", "x")


def test_render_png_explicit_size_survives_target_aspect(tmp_path):
    pytest.importorskip("rdkit")
    pytest.importorskip("pycdxml")
    from retain_cdx.render.preview import render_png

    png = tmp_path / "preview.png"
    render_png(
        Path("tests/fixtures/test.cdx"), png, width=600, height=400, target_aspect=1.3
    )
    image = Image.open(png)
    # Canvas keeps the requested resolution, trimmed to the target aspect.
    assert image.size == (520, 400)
    # The molecule must fill the canvas, not float tiny inside it
    # (ACS 1996 pins bond length; the drawing is supersampled to fit).
    grayscale = image.convert("L")
    bbox = grayscale.point(lambda v: 255 if v < 245 else 0).getbbox()
    assert bbox is not None
    content_w = bbox[2] - bbox[0]
    content_h = bbox[3] - bbox[1]
    assert max(content_w / 520, content_h / 400) > 0.7
