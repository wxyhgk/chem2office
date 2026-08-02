"""单元格对齐：让标签与旁边的结构在垂直方向对齐。

为结构留出的行高达 110pt，而 Excel 默认底端对齐，所以名称会沉到格子底部、
明显低于分子。对齐必须走 `cellXfs` 里的样式记录，不能写在单元格上。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile

import pytest

from ooxml_checks import check_package
from retain_cdx import ExcelCDXWriter, ExcelCellAlignment
from retain_cdx.office.excel._ooxml.xml_io import NS

BLANK = "tests/fixtures/blank.xlsx"


def _styles(archive: zipfile.ZipFile) -> ET.Element:
    return ET.fromstring(archive.read("xl/styles.xml"))


def _cell(archive: zipfile.ZipFile, ref: str) -> ET.Element:
    root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    cell = root.find(f'.//s:c[@r="{ref}"]', NS)
    assert cell is not None, f"{ref} 不存在"
    return cell


def test_alignment_registers_a_style_and_points_the_cell_at_it(tmp_path):
    output = tmp_path / "aligned.xlsx"
    with ExcelCDXWriter(BLANK) as writer:
        writer.set_cell_text(
            "Sheet1", "A2", "苯", align=ExcelCellAlignment(vertical="center")
        )
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        index = _cell(archive, "A2").get("s")
        assert index is not None, "单元格没有指向样式"
        xfs = _styles(archive).findall("s:cellXfs/s:xf", NS)
        alignment = xfs[int(index)].find("s:alignment", NS)
        assert alignment is not None and alignment.get("vertical") == "center"
    assert check_package(output) == []


def test_same_alignment_is_registered_once(tmp_path):
    """几千个结构不该长出几千条 xf。"""
    output = tmp_path / "many.xlsx"
    align = ExcelCellAlignment(vertical="center")
    with ExcelCDXWriter(BLANK) as writer:
        writer.set_cell_texts("Sheet1", {f"A{n}": str(n) for n in range(2, 30)}, align=align)
        writer.set_cell_text("Sheet1", "B2", "again", align=align)
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        xfs = _styles(archive).findall("s:cellXfs/s:xf", NS)
        aligned = [xf for xf in xfs if xf.find("s:alignment", NS) is not None]
        assert len(aligned) == 1, f"重复登记了 {len(aligned)} 条对齐样式"
        assert _cell(archive, "A5").get("s") == _cell(archive, "B2").get("s")


def test_different_alignments_get_separate_styles(tmp_path):
    output = tmp_path / "two.xlsx"
    with ExcelCDXWriter(BLANK) as writer:
        writer.set_cell_text("Sheet1", "A1", "中", align=ExcelCellAlignment(vertical="center"))
        writer.set_cell_text(
            "Sheet1", "A2", "中中", align=ExcelCellAlignment(vertical="center", horizontal="center")
        )
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        assert _cell(archive, "A1").get("s") != _cell(archive, "A2").get("s")


def test_no_alignment_keeps_the_cell_unstyled(tmp_path):
    output = tmp_path / "plain.xlsx"
    with ExcelCDXWriter(BLANK) as writer:
        writer.set_cell_text("Sheet1", "A1", "plain")
        writer.save(output)
    with zipfile.ZipFile(output) as archive:
        assert _cell(archive, "A1").get("s") is None


def test_alignment_overrides_an_inherited_style(tmp_path):
    """显式给了对齐就用它；不给才沿用被覆盖单元格原有的样式。"""
    seeded = tmp_path / "seeded.xlsx"
    with ExcelCDXWriter(BLANK) as writer:
        writer.set_cell_text("Sheet1", "A1", "旧", align=ExcelCellAlignment(horizontal="right"))
        writer.save(seeded)

    output = tmp_path / "out.xlsx"
    with ExcelCDXWriter(seeded) as writer:
        writer.set_cell_text("Sheet1", "A1", "新", align=ExcelCellAlignment(vertical="center"))
        writer.set_cell_text("Sheet1", "A2", "无对齐")
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        xfs = _styles(archive).findall("s:cellXfs/s:xf", NS)
        alignment = xfs[int(_cell(archive, "A1").get("s"))].find("s:alignment", NS)
        assert alignment.get("vertical") == "center"
        assert alignment.get("horizontal") is None


@pytest.mark.parametrize("kwargs", [{}, {"wrap_text": False}])
def test_empty_alignment_is_a_no_op(tmp_path, kwargs):
    output = tmp_path / "empty.xlsx"
    with ExcelCDXWriter(BLANK) as writer:
        writer.set_cell_text("Sheet1", "A1", "x", align=ExcelCellAlignment(**kwargs))
        writer.save(output)
    with zipfile.ZipFile(output) as archive:
        assert _cell(archive, "A1").get("s") in (None, "0")
