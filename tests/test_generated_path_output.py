"""Generated 路径的产物构成——重构时的字节级护栏。

generated 路径已改为直写目标包（`_insert_ole`），
这些测试锁住「产出什么部件、里面有什么」，好让重构只改路径不改结果。
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest

from retain_cdx import ExcelCDXWriter, ExcelPlacement

from test_excel_shell_layout import patch_preview_renderers

CDX = "tests/fixtures/test.cdx"
BLANK = "tests/fixtures/blank.xlsx"

# creationId 每次运行重新生成，比对前要归一化掉。
_GUID = re.compile(r"\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}")


def _insert(output: Path, **kwargs) -> zipfile.ZipFile:
    with ExcelCDXWriter(BLANK) as writer:
        # 显式钉住 png：默认的 "auto" 取决于机器上有没有 Inkscape，
        # 而这个测试比对的是产物构成，必须确定。
        kwargs.setdefault("preview_format", "png")
        writer.insert_cdx(CDX, sheet="Sheet1", **kwargs)
        writer.save(output)
    return zipfile.ZipFile(output)


def test_generated_insert_produces_exactly_these_parts(tmp_path, monkeypatch):
    patch_preview_renderers(monkeypatch)
    archive = _insert(tmp_path / "out.xlsx", cell="B2")
    added = set(archive.namelist()) - set(zipfile.ZipFile(BLANK).namelist())
    assert added == {
        "xl/embeddings/sheet1_retain_oleObject1.bin",
        "xl/media/sheet1_retain_image1.png",
        "xl/drawings/sheet1_retain_drawing.xml",
        "xl/drawings/sheet1_retain_vmlDrawing.vml",
        "xl/drawings/_rels/sheet1_retain_vmlDrawing.vml.rels",
        "xl/worksheets/_rels/sheet1.xml.rels",
    }, "generated 路径的产物构成变了"


def test_staging_workbook_contents_do_not_leak(tmp_path, monkeypatch):
    """临时表里的单元格/共享字符串不该出现在产物里。

    Merge 只搬 bin、media、drawing、vml、``<oleObjects>``；曾经往临时表写
    文件名标签、sharedStrings 和行高，全被丢弃——纯粹的空转。
    """
    patch_preview_renderers(monkeypatch)
    archive = _insert(tmp_path / "out.xlsx", cell="B2")

    assert "xl/sharedStrings.xml" not in archive.namelist()
    sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    assert "test" not in sheet, "临时表的文件名标签泄漏进了产物"
    assert "<c " not in sheet, "generated 路径不应写入任何单元格"


def test_target_own_cells_survive(tmp_path, monkeypatch):
    """插入不能动目标表原有的内容。"""
    patch_preview_renderers(monkeypatch)
    prepared = tmp_path / "prepared.xlsx"
    with ExcelCDXWriter(BLANK) as writer:
        writer.set_cell_text("Sheet1", "A1", "keep me")
        writer.save(prepared)

    out = tmp_path / "out.xlsx"
    with ExcelCDXWriter(prepared) as writer:
        writer.insert_cdx(CDX, sheet="Sheet1", cell="B2")
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        assert "keep me" in archive.read("xl/worksheets/sheet1.xml").decode()


@pytest.mark.parametrize(
    "kwargs",
    [
        {"cell": "B2"},
        {"cell": "C3", "placement": ExcelPlacement("C3", offset_x="2mm")},
        {"cell": "B2", "width_px": 200, "height_px": 150},
        {"cell": "B2", "placement": ExcelPlacement("B2", to_cell="E8")},
    ],
    ids=["plain", "offset", "sized", "two-point"],
)
def test_output_is_deterministic_apart_from_creation_ids(tmp_path, monkeypatch, kwargs):
    """同样输入两次运行必须产出相同字节（GUID 除外）。

    重构期间这是最有用的一条：任何非预期的输出变化都会在这里暴露。
    """
    patch_preview_renderers(monkeypatch)
    digests = []
    for run in range(2):
        archive = _insert(tmp_path / f"run{run}.xlsx", **kwargs)
        digests.append(
            {
                name: _GUID.sub("{GUID}", archive.read(name).decode("utf-8", "replace"))
                for name in sorted(archive.namelist())
            }
        )
    assert digests[0] == digests[1]
