"""Shell 路径的产物构成。

Shell 路径是三条路径里唯一还走 staging 目录的：解包一份原生 shell 包、
按显式顺序改造、再由 `_parts` 搬进目标表。这些测试锁住**结果**，
好让接下来的搬家只改代码位置、不改产出。

行为断言（居中、锚点一致、去边框等）在 `test_excel_shell_layout.py`；
这里只管「产出哪些部件、内容稳不稳定、编排顺序有没有变」。
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest

from ooxml_checks import check_package
from retain_cdx import ExcelCDXWriter, ExcelShellTemplate

from test_excel_shell_layout import patch_preview_renderers

CDX = "tests/fixtures/test.cdx"
BLANK = "tests/fixtures/blank.xlsx"
SHELL = "tests/fixtures/shell_large.xlsx"

# creationId 每次运行重新生成，比对前归一化掉。
_GUID = re.compile(r"\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}")


def _insert(output: Path, **kwargs) -> zipfile.ZipFile:
    with ExcelCDXWriter(BLANK) as writer:
        writer.insert_cdx(
            CDX, sheet="Sheet1", cell="A1", shell=ExcelShellTemplate(SHELL), **kwargs
        )
        writer.save(output)
    return zipfile.ZipFile(output)


def test_shell_insert_produces_exactly_these_parts(tmp_path, monkeypatch):
    patch_preview_renderers(monkeypatch)
    archive = _insert(tmp_path / "out.xlsx")
    added = set(archive.namelist()) - set(zipfile.ZipFile(BLANK).namelist())
    assert added == {
        "xl/embeddings/sheet1_retain_oleObject1.bin",
        "xl/media/sheet1_retain_image1.png",
        "xl/drawings/sheet1_retain_drawing.xml",
        "xl/drawings/sheet1_retain_vmlDrawing.vml",
        "xl/drawings/_rels/sheet1_retain_vmlDrawing.vml.rels",
        "xl/worksheets/_rels/sheet1.xml.rels",
    }, "shell 路径的产物构成变了"
    assert check_package(tmp_path / "out.xlsx") == []


def test_shell_output_is_deterministic(tmp_path, monkeypatch):
    """同输入两次运行字节一致（GUID 除外）。

    重构期间最有用的一条：任何非预期的产出变化都会在这里暴露。
    """
    patch_preview_renderers(monkeypatch)
    runs = []
    for index in range(2):
        archive = _insert(tmp_path / f"run{index}.xlsx")
        runs.append(
            {
                name: _GUID.sub("{GUID}", archive.read(name).decode("utf-8", "replace"))
                for name in sorted(archive.namelist())
            }
        )
    assert runs[0] == runs[1]


def test_shell_staging_does_not_leak_template_content(tmp_path, monkeypatch):
    """shell 模板自己的行高列宽会被复制过来，但它的单元格内容不该跟过来。"""
    patch_preview_renderers(monkeypatch)
    archive = _insert(tmp_path / "out.xlsx")

    assert "xl/sharedStrings.xml" not in archive.namelist()
    sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    # shell 模板里带的是它自己的示例对象，不应留下第二个 OLE
    assert sheet.count("<oleObject ") == 2, "应恰好一个对象（Choice + Fallback）"
    assert len([n for n in archive.namelist() if "/embeddings/" in n]) == 1


def test_display_box_hugs_the_molecule(tmp_path, monkeypatch):
    """D' tighten：显示框贴合分子自身的宽高比，而不是模板原始比例。

    跳过 tighten 时框会保持 shell 模板的比例，这条断言会红。
    """
    from retain_cdx.structure.geometry import cdx_content_aspect

    patch_preview_renderers(monkeypatch)
    archive = _insert(tmp_path / "out.xlsx")
    vml = archive.read("xl/drawings/sheet1_retain_vmlDrawing.vml").decode()

    size = re.search(r"width:([0-9.]+)pt;\s*height:([0-9.]+)pt", vml)
    assert size, "VML 缺尺寸"
    box_aspect = float(size.group(1)) / float(size.group(2))
    molecule_aspect = cdx_content_aspect(Path(CDX).read_bytes())
    # 像素取整会带来千分之几的偏差
    assert abs(box_aspect - molecule_aspect) < 0.02, (
        f"框比 {box_aspect:.4f} 没有贴合分子比 {molecule_aspect:.4f}"
    )


def test_no_border_survives_on_either_layer(tmp_path, monkeypatch):
    """chrome：VML 与 DrawingML 两侧都不留描边。

    注意这条**不**证明几何改写之后那次 chrome 是必要的——用当前夹具把它去掉，
    产物一字节不变。它是防御性的（见 `_shell/geometry.py` 的注释：
    patch_vml_shape 可能留下 stroked=t），保留，但别假称测到了。
    """
    patch_preview_renderers(monkeypatch)
    archive = _insert(tmp_path / "out.xlsx")

    vml = archive.read("xl/drawings/sheet1_retain_vmlDrawing.vml").decode()
    assert 'stroked="t"' not in vml
    assert "strokecolor" not in vml
    drawing = archive.read("xl/drawings/sheet1_retain_drawing.xml").decode()
    assert "<a:noFill/>" in drawing or "<a:noFill />" in drawing


def test_shell_and_target_layout_coexist(tmp_path, monkeypatch):
    """shell 会复制自己的行高列宽；目标表原有的单元格必须留着。"""
    patch_preview_renderers(monkeypatch)
    prepared = tmp_path / "prepared.xlsx"
    with ExcelCDXWriter(BLANK) as writer:
        writer.set_cell_text("Sheet1", "D5", "keep me")
        writer.save(prepared)

    output = tmp_path / "out.xlsx"
    with ExcelCDXWriter(prepared) as writer:
        writer.insert_cdx(CDX, sheet="Sheet1", cell="A1", shell=ExcelShellTemplate(SHELL))
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        assert "keep me" in archive.read("xl/worksheets/sheet1.xml").decode()
    assert check_package(output) == []


@pytest.mark.parametrize("allow_resize", [True, False])
def test_allow_resize_controls_whether_the_box_moves(tmp_path, monkeypatch, allow_resize):
    """`allow_resize` 决定是否把对象挪到目标单元格（D 步）。"""
    patch_preview_renderers(monkeypatch)
    output = tmp_path / f"resize_{allow_resize}.xlsx"
    with ExcelCDXWriter(BLANK) as writer:
        writer.insert_cdx(
            CDX,
            sheet="Sheet1",
            cell="C5",
            shell=ExcelShellTemplate(SHELL, allow_resize=allow_resize),
        )
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        vml = archive.read("xl/drawings/sheet1_retain_vmlDrawing.vml").decode()
    anchor = re.search(r"<x:Anchor>([^<]+)</x:Anchor>", vml)
    assert anchor, "VML 缺锚点"
    from_col = int(anchor.group(1).split(",")[0])
    if allow_resize:
        assert from_col == 2, f"应挪到 C 列（索引 2），实际 {from_col}"
    else:
        assert from_col != 2, "allow_resize=False 时不该挪动"
    assert check_package(output) == []
