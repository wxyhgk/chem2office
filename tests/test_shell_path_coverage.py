"""Shell 路径里测试套件原本走不到的两个分支。

一次 sys.settrace 调查显示：`fit_cdx=True` 和「目标表已有 drawing/VML」这两条
分支是活代码但零覆盖。删除周边死代码之前先把它们钉住，否则动刀时没有护栏。
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

import olefile

from ooxml_checks import check_package
from retain_cdx import ExcelCDXWriter, ExcelShellTemplate

from test_excel_shell_layout import patch_preview_renderers

CDX = "tests/fixtures/test.cdx"
BLANK = "tests/fixtures/blank.xlsx"
SHELL = "tests/fixtures/shell_large.xlsx"


def _contents(archive: zipfile.ZipFile, name: str) -> bytes:
    return olefile.OleFileIO(io.BytesIO(archive.read(name))).openstream("CONTENTS").read()


def test_fit_cdx_normalises_the_structure_to_the_shell_box(tmp_path, monkeypatch):
    """`fit_cdx=True` 会把 CDX 重排成 shell 显示框的比例。

    这是公开 flag `ExcelShellTemplate(fit_cdx=True)` 的唯一实现
    (`_shell/staging._fit_cdx_to_shell`)，此前无人测。
    """
    patch_preview_renderers(monkeypatch)
    plain, fitted = tmp_path / "plain.xlsx", tmp_path / "fitted.xlsx"

    for output, fit in ((plain, False), (fitted, True)):
        with ExcelCDXWriter(BLANK) as writer:
            writer.insert_cdx(
                CDX,
                sheet="Sheet1",
                cell="A1",
                shell=ExcelShellTemplate(SHELL, fit_cdx=fit),
            )
            writer.save(output)

    original = Path(CDX).read_bytes()
    with zipfile.ZipFile(plain) as archive:
        name = [n for n in archive.namelist() if "/embeddings/" in n][0]
        assert _contents(archive, name) == original, "不开 fit_cdx 时应原样注入"

    with zipfile.ZipFile(fitted) as archive:
        name = [n for n in archive.namelist() if "/embeddings/" in n][0]
        payload = _contents(archive, name)
    assert payload[:8] == b"VjCD0100", "重排后仍须是合法 CDX"
    assert payload != original, "fit_cdx 应当改写 CDX 的排布"
    assert check_package(fitted) == []


def test_shell_insert_into_a_sheet_that_already_has_a_drawing(tmp_path, monkeypatch):
    """目标表已有 drawing/VML 时，shell 插入要合并而不是覆盖。

    走的是 `_parts/surface.py` 里 append/remap 那一组函数——它们修过一个真实
    缺陷（研究报告 C7：shapeId 只重映射了一层，导致 OLE 绑定悬空），但常规
    用例走不到。
    """
    patch_preview_renderers(monkeypatch)

    seeded = tmp_path / "seeded.xlsx"
    with ExcelCDXWriter(BLANK) as writer:
        writer.insert_pictures_many([CDX] * 2, sheet="Sheet1", start_cell="D2")
        writer.save(seeded)
    with zipfile.ZipFile(seeded) as archive:
        assert [n for n in archive.namelist() if "/drawings/" in n], "前置条件：应已有 drawing"

    output = tmp_path / "combined.xlsx"
    with ExcelCDXWriter(seeded) as writer:
        writer.insert_cdx(CDX, sheet="Sheet1", cell="A1", shell=ExcelShellTemplate(SHELL))
        writer.save(output)

    assert check_package(output) == []
    with zipfile.ZipFile(output) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode()
        drawings = [
            archive.read(n).decode()
            for n in archive.namelist()
            if n.startswith("xl/drawings/") and n.endswith(".xml")
        ]
        vml = "".join(
            archive.read(n).decode()
            for n in archive.namelist()
            if n.endswith(".vml")
        )

    ole_shape_ids = set(re.findall(r'<oleObject[^>]*shapeId="(\d+)"', sheet))
    assert ole_shape_ids, "shell 对象应出现在 worksheet 上"
    # OLE 声明的 shapeId 必须在 VML 里找得到对应形状，否则绑定悬空
    vml_ids = set(re.findall(r'id="_x0000_s(\d+)"', vml))
    assert ole_shape_ids <= vml_ids, f"OLE {ole_shape_ids} 在 VML {vml_ids} 里没有对应形状"

    # 原有的两张图片必须还在
    pictures = sum(d.count("<xdr:pic>") for d in drawings)
    assert pictures == 2, f"目标表原有的图片被破坏了（剩 {pictures} 张）"


def test_shell_insert_after_generated_keeps_both(tmp_path, monkeypatch):
    """先 generated 再 shell：两种对象共存，shapeId 不得冲突。"""
    patch_preview_renderers(monkeypatch)

    seeded = tmp_path / "seeded.xlsx"
    with ExcelCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, sheet="Sheet1", cell="D2")
        writer.save(seeded)

    output = tmp_path / "both.xlsx"
    with ExcelCDXWriter(seeded) as writer:
        writer.insert_cdx(CDX, sheet="Sheet1", cell="A1", shell=ExcelShellTemplate(SHELL))
        writer.save(output)

    assert check_package(output) == []
    with zipfile.ZipFile(output) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode()
        embeddings = len([n for n in archive.namelist() if "/embeddings/" in n])
    assert embeddings == 2

    # 每个对象在 mc:Choice 与 mc:Fallback 各出现一次，所以按去重后的数量比。
    ids = re.findall(r'<oleObject[^>]*shapeId="(\d+)"', sheet)
    assert len(set(ids)) == embeddings, f"shapeId 冲突: {ids}"
    assert len(ids) == embeddings * 2, f"Choice/Fallback 不成对: {ids}"
