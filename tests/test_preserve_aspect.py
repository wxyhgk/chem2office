"""显示框的比例必须等于 CDX 内容的比例，否则对象在别处会被拉伸。

Excel 里看不出来——预览图已经 letterbox 成了框的比例。一旦复制到 PowerPoint，
目标程序按 CDX 重新渲染、填满框，结构就变形了。用户报的就是这个：
「复制到 ppt 里面后……会有一些形变」。

原来的实现只算 `sources[0].content_aspect` 然后套给整批，所以第一个分子是对的、
后面全错——恰好是最难发现的那种错法。
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

import olefile
import pytest

from retain_cdx import ExcelCDXWriter, StructureSource
from retain_cdx.office.excel._insert.plan import _fit_box
from retain_cdx.structure.geometry import cdx_content_aspect

from test_excel_shell_layout import patch_preview_renderers

BLANK = "tests/fixtures/blank.xlsx"
# 形状差别明显的三个：细长、方、扁宽
SMILES = [
    ("CN1C=NC2=C1C(=O)N(C)C(=O)N2C", "咖啡因"),
    ("C1CCCCC1", "环己烷"),
    ("c1ccc(-c2ccccc2)cc1", "联苯"),
]


def _boxes_and_aspects(path: Path) -> list[tuple[float, float]]:
    """返回每个对象的 (框比例, CDX 内容比例)。"""
    with zipfile.ZipFile(path) as archive:
        vml = "".join(
            archive.read(n).decode() for n in archive.namelist() if n.endswith(".vml")
        )
        sizes = re.findall(r"width:([\d.]+)pt;height:([\d.]+)pt", vml)
        embeds = sorted(n for n in archive.namelist() if "/embeddings/" in n)
        out = []
        for index, name in enumerate(embeds):
            cdx = olefile.OleFileIO(io.BytesIO(archive.read(name))).openstream(
                "CONTENTS"
            ).read()
            width, height = (float(v) for v in sizes[index])
            out.append((width / height, cdx_content_aspect(cdx)))
    return out


def _insert(path: Path, **kwargs) -> None:
    with ExcelCDXWriter(BLANK) as writer:
        writer.insert_cdx_many(
            [StructureSource.from_smiles(s, name=n) for s, n in SMILES],
            sheet="Sheet1",
            start_cell="B2",
            **kwargs,
        )
        writer.save(path)


def test_every_object_gets_its_own_shape(tmp_path, monkeypatch):
    """整批共用一个比例是原来的 bug：只有第一个分子不变形。"""
    patch_preview_renderers(monkeypatch)
    output = tmp_path / "out.xlsx"
    _insert(output)

    pairs = _boxes_and_aspects(output)
    assert len(pairs) == len(SMILES)
    for index, (box, content) in enumerate(pairs):
        assert box == pytest.approx(content, rel=0.02), (
            f"第 {index + 1} 个分子被拉伸 {box / content:.2f}×"
        )

    shapes = {round(box, 2) for box, _ in pairs}
    assert len(shapes) == len(SMILES), f"三个形状不同的分子拿到了相同的框: {shapes}"


def test_boxes_stay_inside_the_requested_size(tmp_path, monkeypatch):
    """width_px/height_px 是上限，框只在里面缩放。"""
    patch_preview_renderers(monkeypatch)
    output = tmp_path / "out.xlsx"
    _insert(output, width_px=145, height_px=100)

    with zipfile.ZipFile(output) as archive:
        vml = "".join(
            archive.read(n).decode() for n in archive.namelist() if n.endswith(".vml")
        )
    for width, height in re.findall(r"width:([\d.]+)pt;height:([\d.]+)pt", vml):
        assert float(width) <= 145 * 0.75 + 0.01
        assert float(height) <= 100 * 0.75 + 0.01


def test_row_height_is_left_to_the_caller(tmp_path, monkeypatch):
    """保比例不得顺手改行高——那是替调用方改排版。"""
    patch_preview_renderers(monkeypatch)
    output = tmp_path / "out.xlsx"
    _insert(output, row_height=42.5)

    with zipfile.ZipFile(output) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    # \b 是必要的：贪婪匹配会滑到 customHeight="1" 的尾巴上去
    heights = re.findall(r'<row [^>]*\bht="([\d.]+)"', sheet)
    assert heights, "没有写出行高"
    assert set(heights) == {"42.5"}, f"行高被改了: {set(heights)}"


def test_opting_out_restores_the_literal_box(tmp_path, monkeypatch):
    """preserve_aspect=False 时框就是 width_px×height_px，全体一致。"""
    patch_preview_renderers(monkeypatch)
    output = tmp_path / "out.xlsx"
    _insert(output, width_px=145, height_px=100, preserve_aspect=False)

    with zipfile.ZipFile(output) as archive:
        vml = "".join(
            archive.read(n).decode() for n in archive.namelist() if n.endswith(".vml")
        )
    sizes = set(re.findall(r"width:([\d.]+)pt;height:([\d.]+)pt", vml))
    assert sizes == {("108.75", "75")}


@pytest.mark.parametrize(
    "aspect,expected",
    [
        (1.45, (145, 100)),  # 正好等于框比例
        (1.0, (100, 100)),  # 高受限，宽收窄
        (0.5, (50, 100)),  # 细长
        (4.0, (145, 36)),  # 宽受限，高压扁
        (None, (145, 100)),  # 比例不可知时原样返回
    ],
)
def test_fit_box(aspect, expected):
    assert _fit_box(aspect, 145, 100) == expected
