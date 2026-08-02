from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

import pytest

from retain_cdx import ExcelDisplayOptions, ExcelCDXWriter, ExcelPlacement, PreviewOptions
from retain_cdx.office.excel._ooxml.xml_io import (
    NS,
    NS_A,
    NS_R,
    NS_XDR,
    WORKSHEET_CHILD_ORDER,
    local_name,
)

from test_excel_shell_layout import patch_preview_renderers

_RANK = {name: idx for idx, name in enumerate(WORKSHEET_CHILD_ORDER)}
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _insert_pictures(output: Path, monkeypatch, *, count: int = 3, **kwargs):
    patch_preview_renderers(monkeypatch)
    sources = ["tests/fixtures/test.cdx"] * count
    # 钉住 png：默认 "auto" 随机器上有无 Inkscape 而变，这些测试断言部件名和内容类型。
    kwargs.setdefault(
        "display", ExcelDisplayOptions(preview=PreviewOptions(format="png"))
    )
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_pictures_many(
            sources, sheet="Sheet1", start_cell="B2", **kwargs
        )
        writer.save(output)
    return zipfile.ZipFile(output)


def _drawing_part(archive: zipfile.ZipFile) -> str:
    names = [
        n
        for n in archive.namelist()
        if n.startswith("xl/drawings/") and n.endswith(".xml")
    ]
    assert len(names) == 1, names
    return names[0]


def test_picture_mode_writes_no_ole_parts(tmp_path, monkeypatch):
    archive = _insert_pictures(tmp_path / "pics.xlsx", monkeypatch)
    names = archive.namelist()

    assert not [n for n in names if "/embeddings/" in n]
    assert not [n for n in names if n.lower().endswith(".vml")]
    media = [n for n in names if "/media/" in n]
    assert len(media) == 3

    sheet = archive.read("xl/worksheets/sheet1.xml").decode()
    assert "<oleObjects" not in sheet
    assert "legacyDrawing" not in sheet
    assert "<drawing " in sheet


def test_picture_anchors_and_rels_are_consistent(tmp_path, monkeypatch):
    archive = _insert_pictures(tmp_path / "pics.xlsx", monkeypatch)
    drawing_part = _drawing_part(archive)
    root = ET.fromstring(archive.read(drawing_part))

    anchors = root.findall(f"{{{NS_XDR}}}twoCellAnchor")
    assert len(anchors) == 3

    rels = ET.fromstring(
        archive.read(f"xl/drawings/_rels/{Path(drawing_part).name}.rels")
    )
    rel_targets = {
        rel.get("Id"): rel.get("Target")
        for rel in rels.findall(f"{{{_REL_NS}}}Relationship")
    }

    shape_ids, embeds, rows = [], [], []
    for anchor in anchors:
        assert anchor.get("editAs") == "oneCell"
        pic = anchor.find(f"{{{NS_XDR}}}pic")
        assert pic is not None
        shape_ids.append(pic.find(f".//{{{NS_XDR}}}cNvPr").get("id"))
        blip = pic.find(f".//{{{NS_A}}}blip")
        embeds.append(blip.get(f"{{{NS_R}}}embed"))
        rows.append(int(anchor.find(f"{{{NS_XDR}}}from/{{{NS_XDR}}}row").text))

    assert len(set(shape_ids)) == 3, "shape ids must be unique"
    assert len(set(embeds)) == 3, "each picture needs its own image rel"
    for embed in embeds:
        assert embed in rel_targets, f"{embed} missing from drawing rels"
        assert rel_targets[embed].startswith("../media/")
    # One structure per row, starting at B2 (0-based row 1).
    assert rows == [1, 2, 3]


def test_picture_mode_sheet_order_is_legal(tmp_path, monkeypatch):
    archive = _insert_pictures(tmp_path / "pics.xlsx", monkeypatch)
    kids = [local_name(c.tag) for c in ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))]
    ranks = [_RANK.get(k, -1) for k in kids]
    assert [
        (kids[i], kids[i + 1])
        for i in range(len(kids) - 1)
        if ranks[i] > ranks[i + 1] >= 0
    ] == []


def test_picture_content_types_declare_media_and_drawing(tmp_path, monkeypatch):
    archive = _insert_pictures(tmp_path / "pics.xlsx", monkeypatch)
    content_types = archive.read("[Content_Types].xml").decode()
    drawing_part = _drawing_part(archive)
    assert 'Extension="png"' in content_types
    assert f'PartName="/{drawing_part}"' in content_types


def test_second_picture_insert_reuses_drawing_part(tmp_path, monkeypatch):
    """A second call must append to the sheet's drawing, not add another one."""
    patch_preview_renderers(monkeypatch)
    output = tmp_path / "twice.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_pictures_many(
            ["tests/fixtures/test.cdx"], sheet="Sheet1", start_cell="B2"
        )
        writer.insert_pictures_many(
            ["tests/fixtures/test.cdx"], sheet="Sheet1", start_cell="B5"
        )
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        drawing_part = _drawing_part(archive)  # asserts exactly one drawing part
        root = ET.fromstring(archive.read(drawing_part))
        assert len(root.findall(f"{{{NS_XDR}}}twoCellAnchor")) == 2
        assert len([n for n in archive.namelist() if "/media/" in n]) == 2
        sheet_root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        assert len(sheet_root.findall("s:drawing", NS)) == 1


def test_picture_mode_honours_display_box(tmp_path, monkeypatch):
    archive = _insert_pictures(
        tmp_path / "pics.xlsx",
        monkeypatch,
        count=1,
        display=ExcelDisplayOptions(width=200, height=150, row_height=120, col_width=30),
        placement=ExcelPlacement("B2", offset_x=8, offset_y=6),
    )
    root = ET.fromstring(archive.read(_drawing_part(archive)))
    ext = root.find(f".//{{{NS_A}}}ext")
    emu_per_px = 9525
    assert int(ext.get("cx")) == 200 * emu_per_px
    assert int(ext.get("cy")) == 150 * emu_per_px
    corner = root.find(f".//{{{NS_XDR}}}from")
    assert int(corner.find(f"{{{NS_XDR}}}colOff").text) == 8 * emu_per_px
    assert int(corner.find(f"{{{NS_XDR}}}rowOff").text) == 6 * emu_per_px


def test_picture_mode_empty_sources_is_noop(tmp_path, monkeypatch):
    patch_preview_renderers(monkeypatch)
    output = tmp_path / "empty.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_pictures_many([], sheet="Sheet1", start_cell="B2")
        writer.save(output)
    with zipfile.ZipFile(output) as archive:
        assert not [n for n in archive.namelist() if "/drawings/" in n]


def test_picture_mode_unknown_sheet_raises(monkeypatch):
    patch_preview_renderers(monkeypatch)
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        with pytest.raises(ValueError, match="Sheet not found"):
            writer.insert_pictures_many(
                ["tests/fixtures/test.cdx"], sheet="Nope", start_cell="A1"
            )
