from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import olefile
import pytest

from ooxml_checks import check_package
from retain_cdx import (
    PowerPointCDXWriter,
    PreviewOptions,
    SlideGrid,
    SlidePlacement,
    StructureSource,
)
from retain_cdx.office.powerpoint._pml import NS_A, NS_P, NS_R

NS = {"p": NS_P, "a": NS_A, "r": NS_R}
NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
CDX = "tests/fixtures/test.cdx"
BLANK = "tests/fixtures/blank.pptx"


@pytest.fixture
def png_preview(monkeypatch):
    """Render tiny PNGs instead of shelling out to Inkscape for EMF."""
    from PIL import Image

    import retain_cdx.office.powerpoint.writer as writer_module

    def fake_render(cdx_path, output_path, **kwargs):
        Image.new("RGB", (32, 24), "white").save(output_path)
        return "png"

    monkeypatch.setattr(writer_module, "render_preview", fake_render)
    monkeypatch.setattr(writer_module, "preview_extension", lambda fmt: "png")
    return PreviewOptions(format="png")


def _slide_root(archive: zipfile.ZipFile) -> ET.Element:
    return ET.fromstring(archive.read("ppt/slides/slide1.xml"))


def _frames(archive: zipfile.ZipFile) -> list[ET.Element]:
    return _slide_root(archive).findall(".//" + f"{{{NS_P}}}graphicFrame")


def test_insert_one_structure_builds_a_valid_package(tmp_path, png_preview):
    out = tmp_path / "one.pptx"
    with PowerPointCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, preview=png_preview)
        writer.save(out)

    assert check_package(out) == []
    with zipfile.ZipFile(out) as archive:
        assert len(_frames(archive)) == 1
        assert len([n for n in archive.namelist() if "/embeddings/" in n]) == 1
        assert len([n for n in archive.namelist() if "/media/" in n]) == 1


def test_inserted_container_carries_our_cdx(tmp_path, png_preview):
    out = tmp_path / "clean.pptx"
    with PowerPointCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, preview=png_preview)
        writer.save(out)

    payload = Path(CDX).read_bytes()
    with zipfile.ZipFile(out) as archive:
        bins = [n for n in archive.namelist() if "/embeddings/" in n]
        assert len(bins) == 1
        contents = olefile.OleFileIO(
            io.BytesIO(archive.read(bins[0]))
        ).openstream("CONTENTS").read()
        assert contents == payload


def test_insert_into_deck_that_already_has_ole_appends(tmp_path, png_preview):
    """A deck already holding ChemDraw objects keeps them; ours are added."""
    from retain_cdx.resources import bundled_template_pptx

    out = tmp_path / "appended_to_existing.pptx"
    with bundled_template_pptx() as existing:
        with zipfile.ZipFile(existing) as archive:
            before = len([n for n in archive.namelist() if "/embeddings/" in n])
        with PowerPointCDXWriter(existing) as writer:
            writer.insert_cdx(CDX, preview=png_preview)
            writer.save(out)

    assert before == 1, "sample deck should ship exactly one ChemDraw object"
    with zipfile.ZipFile(out) as archive:
        assert len([n for n in archive.namelist() if "/embeddings/" in n]) == 2
        assert len(_frames(archive)) == 2
    assert check_package(out) == []


def test_mc_requires_prefix_stays_declared(tmp_path, png_preview):
    """Regression: ElementTree drops the unused xmlns:v that Requires="v" needs.

    Losing it makes PowerPoint offer to repair the deck.
    """
    out = tmp_path / "mce.pptx"
    with PowerPointCDXWriter(BLANK) as writer:
        writer.insert_cdx_many([CDX, CDX], preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        slide = archive.read("ppt/slides/slide1.xml").decode()
    assert 'Requires="v"' in slide
    assert 'xmlns:v="urn:schemas-microsoft-com:vml"' in slide
    assert check_package(out) == []


def test_shape_ids_are_unique_across_frames(tmp_path, png_preview):
    """Regression: the cloned fallback picture kept cNvPr id="0" every time."""
    out = tmp_path / "ids.pptx"
    with PowerPointCDXWriter(BLANK) as writer:
        writer.insert_cdx_many([CDX] * 4, preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        ids = re.findall(
            r'<p:cNvPr id="(\d+)"', archive.read("ppt/slides/slide1.xml").decode()
        )
    assert len(ids) == len(set(ids)), f"duplicate shape ids: {ids}"


def test_each_structure_gets_its_own_parts_and_rels(tmp_path, png_preview):
    out = tmp_path / "many.pptx"
    with PowerPointCDXWriter(BLANK) as writer:
        writer.insert_cdx_many([CDX] * 3, preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        assert len(_frames(archive)) == 3
        assert len([n for n in archive.namelist() if "/embeddings/" in n]) == 3
        assert len([n for n in archive.namelist() if "/media/" in n]) == 3

        rels = ET.fromstring(archive.read("ppt/slides/_rels/slide1.xml.rels"))
        ids = [r.get("Id") for r in rels.findall(f"{{{NS_RELS}}}Relationship")]
        assert len(ids) == len(set(ids))

        embeds = {
            blip.get(f"{{{NS_R}}}embed")
            for blip in _slide_root(archive).iter(f"{{{NS_A}}}blip")
        }
        ole_ids = {
            obj.get(f"{{{NS_R}}}id")
            for obj in _slide_root(archive).iter(f"{{{NS_P}}}oleObj")
        }
        assert len(embeds) == 3 and len(ole_ids) == 3
        assert not (embeds & ole_ids), "image and OLE rels must not collide"
    assert check_package(out) == []


def test_grid_places_structures_in_rows(tmp_path, png_preview):
    out = tmp_path / "grid.pptx"
    grid = SlideGrid(
        columns=2, x="1cm", y="1cm", cell_width="4cm", cell_height="3cm",
        gap_x="0cm", gap_y="0cm",
    )
    with PowerPointCDXWriter(BLANK) as writer:
        writer.insert_cdx_many([CDX] * 4, grid=grid, preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        offsets = [
            (int(off.get("x")), int(off.get("y")))
            for frame in _frames(archive)
            for xfrm in [frame.find(f"{{{NS_P}}}xfrm")]
            for off in [xfrm.find(f"{{{NS_A}}}off")]
        ]
    xs = sorted({x for x, _ in offsets})
    ys = sorted({y for _, y in offsets})
    assert len(xs) == 2 and len(ys) == 2, f"expected a 2x2 grid, got {offsets}"


def test_placement_units_and_aspect_fill(tmp_path, png_preview):
    out = tmp_path / "place.pptx"
    with PowerPointCDXWriter(BLANK) as writer:
        writer.insert_cdx(
            CDX,
            placement=SlidePlacement(x="2cm", y="1cm", width="6cm"),
            preview=png_preview,
        )
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        frame = _frames(archive)[0]
        xfrm = frame.find(f"{{{NS_P}}}xfrm")
        off, ext = xfrm.find(f"{{{NS_A}}}off"), xfrm.find(f"{{{NS_A}}}ext")
        assert int(off.get("x")) == 2 * 360000
        assert int(off.get("y")) == 1 * 360000
        assert int(ext.get("cx")) == 6 * 360000
        # height derived from the molecule's aspect, not left at the template's
        assert int(ext.get("cy")) > 0
        # The fallback picture must track the frame box.
        pic_ext = frame.find(f".//{{{NS_P}}}pic//{{{NS_A}}}ext")
        assert pic_ext.get("cx") == ext.get("cx")
        assert pic_ext.get("cy") == ext.get("cy")


def test_insert_into_user_presentation_keeps_existing_slides(tmp_path, png_preview):
    base = tmp_path / "base.pptx"
    with PowerPointCDXWriter(BLANK) as writer:
        writer.save(base)  # a deck the user might hand us

    out = tmp_path / "appended.pptx"
    with PowerPointCDXWriter(base) as writer:
        assert writer.slide_count == 1
        writer.insert_cdx(CDX, slide=1, preview=png_preview)
        writer.save(out)

    assert check_package(out) == []
    with zipfile.ZipFile(out) as archive:
        assert len(_frames(archive)) == 1


def test_slide_out_of_range_raises(png_preview):
    with PowerPointCDXWriter(BLANK) as writer:
        with pytest.raises(ValueError, match="out of range"):
            writer.insert_cdx(CDX, slide=9, preview=png_preview)


def test_empty_sources_is_noop(tmp_path, png_preview):
    out = tmp_path / "empty.pptx"
    with PowerPointCDXWriter(BLANK) as writer:
        writer.insert_cdx_many([], preview=png_preview)
        writer.save(out)
    with zipfile.ZipFile(out) as archive:
        assert _frames(archive) == []


def test_bundled_template_itself_is_valid():
    """Guards the shipped template against a bad re-export."""
    from retain_cdx.resources import bundled_template_pptx

    with bundled_template_pptx() as template:
        assert check_package(template) == []


def test_structure_from_smiles_roundtrips_into_the_deck(tmp_path, png_preview):
    pytest.importorskip("rdkit")
    pytest.importorskip("pycdxml")
    out = tmp_path / "smiles.pptx"
    source = StructureSource.from_smiles("c1ccccc1", name="benzene")
    with PowerPointCDXWriter(BLANK) as writer:
        writer.insert_cdx(source, preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        name = [n for n in archive.namelist() if "/embeddings/" in n][0]
        contents = olefile.OleFileIO(
            io.BytesIO(archive.read(name))
        ).openstream("CONTENTS").read()
    assert contents[:8] == b"VjCD0100"
    assert contents == source.cdx_data
