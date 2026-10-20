"""Structura compound-table packages → CDXML structures → patent-style Word tables."""

from __future__ import annotations

import io
import json
import re
import shutil
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import replace
from pathlib import Path

import pytest

from ooxml_checks import check_package
from retain_cdx import CompoundTable, PreviewOptions, StructureSource, WordCDXWriter
from retain_cdx.office.word._wml import NS_O, NS_R, NS_V, NS_W

NS = {"w": NS_W, "v": NS_V, "o": NS_O, "r": NS_R}
PACKAGE = Path("tests/fixtures/structura_table")
BLANK = "tests/fixtures/blank.docx"


# -- StructureSource.from_cdxml -------------------------------------------------


def _cdx_as_cdxml(source: StructureSource) -> ET.Element:
    from pycdxml.cdxml_converter.chemdraw_io import read_cdx

    return ET.fromstring(read_cdx(io.BytesIO(source.cdx_data)).to_cdxml().encode())


def test_from_cdxml_keeps_atoms_bonds_coordinates_and_labels():
    """The drawing survives as authored — the reason the hand-over is CDXML."""
    path = PACKAGE / "cdxml" / "C-1.cdxml"
    original = ET.parse(path).getroot()
    converted = _cdx_as_cdxml(StructureSource.from_cdxml(path))

    def atoms(root):
        return [
            (n.get("id"), [float(v) for v in n.get("p").split()], n.get("Element"), n.get("Charge"))
            for n in root.iter("n")
        ]

    def bonds(root):
        return [(b.get("B"), b.get("E"), b.get("Order")) for b in root.iter("b")]

    def labels(root):
        # pycdxml spells out the default face="0"; superscripts (64) must survive.
        return [(s.text, s.get("face") or "0") for s in root.iter("s")]

    assert atoms(converted) == atoms(original)
    assert bonds(converted) == bonds(original)
    assert labels(converted) == labels(original)
    assert ("+", "64") in labels(converted)  # the Na⁺ charge sign
    assert len(converted.findall(".//fragment")) == 2  # anion and counter-ion
    assert float(converted.get("BondLength")) == 14.4


def test_from_cdxml_accepts_path_text_and_bytes():
    path = PACKAGE / "cdxml" / "C-2.cdxml"
    from_path = StructureSource.from_cdxml(path)
    from_text = StructureSource.from_cdxml(path.read_text(encoding="utf-8"), name="C-2")
    from_bytes = StructureSource.from_cdxml(path.read_bytes(), name="C-2")
    assert from_path.name == "C-2"
    assert from_path.cdx_data == from_text.cdx_data == from_bytes.cdx_data
    assert from_path.cdx_data.startswith(b"VjCD0100")


def test_content_aspect_without_fragment_bounding_boxes():
    """Structura writes no fragment BoundingBox; atom positions stand in."""
    path = PACKAGE / "cdxml" / "C-1.cdxml"
    # The root BoundingBox is the atom box (the page box is the whole page).
    x0, y0, x1, y1 = (float(v) for v in ET.parse(path).getroot().get("BoundingBox").split())
    assert StructureSource.from_cdxml(path).content_aspect == pytest.approx((x1 - x0) / (y1 - y0))


# -- CompoundTable.load ---------------------------------------------------------


def _zip_package(tmp_path: Path, *, nested: bool) -> Path:
    archive = tmp_path / ("nested.zip" if nested else "flat.zip")
    with zipfile.ZipFile(archive, "w") as z:
        for file in PACKAGE.rglob("*"):
            if file.is_file():
                name = file.relative_to(PACKAGE).as_posix()
                z.write(file, f"某专利 化合物表/{name}" if nested else name)
    return archive


@pytest.mark.parametrize("kind", ["directory", "flat zip", "nested zip"])
def test_load_reads_directory_and_both_zip_layouts(tmp_path, kind):
    path = {
        "directory": PACKAGE,
        "flat zip": _zip_package(tmp_path, nested=False),
        "nested zip": _zip_package(tmp_path, nested=True),
    }[kind]
    table = CompoundTable.load(path)
    assert table.columns == 3
    assert [c.id for c in table.compounds] == ["C-1", "C-2", "C-3"]
    first = table.compounds[0]
    assert first.formula == "C7H5NaO2"
    assert first.mw == pytest.approx(144.1)
    assert first.fields == {"R1": "H"}
    assert first.smiles is None  # optional in v1
    assert first.structure.name == "C-1"


def _edited_package(tmp_path: Path, edit) -> Path:
    target = tmp_path / "package"
    shutil.copytree(PACKAGE, target)
    manifest = json.loads((target / "manifest.json").read_text(encoding="utf-8"))
    edit(manifest)
    (target / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
    return target


@pytest.mark.parametrize(
    "edit,message",
    [
        (lambda m: m.update(format="something-else"), "format"),
        (lambda m: m.update(version=2), "unsupported"),
        (lambda m: m.update(columns=0), "columns"),
        (lambda m: m.update(compounds=[]), "no compounds"),
        (lambda m: m["compounds"][0].update(file="cdxml/missing.cdxml"), "not in the package"),
        (lambda m: m["compounds"][0].update(file="../../etc/passwd"), "Unsafe"),
        (lambda m: m["compounds"][1].update(id="C-1"), "duplicate"),
    ],
)
def test_load_rejects_bad_manifests(tmp_path, edit, message):
    with pytest.raises(ValueError, match=message):
        CompoundTable.load(_edited_package(tmp_path, edit))


def test_whole_table_cdxml_is_ignored():
    """table.cdxml is for people working in ChemDraw directly; we never read it."""
    assert (PACKAGE / "table.cdxml").exists()
    table = CompoundTable.load(PACKAGE)
    assert [c.id for c in table.compounds] == ["C-1", "C-2", "C-3"]


def test_file_field_is_used_not_the_id(tmp_path):
    """Ids may hold characters the exporter replaced in the file name."""
    def edit(manifest):
        manifest["compounds"][0]["id"] = "C-1/a"

    table = CompoundTable.load(_edited_package(tmp_path, edit))
    assert table.compounds[0].id == "C-1/a"


# -- Word compound table ----------------------------------------------------------

# Natural sizes the fake renderer reports, per compound file (pt).
_SIZES = {"structure1": (102.0, 27.0), "structure2": (102.0, 54.0), "structure3": (60.0, 27.0)}


@pytest.fixture
def fake_natural(monkeypatch):
    """Stand in for RDKit/Inkscape: fixed natural sizes, tiny PNG previews."""
    import retain_cdx.office.word.writer as writer_module

    def fake(jobs, **kwargs):
        from PIL import Image

        sizes = []
        for cdx, out in jobs:
            Image.new("RGB", (8, 8), "white").save(out)
            sizes.append(_SIZES.get(Path(cdx).stem, (80.0, 40.0)))
        return "png", sizes

    monkeypatch.setattr(writer_module, "render_previews_at_natural_scale", fake)
    monkeypatch.setattr(writer_module, "preview_extension", lambda fmt: "png")
    return PreviewOptions(format="png")


def _table(archive: zipfile.ZipFile) -> ET.Element:
    return ET.fromstring(archive.read("word/document.xml")).find(".//w:tbl", NS)


def _shape_width_pt(obj: ET.Element) -> float:
    style = obj.find("v:shape", NS).get("style")
    return float(re.search(r"width:([\d.]+)pt", style).group(1))


def _shape_height_pt(obj: ET.Element) -> float:
    style = obj.find("v:shape", NS).get("style")
    return float(re.search(r"height:([\d.]+)pt", style).group(1))


def _seven_compounds():
    table = CompoundTable.load(PACKAGE)
    return [replace(table.compounds[i % 3], id=f"C-{i + 1}") for i in range(7)]


def test_table_layout_alternates_structure_and_number_rows(tmp_path, fake_natural):
    out = tmp_path / "table.docx"
    with WordCDXWriter() as writer:
        writer.insert_compound_table(_seven_compounds(), columns=3, preview=fake_natural)
        writer.save(out)

    assert check_package(out) == []
    with zipfile.ZipFile(out) as archive:
        rows = _table(archive).findall("w:tr", NS)
    assert len(rows) == 6  # ceil(7 / 3) groups × 2 rows
    for structure_row, label_row in zip(rows[::2], rows[1::2]):
        assert len(structure_row.findall("w:tc", NS)) == 3
        assert len(label_row.findall("w:tc", NS)) == 3
        for row in (structure_row, label_row):
            assert row.find("w:trPr/w:cantSplit", NS) is not None
        for paragraph in structure_row.findall("w:tc/w:p", NS):
            assert paragraph.find("w:pPr/w:keepNext", NS) is not None
    labels = [
        "".join(t.text for t in cell.iter(f"{{{NS_W}}}t"))
        for row in rows[1::2]
        for cell in row.findall("w:tc", NS)
    ]
    assert labels == ["C-1", "C-2", "C-3", "C-4", "C-5", "C-6", "C-7", "", ""]
    objects_per_row = [len(row.findall(".//w:object", NS)) for row in rows[::2]]
    assert objects_per_row == [3, 3, 1]


def test_one_scale_for_the_whole_table(tmp_path, fake_natural):
    """Every object is its natural size × the same factor: equal bond lengths."""
    out = tmp_path / "scale.docx"
    with WordCDXWriter() as writer:
        scale = writer.insert_compound_table(CompoundTable.load(PACKAGE), preview=fake_natural)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        objects = _table(archive).findall(".//w:object", NS)
    natural = [_SIZES["structure1"], _SIZES["structure2"], _SIZES["structure3"]]
    for obj, (width, height) in zip(objects, natural):
        assert _shape_width_pt(obj) == pytest.approx(width * scale, abs=0.01)
        assert _shape_height_pt(obj) == pytest.approx(height * scale, abs=0.01)


def test_default_scale_fits_the_widest_structure_in_a_column(tmp_path, fake_natural):
    out = tmp_path / "fit.docx"
    with WordCDXWriter(BLANK) as writer:
        scale = writer.insert_compound_table(CompoundTable.load(PACKAGE), preview=fake_natural)
        writer.save(out)

    # blank.docx: A4 (11906) minus 1800 + 1800 margins = 8306 twips of text.
    column_twips = 8306 // 3
    usable_pt = (column_twips - 2 * 108) / 20
    assert scale == pytest.approx(usable_pt / 102.0)
    with zipfile.ZipFile(out) as archive:
        table = _table(archive)
        widest = max(_shape_width_pt(obj) for obj in table.findall(".//w:object", NS))
        assert widest == pytest.approx(usable_pt, abs=0.01)
        assert table.find("w:tblGrid/w:gridCol", NS).get(f"{{{NS_W}}}w") == str(column_twips)


def test_explicit_scale_and_columns(tmp_path, fake_natural):
    out = tmp_path / "explicit.docx"
    with WordCDXWriter() as writer:
        scale = writer.insert_compound_table(
            CompoundTable.load(PACKAGE), columns=2, scale=0.5, preview=fake_natural
        )
        writer.save(out)

    assert scale == 0.5
    with zipfile.ZipFile(out) as archive:
        table = _table(archive)
        assert len(table.findall("w:tblGrid/w:gridCol", NS)) == 2
        assert _shape_width_pt(table.find(".//w:object", NS)) == pytest.approx(51.0)


def test_table_replaces_its_placeholder_paragraph(tmp_path, fake_natural):
    source = tmp_path / "input.docx"
    with zipfile.ZipFile(BLANK) as zin, zipfile.ZipFile(source, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                body = (
                    "<w:p><w:r><w:t>前言</w:t></w:r></w:p>"
                    "<w:p><w:r><w:t>{{com</w:t></w:r><w:r><w:t>pounds}}</w:t></w:r></w:p>"
                    "<w:p><w:r><w:t>后记</w:t></w:r></w:p>"
                )
                data = re.sub(r"<w:body>.*?(<w:sectPr)", lambda m: "<w:body>" + body + m.group(1), data.decode()).encode()
            zout.writestr(info, data)

    out = tmp_path / "placed.docx"
    with WordCDXWriter(source) as writer:
        writer.insert_compound_table(
            CompoundTable.load(PACKAGE), placeholder="{{compounds}}", preview=fake_natural
        )
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        body = ET.fromstring(archive.read("word/document.xml")).find("w:body", NS)
    kinds = [child.tag.split("}")[1] for child in body]
    texts = ["".join(t.text or "" for t in child.iter(f"{{{NS_W}}}t")) for child in body]
    assert kinds[:3] == ["p", "tbl", "p"]
    assert texts[0] == "前言"
    assert texts[2] == "后记"
    assert "{{" not in "".join(texts)


def test_placeholder_must_have_its_own_paragraph(tmp_path, monkeypatch, fake_natural):
    source = tmp_path / "input.docx"
    with zipfile.ZipFile(BLANK) as zin, zipfile.ZipFile(source, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                data = data.replace(b"</w:p>", b"<w:r><w:t>see {{t}}</w:t></w:r></w:p>", 1)
            zout.writestr(info, data)
    with WordCDXWriter(source) as writer, pytest.raises(ValueError, match="paragraph of its own"):
        writer.insert_compound_table(CompoundTable.load(PACKAGE), placeholder="{{t}}", preview=fake_natural)


def test_table_in_a_blank_document_has_no_empty_line_above(tmp_path, fake_natural):
    out = tmp_path / "blank.docx"
    with WordCDXWriter() as writer:
        writer.insert_compound_table(CompoundTable.load(PACKAGE), preview=fake_natural)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        body = ET.fromstring(archive.read("word/document.xml")).find("w:body", NS)
    assert [child.tag.split("}")[1] for child in body] == ["tbl", "p", "sectPr"]


def test_containers_carry_the_cdxml_structures(tmp_path, fake_natural):
    import olefile

    out = tmp_path / "payload.docx"
    table = CompoundTable.load(PACKAGE)
    with WordCDXWriter() as writer:
        writer.insert_compound_table(table, preview=fake_natural)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        for index, compound in enumerate(table.compounds, start=1):
            ole = olefile.OleFileIO(io.BytesIO(archive.read(f"word/embeddings/oleObject{index}.bin")))
            assert ole.openstream("CONTENTS").read() == compound.structure.cdx_data


# -- blank documents ----------------------------------------------------------------


def test_writer_without_input_starts_an_empty_document(tmp_path):
    out = tmp_path / "empty.docx"
    with WordCDXWriter() as writer:
        assert writer.paragraph_count == 1
        writer.save(out)

    assert check_package(out) == []
    with zipfile.ZipFile(out) as archive:
        names = archive.namelist()
        assert not [n for n in names if "/embeddings/" in n or "/media/" in n]
        core = archive.read("docProps/core.xml").decode()
    assert "<dc:creator></dc:creator>" in core
    assert "<cp:lastModifiedBy></cp:lastModifiedBy>" in core


# -- natural-scale rendering (real RDKit) ------------------------------------------------


def test_natural_size_tracks_the_drawing(tmp_path):
    pytest.importorskip("rdkit")
    from retain_cdx.render import natural_size

    table = CompoundTable.load(PACKAGE)
    for compound in table.compounds:
        path = compound.structure.write_cdx(tmp_path / f"{compound.id}.cdx")
        width, height = natural_size(path)
        root = ET.parse(PACKAGE / f"cdxml/{compound.id}.cdxml").getroot()
        x0, y0, x1, y1 = (float(v) for v in root.get("BoundingBox").split())
        # ACS 1996 draws 14.4 pt bonds, matching Structura's BondLength="14.4":
        # the atom box comes out at its own size plus the labels sticking out
        # past the edge atoms (Na⁺, O⁻) and padding — an absolute allowance,
        # not a fraction of the molecule.
        assert 0 < width - (x1 - x0) < 30
        assert 0 < height - (y1 - y0) < 20


def test_png_previews_at_natural_scale_share_one_resolution(tmp_path):
    pytest.importorskip("rdkit")
    from PIL import Image

    from retain_cdx.render import render_previews_at_natural_scale

    table = CompoundTable.load(PACKAGE)
    jobs = [
        (c.structure.write_cdx(tmp_path / f"{c.id}.cdx"), tmp_path / f"{c.id}.png")
        for c in table.compounds
    ]
    resolved, sizes = render_previews_at_natural_scale(jobs, format="png")
    assert resolved == "png"
    for (_, png), (width, height) in zip(jobs, sizes):
        assert Image.open(png).size == (round(width * 4), round(height * 4))


# -- CLI --------------------------------------------------------------------------------


def test_cli_builds_a_table_from_a_zip(tmp_path, fake_natural, capsys):
    from scripts.cli.structura_table_to_docx import main

    out = tmp_path / "cli.docx"
    assert main(["--input", str(_zip_package(tmp_path, nested=True)), "--output", str(out), "--columns", "2"]) == 0
    assert "3 compounds, 2 per row" in capsys.readouterr().out
    assert check_package(out) == []
    with zipfile.ZipFile(out) as archive:
        assert len(_table(archive).findall("w:tblGrid/w:gridCol", NS)) == 2


def test_cli_placeholder_needs_a_document(tmp_path):
    from scripts.cli.structura_table_to_docx import main

    with pytest.raises(SystemExit, match="--document"):
        main(["--input", str(PACKAGE), "--output", str(tmp_path / "x.docx"), "--placeholder", "{{t}}"])


def test_preview_keeps_every_fragment(tmp_path):
    """A salt's counter-ion used to vanish: only RDKit's first molecule was drawn."""
    pytest.importorskip("rdkit")
    from rdkit import Chem

    from retain_cdx.render import mol_from_cdx

    salt = StructureSource.from_cdxml(PACKAGE / "cdxml" / "C-1.cdxml")  # sodium benzoate
    mol = mol_from_cdx(salt.write_cdx(tmp_path / "salt.cdx"))
    assert Chem.MolToSmiles(mol) == "O=C([O-])c1ccccc1.[Na+]"
    assert mol.GetNumConformers() == 1


# -- EMF frame -----------------------------------------------------------------------


def _emf_header(frame_mm: tuple[float, float]) -> bytes:
    """A bare EMF header record as Inkscape writes it (1200 dpi device)."""
    import struct

    header = bytearray(88)
    struct.pack_into("<II", header, 0, 1, 88)
    struct.pack_into("<4i", header, 8, 0, 0, 5006, 1368)
    struct.pack_into("<4i", header, 24, 0, 0, round(frame_mm[0] * 100) - 1, round(frame_mm[1] * 100) - 1)
    header[40:44] = b" EMF"
    struct.pack_into("<ii", header, 72, 10205, 13181)
    struct.pack_into("<ii", header, 80, 216, 279)
    return bytes(header)


def test_set_emf_frame_fits_the_frame_to_the_page(tmp_path):
    import struct

    from retain_cdx.render.preview import set_emf_frame

    emf = tmp_path / "x.emf"
    emf.write_bytes(_emf_header((106.0, 29.0)))  # Inkscape's rounded-up frame
    set_emf_frame(emf, 105.833, 28.015)
    data = emf.read_bytes()
    assert struct.unpack_from("<4i", data, 24) == (0, 0, 10582, 2801)
    right, bottom = struct.unpack_from("<4i", data, 8)[2:]
    assert (right + 1) / (bottom + 1) == pytest.approx(105.833 / 28.015, rel=1e-3)


def test_set_emf_frame_rejects_other_files(tmp_path):
    from retain_cdx.render.preview import set_emf_frame

    png = tmp_path / "x.emf"
    png.write_bytes(b"\x89PNG" + b"\x00" * 100)
    with pytest.raises(ValueError, match="Not an EMF"):
        set_emf_frame(png, 10, 10)


def test_emf_previews_at_natural_scale_have_the_natural_aspect(tmp_path):
    """End to end with Inkscape: frame aspect = drawing aspect, every molecule."""
    pytest.importorskip("rdkit")
    import struct

    from retain_cdx.render import inkscape_available, render_previews_at_natural_scale

    if not inkscape_available():
        pytest.skip("Inkscape not installed")
    table = CompoundTable.load(PACKAGE)
    jobs = [
        (c.structure.write_cdx(tmp_path / f"{c.id}.cdx"), tmp_path / f"{c.id}.emf")
        for c in table.compounds
    ]
    resolved, sizes = render_previews_at_natural_scale(jobs, format="emf")
    assert resolved == "emf"
    for (_, emf), (width, height) in zip(jobs, sizes):
        frame = struct.unpack_from("<4i", emf.read_bytes(), 24)
        assert (frame[2] + 1) / (frame[3] + 1) == pytest.approx(width / height, rel=1e-3)
