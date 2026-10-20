from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import olefile
import pytest

from ooxml_checks import check_package
from retain_cdx import PreviewOptions, StructureSource, WordCDXWriter
from retain_cdx.ole import write_cdx_to_ole
from retain_cdx.office.word._wml import NS_O, NS_R, NS_V, NS_W
from retain_cdx.resources import bundled_template_docx

NS = {"w": NS_W, "v": NS_V, "o": NS_O, "r": NS_R}
NS_XML = "http://www.w3.org/XML/1998/namespace"
CDX = "tests/fixtures/test.cdx"
BLANK = "tests/fixtures/blank.docx"
FAKE_EMF = b"\x01\x00\x00\x00fake-emf-preview"


def _fake_renderer(monkeypatch, extension: str, data: bytes | None = None) -> list[Path]:
    """Stand in for RDKit/Inkscape; returns the list of rendered paths."""
    import retain_cdx.office.word.writer as writer_module

    calls: list[Path] = []

    def fake_render(cdx_path, output_path, **kwargs):
        calls.append(Path(output_path))
        if data is not None:
            Path(output_path).write_bytes(data)
        else:
            from PIL import Image

            Image.new("RGB", (32, 24), "white").save(output_path)
        return extension

    monkeypatch.setattr(writer_module, "render_preview", fake_render)
    monkeypatch.setattr(writer_module, "preview_extension", lambda fmt: extension)
    return calls


@pytest.fixture
def png_preview(monkeypatch):
    _fake_renderer(monkeypatch, "png")
    return PreviewOptions(format="png")


@pytest.fixture
def emf_preview(monkeypatch):
    _fake_renderer(monkeypatch, "emf", FAKE_EMF)
    return PreviewOptions(format="emf")


def _docx_with_body(tmp_path: Path, paragraphs_xml: str) -> Path:
    """blank.docx with *paragraphs_xml* as the body (sectPr kept last)."""
    path = tmp_path / "input.docx"
    with zipfile.ZipFile(BLANK) as zin, zipfile.ZipFile(path, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                text = data.decode()
                text = re.sub(r"<w:body>.*?(<w:sectPr)", rf"<w:body>{paragraphs_xml}\1", text)
                data = text.encode()
            zout.writestr(info, data)
    return path


def _document(archive: zipfile.ZipFile) -> ET.Element:
    return ET.fromstring(archive.read("word/document.xml"))


def _objects(archive: zipfile.ZipFile) -> list[ET.Element]:
    return _document(archive).findall(".//w:object", NS)


def _tokens(paragraph: ET.Element) -> list[str]:
    """Paragraph content in order: text pieces and ``[OBJ]`` for objects."""
    tokens = []
    for run in paragraph.iter(f"{{{NS_W}}}r"):
        for child in run:
            if child.tag == f"{{{NS_W}}}t" and child.text:
                tokens.append(child.text)
            elif child.tag == f"{{{NS_W}}}object":
                tokens.append("[OBJ]")
    return tokens


def _container(archive: zipfile.ZipFile, name: str) -> olefile.OleFileIO:
    return olefile.OleFileIO(io.BytesIO(archive.read(name)))


# -- package ------------------------------------------------------------------


def test_insert_one_structure_builds_a_valid_package(tmp_path, png_preview):
    out = tmp_path / "one.docx"
    with WordCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, preview=png_preview)
        writer.save(out)

    assert check_package(out) == []
    with zipfile.ZipFile(out) as archive:
        assert len(_objects(archive)) == 1
        assert [n for n in archive.namelist() if "/embeddings/" in n] == [
            "word/embeddings/oleObject1.bin"
        ]
        assert [n for n in archive.namelist() if "/media/" in n] == ["word/media/image1.png"]


def test_object_references_its_own_parts(tmp_path, png_preview):
    out = tmp_path / "rels.docx"
    with WordCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        rels = {
            el.get("Id"): el.get("Target")
            for el in ET.fromstring(archive.read("word/_rels/document.xml.rels"))
        }
        (obj,) = _objects(archive)
        shape = obj.find("v:shape", NS)
        ole = obj.find("o:OLEObject", NS)
        assert rels[shape.find("v:imagedata", NS).get(f"{{{NS_R}}}id")] == "media/image1.png"
        assert rels[ole.get(f"{{{NS_R}}}id")] == "embeddings/oleObject1.bin"
        assert ole.get("ShapeID") == shape.get("id")
        assert ole.get("ProgID") == "ChemDraw.Document.6.0"


def test_second_object_gets_the_next_part_number(tmp_path, png_preview):
    """Word numbers parts oleObject1, oleObject2…, not oleObject1_2."""
    out = tmp_path / "numbered.docx"
    with WordCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, preview=png_preview)
        writer.insert_cdx(CDX, preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        assert sorted(n for n in archive.namelist() if "/embeddings/" in n) == [
            "word/embeddings/oleObject1.bin",
            "word/embeddings/oleObject2.bin",
        ]


def test_mc_ignorable_prefixes_stay_declared(tmp_path, png_preview):
    """Regression guard: Word's root declares w15, w16cid… only for mc:Ignorable.

    ElementTree drops declarations no element uses; losing them leaves
    Ignorable naming undeclared prefixes, which Word answers with a repair.
    """
    out = tmp_path / "mce.docx"
    with bundled_template_docx() as template:
        with zipfile.ZipFile(template) as archive:
            before = archive.read("word/document.xml").decode()
        with WordCDXWriter(template) as writer:
            writer.insert_cdx(CDX, preview=png_preview)
            writer.save(out)

    with zipfile.ZipFile(out) as archive:
        after = archive.read("word/document.xml").decode()
    ignorable = re.search(r'mc:Ignorable="([^"]+)"', after).group(1).split()
    assert "w16du" in ignorable
    for prefix in re.findall(r"xmlns:([\w.-]+)=", before):
        assert f"xmlns:{prefix}=" in after, prefix
    assert check_package(out) == []


# -- OLE container --------------------------------------------------------------


def test_container_has_the_streams_word_writes(tmp_path, emf_preview):
    """Same stream set as Windows Word: CompObj, Ole, EPRINT, ObjInfo, CONTENTS."""
    out = tmp_path / "streams.docx"
    with WordCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, preview=emf_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        ole = _container(archive, "word/embeddings/oleObject1.bin")
        assert sorted("/".join(e) for e in ole.listdir()) == sorted(
            ["\x01CompObj", "\x01Ole", "\x03EPRINT", "\x03ObjInfo", "CONTENTS"]
        )
        assert ole.openstream("CONTENTS").read() == Path(CDX).read_bytes()
        # Word's cache is the preview EMF itself, byte for byte.
        assert ole.openstream("\x03EPRINT").read() == archive.read("word/media/image1.emf")


def test_fixed_streams_match_the_native_word_container(tmp_path, emf_preview):
    out = tmp_path / "native.docx"
    with WordCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, preview=emf_preview)
        writer.save(out)

    with bundled_template_docx() as template, zipfile.ZipFile(template) as archive:
        native = _container(archive, "word/embeddings/oleObject1.bin")
        native_streams = {
            name: native.openstream(name).read()
            for name in ("\x01CompObj", "\x01Ole", "\x03ObjInfo")
        }
    with zipfile.ZipFile(out) as archive:
        ours = _container(archive, "word/embeddings/oleObject1.bin")
        for name, data in native_streams.items():
            assert ours.openstream(name).read() == data, repr(name)
        assert ours.root.clsid == native.root.clsid


def test_png_preview_carries_no_eprint(tmp_path, png_preview):
    """EPRINT is an EMF; a bitmap preview has nothing to put there."""
    out = tmp_path / "png.docx"
    with WordCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        ole = _container(archive, "word/embeddings/oleObject1.bin")
        assert not ole.exists("\x03EPRINT")
        assert ole.exists("\x03ObjInfo")


def test_presentation_cache_false_drops_eprint(tmp_path, monkeypatch):
    _fake_renderer(monkeypatch, "emf", FAKE_EMF)
    out = tmp_path / "nocache.docx"
    with WordCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, preview=PreviewOptions(format="emf", presentation_cache=False))
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        assert not _container(archive, "word/embeddings/oleObject1.bin").exists("\x03EPRINT")


def test_presentation_cache_with_bitmap_preview_is_an_error(monkeypatch):
    _fake_renderer(monkeypatch, "png")
    with WordCDXWriter(BLANK) as writer, pytest.raises(ValueError, match="vector preview"):
        writer.insert_cdx(CDX, preview=PreviewOptions(format="png", presentation_cache=True))


def test_patching_a_word_container_drops_its_stale_eprint(tmp_path):
    """EPRINT shows the template's molecule; an in-place patch must not keep it."""
    with bundled_template_docx() as template, zipfile.ZipFile(template) as archive:
        native = archive.read("word/embeddings/oleObject1.bin")
    target = tmp_path / "patched.bin"
    write_cdx_to_ole(target, Path(CDX), template_ole=native)
    ole = olefile.OleFileIO(str(target))
    assert not ole.exists("\x03EPRINT")
    assert ole.openstream("CONTENTS").read() == Path(CDX).read_bytes()


# -- ids ------------------------------------------------------------------------


def test_many_objects_get_unique_ids_and_one_shapetype(tmp_path, png_preview):
    out = tmp_path / "many.docx"
    with WordCDXWriter(BLANK) as writer:
        writer.insert_cdx_many([CDX, CDX, CDX], preview=png_preview)
        writer.save(out)

    assert check_package(out) == []
    with zipfile.ZipFile(out) as archive:
        objects = _objects(archive)
        assert len(objects) == 3
        assert len(_document(archive).findall(".//v:shapetype", NS)) == 1
        # The definition precedes every reference to it.
        assert objects[0].find("v:shapetype", NS) is not None
        for obj in objects:
            assert obj.find("v:shape", NS).get("type") == "#_x0000_t75"
            assert obj.get("{http://schemas.microsoft.com/office/word/2010/wordml}anchorId") is None


def test_insert_into_document_with_existing_object_appends(tmp_path, png_preview):
    """A document already holding a ChemDraw object keeps it; ours is added."""
    out = tmp_path / "existing.docx"
    with bundled_template_docx() as template:
        with WordCDXWriter(template) as writer:
            writer.insert_cdx(CDX, preview=png_preview)
            writer.save(out)

    assert check_package(out) == []
    with zipfile.ZipFile(out) as archive:
        objects = _objects(archive)
        assert len(objects) == 2
        # The existing shapetype is reused, not redefined.
        assert objects[1].find("v:shapetype", NS) is None
        native, ours = (obj.find("o:OLEObject", NS) for obj in objects)
        assert native.get("ObjectID") == "_1853394678"
        assert ours.get("ShapeID") == "_x0000_i1026"
        assert ours.get("ObjectID") == "_1853394679"
        # The native object's parts are untouched.
        assert archive.read("word/embeddings/oleObject1.bin")
        assert "word/embeddings/oleObject2.bin" in archive.namelist()


# -- where it goes --------------------------------------------------------------


def test_new_paragraph_goes_before_sectpr_with_alignment(tmp_path, png_preview):
    out = tmp_path / "align.docx"
    with WordCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, align="center", preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        body = _document(archive).find("w:body", NS)
        children = list(body)
        assert children[-1].tag == f"{{{NS_W}}}sectPr"
        paragraph = children[-2]
        assert paragraph.find("w:pPr/w:jc", NS).get(f"{{{NS_W}}}val") == "center"
        assert paragraph.find("w:r/w:object", NS) is not None


def test_bad_alignment_is_rejected_before_rendering(monkeypatch):
    calls = _fake_renderer(monkeypatch, "png")
    with WordCDXWriter(BLANK) as writer, pytest.raises(ValueError, match="align"):
        writer.insert_cdx(CDX, align="middle", preview=PreviewOptions(format="png"))
    assert calls == []


def test_placeholder_in_one_run_keeps_the_text_around_it(tmp_path, png_preview):
    source = _docx_with_body(
        tmp_path,
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">结构 {{mol}} 如图</w:t></w:r></w:p>',
    )
    out = tmp_path / "placeholder.docx"
    with WordCDXWriter(source) as writer:
        writer.insert_cdx(CDX, placeholder="{{mol}}", preview=png_preview)
        writer.save(out)

    assert check_package(out) == []
    with zipfile.ZipFile(out) as archive:
        paragraph = _document(archive).find("w:body/w:p", NS)
        assert _tokens(paragraph) == ["结构 ", "[OBJ]", " 如图"]
        runs = paragraph.findall("w:r", NS)
        # Both halves of the text keep the run's formatting.
        assert runs[0].find("w:rPr/w:b", NS) is not None
        assert runs[2].find("w:rPr/w:b", NS) is not None
        for t in paragraph.iter(f"{{{NS_W}}}t"):
            assert t.get(f"{{{NS_XML}}}space") == "preserve"


def test_placeholder_split_across_runs(tmp_path, png_preview):
    """Word breaks hand-typed text into runs; the placeholder still matches."""
    source = _docx_with_body(
        tmp_path,
        "<w:p>"
        "<w:r><w:t>见{{</w:t></w:r>"
        "<w:r><w:rPr><w:i/></w:rPr><w:t>mol</w:t></w:r>"
        '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>}}后文</w:t></w:r>'
        "</w:p>",
    )
    out = tmp_path / "split.docx"
    with WordCDXWriter(source) as writer:
        writer.insert_cdx(CDX, placeholder="{{mol}}", preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        paragraph = _document(archive).find("w:body/w:p", NS)
        assert _tokens(paragraph) == ["见", "[OBJ]", "后文"]
        # The trailing text stays in its own run, underline and all.
        tail_run = next(
            run for run in paragraph.findall("w:r", NS) if run.findtext("w:t", namespaces=NS) == "后文"
        )
        assert tail_run.find("w:rPr/w:u", NS) is not None
    assert check_package(out) == []


def test_placeholder_inside_a_table_cell(tmp_path, png_preview):
    source = _docx_with_body(
        tmp_path,
        "<w:tbl><w:tr>"
        "<w:tc><w:p><w:r><w:t>咖啡因</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>{{caffeine}}</w:t></w:r></w:p></w:tc>"
        "</w:tr></w:tbl><w:p/>",
    )
    out = tmp_path / "table.docx"
    with WordCDXWriter(source) as writer:
        writer.insert_cdx_many([CDX], placeholders=["{{caffeine}}"], preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        cells = _document(archive).findall(".//w:tc", NS)
        assert _tokens(cells[0]) == ["咖啡因"]
        assert _tokens(cells[1]) == ["[OBJ]"]


def test_placeholders_fill_in_order(tmp_path, png_preview):
    source = _docx_with_body(
        tmp_path, "<w:p><w:r><w:t>{{a}} 与 {{b}}</w:t></w:r></w:p>"
    )
    out = tmp_path / "two.docx"
    with WordCDXWriter(source) as writer:
        writer.insert_cdx_many(
            [CDX, "tests/fixtures/san.cdx"], placeholders=["{{a}}", "{{b}}"], preview=png_preview
        )
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        paragraph = _document(archive).find("w:body/w:p", NS)
        assert _tokens(paragraph) == ["[OBJ]", " 与 ", "[OBJ]"]
        first, second = (obj.find("o:OLEObject", NS) for obj in _objects(archive))
        rels = {
            el.get("Id"): el.get("Target")
            for el in ET.fromstring(archive.read("word/_rels/document.xml.rels"))
        }
        payload = _container(archive, "word/" + rels[second.get(f"{{{NS_R}}}id")])
        assert payload.openstream("CONTENTS").read() == Path("tests/fixtures/san.cdx").read_bytes()


def test_missing_placeholder_fails_before_writing_anything(tmp_path, monkeypatch):
    """A failure halfway would leave orphan parts in the package."""
    calls = _fake_renderer(monkeypatch, "png")
    source = _docx_with_body(tmp_path, "<w:p><w:r><w:t>{{a}}</w:t></w:r></w:p>")
    with WordCDXWriter(source) as writer:
        with pytest.raises(ValueError, match="not found"):
            writer.insert_cdx_many(
                [CDX, CDX], placeholders=["{{a}}", "{{nowhere}}"], preview=PreviewOptions(format="png")
            )
        with pytest.raises(ValueError, match="occurs 1 times"):
            writer.insert_cdx_many(
                [CDX, CDX], placeholders=["{{a}}", "{{a}}"], preview=PreviewOptions(format="png")
            )
        assert not (writer.package_dir / "word" / "embeddings").exists()
    assert calls == []


def test_repeated_placeholder_fills_each_occurrence(tmp_path, png_preview):
    source = _docx_with_body(tmp_path, "<w:p><w:r><w:t>{{m}}，{{m}}</w:t></w:r></w:p>")
    out = tmp_path / "repeat.docx"
    with WordCDXWriter(source) as writer:
        writer.insert_cdx_many([CDX, CDX], placeholders=["{{m}}", "{{m}}"], preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        assert _tokens(_document(archive).find("w:body/w:p", NS)) == ["[OBJ]", "，", "[OBJ]"]


def test_placeholder_count_must_match(png_preview):
    with WordCDXWriter(BLANK) as writer, pytest.raises(ValueError, match="placeholders"):
        writer.insert_cdx_many([CDX, CDX], placeholders=["{{a}}"], preview=png_preview)


def test_paragraph_index_appends_side_by_side(tmp_path, png_preview):
    source = _docx_with_body(
        tmp_path, "<w:p><w:r><w:t>第一段</w:t></w:r></w:p><w:p><w:r><w:t>第二段：</w:t></w:r></w:p>"
    )
    out = tmp_path / "paragraph.docx"
    with WordCDXWriter(source) as writer:
        assert writer.paragraph_count == 2
        writer.insert_cdx_many([CDX, CDX], paragraph=1, preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        first, second = _document(archive).findall("w:body/w:p", NS)
        assert _tokens(first) == ["第一段"]
        assert _tokens(second) == ["第二段：", "[OBJ]", "[OBJ]"]


def test_paragraph_out_of_range_fails_before_rendering(monkeypatch):
    calls = _fake_renderer(monkeypatch, "png")
    with WordCDXWriter(BLANK) as writer, pytest.raises(ValueError, match="out of range"):
        writer.insert_cdx(CDX, paragraph=5, preview=PreviewOptions(format="png"))
    assert calls == []


def test_location_options_are_exclusive(png_preview):
    with WordCDXWriter(BLANK) as writer:
        with pytest.raises(ValueError, match="not both"):
            writer.insert_cdx(CDX, placeholder="{{a}}", paragraph=0, preview=png_preview)
        with pytest.raises(ValueError, match="new paragraphs"):
            writer.insert_cdx(CDX, paragraph=0, align="center", preview=png_preview)


# -- size -----------------------------------------------------------------------


def _shape_size(obj: ET.Element) -> tuple[float, float]:
    style = obj.find("v:shape", NS).get("style")
    width = float(re.search(r"width:([\d.]+)pt", style).group(1))
    height = float(re.search(r"height:([\d.]+)pt", style).group(1))
    return width, height


@pytest.mark.parametrize(
    "kwargs",
    [{}, {"width": "3cm"}, {"height": "2cm"}, {"width": "10cm", "height": "1cm"}, {"width": "1cm", "height": "10cm"}],
)
def test_size_keeps_the_molecule_aspect(tmp_path, png_preview, kwargs):
    """Any box we write matches the CDX aspect, so nothing stretches on re-render."""
    out = tmp_path / "size.docx"
    with WordCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, preview=png_preview, **kwargs)
        writer.save(out)

    aspect = StructureSource.from_cdx(CDX).content_aspect
    with zipfile.ZipFile(out) as archive:
        (obj,) = _objects(archive)
    width, height = _shape_size(obj)
    assert width / height == pytest.approx(aspect, rel=0.01)
    # dxaOrig/dyaOrig (twips) describe the same box, so Word reports 100 %.
    assert int(obj.get(f"{{{NS_W}}}dxaOrig")) == pytest.approx(width * 20, abs=1)
    assert int(obj.get(f"{{{NS_W}}}dyaOrig")) == pytest.approx(height * 20, abs=1)


def test_size_fits_inside_the_given_box(tmp_path, png_preview):
    out = tmp_path / "fit.docx"
    with WordCDXWriter(BLANK) as writer:
        writer.insert_cdx(CDX, width="10cm", height="1cm", preview=png_preview)
        writer.save(out)

    with zipfile.ZipFile(out) as archive:
        width, height = _shape_size(_objects(archive)[0])
    assert height == pytest.approx(72 / 2.54, abs=0.01)  # 1 cm in pt
    assert width < 10 * 72 / 2.54
