from __future__ import annotations

from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

import retain_cdx.office.excel._shell.preview as shell_preview
import retain_cdx.office.excel._source.fill_parts as fill_parts
from retain_cdx import ExcelCDXWriter, ShellTemplate, StructureSource
from retain_cdx.office.excel._ooxml import extract_xlsx, pack_xlsx
from retain_cdx.render.preview import inkscape_available, resolve_preview_format
from PIL import Image
import pytest


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {
    "s": NS_MAIN,
    "rel": NS_REL,
    "xdr": NS_XDR,
    "mc": NS_MC,
}


# Minimal EMF header so `file` / zip packaging can treat the payload as binary media.
_PLACEHOLDER_EMF = bytes.fromhex(
    "01000000d80000000000000000000000"
    "c7090000c70900000000000000000000"
    "b3140000b314000020454d4600000100"
    "14000000000000000000000000000000"
    "00000000000000000000000000000000"
    "00000000000000000000000000000000"
    "00000000000000000000000000000000"
    "00000000000000000000000000000000"
    "00000000000000000000000000000000"
    "00000000000000000000000000000000"
    "00000000000000000000000000000000"
    "00000000000000000000000000000000"
    "00000000000000000000000000000000"
    "0e000000140000000000000010000000"
    "14000000"
)


def write_placeholder_preview(cdx_path: Path, output_path: Path, **kwargs) -> str:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    requested = str(kwargs.get("format", "png")).lower()
    # Mocks should not depend on host tooling; treat auto as png unless path is .emf.
    if requested == "emf" or output_path.suffix.lower() == ".emf":
        output_path.write_bytes(_PLACEHOLDER_EMF)
        return "emf"
    Image.new("RGB", (32, 24), "white").save(output_path)
    return "png"


def write_placeholder_emf_many(jobs, **kwargs) -> list[str]:
    used: list[str] = []
    for cdx_path, emf_path in jobs:
        used.append(write_placeholder_preview(cdx_path, emf_path, format="emf"))
    return used


def patch_preview_renderers(monkeypatch) -> None:
    # Generated path renders via fill_parts; shell path via _shell.preview.
    monkeypatch.setattr(fill_parts, "render_preview", write_placeholder_preview)
    monkeypatch.setattr(fill_parts, "render_emf_many", write_placeholder_emf_many)
    monkeypatch.setattr(shell_preview, "render_preview", write_placeholder_preview)


def force_inkscape_available(monkeypatch, available: bool = True) -> None:
    monkeypatch.setattr(
        fill_parts,
        "preview_extension",
        lambda fmt: "emf" if _wants_emf(fmt, available) else "png",
    )
    monkeypatch.setattr(
        fill_parts,
        "resolve_preview_format",
        lambda fmt: "emf" if _wants_emf(fmt, available) else "png",
    )
    monkeypatch.setattr(
        shell_preview,
        "preview_extension",
        lambda fmt: "emf" if _wants_emf(fmt, available) else "png",
    )
    monkeypatch.setattr(
        "retain_cdx.render.preview.inkscape_available",
        lambda: available,
    )


def _wants_emf(fmt: str, available: bool) -> bool:
    value = str(fmt or "png").lower()
    if value == "emf":
        return True
    if value == "auto":
        return available
    return False


def add_inline_text(xlsx_path: Path, output: Path, cell_ref: str, text_value: str) -> None:
    package_dir = output.with_suffix("")
    package_dir.mkdir()
    extract_xlsx(xlsx_path, package_dir)

    sheet_path = package_dir / "xl" / "worksheets" / "sheet1.xml"
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    sheet_data = root.find("s:sheetData", NS)
    assert sheet_data is not None
    row_ref = "".join(ch for ch in cell_ref if ch.isdigit())
    row = sheet_data.find(f"s:row[@r='{row_ref}']", NS)
    if row is None:
        row = ET.SubElement(sheet_data, f"{{{NS_MAIN}}}row", r=row_ref)
    cell = ET.SubElement(row, f"{{{NS_MAIN}}}c", r=cell_ref, t="inlineStr")
    inline = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
    text = ET.SubElement(inline, f"{{{NS_MAIN}}}t")
    text.text = text_value
    tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)
    pack_xlsx(package_dir, output)


def read_zip_xml(xlsx_path: Path, member: str) -> ET.Element:
    with zipfile.ZipFile(xlsx_path) as z:
        return ET.fromstring(z.read(member))


def relationship_ids(root: ET.Element) -> list[str]:
    return [
        rel.get("Id", "")
        for rel in root.findall("rel:Relationship", NS)
    ]


def test_public_api_exports() -> None:
    import retain_cdx
    import retain_cdx.office.excel as excel
    from retain_cdx import DisplayOptions, Placement, PreviewOptions

    assert retain_cdx.__version__
    assert retain_cdx.ExcelCDXWriter is ExcelCDXWriter
    assert retain_cdx.ShellTemplate is ShellTemplate
    assert retain_cdx.StructureSource is StructureSource
    assert excel.ExcelCDXWriter is ExcelCDXWriter
    assert excel.DisplayOptions is DisplayOptions
    assert excel.Placement is Placement
    assert excel.PreviewOptions is PreviewOptions
    assert StructureSource.from_cdx("tests/fixtures/test.cdx").cdx_data


def test_extract_xlsx_rejects_unsafe_member(tmp_path: Path) -> None:
    unsafe = tmp_path / "unsafe.xlsx"
    with zipfile.ZipFile(unsafe, "w") as z:
        z.writestr("../escape.txt", "bad")

    try:
        extract_xlsx(unsafe, tmp_path / "out")
    except ValueError as exc:
        assert "Unsafe XLSX member path" in str(exc)
    else:
        raise AssertionError("unsafe XLSX member was not rejected")


def test_generated_insert_path_preserves_cells_and_applies_layout(
    tmp_path: Path, monkeypatch
) -> None:
    patch_preview_renderers(monkeypatch)

    target = tmp_path / "target.xlsx"
    add_inline_text(
        Path("tests/fixtures/blank.xlsx"),
        target,
        cell_ref="A1",
        text_value="keep me",
    )

    output = tmp_path / "output.xlsx"
    with ExcelCDXWriter(target) as writer:
        writer.insert_cdx(
            Path("tests/fixtures/test.cdx"),
            sheet="Sheet1",
            cell="B2",
            row_height=42.5,
            col_width=19.25,
        )
        writer.save(output)

    sheet_root = read_zip_xml(output, "xl/worksheets/sheet1.xml")
    kept_text = sheet_root.find(".//s:c[@r='A1']/s:is/s:t", NS)
    assert kept_text is not None
    assert kept_text.text == "keep me"

    inserted_row = sheet_root.find("s:sheetData/s:row[@r='2']", NS)
    assert inserted_row is not None
    assert inserted_row.get("ht") == "42.5"
    assert inserted_row.get("customHeight") == "1"
    inserted_col = sheet_root.find("s:cols/s:col[@min='2'][@max='2']", NS)
    assert inserted_col is not None
    assert inserted_col.get("width") == "19.25"
    assert inserted_col.get("customWidth") == "1"

    ole_objects = sheet_root.findall("s:oleObjects/mc:AlternateContent", NS)
    assert len(ole_objects) == 1


def test_generated_insert_path_creates_ole_and_preview(
    tmp_path: Path, monkeypatch
) -> None:
    patch_preview_renderers(monkeypatch)

    output = tmp_path / "output.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx(
            Path("tests/fixtures/test.cdx"),
            sheet="Sheet1",
            cell="A1",
        )
        writer.save(output)

    with zipfile.ZipFile(output) as z:
        names = set(z.namelist())
        assert "xl/embeddings/sheet1_retain_oleObject1.bin" in names
        assert "xl/media/sheet1_retain_image1.png" in names
        assert "xl/drawings/sheet1_retain_drawing.xml" in names
        assert "xl/drawings/sheet1_retain_vmlDrawing.vml" in names


def test_object_pr_pins_box_size_on_both_paths(tmp_path: Path, monkeypatch) -> None:
    """objectPr must carry autoPict="0"/defaultSize="0".

    Without autoPict="0" the spec default (true) tells Excel to size each object
    from its preview image — wrong for supersampled previews and expensive when
    a sheet holds thousands of objects.
    """
    patch_preview_renderers(monkeypatch)

    for name, kwargs in (
        ("generated.xlsx", {}),
        ("shell.xlsx", {"shell": ShellTemplate("tests/fixtures/shell_large.xlsx")}),
    ):
        output = tmp_path / name
        with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
            writer.insert_cdx(
                Path("tests/fixtures/test.cdx"), sheet="Sheet1", cell="A1", **kwargs
            )
            writer.save(output)

        with zipfile.ZipFile(output) as z:
            root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        object_prs = root.findall(".//{%s}objectPr" % NS_MAIN)
        assert object_prs, f"{name}: no objectPr found"
        for object_pr in object_prs:
            assert object_pr.get("autoPict") == "0", name
            assert object_pr.get("defaultSize") == "0", name


def test_generated_insert_many_creates_multiple_objects(
    tmp_path: Path, monkeypatch
) -> None:
    patch_preview_renderers(monkeypatch)

    output = tmp_path / "output.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx_many(
            [
                "tests/fixtures/test.cdx",
                Path("tests/fixtures/test.cdx"),
            ],
            sheet="Sheet1",
            start_cell="A1",
        )
        writer.save(output)

    with zipfile.ZipFile(output) as z:
        names = set(z.namelist())
        assert "xl/embeddings/sheet1_retain_oleObject1.bin" in names
        assert "xl/embeddings/sheet1_retain_oleObject2.bin" in names
        assert "xl/media/sheet1_retain_image1.png" in names
        assert "xl/media/sheet1_retain_image2.png" in names

    sheet_root = read_zip_xml(output, "xl/worksheets/sheet1.xml")
    ole_objects = sheet_root.findall("s:oleObjects/mc:AlternateContent", NS)
    assert len(ole_objects) == 2
    rows = [
        obj.find(
            "mc:Choice/s:oleObject/s:objectPr/s:anchor/s:from/xdr:row",
            NS,
        ).text
        for obj in ole_objects
    ]
    assert rows == ["0", "1"]

    drawing_root = read_zip_xml(output, "xl/drawings/sheet1_retain_drawing.xml")
    assert len(drawing_root.findall(".//xdr:twoCellAnchor", NS)) == 2

    sheet_rels = read_zip_xml(
        output,
        "xl/worksheets/_rels/sheet1.xml.rels",
    )
    relationships = sheet_rels.findall("rel:Relationship", NS)
    ole_rels = [
        rel
        for rel in relationships
        if rel.get("Type", "").endswith("/oleObject")
    ]
    image_rels = [
        rel
        for rel in relationships
        if rel.get("Type", "").endswith("/image")
    ]
    assert [rel.get("Target") for rel in ole_rels] == [
        "../embeddings/sheet1_retain_oleObject1.bin",
        "../embeddings/sheet1_retain_oleObject2.bin",
    ]
    assert [rel.get("Target") for rel in image_rels] == [
        "../media/sheet1_retain_image1.png",
        "../media/sheet1_retain_image2.png",
    ]


def test_repeated_insert_same_sheet_appends_objects(
    tmp_path: Path, monkeypatch
) -> None:
    patch_preview_renderers(monkeypatch)

    output = tmp_path / "output.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx("tests/fixtures/test.cdx", sheet="Sheet1", cell="A1")
        writer.insert_cdx("tests/fixtures/test.cdx", sheet="Sheet1", cell="A2")
        writer.save(output)

    with zipfile.ZipFile(output) as z:
        names = set(z.namelist())
        embeddings = sorted(
            name for name in names if name.startswith("xl/embeddings/sheet1_retain_")
        )
        media = sorted(name for name in names if name.startswith("xl/media/sheet1_retain_"))
    assert embeddings == [
        "xl/embeddings/sheet1_retain_oleObject1.bin",
        "xl/embeddings/sheet1_retain_oleObject1_2.bin",
    ]
    assert media == [
        "xl/media/sheet1_retain_image1.png",
        "xl/media/sheet1_retain_image1_2.png",
    ]

    sheet_root = read_zip_xml(output, "xl/worksheets/sheet1.xml")
    assert len(sheet_root.findall("s:drawing", NS)) == 1
    assert len(sheet_root.findall("s:legacyDrawing", NS)) == 1
    assert len(sheet_root.findall("s:cols/s:col[@min='1'][@max='1']", NS)) == 1
    ole_objects = sheet_root.findall("s:oleObjects/mc:AlternateContent", NS)
    assert len(ole_objects) == 2
    shape_ids = [
        obj.find("mc:Choice/s:oleObject", NS).get("shapeId")
        for obj in ole_objects
    ]
    assert shape_ids == ["1025", "1026"]

    sheet_rels = read_zip_xml(output, "xl/worksheets/_rels/sheet1.xml.rels")
    rids = relationship_ids(sheet_rels)
    assert len(rids) == len(set(rids))
    targets = [rel.get("Target") for rel in sheet_rels.findall("rel:Relationship", NS)]
    assert "../embeddings/sheet1_retain_oleObject1.bin" in targets
    assert "../embeddings/sheet1_retain_oleObject1_2.bin" in targets
    assert "../media/sheet1_retain_image1.png" in targets
    assert "../media/sheet1_retain_image1_2.png" in targets

    drawing_root = read_zip_xml(output, "xl/drawings/sheet1_retain_drawing.xml")
    assert len(drawing_root.findall(".//xdr:twoCellAnchor", NS)) == 2

    vml_root = read_zip_xml(output, "xl/drawings/sheet1_retain_vmlDrawing.vml")
    assert len(vml_root.findall("{urn:schemas-microsoft-com:vml}shape")) == 2
    vml_rels = read_zip_xml(
        output,
        "xl/drawings/_rels/sheet1_retain_vmlDrawing.vml.rels",
    )
    assert relationship_ids(vml_rels) == ["rId1", "rId2"]


def test_insert_preserves_existing_sheet_relationships(
    tmp_path: Path, monkeypatch
) -> None:
    patch_preview_renderers(monkeypatch)

    target = tmp_path / "target.xlsx"
    package_dir = tmp_path / "target_package"
    package_dir.mkdir()
    extract_xlsx(Path("tests/fixtures/blank.xlsx"), package_dir)
    rels_path = package_dir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels"
    rels_path.parent.mkdir(parents=True)
    rels_path.write_text(
        """<?xml version='1.0' encoding='UTF-8'?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="customRel" Type="http://example.com/custom" Target="../custom/item.xml" />
</Relationships>
""",
        encoding="utf-8",
    )
    pack_xlsx(package_dir, target)

    output = tmp_path / "output.xlsx"
    with ExcelCDXWriter(target) as writer:
        writer.insert_cdx("tests/fixtures/test.cdx", sheet="Sheet1", cell="A1")
        writer.save(output)

    sheet_rels = read_zip_xml(output, "xl/worksheets/_rels/sheet1.xml.rels")
    custom = [
        rel
        for rel in sheet_rels.findall("rel:Relationship", NS)
        if rel.get("Id") == "customRel"
    ]
    assert len(custom) == 1
    assert custom[0].get("Target") == "../custom/item.xml"


def test_shell_insert_preserves_existing_sheet_cells(tmp_path: Path) -> None:
    target = tmp_path / "target.xlsx"
    add_inline_text(
        Path("tests/fixtures/blank.xlsx"),
        target,
        cell_ref="A1",
        text_value="keep me",
    )

    output = tmp_path / "output.xlsx"
    with ExcelCDXWriter(target) as writer:
        writer.insert_cdx(
            Path("tests/fixtures/test.cdx"),
            sheet="Sheet1",
            cell="A1",
            shell=ShellTemplate(Path("tests/fixtures/shell_large.xlsx")),
        )
        writer.save(output)

    output_dir = tmp_path / "output_package"
    output_dir.mkdir()
    extract_xlsx(output, output_dir)
    output_sheet = ET.parse(output_dir / "xl" / "worksheets" / "sheet1.xml").getroot()

    kept_text = output_sheet.find(".//s:c[@r='A1']/s:is/s:t", NS)
    assert kept_text is not None
    assert kept_text.text == "keep me"

    output_row = output_sheet.find("s:sheetData/s:row[@r='1']", NS)
    assert output_row is not None
    assert output_row.get("ht") == "95.25"
    assert output_row.get("customHeight") == "1"


def test_shell_template_does_not_fit_cdx_by_default() -> None:
    assert ShellTemplate(Path("tests/fixtures/shell_large.xlsx")).fit_cdx is False


def test_shell_object_is_centered_in_target_cell(tmp_path: Path) -> None:
    output = tmp_path / "output.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx(
            Path("tests/fixtures/test.cdx"),
            sheet="Sheet1",
            cell="A1",
            shell=ShellTemplate(Path("tests/fixtures/shell_large.xlsx")),
            preview_format="png",
        )
        writer.save(output)

    with zipfile.ZipFile(output) as z:
        vml = z.read("xl/drawings/sheet1_retain_vmlDrawing.vml").decode("utf-8")
        drawing = z.read("xl/drawings/sheet1_retain_drawing.xml").decode("utf-8")

    # A1 is much larger than the tightened OLE box; centered placement must
    # use non-zero in-cell offsets (not stuck at 0,0).
    from_match = re.search(
        r"<xdr:from>\s*"
        r"<xdr:col>(\d+)</xdr:col>\s*"
        r"<xdr:colOff>(\d+)</xdr:colOff>\s*"
        r"<xdr:row>(\d+)</xdr:row>\s*"
        r"<xdr:rowOff>(\d+)</xdr:rowOff>",
        drawing,
        flags=re.S,
    )
    assert from_match is not None
    assert from_match.group(1) == "0"
    assert from_match.group(3) == "0"
    assert int(from_match.group(2)) > 0  # colOff
    assert int(from_match.group(4)) > 0  # rowOff

    style = re.search(
        r"margin-left:([0-9.]+)pt;margin-top:([0-9.]+)pt",
        vml,
    )
    assert style is not None
    assert float(style.group(1)) > 0
    assert float(style.group(2)) > 0


def test_shell_sheet_ole_anchor_matches_drawing_anchor(tmp_path: Path) -> None:
    """Regression: sheet objectPr anchor must track drawing/VML or Excel stretches."""
    output = tmp_path / "output.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx(
            Path("tests/fixtures/test.cdx"),
            sheet="Sheet1",
            cell="A1",
            shell=ShellTemplate(Path("tests/fixtures/shell_large.xlsx")),
            preview_format="png",
        )
        writer.save(output)

    with zipfile.ZipFile(output) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        drawing = z.read("xl/drawings/sheet1_retain_drawing.xml").decode("utf-8")

    def corner(text: str, open_tag: str) -> tuple[str, str, str, str]:
        frag = re.search(rf"<{open_tag}>[\s\S]*?</{open_tag}>", text)
        assert frag is not None
        body = frag.group(0)
        return (
            re.search(r"<xdr:col>(\d+)</xdr:col>", body).group(1),
            re.search(r"<xdr:colOff>(\d+)</xdr:colOff>", body).group(1),
            re.search(r"<xdr:row>(\d+)</xdr:row>", body).group(1),
            re.search(r"<xdr:rowOff>(\d+)</xdr:rowOff>", body).group(1),
        )

    assert corner(sheet, "from") == corner(drawing, "xdr:from")
    assert corner(sheet, "to") == corner(drawing, "xdr:to")


def test_shell_preview_matches_display_box_aspect_and_has_no_border(
    tmp_path: Path,
) -> None:
    output = tmp_path / "output.xlsx"
    shell = Path("tests/fixtures/shell_large.xlsx")
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx(
            Path("tests/fixtures/test.cdx"),
            sheet="Sheet1",
            cell="A1",
            shell=ShellTemplate(shell),
            preview_width=700,
            preview_height=700,
            preview_no_margin=True,
        )
        writer.save(output)

    with zipfile.ZipFile(output) as z:
        vml = z.read("xl/drawings/sheet1_retain_vmlDrawing.vml").decode("utf-8")
        drawing = z.read("xl/drawings/sheet1_retain_drawing.xml").decode("utf-8")
        preview = tmp_path / "preview.png"
        preview.write_bytes(z.read("xl/media/sheet1_retain_image1.png"))

    # Outer box is tightened to content; preview aspect must match *output* VML box.
    match = re.search(r"width:([0-9.]+)pt;.*?height:([0-9.]+)pt", vml, re.S)
    assert match is not None
    box_aspect = float(match.group(1)) / float(match.group(2))
    with Image.open(preview) as image:
        preview_aspect = image.width / image.height
    assert abs(preview_aspect - box_aspect) < 0.02

    # No black OLE chrome.
    assert 'stroked="f"' in vml or "stroked='f'" in vml
    assert 'stroked="t"' not in vml
    assert "<a:noFill/>" in drawing
    assert 'val="000000"' not in drawing or "<a:ln><a:noFill/></a:ln>" in drawing


def test_resolve_preview_format_defaults_and_auto(monkeypatch) -> None:
    monkeypatch.setattr("retain_cdx.render.preview.inkscape_available", lambda: False)
    assert resolve_preview_format("png") == "png"
    assert resolve_preview_format("auto") == "png"

    monkeypatch.setattr("retain_cdx.render.preview.inkscape_available", lambda: True)
    assert resolve_preview_format("auto") == "emf"
    assert resolve_preview_format("emf") == "emf"


def test_resolve_preview_format_emf_requires_inkscape(monkeypatch) -> None:
    monkeypatch.setattr("retain_cdx.render.preview.inkscape_available", lambda: False)
    with pytest.raises(SystemExit, match="Inkscape is required"):
        resolve_preview_format("emf")


def test_generated_insert_path_can_use_emf_preview(
    tmp_path: Path, monkeypatch
) -> None:
    force_inkscape_available(monkeypatch, True)
    patch_preview_renderers(monkeypatch)

    output = tmp_path / "output.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx(
            Path("tests/fixtures/test.cdx"),
            sheet="Sheet1",
            cell="A1",
            preview_format="emf",
        )
        writer.save(output)

    with zipfile.ZipFile(output) as z:
        names = set(z.namelist())
        assert "xl/media/sheet1_retain_image1.emf" in names
        assert "xl/media/sheet1_retain_image1.png" not in names
        emf_bytes = z.read("xl/media/sheet1_retain_image1.emf")
    assert emf_bytes.startswith(b"\x01\x00\x00\x00")

    sheet_rels = read_zip_xml(output, "xl/worksheets/_rels/sheet1.xml.rels")
    image_targets = [
        rel.get("Target")
        for rel in sheet_rels.findall("rel:Relationship", NS)
        if rel.get("Type", "").endswith("/image")
    ]
    assert image_targets == ["../media/sheet1_retain_image1.emf"]


def test_shell_insert_can_use_emf_preview(tmp_path: Path, monkeypatch) -> None:
    force_inkscape_available(monkeypatch, True)
    patch_preview_renderers(monkeypatch)

    output = tmp_path / "output.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx(
            Path("tests/fixtures/test.cdx"),
            sheet="Sheet1",
            cell="A1",
            shell=ShellTemplate(Path("tests/fixtures/shell_large.xlsx")),
            preview_format="emf",
        )
        writer.save(output)

    with zipfile.ZipFile(output) as z:
        names = set(z.namelist())
        assert "xl/media/sheet1_retain_image1.emf" in names
        # Default shell had image1.emf; after replace it should still be emf, not png.
        assert not any(name.endswith(".png") and "media" in name for name in names)
        assert z.read("xl/media/sheet1_retain_image1.emf").startswith(b"\x01\x00\x00\x00")


@pytest.mark.skipif(not inkscape_available(), reason="Inkscape not installed")
def test_render_emf_real_pipeline(tmp_path: Path) -> None:
    from retain_cdx.render.preview import render_emf

    emf_path = tmp_path / "structure.emf"
    render_emf(Path("tests/fixtures/test.cdx"), emf_path)
    assert emf_path.exists()
    assert emf_path.stat().st_size > 100
    assert emf_path.read_bytes()[40:44] == b" EMF" or emf_path.read_bytes()[:4] == b"\x01\x00\x00\x00"


@pytest.mark.skipif(not inkscape_available(), reason="Inkscape not installed")
def test_svg_to_emf_batch_shell(tmp_path: Path) -> None:
    from retain_cdx.render.preview import render_svg, svg_to_emf_batch

    cdx = Path("tests/fixtures/test.cdx")
    pairs = []
    for i in range(3):
        svg = tmp_path / f"m{i}.svg"
        emf = tmp_path / f"m{i}.emf"
        render_svg(cdx, svg, backend="auto", style="acs1996", mono=True)
        pairs.append((svg, emf))
    svg_to_emf_batch(pairs, chunk_size=2)
    for _, emf in pairs:
        assert emf.exists() and emf.stat().st_size > 100


def test_render_png_defaults_to_monochrome(tmp_path: Path) -> None:
    from retain_cdx.render.preview import _obabel_cmd, render_png, resolve_preview_backend

    cdx = Path("tests/fixtures/test.cdx")
    cmd = _obabel_cmd(
        cdx,
        tmp_path / "out.png",
        output_format="png",
        width=None,
        height=None,
        image_size=None,
        thick_lines=False,
        no_margin=True,
        no_title=True,
        mono=True,
    )
    assert "-xu" in cmd and "-xU" in cmd

    color_cmd = _obabel_cmd(
        cdx,
        tmp_path / "color.png",
        output_format="png",
        width=None,
        height=None,
        image_size=None,
        thick_lines=False,
        no_margin=True,
        no_title=True,
        mono=False,
    )
    assert "-xu" not in color_cmd and "-xU" not in color_cmd

    png_path = tmp_path / "mono.png"
    used = render_png(cdx, png_path, no_margin=True, backend="auto")
    assert used in {"rdkit", "obabel"}
    if resolve_preview_backend("auto") == "rdkit":
        assert used == "rdkit"
    image = Image.open(png_path).convert("RGB")
    chromatic = [
        pixel
        for pixel in image.getdata()
        if max(pixel) - min(pixel) > 15 and not (pixel[0] > 245 and pixel[1] > 245 and pixel[2] > 245)
    ]
    assert chromatic == []


def test_rdkit_acs1996_backend(tmp_path: Path) -> None:
    from retain_cdx.render.preview import (
        pycdxml_available,
        rdkit_available,
        render_png,
        resolve_preview_backend,
    )

    if not (rdkit_available() and pycdxml_available()):
        pytest.skip("RDKit + pycdxml required")

    assert resolve_preview_backend("auto") == "rdkit"
    png_path = tmp_path / "acs.png"
    used = render_png(
        Path("tests/fixtures/test.cdx"),
        png_path,
        backend="rdkit",
        style="acs1996",
        mono=True,
    )
    assert used == "rdkit"
    assert png_path.exists() and png_path.stat().st_size > 100
    image = Image.open(png_path).convert("RGB")
    chromatic = [
        pixel
        for pixel in image.getdata()
        if max(pixel) - min(pixel) > 15 and not (pixel[0] > 245 and pixel[1] > 245 and pixel[2] > 245)
    ]
    assert chromatic == []
