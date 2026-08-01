"""Two-point anchor placement (``Placement.to_cell``).

The display box spans ``cell + offset → to_cell + to_offset`` and is resolved
against the target sheet's real column widths / row heights (Excel defaults
64px/19px for undeclared ones), overriding ``width_px``/``height_px``.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest
from PIL import Image

import retain_cdx.office.excel._shell.preview as shell_preview
import retain_cdx.office.excel._source.fill_parts as fill_parts
from retain_cdx import ExcelCDXWriter, Placement, ShellTemplate
from retain_cdx.office._package import extract_ooxml, pack_ooxml
from retain_cdx.office.excel._utils import EMU_PER_PIXEL

CDX = Path("tests/fixtures/test.cdx")
BLANK = Path("tests/fixtures/blank.xlsx")
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


# --- placeholder renderers (same pattern as test_excel_shell_layout) ---

_PLACEHOLDER_EMF = bytes.fromhex(
    "01000000d80000000000000000000000"
    "c7090000c70900000000000000000000"
    "b3140000b314000020454d4600000100"
    "1400000000000000000000000000000000000000"
)


def write_placeholder_preview(cdx_path: Path, output_path: Path, **kwargs) -> str:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    requested = str(kwargs.get("format", "png")).lower()
    if requested == "emf" or output_path.suffix.lower() == ".emf":
        output_path.write_bytes(_PLACEHOLDER_EMF)
        return "emf"
    Image.new("RGB", (32, 24), "white").save(output_path)
    return "png"


def write_placeholder_emf_many(jobs, **kwargs) -> list[str]:
    return [
        write_placeholder_preview(cdx_path, emf_path, format="emf")
        for cdx_path, emf_path in jobs
    ]


@pytest.fixture()
def placeholder_previews(monkeypatch) -> None:
    monkeypatch.setattr(fill_parts, "render_preview", write_placeholder_preview)
    monkeypatch.setattr(fill_parts, "render_emf_many", write_placeholder_emf_many)
    monkeypatch.setattr(shell_preview, "render_preview", write_placeholder_preview)


# --- helpers ---


def default_grid_workbook(tmp_path: Path) -> Path:
    """blank.xlsx without sheetFormatPr → grid falls back to 64px/19px."""
    target = tmp_path / "default_grid.xlsx"
    package = tmp_path / "default_grid_pkg"
    package.mkdir()
    extract_ooxml(BLANK, package)
    sheet = package / "xl" / "worksheets" / "sheet1.xml"
    text = sheet.read_text(encoding="utf-8")
    stripped = re.sub(r"<sheetFormatPr[^>]*/>", "", text)
    assert stripped != text  # fixture must have carried a sheetFormatPr
    sheet.write_text(stripped, encoding="utf-8")
    pack_ooxml(package, target)
    return target


def drawing_anchors(output: Path) -> list[dict[str, tuple[int, int, int, int]]]:
    """Extract (col, colOff, row, rowOff) for each drawing from/to corner."""
    with zipfile.ZipFile(output) as z:
        drawing = z.read("xl/drawings/sheet1_retain_drawing.xml").decode("utf-8")

    corner_re = (
        r"<xdr:col>(\d+)</xdr:col>\s*"
        r"<xdr:colOff>(\d+)</xdr:colOff>\s*"
        r"<xdr:row>(\d+)</xdr:row>\s*"
        r"<xdr:rowOff>(\d+)</xdr:rowOff>"
    )
    anchors = []
    for block in re.findall(
        r"<xdr:twoCellAnchor[\s\S]*?</xdr:twoCellAnchor>", drawing
    ):
        from_match = re.search(rf"<xdr:from>\s*{corner_re}", block)
        to_match = re.search(rf"<xdr:to>\s*{corner_re}", block)
        assert from_match and to_match
        anchors.append(
            {
                "from": tuple(int(g) for g in from_match.groups()),
                "to": tuple(int(g) for g in to_match.groups()),
            }
        )
    return anchors


def anchor_size_emu(
    anchor: dict[str, tuple[int, int, int, int]], workbook: Path | None = None
) -> tuple[int, int]:
    """Box size in EMU, measured across the grid.

    Anchors are normalised, so the to-corner sits in whichever cell the box
    actually reaches; the span has to be walked rather than subtracted.
    """
    from_c, to_c = anchor["from"], anchor["to"]
    if to_c[0] == from_c[0] and to_c[2] == from_c[2]:
        return to_c[1] - from_c[1], to_c[3] - from_c[3]

    assert workbook is not None, "cross-cell anchor needs the workbook grid"
    metrics = grid_metrics(workbook)
    width = (
        metrics.cell_origin_x_px(to_c[0]) + to_c[1] // EMU_PER_PIXEL
    ) - (metrics.cell_origin_x_px(from_c[0]) + from_c[1] // EMU_PER_PIXEL)
    height = (
        metrics.cell_origin_y_px(to_c[2]) + to_c[3] // EMU_PER_PIXEL
    ) - (metrics.cell_origin_y_px(from_c[2]) + from_c[3] // EMU_PER_PIXEL)
    return width * EMU_PER_PIXEL, height * EMU_PER_PIXEL


def grid_metrics(workbook: Path):
    """ExcelGridMetrics for the saved workbook's first sheet."""
    import tempfile

    from retain_cdx.office._package import extract_ooxml
    from retain_cdx.office.excel._grid import ExcelGridMetrics

    with tempfile.TemporaryDirectory() as tmp:
        package = Path(tmp) / "pkg"
        package.mkdir()
        extract_ooxml(workbook, package)
        return ExcelGridMetrics.from_worksheet(
            package / "xl" / "worksheets" / "sheet1.xml"
        )


def insert_two_point(
    workbook: Path,
    output: Path,
    placement: Placement,
    **kwargs,
) -> None:
    with ExcelCDXWriter(workbook) as writer:
        writer.insert_cdx(CDX, sheet="Sheet1", placement=placement, **kwargs)
        writer.save(output)


# --- tests ---


def test_two_point_spans_custom_grid(tmp_path: Path, placeholder_previews) -> None:
    """(a) Box matches the real grid span with custom col widths / row heights."""
    output = tmp_path / "out.xlsx"
    with ExcelCDXWriter(BLANK) as writer:
        # Custom grid: A=20ch → 145px, B=10ch → 75px; row1=30pt → 40px,
        # row2=15pt → 20px (px = int(width*7+5), px = pt/0.75).
        writer.set_column_width("Sheet1", "A", 20)
        writer.set_column_width("Sheet1", "B", 10)
        writer.set_row_height("Sheet1", 1, 30)
        writer.set_row_height("Sheet1", 2, 15)
        writer.insert_cdx(
            CDX,
            sheet="Sheet1",
            placement=Placement("A1", to_cell="C3"),
        )
        writer.save(output)

    (anchor,) = drawing_anchors(output)
    assert anchor["from"] == (0, 0, 0, 0)
    width_emu, height_emu = anchor_size_emu(anchor, output)
    assert width_emu == (145 + 75) * EMU_PER_PIXEL
    assert height_emu == (40 + 20) * EMU_PER_PIXEL

    # Two-point placement must not reshape the grid it was resolved against.
    with zipfile.ZipFile(output) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    col_a = re.search(r'<col[^>]*min="1"[^>]*/>', sheet)
    assert col_a is not None and 'width="20"' in col_a.group(0)
    row_1 = re.search(r'<row[^>]*r="1"[^>]*>', sheet)
    assert row_1 is not None and 'ht="30"' in row_1.group(0)


def test_to_cell_overrides_width_px(tmp_path: Path, placeholder_previews) -> None:
    """(b) to_cell wins over explicit width_px/height_px."""
    output = tmp_path / "out.xlsx"
    insert_two_point(
        default_grid_workbook(tmp_path),
        output,
        Placement("A1", to_cell="B2"),
        width_px=999,
        height_px=888,
    )

    (anchor,) = drawing_anchors(output)
    width_emu, height_emu = anchor_size_emu(anchor, output)
    assert width_emu == 64 * EMU_PER_PIXEL != 999 * EMU_PER_PIXEL
    assert height_emu == 19 * EMU_PER_PIXEL != 888 * EMU_PER_PIXEL


def test_default_grid_resolves_with_excel_defaults(
    tmp_path: Path, placeholder_previews
) -> None:
    """(c) Undeclared columns/rows fall back to 64px/19px."""
    output = tmp_path / "out.xlsx"
    insert_two_point(
        default_grid_workbook(tmp_path),
        output,
        Placement("B2", to_cell="D5"),
    )

    (anchor,) = drawing_anchors(output)
    assert anchor["from"] == (1, 0, 1, 0)
    width_emu, height_emu = anchor_size_emu(anchor, output)
    assert width_emu == 2 * 64 * EMU_PER_PIXEL  # cols B..C
    assert height_emu == 3 * 19 * EMU_PER_PIXEL  # rows 2..4


def test_two_point_offsets_shift_both_corners(
    tmp_path: Path, placeholder_previews
) -> None:
    """Offsets use the same px/pt/in/mm parsing as the single-anchor offsets."""
    output = tmp_path / "out.xlsx"
    insert_two_point(
        default_grid_workbook(tmp_path),
        output,
        Placement(
            "A1",
            offset_x="2mm",  # → 8 px
            offset_y="10pt",  # → 13 px
            to_cell="B2",
            to_offset_x=7,
            to_offset_y="5px",
        ),
    )

    (anchor,) = drawing_anchors(output)
    assert anchor["from"] == (0, 8 * EMU_PER_PIXEL, 0, 13 * EMU_PER_PIXEL)
    width_emu, height_emu = anchor_size_emu(anchor, output)
    assert width_emu == (64 + 7 - 8) * EMU_PER_PIXEL
    assert height_emu == (19 + 5 - 13) * EMU_PER_PIXEL


def test_two_point_applies_per_row_for_many(
    tmp_path: Path, placeholder_previews
) -> None:
    """insert_cdx_many keeps the resolved box size per row (like width_px)."""
    output = tmp_path / "out.xlsx"
    with ExcelCDXWriter(default_grid_workbook(tmp_path)) as writer:
        writer.insert_cdx_many(
            [CDX, CDX],
            sheet="Sheet1",
            start_cell="A1",
            placement=Placement("A1", to_cell="C3"),
        )
        writer.save(output)

    anchors = drawing_anchors(output)
    assert len(anchors) == 2
    assert [a["from"][2] for a in anchors] == [0, 1]
    for anchor in anchors:
        width_emu, height_emu = anchor_size_emu(anchor, output)
        assert width_emu == 2 * 64 * EMU_PER_PIXEL
        assert height_emu == 2 * 19 * EMU_PER_PIXEL


def test_invalid_to_cell_before_from_cell_raises(
    tmp_path: Path, placeholder_previews
) -> None:
    """(d) A to_cell above/left of the from cell is a clear error."""
    with ExcelCDXWriter(BLANK) as writer:
        with pytest.raises(ValueError, match="Two-point placement"):
            writer.insert_cdx(
                CDX,
                sheet="Sheet1",
                placement=Placement("C3", to_cell="A1"),
            )


def test_malformed_to_cell_raises(tmp_path: Path, placeholder_previews) -> None:
    with ExcelCDXWriter(BLANK) as writer:
        with pytest.raises(ValueError, match="Invalid cell reference"):
            writer.insert_cdx(
                CDX,
                sheet="Sheet1",
                placement=Placement("A1", to_cell="notacell"),
            )


def test_two_point_feeds_shell_geometry(tmp_path: Path, placeholder_previews) -> None:
    """Shell path: resolved box size reaches patch_shell_geometry (VML size)."""
    output = tmp_path / "out.xlsx"
    insert_two_point(
        default_grid_workbook(tmp_path),
        output,
        Placement("A1", to_cell="C3"),  # 128 x 38 px on the 64/19 default grid
        shell=ShellTemplate(
            Path("tests/fixtures/shell_large.xlsx"),
            preserve_aspect=False,
        ),
        preview_format="png",
    )

    with zipfile.ZipFile(output) as z:
        vml = z.read("xl/drawings/sheet1_retain_vmlDrawing.vml").decode("utf-8")
    style = re.search(r"width:([0-9.]+)pt;.*?height:([0-9.]+)pt", vml, re.S)
    assert style is not None
    assert float(style.group(1)) == pytest.approx(2 * 64 * 0.75)  # 96pt
    assert float(style.group(2)) == pytest.approx(2 * 19 * 0.75)  # 28.5pt


def test_single_anchor_placement_unchanged(tmp_path: Path, placeholder_previews) -> None:
    """Backward compat: no to_cell → width_px/height_px still size the box."""
    output = tmp_path / "out.xlsx"
    insert_two_point(
        default_grid_workbook(tmp_path),
        output,
        Placement("A1"),
        width_px=145,
        height_px=100,
    )

    (anchor,) = drawing_anchors(output)
    width_emu, height_emu = anchor_size_emu(anchor, output)
    assert width_emu == 145 * EMU_PER_PIXEL
    assert height_emu == 100 * EMU_PER_PIXEL
