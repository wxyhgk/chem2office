from __future__ import annotations

from pathlib import Path
import re

from retain_cdx.ole.embed import write_cdx_to_ole

from .units import EMU_PER_PIXEL, POINTS_PER_PIXEL
from .package import extract_xlsx, pack_xlsx


def build_source_ole_workbook(
    cdx_paths: list[Path],
    output: Path,
    *,
    col_letter: str,
    start_row: int,
    offset_x_px: int,
    offset_y_px: int,
    width_px: int,
    height_px: int,
    row_height: float,
    preview_width: int | None,
    preview_height: int | None,
    preview_thick_lines: bool,
    preview_no_margin: bool,
) -> None:
    from scripts.cli.inject_cdx_into_xlsx import _build_generate_mode

    # XLSX anchors use 96-dpi pixels. The existing generator accepts metric
    # lengths, so convert pixels to points for a stable display box.
    offset_x_pt = offset_x_px * 0.75
    offset_y_pt = offset_y_px * 0.75
    width_pt = width_px * 0.75
    height_pt = height_px * 0.75
    _build_generate_mode(
        cdx_paths,
        output,
        col=col_letter,
        start_row=start_row,
        margin_left=f"{offset_x_pt}pt",
        margin_top=f"{offset_y_pt}pt",
        width=f"{width_pt}pt",
        height=f"{height_pt}pt",
        from_spec=None,
        to_spec=None,
        row_height=row_height,
        preview_width=preview_width,
        preview_height=preview_height,
        preview_thick_lines=preview_thick_lines,
        preview_no_margin=preview_no_margin,
    )


def build_source_workbook_from_shell(
    cdx_path: Path,
    shell_xlsx: Path,
    output: Path,
    *,
    start_col: int | None = None,
    start_row: int | None = None,
    offset_x_px: int = 0,
    offset_y_px: int = 0,
    width_px: int | None = None,
    height_px: int | None = None,
    preserve_aspect: bool = True,
    fit_cdx: bool = False,
) -> None:
    """Build a one-object source workbook by reusing a native OLE shell."""
    import tempfile

    with tempfile.TemporaryDirectory(prefix="retain_cdx_shell_") as tmp:
        package_dir = Path(tmp) / "package"
        package_dir.mkdir()
        extract_xlsx(shell_xlsx, package_dir)

        embeddings = sorted((package_dir / "xl" / "embeddings").glob("*.bin"))
        if not embeddings:
            raise ValueError(f"Shell template has no OLE embedding: {shell_xlsx}")

        if fit_cdx:
            cdx_path = _fit_cdx_to_shell(cdx_path, package_dir)
        write_cdx_to_ole(embeddings[0], cdx_path, stream_name="CONTENTS")
        if start_col is not None and start_row is not None:
            _patch_shell_geometry(
                package_dir,
                start_col=start_col,
                start_row=start_row,
                offset_x_px=offset_x_px,
                offset_y_px=offset_y_px,
                width_px=width_px,
                height_px=height_px,
                preserve_aspect=preserve_aspect,
            )
        pack_xlsx(package_dir, output)


def _fit_cdx_to_shell(cdx_path: Path, package_dir: Path) -> Path:
    from retain_cdx.structure import StructureSource

    vml_path = package_dir / "xl" / "drawings" / "vmlDrawing1.vml"
    vml_text = vml_path.read_text(encoding="utf-8")
    shell_width_px, shell_height_px = _read_vml_size_px(vml_text)
    if shell_width_px <= 0 or shell_height_px <= 0:
        return cdx_path
    fitted = StructureSource.from_cdx(cdx_path).normalized(
        target_aspect=shell_width_px / shell_height_px,
        target_width=shell_width_px * POINTS_PER_PIXEL,
        target_height=shell_height_px * POINTS_PER_PIXEL,
        fill_ratio=0.8,
    )
    output = cdx_path.with_name(f"{cdx_path.stem}_fit_shell.cdx")
    fitted.write_cdx(output)
    return output


def _patch_shell_geometry(
    package_dir: Path,
    *,
    start_col: int,
    start_row: int,
    offset_x_px: int,
    offset_y_px: int,
    width_px: int | None,
    height_px: int | None,
    preserve_aspect: bool,
) -> None:
    sheet_path = package_dir / "xl" / "worksheets" / "sheet1.xml"
    drawing_path = package_dir / "xl" / "drawings" / "drawing1.xml"
    vml_path = package_dir / "xl" / "drawings" / "vmlDrawing1.vml"

    vml_text = vml_path.read_text(encoding="utf-8")
    shell_width_px, shell_height_px = _read_vml_size_px(vml_text)
    target_width_px, target_height_px = _target_size(
        shell_width_px,
        shell_height_px,
        width_px=width_px,
        height_px=height_px,
        preserve_aspect=preserve_aspect,
    )

    from_col, from_col_off_px = _split_col(start_col - 1, offset_x_px)
    from_row, from_row_off_px = _split_row(start_row - 1, offset_y_px)
    to_col, to_col_off_px = _advance_cols(from_col, from_col_off_px + target_width_px)
    to_row, to_row_off_px = _advance_rows(from_row, from_row_off_px + target_height_px)

    from_xml = _anchor_corner_xml(from_col, from_col_off_px, from_row, from_row_off_px)
    to_xml = _anchor_corner_xml(to_col, to_col_off_px, to_row, to_row_off_px)
    for path in [sheet_path, drawing_path]:
        _replace_first_corner(path, "from", from_xml)
        _replace_first_corner(path, "to", to_xml)

    left_px = _absolute_x_px(from_col, from_col_off_px)
    top_px = _absolute_y_px(from_row, from_row_off_px)
    style = (
        "position:absolute;"
        f"margin-left:{_px_to_pt(left_px):g}pt;"
        f"margin-top:{_px_to_pt(top_px):g}pt;"
        f"width:{_px_to_pt(target_width_px):g}pt;"
        f"height:{_px_to_pt(target_height_px):g}pt;"
        "z-index:1"
    )
    if "mso-wrap-style:tight" in vml_text:
        style += ";mso-wrap-style:tight"
    vml_text = re.sub(
        r"(<v:shape\b[^>]*\sstyle=)(['\"])(.*?)\2",
        lambda match: f"{match.group(1)}{match.group(2)}{style}{match.group(2)}",
        vml_text,
        count=1,
        flags=re.S,
    )
    vml_anchor = (
        f"{from_col}, {from_col_off_px}, {from_row}, {from_row_off_px}, "
        f"{to_col}, {to_col_off_px}, {to_row}, {to_row_off_px}"
    )
    vml_text = re.sub(
        r"<x:Anchor>.*?</x:Anchor>",
        f"<x:Anchor>\n    {vml_anchor}</x:Anchor>",
        vml_text,
        count=1,
        flags=re.S,
    )
    vml_path.write_text(vml_text, encoding="utf-8")


def _read_vml_size_px(vml_text: str) -> tuple[int, int]:
    match = re.search(r"width:([0-9.]+)pt;.*?height:([0-9.]+)pt", vml_text, re.S)
    if not match:
        raise ValueError("Shell VML shape has no width/height style.")
    return round(float(match.group(1)) / POINTS_PER_PIXEL), round(
        float(match.group(2)) / POINTS_PER_PIXEL
    )


def _target_size(
    shell_width_px: int,
    shell_height_px: int,
    *,
    width_px: int | None,
    height_px: int | None,
    preserve_aspect: bool,
) -> tuple[int, int]:
    if width_px is None and height_px is None:
        return shell_width_px, shell_height_px
    if not preserve_aspect:
        return width_px or shell_width_px, height_px or shell_height_px

    aspect = shell_width_px / shell_height_px
    if width_px is not None and height_px is not None:
        scale = min(width_px / shell_width_px, height_px / shell_height_px)
        return max(1, round(shell_width_px * scale)), max(1, round(shell_height_px * scale))
    if width_px is not None:
        return width_px, max(1, round(width_px / aspect))
    height = height_px or shell_height_px
    return max(1, round(height * aspect)), height


def _replace_first_corner(path: Path, corner: str, replacement: str) -> None:
    text = path.read_text(encoding="utf-8")
    text = re.sub(
        rf"<xdr:{corner}>.*?</xdr:{corner}>",
        f"<xdr:{corner}>{replacement}</xdr:{corner}>",
        text,
        count=1,
        flags=re.S,
    )
    path.write_text(text, encoding="utf-8")


def _anchor_corner_xml(col: int, col_off_px: int, row: int, row_off_px: int) -> str:
    return (
        f"<xdr:col>{col}</xdr:col>"
        f"<xdr:colOff>{col_off_px * EMU_PER_PIXEL}</xdr:colOff>"
        f"<xdr:row>{row}</xdr:row>"
        f"<xdr:rowOff>{row_off_px * EMU_PER_PIXEL}</xdr:rowOff>"
    )


def _split_col(col: int, offset_px: int) -> tuple[int, int]:
    while offset_px >= _col_width_px(col):
        offset_px -= _col_width_px(col)
        col += 1
    return col, offset_px


def _split_row(row: int, offset_px: int) -> tuple[int, int]:
    while offset_px >= _row_height_px(row):
        offset_px -= _row_height_px(row)
        row += 1
    return row, offset_px


def _advance_cols(col: int, width_px: int) -> tuple[int, int]:
    return _split_col(col, width_px)


def _advance_rows(row: int, height_px: int) -> tuple[int, int]:
    return _split_row(row, height_px)


def _absolute_x_px(col: int, col_off_px: int) -> int:
    return sum(_col_width_px(idx) for idx in range(col)) + col_off_px


def _absolute_y_px(row: int, row_off_px: int) -> int:
    return sum(_row_height_px(idx) for idx in range(row)) + row_off_px


def _col_width_px(_col: int) -> int:
    return 64


def _row_height_px(_row: int) -> int:
    return 19


def _px_to_pt(px: int) -> float:
    return px * POINTS_PER_PIXEL
