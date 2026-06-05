from __future__ import annotations

from pathlib import Path

from retain_cdx.ole.embed import write_cdx_to_ole

from retain_cdx.office.excel._ooxml import extract_xlsx, pack_xlsx
from retain_cdx.office.excel._shell import (
    patch_shell_geometry,
    read_vml_size_px,
    replace_shell_preview,
)
from retain_cdx.office.excel._utils import POINTS_PER_PIXEL

from .generated import build_generated_source_workbook


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
    # XLSX anchors use 96-dpi pixels. The existing generator accepts metric
    # lengths, so convert pixels to points for a stable display box.
    offset_x_pt = offset_x_px * 0.75
    offset_y_pt = offset_y_px * 0.75
    width_pt = width_px * 0.75
    height_pt = height_px * 0.75
    build_generated_source_workbook(
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
    preview_width: int | None = None,
    preview_height: int | None = None,
    preview_thick_lines: bool = False,
    preview_no_margin: bool = True,
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
        replace_shell_preview(
            cdx_path,
            package_dir,
            width=preview_width,
            height=preview_height,
            thick_lines=preview_thick_lines,
            no_margin=preview_no_margin,
        )
        if start_col is not None and start_row is not None:
            patch_shell_geometry(
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
    shell_width_px, shell_height_px = read_vml_size_px(vml_text)
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
