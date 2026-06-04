from __future__ import annotations

from pathlib import Path

from retain_cdx.ole.embed import write_cdx_to_ole

from .package import extract_xlsx, pack_xlsx
from scripts.xlsx.position import compute_anchors


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

        write_cdx_to_ole(embeddings[0], cdx_path, stream_name="CONTENTS")
        pack_xlsx(package_dir, output)
