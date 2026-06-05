from __future__ import annotations

import tempfile
from pathlib import Path

from retain_cdx.ole.embed import write_cdx_to_ole
from retain_cdx.office.excel._ooxml import (
    build_sheet_relationships,
    build_vml_relationships,
    ensure_png_content_type,
    extract_xlsx,
    pack_xlsx,
)
from retain_cdx.render import render_png

from .anchors import compute_anchors


def build_generated_source_workbook(
    cdx_paths: list[Path],
    output: Path,
    *,
    col: str,
    start_row: int,
    margin_left: str,
    margin_top: str,
    width: str,
    height: str,
    row_height: float,
    preview_width: int | None = None,
    preview_height: int | None = None,
    preview_thick_lines: bool = False,
    preview_no_margin: bool = False,
) -> None:
    bundled_template = Path(__file__).resolve().parents[4] / "template.xlsx"
    anchors = compute_anchors(
        len(cdx_paths),
        col=col,
        start_row=start_row,
        margin_left=margin_left,
        margin_top=margin_top,
        width=width,
        height=height,
    )
    build_template_source_workbook(
        bundled_template,
        cdx_paths,
        output,
        anchors=anchors,
        row_height=row_height,
        preview_width=preview_width,
        preview_height=preview_height,
        preview_thick_lines=preview_thick_lines,
        preview_no_margin=preview_no_margin,
    )


def build_template_source_workbook(
    template: Path,
    cdx_paths: list[Path],
    output: Path,
    *,
    anchors,
    row_height: float,
    preview_width: int | None = None,
    preview_height: int | None = None,
    preview_thick_lines: bool = False,
    preview_no_margin: bool = False,
) -> None:
    from .template_namespaces import register_namespaces
    from .template_xml import (
        extract_base_anchor,
        update_drawing,
        update_ole_objects,
        update_row_heights,
        update_sheet_labels,
        update_vml,
        write_shared_strings,
    )

    register_namespaces()
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        extract_xlsx(template, tmpdir)

        ole_dir = tmpdir / "xl" / "embeddings"
        ole_dir.mkdir(parents=True, exist_ok=True)
        base_ole_path = ole_dir / "oleObject1.bin"
        if not base_ole_path.exists():
            raise ValueError(f"OLE object not found: {base_ole_path}")
        base_ole_bytes = base_ole_path.read_bytes()

        sheet_path = tmpdir / "xl" / "worksheets" / "sheet1.xml"
        drawing_path = tmpdir / "xl" / "drawings" / "drawing1.xml"
        vml_path = tmpdir / "xl" / "drawings" / "vmlDrawing1.vml"
        sheet_rels_path = tmpdir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels"
        vml_rels_path = tmpdir / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels"
        shared_strings_path = tmpdir / "xl" / "sharedStrings.xml"
        content_types_path = tmpdir / "[Content_Types].xml"

        base_shape_id, from_row, to_row = extract_base_anchor(sheet_path)
        labels = [path.stem for path in cdx_paths]
        total_rows = max(anchor.to_pt.row + 1 for anchor in anchors)
        update_sheet_labels(sheet_path, labels, total_rows=total_rows)
        update_row_heights(sheet_path, row_height, total_rows)
        write_shared_strings(shared_strings_path, labels)

        ole_targets = []
        image_targets = []
        for idx, cdx_path in enumerate(cdx_paths, start=1):
            ole_path = ole_dir / f"oleObject{idx}.bin"
            if idx == 1:
                if not ole_path.exists():
                    raise ValueError(f"OLE object not found: {ole_path}")
            else:
                ole_path.write_bytes(base_ole_bytes)

            write_cdx_to_ole(ole_path, cdx_path, stream_name="CONTENTS")

            image_name = f"image{idx}.png"
            image_path = tmpdir / "xl" / "media" / image_name
            image_path.parent.mkdir(parents=True, exist_ok=True)
            render_png(
                cdx_path,
                image_path,
                width=preview_width,
                height=preview_height,
                thick_lines=preview_thick_lines,
                no_margin=preview_no_margin,
            )

            ole_targets.append(f"../embeddings/oleObject{idx}.bin")
            image_targets.append(f"../media/{image_name}")

        emf_path = tmpdir / "xl" / "media" / "image1.emf"
        if emf_path.exists():
            emf_path.unlink()

        rel_pairs = build_sheet_relationships(sheet_rels_path, ole_targets, image_targets)
        vml_rids = build_vml_relationships(vml_rels_path, image_targets)
        update_ole_objects(sheet_path, base_shape_id, from_row, to_row, rel_pairs, anchors=anchors)
        update_drawing(drawing_path, base_shape_id, from_row, to_row, len(cdx_paths), anchors=anchors)
        update_vml(
            vml_path,
            base_shape_id,
            from_row,
            to_row,
            vml_rids,
            anchors=anchors,
            row_height=row_height,
        )
        ensure_png_content_type(content_types_path)
        pack_xlsx(tmpdir, output)
