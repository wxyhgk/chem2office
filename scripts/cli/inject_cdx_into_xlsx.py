#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
import tempfile
import zipfile
from pathlib import Path

from scripts.ole.inject import inject_cdx_stream, patch_direntry_size
from scripts.ole.ole_builder import build_chemdraw_ole
from scripts.render.preview import render_png
from scripts.xlsx.namespaces import register_namespaces
from scripts.xlsx.position import compute_anchors
from scripts.xlsx.xlsx_builder import build_xlsx_skeleton

# -- legacy template-mode imports (only used when --template is given) --
from scripts.xlsx.content_types import ensure_png_content_type
from scripts.xlsx.relationships import build_sheet_relationships, build_vml_relationships
from scripts.xlsx.sheet import (
    extract_base_anchor,
    update_drawing,
    update_ole_objects,
    update_sheet_labels,
    update_vml,
    write_shared_strings,
)


def _build_template_mode(
    template: Path,
    cdx_paths: list[Path],
    output: Path,
    stream_name: str,
) -> None:
    """Legacy path: inject CDX into a user-provided XLSX template."""
    register_namespaces()

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        with zipfile.ZipFile(template, "r") as z:
            z.extractall(tmpdir)

        ole_dir = tmpdir / "xl" / "embeddings"
        ole_dir.mkdir(parents=True, exist_ok=True)
        base_ole_path = ole_dir / "oleObject1.bin"
        if not base_ole_path.exists():
            raise SystemExit(f"OLE object not found: {base_ole_path}")
        base_ole_bytes = base_ole_path.read_bytes()

        sheet_path = tmpdir / "xl" / "worksheets" / "sheet1.xml"
        drawing_path = tmpdir / "xl" / "drawings" / "drawing1.xml"
        vml_path = tmpdir / "xl" / "drawings" / "vmlDrawing1.vml"
        sheet_rels_path = tmpdir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels"
        vml_rels_path = tmpdir / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels"
        shared_strings_path = tmpdir / "xl" / "sharedStrings.xml"
        content_types_path = tmpdir / "[Content_Types].xml"

        base_shape_id, from_row, to_row = extract_base_anchor(sheet_path)

        labels = [p.stem for p in cdx_paths]
        update_sheet_labels(sheet_path, labels)
        write_shared_strings(shared_strings_path, labels)

        ole_targets = []
        image_targets = []
        vml_image_targets = []

        for idx, cdx_path in enumerate(cdx_paths, start=1):
            ole_path = ole_dir / f"oleObject{idx}.bin"
            if idx == 1:
                if not ole_path.exists():
                    raise SystemExit(f"OLE object not found: {ole_path}")
            else:
                ole_path.write_bytes(base_ole_bytes)

            cdx_data = cdx_path.read_bytes()
            try:
                inject_cdx_stream(ole_path, stream_name, cdx_data)
                patch_direntry_size(ole_path, stream_name, len(cdx_data))
            except SystemExit as exc:
                if "CDX is larger than existing stream" not in str(exc):
                    raise
                # Keep the stable template XML/VML path, but swap in a
                # freshly built OLE container that can hold a large
                # CONTENTS stream.
                ole_path.write_bytes(build_chemdraw_ole(cdx_data))

            image_name = f"image{idx}.png"
            image_path = tmpdir / "xl" / "media" / image_name
            image_path.parent.mkdir(parents=True, exist_ok=True)
            render_png(cdx_path, image_path)

            ole_targets.append(f"../embeddings/oleObject{idx}.bin")
            image_targets.append(f"../media/{image_name}")
            vml_image_targets.append(f"../media/{image_name}")

        emf_path = tmpdir / "xl" / "media" / "image1.emf"
        if emf_path.exists():
            emf_path.unlink()

        rel_pairs = build_sheet_relationships(
            sheet_rels_path, ole_targets, image_targets
        )
        vml_rids = build_vml_relationships(vml_rels_path, vml_image_targets)

        update_ole_objects(sheet_path, base_shape_id, from_row, to_row, rel_pairs)
        update_drawing(drawing_path, base_shape_id, from_row, to_row, len(cdx_paths))
        update_vml(vml_path, base_shape_id, from_row, to_row, vml_rids)
        ensure_png_content_type(content_types_path)

        _pack_xlsx(tmpdir, output)


def _build_generate_mode(
    cdx_paths: list[Path],
    output: Path,
    *,
    col: str,
    start_row: int,
    margin_left: str,
    margin_top: str,
    width: str,
    height: str,
    from_spec: str | None,
    to_spec: str | None,
    row_height: float,
) -> None:
    """Template-free path using the bundled Excel skeleton by default."""
    labels = [p.stem for p in cdx_paths]
    use_bundled_template = (
        col == "B"
        and start_row == 1
        and margin_left == "12mm"
        and margin_top == "5mm"
        and width == "22mm"
        and height == "25mm"
        and from_spec is None
        and to_spec is None
        and row_height == 97.5
    )

    if use_bundled_template:
        bundled_template = Path(__file__).resolve().parents[2] / "template.xlsx"
        _build_template_mode(bundled_template, cdx_paths, output, "CONTENTS")
        return

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)

        anchors = compute_anchors(
            len(cdx_paths),
            col=col,
            start_row=start_row,
            margin_left=margin_left,
            margin_top=margin_top,
            width=width,
            height=height,
            from_spec=from_spec,
            to_spec=to_spec,
        )

        build_xlsx_skeleton(
            tmpdir, labels, anchors, row_height=row_height,
        )

        ole_dir = tmpdir / "xl" / "embeddings"
        media_dir = tmpdir / "xl" / "media"
        ole_dir.mkdir(parents=True, exist_ok=True)
        media_dir.mkdir(parents=True, exist_ok=True)

        for idx, cdx_path in enumerate(cdx_paths, start=1):
            cdx_data = cdx_path.read_bytes()
            ole_bytes = build_chemdraw_ole(cdx_data)
            (ole_dir / f"oleObject{idx}.bin").write_bytes(ole_bytes)

            image_path = media_dir / f"image{idx}.png"
            render_png(cdx_path, image_path)

        _pack_xlsx(tmpdir, output)


def _pack_xlsx(tmpdir: Path, output: Path) -> None:
    """Pack *tmpdir* contents into an XLSX (ZIP) file."""
    if output.exists():
        output.unlink()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(tmpdir):
            for name in files:
                full_path = Path(root) / name
                rel_path = full_path.relative_to(tmpdir)
                z.write(full_path, rel_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Inject CDX into Excel with ChemDraw OLE objects.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
examples:
  # Template-free (default, recommended):
  %(prog)s --cdx mol1.cdx mol2.cdx -o output.xlsx
  %(prog)s --cdx mol.cdx --col C --start-row 2 -o output.xlsx
  %(prog)s --cdx mol.cdx --from A1+0mm,40pt --to A2+20mm,10pt -o output.xlsx

  # Legacy template mode:
  %(prog)s --template template.xlsx --cdx mol.cdx -o output.xlsx
""",
    )
    parser.add_argument("--template", default=None,
                        help="Legacy: XLSX template with pre-made OLE object")
    parser.add_argument("--cdx", required=True, nargs="+",
                        help="CDX file(s) to inject")
    parser.add_argument("-o", "--output", default="output.xlsx",
                        help="Output xlsx (default: output.xlsx)")
    parser.add_argument("--stream", default="CONTENTS",
                        help="OLE stream name (legacy mode only)")

    # Position control (template-free mode only)
    pos = parser.add_argument_group("position control (template-free mode)")
    pos.add_argument("--col", default="B",
                     help="Column for OLE objects (default: B)")
    pos.add_argument("--start-row", type=int, default=1,
                     help="Starting row, 1-indexed (default: 1)")
    pos.add_argument("--margin-left", default="12mm",
                     help="Left offset within cell (default: 12mm)")
    pos.add_argument("--margin-top", default="5mm",
                     help="Top offset within cell (default: 5mm)")
    pos.add_argument("--width", default="22mm",
                     help="OLE object width (default: 22mm)")
    pos.add_argument("--height", default="25mm",
                     help="OLE object height (default: 25mm)")
    pos.add_argument("--from", dest="from_spec", default=None,
                     help="Advanced: precise from-anchor, e.g. 'A1+5mm,3mm'")
    pos.add_argument("--to", dest="to_spec", default=None,
                     help="Advanced: precise to-anchor, e.g. 'A2+25mm,18mm'")
    pos.add_argument("--row-height", type=float, default=97.5,
                     help="Row height in points (default: 97.5)")

    args = parser.parse_args()

    cdx_paths = [Path(p) for p in args.cdx]
    output = Path(args.output)

    for cdx_path in cdx_paths:
        if not cdx_path.exists():
            raise SystemExit(f"CDX not found: {cdx_path}")
        if not cdx_path.read_bytes():
            raise SystemExit(f"CDX is empty: {cdx_path}")

    if args.template:
        template = Path(args.template)
        if not template.exists():
            raise SystemExit(f"Template not found: {template}")
        print(f"Using template mode: {template}")
        _build_template_mode(template, cdx_paths, output, args.stream)
    else:
        _build_generate_mode(
            cdx_paths, output,
            col=args.col,
            start_row=args.start_row,
            margin_left=args.margin_left,
            margin_top=args.margin_top,
            width=args.width,
            height=args.height,
            from_spec=args.from_spec,
            to_spec=args.to_spec,
            row_height=args.row_height,
        )

    print(f"Wrote {output}")


if __name__ == "__main__":
    main()
