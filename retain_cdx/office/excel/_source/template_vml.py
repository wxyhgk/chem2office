from __future__ import annotations

import copy
from pathlib import Path
import xml.etree.ElementTree as ET

from .anchors import OleAnchor
from .template_namespaces import NS


def update_vml(
    vml_path: Path,
    base_shape_id: int,
    from_row: int,
    to_row: int,
    vml_rids: list[str],
    anchors: list[OleAnchor] | None = None,
    row_height: float = 97.5,
) -> None:
    tree = ET.parse(vml_path)
    root = tree.getroot()
    base_shape = root.find("v:shape", NS)
    if base_shape is None:
        raise SystemExit("VML shape not found")

    for shape in list(root.findall("v:shape", NS)):
        root.remove(shape)

    for idx, rid in enumerate(vml_rids, start=1):
        shape = copy.deepcopy(base_shape)
        shape_id = base_shape_id + (idx - 1)
        row_delta = idx - 1
        custom_anchor = anchors[idx - 1] if anchors is not None else None
        shape.set("id", f"_x0000_s{shape_id}")

        imagedata = shape.find("v:imagedata", NS)
        if imagedata is not None:
            imagedata.set("{urn:schemas-microsoft-com:office:office}relid", rid)

        _update_vml_anchor(
            shape,
            from_row=from_row,
            to_row=to_row,
            row_delta=row_delta,
            custom_anchor=custom_anchor,
            row_height=row_height,
        )
        if custom_anchor is not None:
            _sync_vml_style(shape, custom_anchor, z_index=idx, row_height=row_height)

        root.append(shape)

    tree.write(vml_path, encoding="UTF-8", xml_declaration=True)


def _update_vml_anchor(
    shape: ET.Element,
    *,
    from_row: int,
    to_row: int,
    row_delta: int,
    custom_anchor: OleAnchor | None,
    row_height: float,
) -> None:
    anchor = shape.find("x:ClientData/x:Anchor", NS)
    if anchor is None or not anchor.text:
        return
    parts = [part.strip() for part in anchor.text.strip().split(",")]
    if len(parts) != 8:
        return
    if custom_anchor is None:
        parts[2] = str(from_row + row_delta)
        parts[6] = str(to_row + row_delta)
    else:
        from_col, from_col_off = _vml_col_anchor(
            custom_anchor.from_pt.col,
            custom_anchor.from_pt.col_off,
        )
        from_row, from_row_off = _vml_row_anchor(
            custom_anchor.from_pt.row,
            custom_anchor.from_pt.row_off,
            row_height=row_height,
        )
        to_col, to_col_off = _vml_col_anchor(
            custom_anchor.to_pt.col,
            custom_anchor.to_pt.col_off,
        )
        to_row, to_row_off = _vml_row_anchor(
            custom_anchor.to_pt.row,
            custom_anchor.to_pt.row_off,
            row_height=row_height,
        )
        parts[0] = str(from_col)
        parts[1] = str(from_col_off)
        parts[2] = str(from_row)
        parts[3] = str(from_row_off)
        parts[4] = str(to_col)
        parts[5] = str(to_col_off)
        parts[6] = str(to_row)
        parts[7] = str(to_row_off)
    anchor.text = ", ".join(parts)


def _vml_col_anchor(col: int, col_off_emu: int) -> tuple[int, int]:
    return col, round(col_off_emu / 9525)


def _vml_row_anchor(row: int, row_off_emu: int, *, row_height: float) -> tuple[int, int]:
    return row, round(row_off_emu / 9525)


def _sync_vml_style(
    shape: ET.Element,
    anchor: OleAnchor,
    *,
    z_index: int,
    row_height: float,
) -> None:
    left, top, width, height = _vml_box_px(anchor, row_height)
    style = shape.get("style", "")
    wrap = "mso-wrap-style:tight" if "mso-wrap-style:tight" in style else ""
    declarations = [
        "position:absolute",
        f"margin-left:{_px_to_pt(left):g}pt",
        f"margin-top:{_px_to_pt(top):g}pt",
        f"width:{_px_to_pt(width):g}pt",
        f"height:{_px_to_pt(height):g}pt",
        f"z-index:{z_index}",
    ]
    if wrap:
        declarations.append(wrap)
    shape.set("style", ";".join(declarations))
    shape.set("stroked", "t")


def _px_to_pt(px: float) -> float:
    return px * 0.75


def _excel_col_width_to_px(width: float) -> int:
    return int(width * 7 + 5)


def _col_width_px(col: int) -> int:
    if col == 1:
        return _excel_col_width_to_px(21.375)
    return 64


def _row_height_px(row_height: float) -> float:
    return row_height / 0.75


def _absolute_x_px(col: int, col_off_emu: int) -> float:
    return sum(_col_width_px(idx) for idx in range(col)) + col_off_emu / 9525


def _absolute_y_px(row: int, row_off_emu: int, row_height: float) -> float:
    return row * _row_height_px(row_height) + row_off_emu / 9525


def _vml_box_px(anchor: OleAnchor, row_height: float) -> tuple[float, float, float, float]:
    left = _absolute_x_px(anchor.from_pt.col, anchor.from_pt.col_off)
    top = _absolute_y_px(anchor.from_pt.row, anchor.from_pt.row_off, row_height)
    right = _absolute_x_px(anchor.to_pt.col, anchor.to_pt.col_off)
    bottom = _absolute_y_px(anchor.to_pt.row, anchor.to_pt.row_off, row_height)
    return left, top, max(1, right - left), max(1, bottom - top)
