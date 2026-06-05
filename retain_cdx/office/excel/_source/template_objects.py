from __future__ import annotations

import copy
from pathlib import Path
import uuid
import xml.etree.ElementTree as ET

from .anchors import OleAnchor
from .template_common import set_anchor_corner
from .template_namespaces import NS, NS_R, ensure_sheet_namespaces


def update_ole_objects(
    sheet_path: Path,
    base_shape_id: int,
    from_row: int,
    to_row: int,
    rel_pairs: list[tuple[str, str]],
    anchors: list[OleAnchor] | None = None,
) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    ensure_sheet_namespaces(root)

    ole_objects = root.find("s:oleObjects", NS)
    if ole_objects is None:
        raise SystemExit("oleObjects not found")
    base_alt = ole_objects.find("mc:AlternateContent", NS)
    if base_alt is None:
        raise SystemExit("AlternateContent not found")

    for child in list(ole_objects):
        ole_objects.remove(child)

    for idx, (ole_rid, img_rid) in enumerate(rel_pairs, start=1):
        alt = copy.deepcopy(base_alt)
        shape_id = base_shape_id + (idx - 1)
        row_delta = idx - 1
        custom_anchor = anchors[idx - 1] if anchors is not None else None

        choice = alt.find("mc:Choice", NS)
        if choice is not None:
            ole_obj = choice.find("s:oleObject", NS)
            if ole_obj is not None:
                ole_obj.set("shapeId", str(shape_id))
                ole_obj.set(f"{{{NS_R}}}id", ole_rid)
                obj_pr = ole_obj.find("s:objectPr", NS)
                if obj_pr is not None:
                    obj_pr.set(f"{{{NS_R}}}id", img_rid)
                    anchor = obj_pr.find("s:anchor", NS)
                    if anchor is not None:
                        _update_anchor_rows(
                            anchor,
                            from_row=from_row,
                            to_row=to_row,
                            row_delta=row_delta,
                            custom_anchor=custom_anchor,
                            corner_prefix="s",
                        )

        fallback = alt.find("mc:Fallback", NS)
        if fallback is not None:
            fb_ole = fallback.find("s:oleObject", NS)
            if fb_ole is not None:
                fb_ole.set("shapeId", str(shape_id))
                fb_ole.set(f"{{{NS_R}}}id", ole_rid)

        ole_objects.append(alt)

    tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)


def update_drawing(
    drawing_path: Path,
    base_shape_id: int,
    from_row: int,
    to_row: int,
    count: int,
    anchors: list[OleAnchor] | None = None,
) -> None:
    tree = ET.parse(drawing_path)
    root = tree.getroot()
    ac = root.find("mc:AlternateContent", NS)
    if ac is None:
        raise SystemExit("drawing AlternateContent not found")
    choice = ac.find("mc:Choice", NS)
    if choice is None:
        raise SystemExit("drawing Choice not found")
    base_anchor = choice.find("xdr:twoCellAnchor", NS)
    if base_anchor is None:
        raise SystemExit("drawing twoCellAnchor not found")

    for anchor in list(choice.findall("xdr:twoCellAnchor", NS)):
        choice.remove(anchor)

    for idx in range(1, count + 1):
        anchor = copy.deepcopy(base_anchor)
        shape_id = base_shape_id + (idx - 1)
        row_delta = idx - 1
        custom_anchor = anchors[idx - 1] if anchors is not None else None

        _update_anchor_rows(
            anchor,
            from_row=from_row,
            to_row=to_row,
            row_delta=row_delta,
            custom_anchor=custom_anchor,
            corner_prefix="xdr",
        )

        c_nv_pr = anchor.find("xdr:sp/xdr:nvSpPr/xdr:cNvPr", NS)
        if c_nv_pr is not None:
            c_nv_pr.set("id", str(shape_id))
            c_nv_pr.set("name", f"Object {idx}")
            ext = c_nv_pr.find("a:extLst/a:ext/a14:compatExt", NS)
            if ext is not None:
                ext.set("spid", f"_x0000_s{shape_id}")
            creation = c_nv_pr.find("a:extLst/a:ext/a16:creationId", NS)
            if creation is not None:
                creation.set("id", f"{{{str(uuid.uuid4()).upper()}}}")

        choice.append(anchor)

    tree.write(drawing_path, encoding="UTF-8", xml_declaration=True)


def _update_anchor_rows(
    anchor: ET.Element,
    *,
    from_row: int,
    to_row: int,
    row_delta: int,
    custom_anchor: OleAnchor | None,
    corner_prefix: str,
) -> None:
    from_el = anchor.find(f"{corner_prefix}:from", NS)
    to_el = anchor.find(f"{corner_prefix}:to", NS)
    if custom_anchor is None:
        _set_row(from_el, from_row + row_delta)
        _set_row(to_el, to_row + row_delta)
        return

    set_anchor_corner(
        from_el,
        col=custom_anchor.from_pt.col,
        col_off=custom_anchor.from_pt.col_off,
        row=custom_anchor.from_pt.row,
        row_off=custom_anchor.from_pt.row_off,
    )
    set_anchor_corner(
        to_el,
        col=custom_anchor.to_pt.col,
        col_off=custom_anchor.to_pt.col_off,
        row=custom_anchor.to_pt.row,
        row_off=custom_anchor.to_pt.row_off,
    )


def _set_row(corner: ET.Element | None, row: int) -> None:
    if corner is None:
        return
    row_el = corner.find("xdr:row", NS)
    if row_el is not None:
        row_el.text = str(row)
