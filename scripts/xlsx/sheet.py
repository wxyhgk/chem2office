from __future__ import annotations

import copy
from pathlib import Path
import uuid
import xml.etree.ElementTree as ET

from .position import AnchorPoint, OleAnchor
from .namespaces import (
    NS,
    NS_MAIN,
    NS_R,
    NS_X14AC,
    ensure_sheet_namespaces,
)


def extract_base_anchor(sheet_path: Path) -> tuple[int, int, int]:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    base_alt = root.find("s:oleObjects/mc:AlternateContent", NS)
    if base_alt is None:
        raise SystemExit("base OLE object not found in sheet1.xml")
    base_choice = base_alt.find("mc:Choice", NS)
    if base_choice is None:
        raise SystemExit("base OLE choice not found in sheet1.xml")
    base_ole = base_choice.find("s:oleObject", NS)
    if base_ole is None:
        raise SystemExit("base OLE element not found in sheet1.xml")
    base_shape_id = int(base_ole.get("shapeId", "1025"))
    obj_pr = base_ole.find("s:objectPr/s:anchor", NS)
    if obj_pr is None:
        raise SystemExit("objectPr anchor not found in sheet1.xml")
    from_row = int(obj_pr.find("s:from/xdr:row", NS).text)  # type: ignore[union-attr]
    to_row = int(obj_pr.find("s:to/xdr:row", NS).text)  # type: ignore[union-attr]
    return base_shape_id, from_row, to_row


def extract_base_ole_anchor(sheet_path: Path) -> OleAnchor:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    base_alt = root.find("s:oleObjects/mc:AlternateContent", NS)
    if base_alt is None:
        raise SystemExit("base OLE object not found in sheet1.xml")
    base_choice = base_alt.find("mc:Choice", NS)
    if base_choice is None:
        raise SystemExit("base OLE choice not found in sheet1.xml")
    base_ole = base_choice.find("s:oleObject", NS)
    if base_ole is None:
        raise SystemExit("base OLE element not found in sheet1.xml")
    anchor = base_ole.find("s:objectPr/s:anchor", NS)
    if anchor is None:
        raise SystemExit("base OLE anchor not found in sheet1.xml")

    def read_corner(name: str) -> AnchorPoint:
        corner = anchor.find(f"s:{name}", NS)
        if corner is None:
            raise SystemExit(f"anchor corner not found: {name}")
        return AnchorPoint(
            col=int(corner.find("xdr:col", NS).text),  # type: ignore[union-attr]
            col_off=int(corner.find("xdr:colOff", NS).text),  # type: ignore[union-attr]
            row=int(corner.find("xdr:row", NS).text),  # type: ignore[union-attr]
            row_off=int(corner.find("xdr:rowOff", NS).text),  # type: ignore[union-attr]
        )

    return OleAnchor(from_pt=read_corner("from"), to_pt=read_corner("to"))


def _set_anchor_corner(corner: ET.Element | None, *, col: int, col_off: int, row: int, row_off: int) -> None:
    if corner is None:
        return
    mapping = {
        "xdr:col": str(col),
        "xdr:colOff": str(col_off),
        "xdr:row": str(row),
        "xdr:rowOff": str(row_off),
    }
    for tag, value in mapping.items():
        el = corner.find(tag, NS)
        if el is not None:
            el.text = value


def update_sheet_labels(sheet_path: Path, labels: list[str], total_rows: int | None = None) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    ensure_sheet_namespaces(root)

    dim = root.find("s:dimension", NS)
    if dim is not None:
        max_rows = max(total_rows or len(labels), len(labels))
        dim.set("ref", f"A1:B{max_rows}")

    sheet_data = root.find("s:sheetData", NS)
    if sheet_data is None:
        raise SystemExit("sheetData not found")

    template_row = sheet_data.find("s:row", NS)
    if template_row is None:
        raise SystemExit("template row not found")

    for row in list(sheet_data):
        sheet_data.remove(row)

    max_rows = max(total_rows or len(labels), len(labels))
    for idx in range(1, max_rows + 1):
        row = copy.deepcopy(template_row)
        row.set("r", str(idx))
        if f"{{{NS_X14AC}}}dyDescent" in row.attrib:
            row.set(f"{{{NS_X14AC}}}dyDescent", "0.2")

        cell = row.find("s:c", NS)
        if cell is None:
            raise SystemExit("cell not found in template row")
        if idx <= len(labels):
            cell.set("r", f"A{idx}")
            cell.set("t", "s")
            v = cell.find("s:v", NS)
            if v is None:
                v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
            v.text = str(idx - 1)
        else:
            row.remove(cell)
        sheet_data.append(row)

    tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)


def update_row_heights(sheet_path: Path, row_height: float, total_rows: int) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    ensure_sheet_namespaces(root)

    sheet_format = root.find("s:sheetFormatPr", NS)
    if sheet_format is not None:
        sheet_format.set("defaultRowHeight", str(row_height))
        sheet_format.set("customHeight", "1")

    sheet_data = root.find("s:sheetData", NS)
    if sheet_data is None:
        raise SystemExit("sheetData not found")
    for idx in range(1, total_rows + 1):
        row = sheet_data.find(f"s:row[@r='{idx}']", NS)
        if row is not None:
            row.set("ht", str(row_height))
            row.set("customHeight", "1")

    tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)


def write_shared_strings(path: Path, labels: list[str]) -> None:
    root = ET.Element(
        f"{{{NS_MAIN}}}sst", count=str(len(labels)), uniqueCount=str(len(labels))
    )
    for label in labels:
        si = ET.SubElement(root, f"{{{NS_MAIN}}}si")
        t = ET.SubElement(si, f"{{{NS_MAIN}}}t")
        t.text = label
    ET.ElementTree(root).write(path, encoding="UTF-8", xml_declaration=True)


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
                        from_el = anchor.find("s:from", NS)
                        to_el = anchor.find("s:to", NS)
                        if custom_anchor is None:
                            if from_el is not None:
                                row_el = from_el.find("xdr:row", NS)
                                if row_el is not None:
                                    row_el.text = str(from_row + row_delta)
                            if to_el is not None:
                                row_el = to_el.find("xdr:row", NS)
                                if row_el is not None:
                                    row_el.text = str(to_row + row_delta)
                        else:
                            _set_anchor_corner(
                                from_el,
                                col=custom_anchor.from_pt.col,
                                col_off=custom_anchor.from_pt.col_off,
                                row=custom_anchor.from_pt.row,
                                row_off=custom_anchor.from_pt.row_off,
                            )
                            _set_anchor_corner(
                                to_el,
                                col=custom_anchor.to_pt.col,
                                col_off=custom_anchor.to_pt.col_off,
                                row=custom_anchor.to_pt.row,
                                row_off=custom_anchor.to_pt.row_off,
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

        from_el = anchor.find("xdr:from", NS)
        to_el = anchor.find("xdr:to", NS)
        if custom_anchor is None:
            if from_el is not None:
                row_el = from_el.find("xdr:row", NS)
                if row_el is not None:
                    row_el.text = str(from_row + row_delta)
            if to_el is not None:
                row_el = to_el.find("xdr:row", NS)
                if row_el is not None:
                    row_el.text = str(to_row + row_delta)
        else:
            _set_anchor_corner(
                from_el,
                col=custom_anchor.from_pt.col,
                col_off=custom_anchor.from_pt.col_off,
                row=custom_anchor.from_pt.row,
                row_off=custom_anchor.from_pt.row_off,
            )
            _set_anchor_corner(
                to_el,
                col=custom_anchor.to_pt.col,
                col_off=custom_anchor.to_pt.col_off,
                row=custom_anchor.to_pt.row,
                row_off=custom_anchor.to_pt.row_off,
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

        anchor = shape.find("x:ClientData/x:Anchor", NS)
        if anchor is not None and anchor.text:
            parts = [p.strip() for p in anchor.text.strip().split(",")]
            if len(parts) == 8:
                if custom_anchor is None:
                    parts[2] = str(from_row + row_delta)
                    parts[6] = str(to_row + row_delta)
                else:
                    from_col, from_col_off = _vml_col_anchor(
                        custom_anchor.from_pt.col, custom_anchor.from_pt.col_off
                    )
                    from_row, from_row_off = _vml_row_anchor(
                        custom_anchor.from_pt.row,
                        custom_anchor.from_pt.row_off,
                        row_height=row_height,
                    )
                    to_col, to_col_off = _vml_col_anchor(
                        custom_anchor.to_pt.col, custom_anchor.to_pt.col_off
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

        if custom_anchor is not None:
            _sync_vml_style(shape, custom_anchor, z_index=idx, row_height=row_height)

        root.append(shape)

    tree.write(vml_path, encoding="UTF-8", xml_declaration=True)


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
    return sum(_col_width_px(i) for i in range(col)) + col_off_emu / 9525


def _absolute_y_px(row: int, row_off_emu: int, row_height: float) -> float:
    return row * _row_height_px(row_height) + row_off_emu / 9525


def _vml_box_px(anchor: OleAnchor, row_height: float) -> tuple[float, float, float, float]:
    left = _absolute_x_px(anchor.from_pt.col, anchor.from_pt.col_off)
    top = _absolute_y_px(anchor.from_pt.row, anchor.from_pt.row_off, row_height)
    right = _absolute_x_px(anchor.to_pt.col, anchor.to_pt.col_off)
    bottom = _absolute_y_px(anchor.to_pt.row, anchor.to_pt.row_off, row_height)
    return left, top, max(1, right - left), max(1, bottom - top)
