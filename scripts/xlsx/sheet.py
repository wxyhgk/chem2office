from __future__ import annotations

import copy
from pathlib import Path
import uuid
import xml.etree.ElementTree as ET

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


def update_sheet_labels(sheet_path: Path, labels: list[str]) -> None:
    tree = ET.parse(sheet_path)
    root = tree.getroot()
    ensure_sheet_namespaces(root)

    dim = root.find("s:dimension", NS)
    if dim is not None:
        dim.set("ref", f"A1:B{len(labels)}")

    sheet_data = root.find("s:sheetData", NS)
    if sheet_data is None:
        raise SystemExit("sheetData not found")

    template_row = sheet_data.find("s:row", NS)
    if template_row is None:
        raise SystemExit("template row not found")

    for row in list(sheet_data):
        sheet_data.remove(row)

    for idx, _label in enumerate(labels, start=1):
        row = copy.deepcopy(template_row)
        row.set("r", str(idx))
        if f"{{{NS_X14AC}}}dyDescent" in row.attrib:
            row.set(f"{{{NS_X14AC}}}dyDescent", "0.2")

        cell = row.find("s:c", NS)
        if cell is None:
            raise SystemExit("cell not found in template row")
        cell.set("r", f"A{idx}")
        cell.set("t", "s")
        v = cell.find("s:v", NS)
        if v is None:
            v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
        v.text = str(idx - 1)
        sheet_data.append(row)

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
                        if from_el is not None:
                            row_el = from_el.find("xdr:row", NS)
                            if row_el is not None:
                                row_el.text = str(from_row + row_delta)
                        if to_el is not None:
                            row_el = to_el.find("xdr:row", NS)
                            if row_el is not None:
                                row_el.text = str(to_row + row_delta)

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

        from_el = anchor.find("xdr:from", NS)
        to_el = anchor.find("xdr:to", NS)
        if from_el is not None:
            row_el = from_el.find("xdr:row", NS)
            if row_el is not None:
                row_el.text = str(from_row + row_delta)
        if to_el is not None:
            row_el = to_el.find("xdr:row", NS)
            if row_el is not None:
                row_el.text = str(to_row + row_delta)

        c_nv_pr = anchor.find("xdr:sp/xdr:nvSpPr/xdr:cNvPr", NS)
        if c_nv_pr is not None:
            c_nv_pr.set("id", str(shape_id))
            c_nv_pr.set("name", f"Object {idx}")
            ext = c_nv_pr.find("a:extLst/a:ext/a14:compatExt", NS)
            if ext is not None:
                ext.set("spid", f"_x0000_s{shape_id}")
            creation = c_nv_pr.find("a:extLst/a:ext/a16:creationId", NS)
            if creation is not None and idx > 1:
                creation.set("id", f"{{{str(uuid.uuid4()).upper()}}}")

        choice.append(anchor)

    tree.write(drawing_path, encoding="UTF-8", xml_declaration=True)


def update_vml(
    vml_path: Path,
    base_shape_id: int,
    from_row: int,
    to_row: int,
    vml_rids: list[str],
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
        shape.set("id", f"_x0000_s{shape_id}")

        imagedata = shape.find("v:imagedata", NS)
        if imagedata is not None:
            imagedata.set("{urn:schemas-microsoft-com:office:office}relid", rid)

        anchor = shape.find("x:ClientData/x:Anchor", NS)
        if anchor is not None and anchor.text:
            parts = [p.strip() for p in anchor.text.strip().split(",")]
            if len(parts) == 8:
                parts[2] = str(from_row + row_delta)
                parts[6] = str(to_row + row_delta)
                anchor.text = ", ".join(parts)

        root.append(shape)

    tree.write(vml_path, encoding="UTF-8", xml_declaration=True)
