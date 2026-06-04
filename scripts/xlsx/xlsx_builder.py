"""Build a complete XLSX skeleton with OLE embedding support.

Generates all required XML files for an XLSX that contains ChemDraw OLE
objects, using pre-computed anchor positions for precise placement.
"""

from __future__ import annotations

import shutil
import uuid
import xml.etree.ElementTree as ET
from pathlib import Path

from .namespaces import (
    NS_A,
    NS_A14,
    NS_A16,
    NS_MAIN,
    NS_MC,
    NS_R,
    NS_X14,
    NS_X14AC,
    NS_XDR,
    NS_XR,
    NS_XR2,
    NS_XR3,
    register_namespaces,
)
from .position import OleAnchor

# Directory containing static template resources.
_TEMPLATE_DIR = Path(__file__).resolve().parent.parent.parent / "template"


def _write_xml(path: Path, root: ET.Element) -> None:
    ET.ElementTree(root).write(path, encoding="UTF-8", xml_declaration=True)


def _copy_static_files(tmpdir: Path) -> None:
    """Copy static XLSX parts that never change."""
    for rel in [
        "_rels/.rels",
        "docProps/app.xml",
        "docProps/core.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/workbook.xml",
        "xl/styles.xml",
    ]:
        src = _TEMPLATE_DIR / rel
        dst = tmpdir / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)

    # Copy theme directory.
    theme_src = _TEMPLATE_DIR / "xl" / "theme"
    theme_dst = tmpdir / "xl" / "theme"
    if theme_src.exists():
        shutil.copytree(theme_src, theme_dst, dirs_exist_ok=True)


def _write_content_types(path: Path, count: int) -> None:
    """Generate [Content_Types].xml."""
    ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    root = ET.Element(f"{{{ns}}}Types")

    defaults = [
        ("bin", "application/vnd.openxmlformats-officedocument.oleObject"),
        ("png", "image/png"),
        ("rels", "application/vnd.openxmlformats-package.relationships+xml"),
        ("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"),
        ("xml", "application/xml"),
    ]
    for ext, ct in defaults:
        ET.SubElement(root, f"{{{ns}}}Default", Extension=ext, ContentType=ct)

    overrides = [
        ("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        ("/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
        ("/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
        ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
        ("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"),
        ("/xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml"),
        ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
        ("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
    ]
    for pn, ct in overrides:
        ET.SubElement(root, f"{{{ns}}}Override", PartName=pn, ContentType=ct)

    _write_xml(path, root)


def _write_shared_strings(path: Path, labels: list[str]) -> None:
    root = ET.Element(
        f"{{{NS_MAIN}}}sst",
        count=str(len(labels)),
        uniqueCount=str(len(labels)),
    )
    for label in labels:
        si = ET.SubElement(root, f"{{{NS_MAIN}}}si")
        t = ET.SubElement(si, f"{{{NS_MAIN}}}t")
        t.text = label
    _write_xml(path, root)


def _write_sheet(
    path: Path,
    labels: list[str],
    anchors: list[OleAnchor],
    base_shape_id: int,
    row_height: float,
    ole_rids: list[str],
    img_rids: list[str],
) -> None:
    """Generate sheet1.xml with rows, cells, and oleObjects."""
    root = ET.Element(f"{{{NS_MAIN}}}worksheet")
    # Namespaces r, xdr, mc, x14ac, xr are auto-declared by ET via
    # register_namespaces() when used in element/attribute names.
    # Only declare x14, xr2, xr3 explicitly – they appear in
    # mc:Ignorable text but have no element/attribute usage.
    root.set("xmlns:x14", NS_X14)
    root.set("xmlns:xr2", NS_XR2)
    root.set("xmlns:xr3", NS_XR3)
    root.set(f"{{{NS_MC}}}Ignorable", "x14ac xr xr2 xr3")
    root.set(f"{{{NS_XR}}}uid", f"{{{str(uuid.uuid4()).upper()}}}")

    dim = ET.SubElement(root, f"{{{NS_MAIN}}}dimension")
    dim.set("ref", f"A1:B{len(labels)}")

    views = ET.SubElement(root, f"{{{NS_MAIN}}}sheetViews")
    view = ET.SubElement(views, f"{{{NS_MAIN}}}sheetView", tabSelected="1", workbookViewId="0")
    ET.SubElement(view, f"{{{NS_MAIN}}}selection", activeCell="C1", sqref="C1")

    fmt = ET.SubElement(root, f"{{{NS_MAIN}}}sheetFormatPr")
    fmt.set("defaultRowHeight", str(row_height))
    fmt.set("customHeight", "1")
    fmt.set(f"{{{NS_X14AC}}}dyDescent", "0.2")

    cols = ET.SubElement(root, f"{{{NS_MAIN}}}cols")
    ET.SubElement(cols, f"{{{NS_MAIN}}}col", min="2", max="2", width="21.375", customWidth="1")

    sheet_data = ET.SubElement(root, f"{{{NS_MAIN}}}sheetData")
    for idx, _label in enumerate(labels, start=1):
        row = ET.SubElement(sheet_data, f"{{{NS_MAIN}}}row", r=str(idx), spans="1:1")
        row.set("ht", str(row_height))
        row.set("customHeight", "1")
        row.set(f"{{{NS_X14AC}}}dyDescent", "0.2")
        cell = ET.SubElement(row, f"{{{NS_MAIN}}}c", r=f"A{idx}", t="s")
        v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
        v.text = str(idx - 1)

    ET.SubElement(root, f"{{{NS_MAIN}}}phoneticPr", fontId="1", type="noConversion")
    pm = ET.SubElement(root, f"{{{NS_MAIN}}}pageMargins")
    pm.set("left", "0.7")
    pm.set("right", "0.7")
    pm.set("top", "0.75")
    pm.set("bottom", "0.75")
    pm.set("header", "0.3")
    pm.set("footer", "0.3")

    ET.SubElement(root, f"{{{NS_MAIN}}}drawing", **{f"{{{NS_R}}}id": "rId1"})
    ET.SubElement(root, f"{{{NS_MAIN}}}legacyDrawing", **{f"{{{NS_R}}}id": "rId2"})

    # oleObjects
    ole_objects = ET.SubElement(root, f"{{{NS_MAIN}}}oleObjects")
    for idx, (anchor, ole_rid, i_rid) in enumerate(zip(anchors, ole_rids, img_rids)):
        shape_id = base_shape_id + idx
        ac = ET.SubElement(ole_objects, f"{{{NS_MC}}}AlternateContent")
        # Excel's XML parser seems to demand a local xmlns:mc redeclaration here
        # or else it drops the object.
        ac.set("xmlns:mc", NS_MC)

        choice = ET.SubElement(ac, f"{{{NS_MC}}}Choice", Requires="x14")
        ole_el = ET.SubElement(choice, f"{{{NS_MAIN}}}oleObject")
        ole_el.set("progId", "ChemDraw.Document.6.0")
        ole_el.set("shapeId", str(shape_id))
        ole_el.set(f"{{{NS_R}}}id", ole_rid)

        obj_pr = ET.SubElement(ole_el, f"{{{NS_MAIN}}}objectPr")
        obj_pr.set("defaultSize", "0")
        obj_pr.set("autoPict", "0")
        obj_pr.set(f"{{{NS_R}}}id", i_rid)

        anc = ET.SubElement(obj_pr, f"{{{NS_MAIN}}}anchor", moveWithCells="1")
        from_el = ET.SubElement(anc, f"{{{NS_MAIN}}}from")
        ET.SubElement(from_el, f"{{{NS_XDR}}}col").text = str(anchor.from_pt.col)
        ET.SubElement(from_el, f"{{{NS_XDR}}}colOff").text = str(anchor.from_pt.col_off)
        ET.SubElement(from_el, f"{{{NS_XDR}}}row").text = str(anchor.from_pt.row)
        ET.SubElement(from_el, f"{{{NS_XDR}}}rowOff").text = str(anchor.from_pt.row_off)
        to_el = ET.SubElement(anc, f"{{{NS_MAIN}}}to")
        ET.SubElement(to_el, f"{{{NS_XDR}}}col").text = str(anchor.to_pt.col)
        ET.SubElement(to_el, f"{{{NS_XDR}}}colOff").text = str(anchor.to_pt.col_off)
        ET.SubElement(to_el, f"{{{NS_XDR}}}row").text = str(anchor.to_pt.row)
        ET.SubElement(to_el, f"{{{NS_XDR}}}rowOff").text = str(anchor.to_pt.row_off)

        fallback = ET.SubElement(ac, f"{{{NS_MC}}}Fallback")
        fb_ole = ET.SubElement(fallback, f"{{{NS_MAIN}}}oleObject")
        fb_ole.set("progId", "ChemDraw.Document.6.0")
        fb_ole.set("shapeId", str(shape_id))
        fb_ole.set(f"{{{NS_R}}}id", ole_rid)

    _write_xml(path, root)


def _write_sheet_rels(
    path: Path,
    ole_targets: list[str],
    image_targets: list[str],
) -> tuple[list[str], list[str]]:
    """Generate sheet1.xml.rels and return (ole_rids, img_rids)."""
    ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    root = ET.Element("Relationships", xmlns=ns)

    # Fixed rels: drawing + vmlDrawing
    ET.SubElement(root, "Relationship",
                  Id="rId1",
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                  Target="../drawings/drawing1.xml")
    ET.SubElement(root, "Relationship",
                  Id="rId2",
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
                  Target="../drawings/vmlDrawing1.vml")

    ole_rids: list[str] = []
    img_rids: list[str] = []
    next_id = 3
    for ole_t, img_t in zip(ole_targets, image_targets):
        ole_rid = f"rId{next_id}"
        next_id += 1
        img_rid = f"rId{next_id}"
        next_id += 1
        ole_rids.append(ole_rid)
        img_rids.append(img_rid)
        ET.SubElement(root, "Relationship",
                      Id=ole_rid,
                      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                      Target=ole_t)
        ET.SubElement(root, "Relationship",
                      Id=img_rid,
                      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                      Target=img_t)

    _write_xml(path, root)
    return ole_rids, img_rids


def _write_drawing(
    path: Path,
    anchors: list[OleAnchor],
    base_shape_id: int,
) -> None:
    """Generate drawing1.xml."""
    # Namespaces a, mc, xdr, a14 are auto-declared by ET.
    root = ET.Element(f"{{{NS_XDR}}}wsDr")

    ac = ET.SubElement(root, f"{{{NS_MC}}}AlternateContent")
    choice = ET.SubElement(ac, f"{{{NS_MC}}}Choice")
    choice.set("Requires", "a14")

    for idx, anchor in enumerate(anchors):
        shape_id = base_shape_id + idx

        tca = ET.SubElement(choice, f"{{{NS_XDR}}}twoCellAnchor", editAs="oneCell")

        from_el = ET.SubElement(tca, f"{{{NS_XDR}}}from")
        ET.SubElement(from_el, f"{{{NS_XDR}}}col").text = str(anchor.from_pt.col)
        ET.SubElement(from_el, f"{{{NS_XDR}}}colOff").text = str(anchor.from_pt.col_off)
        ET.SubElement(from_el, f"{{{NS_XDR}}}row").text = str(anchor.from_pt.row)
        ET.SubElement(from_el, f"{{{NS_XDR}}}rowOff").text = str(anchor.from_pt.row_off)

        to_el = ET.SubElement(tca, f"{{{NS_XDR}}}to")
        ET.SubElement(to_el, f"{{{NS_XDR}}}col").text = str(anchor.to_pt.col)
        ET.SubElement(to_el, f"{{{NS_XDR}}}colOff").text = str(anchor.to_pt.col_off)
        ET.SubElement(to_el, f"{{{NS_XDR}}}row").text = str(anchor.to_pt.row)
        ET.SubElement(to_el, f"{{{NS_XDR}}}rowOff").text = str(anchor.to_pt.row_off)

        sp = ET.SubElement(tca, f"{{{NS_XDR}}}sp", macro="", textlink="")
        nv = ET.SubElement(sp, f"{{{NS_XDR}}}nvSpPr")
        cnv = ET.SubElement(nv, f"{{{NS_XDR}}}cNvPr")
        cnv.set("id", str(shape_id))
        cnv.set("name", f"Object {idx + 1}")
        cnv.set("hidden", "1")

        ext_lst = ET.SubElement(cnv, f"{{{NS_A}}}extLst")

        ext1 = ET.SubElement(ext_lst, f"{{{NS_A}}}ext")
        ext1.set("uri", "{63B3BB69-23CF-44E3-9099-C40C66FF867C}")
        compat = ET.SubElement(ext1, f"{{{NS_A14}}}compatExt")
        compat.set("spid", f"_x0000_s{shape_id}")

        ext2 = ET.SubElement(ext_lst, f"{{{NS_A}}}ext")
        ext2.set("uri", "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}")
        cid = ET.SubElement(ext2, f"{{{NS_A16}}}creationId")
        cid.set("id", f"{{{str(uuid.uuid4()).upper()}}}")

        ET.SubElement(nv, f"{{{NS_XDR}}}cNvSpPr")

        sp_pr = ET.SubElement(sp, f"{{{NS_XDR}}}spPr", bwMode="auto")
        xfrm = ET.SubElement(sp_pr, f"{{{NS_A}}}xfrm")
        ET.SubElement(xfrm, f"{{{NS_A}}}off", x="0", y="0")
        ET.SubElement(xfrm, f"{{{NS_A}}}ext", cx="0", cy="0")
        geom = ET.SubElement(sp_pr, f"{{{NS_A}}}prstGeom", prst="rect")
        ET.SubElement(geom, f"{{{NS_A}}}avLst")
        fill = ET.SubElement(sp_pr, f"{{{NS_A}}}solidFill")
        clr = ET.SubElement(fill, f"{{{NS_A}}}srgbClr", val="FFFFFF")
        clr.set(f"{{{NS_MC}}}Ignorable", "a14")
        clr.set(f"{{{NS_A14}}}legacySpreadsheetColorIndex", "65")
        ln = ET.SubElement(sp_pr, f"{{{NS_A}}}ln")
        ET.SubElement(ln, f"{{{NS_A}}}noFill")
        ET.SubElement(sp_pr, f"{{{NS_A}}}effectLst")

        # hiddenLine + hiddenEffects extensions (required by Excel)
        sp_ext_lst = ET.SubElement(sp_pr, f"{{{NS_A}}}extLst")

        hl_ext = ET.SubElement(sp_ext_lst, f"{{{NS_A}}}ext")
        hl_ext.set("uri", "{91240B29-F687-4F45-9708-019B960494DF}")
        hl = ET.SubElement(hl_ext, f"{{{NS_A14}}}hiddenLine", w="9525")
        hl_fill = ET.SubElement(hl, f"{{{NS_A}}}solidFill")
        hl_clr = ET.SubElement(hl_fill, f"{{{NS_A}}}srgbClr", val="000000")
        hl_clr.set(f"{{{NS_MC}}}Ignorable", "a14")
        hl_clr.set(f"{{{NS_A14}}}legacySpreadsheetColorIndex", "64")
        ET.SubElement(hl, f"{{{NS_A}}}prstDash", val="solid")
        ET.SubElement(hl, f"{{{NS_A}}}miter", lim="800000")
        ET.SubElement(hl, f"{{{NS_A}}}headEnd")
        ET.SubElement(hl, f"{{{NS_A}}}tailEnd", type="none", w="med", len="med")

        he_ext = ET.SubElement(sp_ext_lst, f"{{{NS_A}}}ext")
        he_ext.set("uri", "{AF507438-7753-43E0-B8FC-AC1667EBCBE1}")
        he = ET.SubElement(he_ext, f"{{{NS_A14}}}hiddenEffects")
        he_eff = ET.SubElement(he, f"{{{NS_A}}}effectLst")
        shadow = ET.SubElement(he_eff, f"{{{NS_A}}}outerShdw")
        shadow.set("dist", "35921")
        shadow.set("dir", "2700000")
        shadow.set("algn", "ctr")
        shadow.set("rotWithShape", "0")
        ET.SubElement(shadow, f"{{{NS_A}}}srgbClr", val="808080")

        ET.SubElement(tca, f"{{{NS_XDR}}}clientData")

    ET.SubElement(ac, f"{{{NS_MC}}}Fallback")
    _write_xml(path, root)


def _write_vml(
    path: Path,
    anchors: list[OleAnchor],
    base_shape_id: int,
    vml_rids: list[str],
    *,
    row_height: float,
) -> None:
    """Generate vmlDrawing1.vml."""
    NS_V = "urn:schemas-microsoft-com:vml"
    NS_O = "urn:schemas-microsoft-com:office:office"
    NS_X = "urn:schemas-microsoft-com:office:excel"

    # VML uses non-standard root tag "xml" which ET won't auto-prefix,
    # but these namespaces are registered globally so they'll be
    # auto-declared when used in child elements.
    root = ET.Element("xml")

    layout = ET.SubElement(root, f"{{{NS_O}}}shapelayout")
    layout.set(f"{{{NS_V}}}ext", "edit")
    idmap = ET.SubElement(layout, f"{{{NS_O}}}idmap")
    idmap.set(f"{{{NS_V}}}ext", "edit")
    idmap.set("data", "1")

    # shapetype
    st = ET.SubElement(root, f"{{{NS_V}}}shapetype")
    st.set("id", "_x0000_t75")
    st.set("coordsize", "21600,21600")
    st.set(f"{{{NS_O}}}spt", "75")
    st.set(f"{{{NS_O}}}preferrelative", "t")
    st.set("path", "m@4@5l@4@11@9@11@9@5xe")
    st.set("filled", "f")
    st.set("stroked", "f")

    stroke = ET.SubElement(st, f"{{{NS_V}}}stroke")
    stroke.set("joinstyle", "miter")

    formulas = ET.SubElement(st, f"{{{NS_V}}}formulas")
    for eq in [
        "if lineDrawn pixelLineWidth 0",
        "sum @0 1 0",
        "sum 0 0 @1",
        "prod @2 1 2",
        "prod @3 21600 pixelWidth",
        "prod @3 21600 pixelHeight",
        "sum @0 0 1",
        "prod @6 1 2",
        "prod @7 21600 pixelWidth",
        "sum @8 21600 0",
        "prod @7 21600 pixelHeight",
        "sum @10 21600 0",
    ]:
        f_el = ET.SubElement(formulas, f"{{{NS_V}}}f")
        f_el.set("eqn", eq)

    vpath = ET.SubElement(st, f"{{{NS_V}}}path")
    vpath.set(f"{{{NS_O}}}extrusionok", "f")
    vpath.set("gradientshapeok", "t")
    vpath.set(f"{{{NS_O}}}connecttype", "rect")

    lock = ET.SubElement(st, f"{{{NS_O}}}lock")
    lock.set(f"{{{NS_V}}}ext", "edit")
    lock.set("aspectratio", "t")

    # shapes
    for idx, (anchor, rid) in enumerate(zip(anchors, vml_rids)):
        shape_id = base_shape_id + idx
        left_px, top_px, width_px, height_px = _vml_box_px(anchor, row_height)
        shape = ET.SubElement(root, f"{{{NS_V}}}shape")
        shape.set("id", f"_x0000_s{shape_id}")
        shape.set("type", "#_x0000_t75")
        shape.set(
            "style",
            "position:absolute;"
            f"margin-left:{_px_to_pt(left_px):g}pt;"
            f"margin-top:{_px_to_pt(top_px):g}pt;"
            f"width:{_px_to_pt(width_px):g}pt;"
            f"height:{_px_to_pt(height_px):g}pt;"
            f"z-index:{idx + 1};"
            "mso-wrap-style:tight",
        )
        shape.set("filled", "t")
        shape.set("fillcolor", "window [65]")
        shape.set("stroked", "t")
        shape.set("strokecolor", "windowText [64]")
        shape.set(f"{{{NS_O}}}insetmode", "auto")

        vfill = ET.SubElement(shape, f"{{{NS_V}}}fill")
        vfill.set("color2", "window [65]")

        imgdata = ET.SubElement(shape, f"{{{NS_V}}}imagedata")
        imgdata.set(f"{{{NS_O}}}relid", rid)
        imgdata.set(f"{{{NS_O}}}title", "")

        cd = ET.SubElement(shape, f"{{{NS_X}}}ClientData")
        cd.set("ObjectType", "Pict")
        ET.SubElement(cd, f"{{{NS_X}}}SizeWithCells")

        anc_el = ET.SubElement(cd, f"{{{NS_X}}}Anchor")
        # VML anchor format: fromCol, fromColOff, fromRow, fromRowOff,
        #                     toCol,   toColOff,   toRow,   toRowOff
        # The offsets here are in a different unit (pixels at 96dpi
        # vs EMU), approximate via division.
        EMU_PER_PX = 9525
        fc = anchor.from_pt.col
        fco = round(anchor.from_pt.col_off / EMU_PER_PX)
        fr = anchor.from_pt.row
        fro = round(anchor.from_pt.row_off / EMU_PER_PX)
        tc = anchor.to_pt.col
        tco = round(anchor.to_pt.col_off / EMU_PER_PX)
        tr = anchor.to_pt.row
        tro = round(anchor.to_pt.row_off / EMU_PER_PX)
        anc_el.text = f"\n     {fc}, {fco}, {fr}, {fro}, {tc}, {tco}, {tr}, {tro}"

        ET.SubElement(cd, f"{{{NS_X}}}CF").text = "Pict"

    _write_xml(path, root)


def _px_to_pt(px: float) -> float:
    return px * 0.75


def _excel_col_width_to_px(width: float) -> int:
    # Excel's documented-ish display conversion for widths at Calibri 11.
    return int(width * 7 + 5)


def _col_width_px(col: int) -> int:
    # _write_sheet gives Excel column B a custom width; other columns keep the
    # normal default width.
    if col == 1:
        return _excel_col_width_to_px(21.375)
    return 64


def _row_height_px(row_height_pt: float) -> float:
    return row_height_pt / 0.75


def _absolute_x_px(col: int, col_off_emu: int) -> float:
    return sum(_col_width_px(i) for i in range(col)) + col_off_emu / 9525


def _absolute_y_px(row: int, row_off_emu: int, row_height_pt: float) -> float:
    return row * _row_height_px(row_height_pt) + row_off_emu / 9525


def _vml_box_px(anchor: OleAnchor, row_height_pt: float) -> tuple[float, float, float, float]:
    left = _absolute_x_px(anchor.from_pt.col, anchor.from_pt.col_off)
    top = _absolute_y_px(anchor.from_pt.row, anchor.from_pt.row_off, row_height_pt)
    right = _absolute_x_px(anchor.to_pt.col, anchor.to_pt.col_off)
    bottom = _absolute_y_px(anchor.to_pt.row, anchor.to_pt.row_off, row_height_pt)
    return left, top, max(1, right - left), max(1, bottom - top)


def _write_vml_rels(path: Path, image_targets: list[str]) -> list[str]:
    """Generate vmlDrawing1.vml.rels and return rIds."""
    ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    root = ET.Element("Relationships", xmlns=ns)
    rids: list[str] = []
    for idx, target in enumerate(image_targets, start=1):
        rid = f"rId{idx}"
        rids.append(rid)
        ET.SubElement(root, "Relationship",
                      Id=rid,
                      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                      Target=target)
    _write_xml(path, root)
    return rids


def build_xlsx_skeleton(
    tmpdir: Path,
    labels: list[str],
    anchors: list[OleAnchor],
    *,
    row_height: float = 97.5,
    base_shape_id: int = 1025,
) -> tuple[list[str], list[str]]:
    """Generate all XLSX XML files in *tmpdir*.

    Parameters
    ----------
    tmpdir:
        Working directory to populate with the XLSX structure.
    labels:
        List of labels for column A (one per OLE object).
    anchors:
        Anchor positions for each OLE object.
    row_height:
        Row height in points.
    base_shape_id:
        Starting shape ID for OLE objects.

    Returns
    -------
    ole_rids, vml_rids:
        Relationship IDs needed for writing OLE binary files.
    """
    register_namespaces()

    count = len(labels)

    _copy_static_files(tmpdir)

    # Ensure directories exist.
    for d in [
        "xl/embeddings",
        "xl/media",
        "xl/worksheets/_rels",
        "xl/drawings/_rels",
    ]:
        (tmpdir / d).mkdir(parents=True, exist_ok=True)

    _write_content_types(tmpdir / "[Content_Types].xml", count)
    _write_shared_strings(tmpdir / "xl" / "sharedStrings.xml", labels)

    ole_targets = [f"../embeddings/oleObject{i}.bin" for i in range(1, count + 1)]
    image_targets = [f"../media/image{i}.png" for i in range(1, count + 1)]

    ole_rids, img_rids = _write_sheet_rels(
        tmpdir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels",
        ole_targets,
        image_targets,
    )

    vml_image_targets = [f"../media/image{i}.png" for i in range(1, count + 1)]
    vml_rids = _write_vml_rels(
        tmpdir / "xl" / "drawings" / "_rels" / "vmlDrawing1.vml.rels",
        vml_image_targets,
    )

    _write_sheet(
        tmpdir / "xl" / "worksheets" / "sheet1.xml",
        labels, anchors, base_shape_id, row_height, ole_rids, img_rids,
    )

    _write_drawing(
        tmpdir / "xl" / "drawings" / "drawing1.xml",
        anchors, base_shape_id,
    )

    _write_vml(
        tmpdir / "xl" / "drawings" / "vmlDrawing1.vml",
        anchors, base_shape_id, vml_rids, row_height=row_height,
    )

    return ole_rids, vml_rids
