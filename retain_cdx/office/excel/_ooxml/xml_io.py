from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET


# Canonical OOXML namespace URIs (single source of truth for the package).
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
NS_A16 = "http://schemas.microsoft.com/office/drawing/2014/main"
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
NS_XR2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
NS_XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
NS_V = "urn:schemas-microsoft-com:vml"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_X = "urn:schemas-microsoft-com:office:excel"

NS = {
    "s": NS_MAIN,
    "r": NS_R,
    "xdr": NS_XDR,
    "x14": NS_X14,
    "mc": NS_MC,
    "a": NS_A,
    "a14": NS_A14,
    "a16": NS_A16,
    "x14ac": NS_X14AC,
    "xr": NS_XR,
    "xr2": NS_XR2,
    "xr3": NS_XR3,
    "v": NS_V,
    "o": NS_O,
    "x": NS_X,
}


# CT_Worksheet child sequence (ECMA-376 §18.3.1.99). Excel prompts to repair —
# and drops the offending drawing/OLE parts — when children are out of order,
# so new elements must be placed by schema rank, never "after pageMargins".
WORKSHEET_CHILD_ORDER = (
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
)

_WORKSHEET_RANK = {name: idx for idx, name in enumerate(WORKSHEET_CHILD_ORDER)}


def local_name(tag: str) -> str:
    """Strip the ``{namespace}`` prefix from an ElementTree tag."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def insert_worksheet_child(root: ET.Element, element: ET.Element) -> None:
    """Insert *element* into a worksheet at its schema-mandated position.

    Places it before the first existing child that must follow it. Children
    with unknown tags (e.g. ``mc:AlternateContent``) carry no rank and are
    skipped when scanning, so they never drag a known element out of order.
    """
    rank = _WORKSHEET_RANK.get(local_name(element.tag))
    if rank is None:
        root.append(element)
        return
    for idx, child in enumerate(root):
        child_rank = _WORKSHEET_RANK.get(local_name(child.tag))
        if child_rank is not None and child_rank > rank:
            root.insert(idx, element)
            return
    root.append(element)


def ensure_sheet_compat_namespaces(root: ET.Element) -> None:
    root.set(f"{{{NS_MC}}}Ignorable", "x14ac xr xr2 xr3")


def write_sheet_xml(tree: ET.ElementTree, sheet_path: Path) -> None:
    register_namespaces()
    tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)
    text = sheet_path.read_text(encoding="UTF-8")
    root_start = text.find("<worksheet")
    if root_start == -1:
        return
    root_end = text.find(">", root_start)
    if root_end == -1:
        return

    root_tag = text[root_start:root_end]
    declarations = {
        "xmlns:x14": NS_X14,
        "xmlns:x14ac": NS_X14AC,
        "xmlns:xr": NS_XR,
        "xmlns:xr2": NS_XR2,
        "xmlns:xr3": NS_XR3,
    }
    for name, uri in declarations.items():
        if f"{name}=" not in root_tag:
            root_tag += f' {name}="{uri}"'

    text = text[:root_start] + root_tag + text[root_end:]
    sheet_path.write_text(text, encoding="UTF-8")


def register_namespaces() -> None:
    ET.register_namespace("", NS_MAIN)
    ET.register_namespace("mc", NS_MC)
    ET.register_namespace("r", NS_R)
    ET.register_namespace("xdr", NS_XDR)
    ET.register_namespace("a", NS_A)
    ET.register_namespace("a14", NS_A14)
    ET.register_namespace("a16", NS_A16)
    ET.register_namespace("x14", NS_X14)
    ET.register_namespace("x14ac", NS_X14AC)
    ET.register_namespace("xr", NS_XR)
    ET.register_namespace("xr2", NS_XR2)
    ET.register_namespace("xr3", NS_XR3)
    ET.register_namespace("v", NS_V)
    ET.register_namespace("o", NS_O)
    ET.register_namespace("x", NS_X)
