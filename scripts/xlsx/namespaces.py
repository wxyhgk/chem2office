from __future__ import annotations

import xml.etree.ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
NS_A16 = "http://schemas.microsoft.com/office/drawing/2014/main"
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


def register_namespaces() -> None:
    ET.register_namespace("", NS_MAIN)
    ET.register_namespace("r", NS_R)
    ET.register_namespace("xdr", NS_XDR)
    ET.register_namespace("x14", NS_X14)
    ET.register_namespace("mc", NS_MC)
    ET.register_namespace("a", NS_A)
    ET.register_namespace("a14", NS_A14)
    ET.register_namespace("a16", NS_A16)
    ET.register_namespace("x14ac", NS_X14AC)
    ET.register_namespace("xr", NS_XR)
    ET.register_namespace("xr2", NS_XR2)
    ET.register_namespace("xr3", NS_XR3)
    ET.register_namespace("v", NS_V)
    ET.register_namespace("o", NS_O)
    ET.register_namespace("x", NS_X)


def ensure_sheet_namespaces(root: ET.Element) -> None:
    # Preserve required prefixes that are not tied to element names.
    root.set("xmlns:x14", NS_X14)
    root.set("xmlns:xr2", NS_XR2)
    root.set("xmlns:xr3", NS_XR3)
