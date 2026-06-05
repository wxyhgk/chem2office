from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
NS_XR2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
NS_XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"

NS = {"s": NS_MAIN}


def ensure_sheet_compat_namespaces(root: ET.Element) -> None:
    root.set(f"{{{NS_MC}}}Ignorable", "x14ac xr xr2 xr3")


def write_sheet_xml(tree: ET.ElementTree, sheet_path: Path) -> None:
    _register_namespaces()
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


def _register_namespaces() -> None:
    ET.register_namespace("", NS_MAIN)
    ET.register_namespace("mc", NS_MC)
    ET.register_namespace("r", NS_R)
    ET.register_namespace("x14", NS_X14)
    ET.register_namespace("x14ac", NS_X14AC)
    ET.register_namespace("xr", NS_XR)
    ET.register_namespace("xr2", NS_XR2)
    ET.register_namespace("xr3", NS_XR3)
    ET.register_namespace(
        "xdr",
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    )
