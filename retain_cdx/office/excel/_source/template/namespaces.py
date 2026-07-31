"""Namespace names for template code (canonical values live in ``_ooxml``)."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from retain_cdx.office.excel._ooxml.xml_io import (
    NS,
    NS_A,
    NS_A14,
    NS_A16,
    NS_MAIN,
    NS_MC,
    NS_O,
    NS_R,
    NS_V,
    NS_X,
    NS_X14,
    NS_X14AC,
    NS_XDR,
    NS_XR,
    NS_XR2,
    NS_XR3,
    register_namespaces,
)

__all__ = [
    "NS",
    "NS_A",
    "NS_A14",
    "NS_A16",
    "NS_MAIN",
    "NS_MC",
    "NS_O",
    "NS_R",
    "NS_V",
    "NS_X",
    "NS_X14",
    "NS_X14AC",
    "NS_XDR",
    "NS_XR",
    "NS_XR2",
    "NS_XR3",
    "ensure_sheet_namespaces",
    "register_namespaces",
]


def ensure_sheet_namespaces(root: ET.Element) -> None:
    # Preserve required prefixes that are not tied to element names.
    root.set("xmlns:x14", NS_X14)
    root.set("xmlns:xr2", NS_XR2)
    root.set("xmlns:xr3", NS_XR3)
