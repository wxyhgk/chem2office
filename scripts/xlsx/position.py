"""OLE anchor-point calculations for XLSX embedding.

Provides helpers to compute twoCellAnchor coordinates (from/to) that
position OLE objects at sub-cell precision using EMU offsets.
"""

from __future__ import annotations

import re
from dataclasses import dataclass


# 1 mm  = 36000 EMU
# 1 pt  = 12700 EMU
# 1 in  = 914400 EMU
_UNIT_TO_EMU = {
    "mm": 36000,
    "pt": 12700,
    "in": 914400,
    "emu": 1,
}


def _parse_length(s: str) -> int:
    """Parse a length string like '5mm', '12pt', '1in' into EMU."""
    m = re.match(r"^\s*([0-9]*\.?[0-9]+)\s*(mm|pt|in|emu)\s*$", s.strip(), re.I)
    if not m:
        raise ValueError(
            f"Invalid length: {s!r}. Use a number followed by mm/pt/in/emu."
        )
    return round(float(m.group(1)) * _UNIT_TO_EMU[m.group(2).lower()])


def _parse_col_letter(col: str) -> int:
    """Convert Excel column letter(s) to 0-indexed number.  A=0, B=1, …"""
    col = col.upper().strip()
    result = 0
    for ch in col:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


@dataclass
class AnchorPoint:
    """One corner of a twoCellAnchor rectangle."""

    col: int  # 0-indexed column number
    col_off: int  # EMU offset within the column
    row: int  # 0-indexed row number
    row_off: int  # EMU offset within the row


@dataclass
class OleAnchor:
    """Full rectangle defined by from/to anchor points."""

    from_pt: AnchorPoint
    to_pt: AnchorPoint


def parse_anchor_spec(spec: str) -> AnchorPoint:
    """Parse a human-readable anchor spec into an :class:`AnchorPoint`.

    Accepted formats::

        B1              → col=1, row=0, offsets=0
        A1+5mm,3mm      → col=0, row=0, col_off=5mm, row_off=3mm
        C3+0pt,40pt     → col=2, row=2, col_off=0,  row_off=40pt

    The cell reference is 1-indexed (matching Excel convention).
    """
    m = re.match(
        r"^([A-Za-z]+)(\d+)"
        r"(?:\+([0-9.]+(?:mm|pt|in|emu)),([0-9.]+(?:mm|pt|in|emu)))?$",
        spec.strip(),
    )
    if not m:
        raise ValueError(
            f"Invalid anchor spec: {spec!r}. "
            "Expected format like 'B1' or 'A1+5mm,3mm'."
        )
    col = _parse_col_letter(m.group(1))
    row = int(m.group(2)) - 1  # 1-indexed → 0-indexed
    col_off = _parse_length(m.group(3)) if m.group(3) else 0
    row_off = _parse_length(m.group(4)) if m.group(4) else 0
    return AnchorPoint(col=col, col_off=col_off, row=row, row_off=row_off)


def compute_anchors(
    count: int,
    *,
    col: str = "B",
    start_row: int = 1,
    margin_left: str = "12mm",
    margin_top: str = "5mm",
    width: str = "22mm",
    height: str = "25mm",
    from_spec: str | None = None,
    to_spec: str | None = None,
) -> list[OleAnchor]:
    """Compute anchor positions for *count* OLE objects.

    If *from_spec* / *to_spec* are given they override the simple
    col/margin/size parameters and apply to the first object; subsequent
    objects are shifted down by one row each.
    """
    if from_spec and to_spec:
        base_from = parse_anchor_spec(from_spec)
        base_to = parse_anchor_spec(to_spec)
        anchors: list[OleAnchor] = []
        for i in range(count):
            anchors.append(
                OleAnchor(
                    from_pt=AnchorPoint(
                        col=base_from.col,
                        col_off=base_from.col_off,
                        row=base_from.row + i,
                        row_off=base_from.row_off,
                    ),
                    to_pt=AnchorPoint(
                        col=base_to.col,
                        col_off=base_to.col_off,
                        row=base_to.row + i,
                        row_off=base_to.row_off,
                    ),
                )
            )
        return anchors

    col_idx = _parse_col_letter(col)
    ml = _parse_length(margin_left)
    mt = _parse_length(margin_top)
    w = _parse_length(width)
    h = _parse_length(height)
    row0 = start_row - 1  # 1-indexed → 0-indexed

    anchors = []
    for i in range(count):
        anchors.append(
            OleAnchor(
                from_pt=AnchorPoint(
                    col=col_idx, col_off=ml, row=row0 + i, row_off=mt,
                ),
                to_pt=AnchorPoint(
                    col=col_idx, col_off=ml + w, row=row0 + i, row_off=mt + h,
                ),
            )
        )
    return anchors
