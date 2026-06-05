from __future__ import annotations

from dataclasses import dataclass
import re


_UNIT_TO_EMU = {
    "mm": 36000,
    "pt": 12700,
    "in": 914400,
    "emu": 1,
}


@dataclass
class AnchorPoint:
    col: int
    col_off: int
    row: int
    row_off: int


@dataclass
class OleAnchor:
    from_pt: AnchorPoint
    to_pt: AnchorPoint


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
    if from_spec and to_spec:
        base_from = parse_anchor_spec(from_spec)
        base_to = parse_anchor_spec(to_spec)
        anchors: list[OleAnchor] = []
        for idx in range(count):
            anchors.append(
                OleAnchor(
                    from_pt=AnchorPoint(
                        col=base_from.col,
                        col_off=base_from.col_off,
                        row=base_from.row + idx,
                        row_off=base_from.row_off,
                    ),
                    to_pt=AnchorPoint(
                        col=base_to.col,
                        col_off=base_to.col_off,
                        row=base_to.row + idx,
                        row_off=base_to.row_off,
                    ),
                )
            )
        return anchors

    col_idx = _parse_col_letter(col)
    left = _parse_length(margin_left)
    top = _parse_length(margin_top)
    width_emu = _parse_length(width)
    height_emu = _parse_length(height)
    row0 = start_row - 1

    anchors = []
    for idx in range(count):
        anchors.append(
            OleAnchor(
                from_pt=AnchorPoint(
                    col=col_idx,
                    col_off=left,
                    row=row0 + idx,
                    row_off=top,
                ),
                to_pt=AnchorPoint(
                    col=col_idx,
                    col_off=left + width_emu,
                    row=row0 + idx,
                    row_off=top + height_emu,
                ),
            )
        )
    return anchors


def parse_anchor_spec(spec: str) -> AnchorPoint:
    match = re.match(
        r"^([A-Za-z]+)(\d+)"
        r"(?:\+([0-9.]+(?:mm|pt|in|emu)),([0-9.]+(?:mm|pt|in|emu)))?$",
        spec.strip(),
    )
    if not match:
        raise ValueError(
            f"Invalid anchor spec: {spec!r}. "
            "Expected format like 'B1' or 'A1+5mm,3mm'."
        )
    return AnchorPoint(
        col=_parse_col_letter(match.group(1)),
        col_off=_parse_length(match.group(3)) if match.group(3) else 0,
        row=int(match.group(2)) - 1,
        row_off=_parse_length(match.group(4)) if match.group(4) else 0,
    )


def _parse_length(value: str) -> int:
    match = re.match(r"^\s*([0-9]*\.?[0-9]+)\s*(mm|pt|in|emu)\s*$", value.strip(), re.I)
    if not match:
        raise ValueError(
            f"Invalid length: {value!r}. Use a number followed by mm/pt/in/emu."
        )
    return round(float(match.group(1)) * _UNIT_TO_EMU[match.group(2).lower()])


def _parse_col_letter(col: str) -> int:
    result = 0
    for ch in col.upper().strip():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1
