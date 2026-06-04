from __future__ import annotations

import re


def col_to_index(col: str) -> int:
    value = 0
    for ch in col.upper():
        value = value * 26 + ord(ch) - ord("A") + 1
    return value


def index_to_col(index: int) -> str:
    out = ""
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


def parse_cell(cell: str) -> tuple[int, int]:
    match = re.match(r"^([A-Za-z]+)(\d+)$", cell.strip())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell!r}")
    return col_to_index(match.group(1)), int(match.group(2))


def shift_cell_ref(ref: str, *, start_col: int, delta: int) -> str:
    match = re.match(r"^([A-Za-z]+)(\d+)$", ref)
    if not match:
        return ref
    col = col_to_index(match.group(1))
    if col >= start_col:
        col += delta
    return f"{index_to_col(col)}{match.group(2)}"
