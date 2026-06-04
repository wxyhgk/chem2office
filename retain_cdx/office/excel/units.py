from __future__ import annotations

import re


EMU_PER_PIXEL = 9525
POINTS_PER_PIXEL = 0.75

_UNIT_TO_PX = {
    "px": 1.0,
    "pt": 1.0 / POINTS_PER_PIXEL,
    "in": 96.0,
    "mm": 96.0 / 25.4,
}


def parse_px(value: int | float | str) -> int:
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return round(value)

    match = re.match(r"^\s*([0-9]*\.?[0-9]+)\s*(px|pt|in|mm)\s*$", value, re.I)
    if not match:
        raise ValueError(f"Invalid length: {value!r}. Use px, pt, in, or mm.")
    number = float(match.group(1))
    unit = match.group(2).lower()
    return round(number * _UNIT_TO_PX[unit])


def px_to_points(px: int | float) -> float:
    return float(px) * POINTS_PER_PIXEL
