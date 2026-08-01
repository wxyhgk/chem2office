"""Length units shared by the OOXML formats.

EMU is the universal OOXML unit; pixels and points are how Excel talks
about its grid. PresentationML positions purely in EMU.
"""

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


# PresentationML positions everything in EMU (no grid to resolve against).
EMU_PER_INCH = 914400
_UNIT_TO_EMU = {
    "emu": 1.0,
    "px": float(EMU_PER_PIXEL),
    "pt": EMU_PER_INCH / 72.0,
    "in": float(EMU_PER_INCH),
    "mm": EMU_PER_INCH / 25.4,
    "cm": EMU_PER_INCH / 2.54,
}


def parse_emu(value: int | float | str) -> int:
    """Parse a length to EMU. Bare numbers are already EMU."""
    if isinstance(value, (int, float)):
        return round(value)

    match = re.match(r"^\s*([0-9]*\.?[0-9]+)\s*(emu|px|pt|in|mm|cm)\s*$", value, re.I)
    if not match:
        raise ValueError(
            f"Invalid length: {value!r}. Use emu, px, pt, in, mm, or cm."
        )
    return round(float(match.group(1)) * _UNIT_TO_EMU[match.group(2).lower()])
