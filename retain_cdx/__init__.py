from __future__ import annotations

from .structure import StructureSource
from .office.excel import (
    DisplayOptions,
    ExcelCDXWriter,
    Placement,
    PreviewOptions,
    ShellTemplate,
)
from .office.powerpoint import PowerPointCDXWriter, SlideGrid, SlidePlacement

__all__ = [
    "DisplayOptions",
    "ExcelCDXWriter",
    "Placement",
    "PowerPointCDXWriter",
    "PreviewOptions",
    "ShellTemplate",
    "SlideGrid",
    "SlidePlacement",
    "StructureSource",
    "__version__",
]

__version__ = "0.0.11"
