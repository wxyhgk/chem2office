from __future__ import annotations

from .structure import Compound, CompoundTable, StructureSource
from .office.excel import (
    ExcelCellAlignment,
    ExcelDisplayOptions,
    ExcelCDXWriter,
    ExcelPlacement,
    PreviewOptions,
    ExcelShellTemplate,
)
from .office.powerpoint import PowerPointCDXWriter, PowerPointGrid, PowerPointPlacement
from .office.word import WordCDXWriter

__all__ = [
    "Compound",
    "CompoundTable",
    "ExcelCellAlignment",
    "ExcelDisplayOptions",
    "ExcelCDXWriter",
    "ExcelPlacement",
    "PowerPointCDXWriter",
    "PreviewOptions",
    "ExcelShellTemplate",
    "PowerPointGrid",
    "PowerPointPlacement",
    "StructureSource",
    "WordCDXWriter",
    "__version__",
]

__version__ = "0.0.14"


# --- 0.0.11 之前的名字 -------------------------------------------------
# Excel 的这批诞生时包里只有 Excel，所以没有格式限定词；PowerPoint 进来后
# 就分不清归属了。旧名保留一个版本，用时告警。
_RENAMED = {
    "Placement": "ExcelPlacement",
    "DisplayOptions": "ExcelDisplayOptions",
    "ShellTemplate": "ExcelShellTemplate",
    "SlidePlacement": "PowerPointPlacement",
    "SlideGrid": "PowerPointGrid",
}


def __getattr__(name: str):
    new = _RENAMED.get(name)
    if new is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    import warnings

    warnings.warn(
        f"retain_cdx.{name} was renamed to {new}; the old name will be removed "
        "in a future release.",
        DeprecationWarning,
        stacklevel=2,
    )
    return globals()[new]
