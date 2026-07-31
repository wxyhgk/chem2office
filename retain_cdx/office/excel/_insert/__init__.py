"""Insert pipeline (internal): plan → CDX files → source workbook → merge.

Public callers use ``ExcelCDXWriter``; this package is the orchestration layer
between options and the ``_source`` / ``_parts`` implementation packages.
"""

from __future__ import annotations

from retain_cdx.office.excel.models import (
    LayoutSpec,
    NormalizeSpec,
    PreviewSpec,
    ShellSourceOptions,
    SourceBuildRequest,
)

from .cdx import write_cdx_sources
from .executor import TargetSheet, execute_insert, resolve_target_sheet
from .plan import InsertPlan, build_insert_plan

__all__ = [
    "InsertPlan",
    "LayoutSpec",
    "NormalizeSpec",
    "PreviewSpec",
    "ShellSourceOptions",
    "SourceBuildRequest",
    "TargetSheet",
    "build_insert_plan",
    "execute_insert",
    "resolve_target_sheet",
    "write_cdx_sources",
]
