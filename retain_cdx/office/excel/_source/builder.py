"""Dispatch a neutral ``SourceBuildRequest`` to the shell or generated path.

builder 只认 ``SourceBuildRequest``（不 import plan）：有 shell 走 shell 路径，
否则走 generated 路径；具体组装在 ``workbook.py``。
"""

from __future__ import annotations

from pathlib import Path

from retain_cdx.office.excel.models import SourceBuildRequest

from .workbook import build_source_ole_workbook, build_source_workbook_from_shell


def build_source_workbook(
    cdx_paths: list[Path],
    output_xlsx: Path,
    request: SourceBuildRequest,
) -> None:
    """Build a temporary source workbook from a neutral request (no plan import)."""
    if request.shell is not None:
        build_source_workbook_from_shell(
            cdx_paths[0],
            output_xlsx,
            layout=request.layout,
            shell=request.shell,
            preview=request.preview,
        )
        return

    if request.layout.row_height is None:
        raise ValueError("SourceBuildRequest.layout.row_height is required without a shell.")

    build_source_ole_workbook(
        cdx_paths,
        output_xlsx,
        layout=request.layout,
        preview=request.preview,
    )
