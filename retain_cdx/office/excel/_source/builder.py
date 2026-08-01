"""把 `SourceBuildRequest` 交给 shell 组装。

generated 路径已改为直接写目标包（`_insert_ole`），所以这里不再分发——
留着这层是为了让 executor 只认一个中性的 request，不必知道 shell 的参数形状。
"""

from __future__ import annotations

from pathlib import Path

from retain_cdx.office.excel.models import SourceBuildRequest

from .workbook import build_source_workbook_from_shell


def build_source_workbook(
    cdx_paths: list[Path],
    source_dir: Path,
    request: SourceBuildRequest,
) -> None:
    """Assemble the shell parts into *source_dir* (unpacked, not zipped)."""
    if request.shell is None:
        raise ValueError("build_source_workbook 现在只服务 shell 路径")
    build_source_workbook_from_shell(
        cdx_paths[0],
        source_dir,
        layout=request.layout,
        shell=request.shell,
        preview=request.preview,
    )
