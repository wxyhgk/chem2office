"""原生 ChemDraw OLE shell 的改造（单对象）。

三条插入路径里只有 shell 走这里：它拿一份 Windows Excel + ChemDraw 真实产出的
shell 包，整份改造后交给 `_parts.merge` 搬进目标表。generated 与 picture 都是
直接往目标包写零件。

子步骤::

  B   preview.replace_shell_preview — 换预览图与 image rels
  D   geometry — 贴合分子比例 / 挪格改尺寸
  chrome — 去掉黑色 OLE 边框

编排顺序在 `staging.py`；底层：`model` / `vml` / `ooxml_patch` / `media_rels`。
"""

from __future__ import annotations

from .chrome import remove_shell_ole_border
from .geometry import (
    patch_shell_geometry,
    read_vml_size_px,
    tighten_shell_display_box,
)
from .preview import replace_shell_preview
from .staging import build_shell_staging

__all__ = [
    "build_shell_staging",
    "patch_shell_geometry",
    "read_vml_size_px",
    "remove_shell_ole_border",
    "replace_shell_preview",
    "tighten_shell_display_box",
]
