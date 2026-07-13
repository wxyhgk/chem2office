from __future__ import annotations

from .chrome import remove_shell_ole_border, tighten_shell_display_box
from .geometry import patch_shell_geometry, read_vml_size_px
from .preview import replace_shell_preview

__all__ = [
    "patch_shell_geometry",
    "read_vml_size_px",
    "remove_shell_ole_border",
    "replace_shell_preview",
    "tighten_shell_display_box",
]
