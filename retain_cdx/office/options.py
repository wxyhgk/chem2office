"""User-facing options shared by the Office writers."""

from __future__ import annotations

from dataclasses import dataclass

from retain_cdx.render.preview import PreviewBackend, PreviewFormat, PreviewStyle


@dataclass(frozen=True)
class PreviewOptions:
    width: int | None = None
    height: int | None = None
    thick_lines: bool = False
    no_margin: bool = False
    mono: bool = True
    # "png" default; "emf" uses SVG + Inkscape; "auto" prefers emf when available.
    format: PreviewFormat = "png"
    # "auto" prefers RDKit ACS1996, then Open Babel.
    backend: PreviewBackend = "auto"
    # RDKit drawing style; ACS1996 is monochrome journal style.
    style: PreviewStyle = "acs1996"
    #: 把渲染缓存进 OLE 容器（`\x02OlePres000`）。没有它，把结构复制到别的
    #: 程序时会现场唤起 ChemDraw 出图；有它则每个对象多 10~15 KB。
    #: ``None`` 表示按格式的默认值：PowerPoint 开，Excel 关。
    presentation_cache: bool | None = None


__all__ = ["PreviewOptions"]
