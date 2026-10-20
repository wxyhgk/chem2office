from __future__ import annotations

from contextlib import contextmanager
from importlib.resources import as_file, files
from pathlib import Path
from typing import Iterator


@contextmanager
def bundled_template_xlsx() -> Iterator[Path]:
    """Yield the packaged default ChemDraw OLE template as a filesystem path."""
    resource = files("retain_cdx.data").joinpath("template.xlsx")
    with as_file(resource) as path:
        yield path


@contextmanager
def bundled_template_pptx() -> Iterator[Path]:
    """Yield the packaged PowerPoint ChemDraw OLE template.

    A single slide carrying one native PowerPoint+ChemDraw object; the writer
    clones its ``p:graphicFrame`` and OLE container for every structure.
    """
    resource = files("retain_cdx.data").joinpath("template.pptx")
    with as_file(resource) as path:
        yield path


@contextmanager
def bundled_template_docx() -> Iterator[Path]:
    """Yield the packaged Word ChemDraw OLE template.

    One paragraph carrying one native Word+ChemDraw object; the writer clones
    its ``w:object`` for every structure.
    """
    resource = files("retain_cdx.data").joinpath("template.docx")
    with as_file(resource) as path:
        yield path
