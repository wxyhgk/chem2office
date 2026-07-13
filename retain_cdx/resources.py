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
