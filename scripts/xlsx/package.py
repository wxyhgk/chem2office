from __future__ import annotations

import os
import zipfile
from pathlib import Path


def extract_xlsx(xlsx_path: Path, output_dir: Path) -> None:
    """Extract an XLSX package into *output_dir*."""
    with zipfile.ZipFile(xlsx_path, "r") as z:
        z.extractall(output_dir)


def pack_xlsx(package_dir: Path, output: Path) -> None:
    """Pack *package_dir* contents into an XLSX file."""
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(package_dir):
            for name in files:
                full_path = Path(root) / name
                rel_path = full_path.relative_to(package_dir)
                z.write(full_path, rel_path)
