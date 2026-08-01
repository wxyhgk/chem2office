"""Zip in / zip out for OOXML packages — format agnostic.

Every OOXML file is a ZIP; Excel, PowerPoint and Word all share this step,
so it lives above the per-format packages.
"""

from __future__ import annotations

import os
import zipfile
from pathlib import Path


def extract_ooxml(package_path: Path, output_dir: Path) -> None:
    """Extract any OOXML package (.xlsx / .pptx / .docx) into *output_dir*."""
    output_root = output_dir.resolve()
    with zipfile.ZipFile(package_path, "r") as z:
        for member in z.infolist():
            target = (output_dir / member.filename).resolve()
            if target != output_root and output_root not in target.parents:
                raise ValueError(f"Unsafe OOXML member path: {member.filename}")
            z.extract(member, output_dir)


def pack_ooxml(package_dir: Path, output: Path) -> None:
    """Pack *package_dir* contents into an OOXML file."""
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(package_dir):
            for name in files:
                full_path = Path(root) / name
                rel_path = full_path.relative_to(package_dir)
                z.write(full_path, rel_path)
