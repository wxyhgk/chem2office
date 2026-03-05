from __future__ import annotations

from pathlib import Path
import subprocess


def render_png(cdx_path: Path, png_path: Path) -> None:
    proc = subprocess.run(
        ["obabel", "-icdx", str(cdx_path), "-opng", "-O", str(png_path)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if proc.returncode != 0:
        raise SystemExit(f"obabel failed: {proc.stderr.strip()}")
