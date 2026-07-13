from __future__ import annotations

from pathlib import Path
import subprocess


def render_png(
    cdx_path: Path,
    png_path: Path,
    *,
    width: int | None = None,
    height: int | None = None,
    image_size: int | None = None,
    thick_lines: bool = False,
    no_margin: bool = False,
    no_title: bool = True,
    mono: bool = True,
) -> None:
    cmd = ["obabel", "-icdx", str(cdx_path), "-opng", "-O", str(png_path)]
    if image_size is not None:
        cmd.extend(["-xp", str(image_size)])
    if width is not None:
        cmd.extend(["-xw", str(width)])
    if height is not None:
        cmd.extend(["-xh", str(height)])
    if thick_lines:
        cmd.append("-xt")
    if no_margin:
        cmd.append("-xm")
    if no_title:
        cmd.append("-xd")
    if mono:
        cmd.extend(["-xu", "-xU"])

    proc = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if proc.returncode != 0:
        raise SystemExit(f"obabel failed: {proc.stderr.strip()}")
