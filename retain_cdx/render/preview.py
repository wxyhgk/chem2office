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
    crop: bool = False,
    crop_padding: int = 2,
    keep_size: bool = True,
    target_aspect: float | None = None,
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

    proc = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if proc.returncode != 0:
        raise SystemExit(f"obabel failed: {proc.stderr.strip()}")
    if crop:
        crop_png_whitespace(
            png_path,
            padding=crop_padding,
            keep_size=keep_size,
            target_aspect=target_aspect,
        )


def crop_png_whitespace(
    png_path: Path,
    *,
    padding: int = 2,
    background: tuple[int, int, int] = (255, 255, 255),
    threshold: int = 245,
    keep_size: bool = True,
    target_aspect: float | None = None,
) -> None:
    from PIL import Image

    image = Image.open(png_path).convert("RGB")
    pixels = image.load()
    width, height = image.size
    left, top = width, height
    right, bottom = -1, -1

    for y in range(height):
        for x in range(width):
            red, green, blue = pixels[x, y]
            if (
                abs(red - background[0]) > 255 - threshold
                or abs(green - background[1]) > 255 - threshold
                or abs(blue - background[2]) > 255 - threshold
            ):
                left = min(left, x)
                top = min(top, y)
                right = max(right, x)
                bottom = max(bottom, y)

    if right < left or bottom < top:
        return
    left = max(0, left - padding)
    top = max(0, top - padding)
    right = min(width, right + padding + 1)
    bottom = min(height, bottom + padding + 1)
    cropped = image.crop((left, top, right, bottom))
    if not keep_size:
        cropped.save(png_path)
        return

    if target_aspect is not None and target_aspect > 0:
        if width / height > target_aspect:
            canvas_height = height
            canvas_width = max(1, round(canvas_height * target_aspect))
        else:
            canvas_width = width
            canvas_height = max(1, round(canvas_width / target_aspect))
    else:
        canvas_width, canvas_height = width, height
    max_width = max(1, canvas_width - (2 * padding))
    max_height = max(1, canvas_height - (2 * padding))
    scale = min(max_width / cropped.width, max_height / cropped.height)
    resized_width = max(1, round(cropped.width * scale))
    resized_height = max(1, round(cropped.height * scale))
    resized = cropped.resize((resized_width, resized_height), Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", (canvas_width, canvas_height), background)
    offset = ((canvas_width - resized_width) // 2, (canvas_height - resized_height) // 2)
    canvas.paste(resized, offset)
    canvas.save(png_path)
