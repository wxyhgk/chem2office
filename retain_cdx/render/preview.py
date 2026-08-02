from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
import shutil
import subprocess
import tempfile
from typing import Any, Literal

PreviewFormat = Literal["auto", "png", "emf"]
ResolvedPreviewFormat = Literal["png", "emf"]
PreviewBackend = Literal["auto", "rdkit", "obabel"]
PreviewStyle = Literal["acs1996", "default"]

# How many SVG→EMF exports to send to one long-lived ``inkscape --shell`` process.
# Keeps memory bounded while avoiding per-file cold starts (~0.5s on macOS).
_INKSCAPE_SHELL_CHUNK = 80


def inkscape_available() -> bool:
    return shutil.which("inkscape") is not None


def obabel_available() -> bool:
    return shutil.which("obabel") is not None


def rdkit_available() -> bool:
    try:
        from rdkit.Chem.Draw import rdMolDraw2D  # noqa: F401

        return True
    except Exception:
        return False


def pycdxml_available() -> bool:
    try:
        from pycdxml.cdxml_converter.chemdraw_io import read_cdx  # noqa: F401

        return True
    except Exception:
        return False


def resolve_preview_format(requested: PreviewFormat | str) -> ResolvedPreviewFormat:
    """Resolve a preview format request to a concrete format."""
    value = (requested or "png").lower()
    if value == "auto":
        return "emf" if inkscape_available() else "png"
    if value == "png":
        return "png"
    if value == "emf":
        if not inkscape_available():
            raise SystemExit(
                "Inkscape is required for EMF preview. "
                "Install it (e.g. `brew install inkscape`) or use preview_format='png'."
            )
        return "emf"
    raise ValueError(f"Unsupported preview format: {requested!r}")


def resolve_preview_backend(requested: PreviewBackend | str = "auto") -> str:
    """Pick a concrete render backend.

    Preference for ``auto``:
      1. RDKit (+ pycdxml for CDX→CDXML) — ACS1996, no CLI
      2. Open Babel CLI — fallback
    """
    value = (requested or "auto").lower()
    if value == "rdkit":
        if not rdkit_available():
            raise SystemExit(
                "RDKit is required for preview_backend='rdkit'. "
                "Install with: pip install -e '.[chem]'"
            )
        if not pycdxml_available():
            raise SystemExit(
                "pycdxml is required to render CDX with RDKit. "
                "Install with: pip install -e '.[chem]'"
            )
        return "rdkit"
    if value == "obabel":
        if not obabel_available():
            raise SystemExit(
                "Open Babel (obabel) is required for preview_backend='obabel'."
            )
        return "obabel"
    if value == "auto":
        if rdkit_available() and pycdxml_available():
            return "rdkit"
        if obabel_available():
            return "obabel"
        raise SystemExit(
            "No preview backend available. Install chem extras "
            "(RDKit + pycdxml) or system Open Babel (obabel)."
        )
    raise ValueError(f"Unsupported preview backend: {requested!r}")


def mol_from_cdx(cdx_path: Path) -> Any:
    """Load the first molecule from a CDX file via pycdxml → CDXML → RDKit."""
    from pycdxml.cdxml_converter.chemdraw_io import read_cdx
    from rdkit import Chem
    from rdkit.Chem import AllChem, Draw

    doc = read_cdx(str(cdx_path))
    cdxml = doc.to_cdxml()
    mols = list(Chem.MolsFromCDXML(cdxml) or [])
    mols = [m for m in mols if m is not None]
    if not mols:
        raise ValueError(f"No molecules parsed from CDX: {cdx_path}")
    mol = mols[0]
    if mol.GetNumConformers() == 0:
        AllChem.Compute2DCoords(mol)
    elif mol.GetNumBonds() > 0:
        # Degenerate CDX coords (mean bond length 0) break ACS1996.
        try:
            if Draw.MeanBondLength(mol) <= 0:
                AllChem.Compute2DCoords(mol)
        except Exception:
            AllChem.Compute2DCoords(mol)
    return mol


def _acs1996_supported(mol: Any) -> bool:
    """ACS1996 needs at least one bond with positive mean bond length."""
    if mol is None or mol.GetNumBonds() <= 0:
        return False
    try:
        from rdkit.Chem import Draw

        return Draw.MeanBondLength(mol) > 0
    except Exception:
        return False


def render_mol_png(
    mol: Any,
    png_path: Path,
    *,
    width: int | None = None,
    height: int | None = None,
    image_size: int | None = None,
    style: PreviewStyle = "acs1996",
    mono: bool = True,
) -> None:
    """Render an RDKit molecule to PNG (ACS1996 by default)."""
    from rdkit.Chem import Draw

    png_path = Path(png_path)
    png_path.parent.mkdir(parents=True, exist_ok=True)

    if image_size is not None:
        width = height = image_size

    use_acs = style == "acs1996" and _acs1996_supported(mol)

    if width is None and height is None:
        # Flexi canvas is recommended for ACS1996 mode.
        drawer = Draw.MolDraw2DCairo(-1, -1)
        if use_acs:
            Draw.DrawMoleculeACS1996(drawer, mol)
        else:
            opts = drawer.drawOptions()
            if mono:
                opts.useBWAtomPalette()
            drawer.DrawMolecule(mol)
        drawer.FinishDrawing()
        png_path.write_bytes(drawer.GetDrawingText())
        return

    w = int(width or height or 300)
    h = int(height or width or 300)
    if use_acs:
        png_path.write_bytes(_acs_png_sized(mol, w, h))
        return
    drawer = Draw.MolDraw2DCairo(w, h)
    opts = drawer.drawOptions()
    if mono:
        opts.useBWAtomPalette()
    drawer.DrawMolecule(mol)
    drawer.FinishDrawing()
    png_path.write_bytes(drawer.GetDrawingText())


def _draw_acs_flexi_png(mol: Any, scale: float = 1.0) -> bytes:
    """Flexi-canvas ACS 1996 PNG; *scale* supersamples the whole drawing."""
    from rdkit.Chem import Draw

    drawer = Draw.MolDraw2DCairo(-1, -1)
    opts = drawer.drawOptions()
    Draw.SetACS1996Mode(opts, Draw.MeanBondLength(mol))
    if scale != 1.0:
        # ACS pins the drawn bond length; scale every fixed quantity so the
        # style's proportions survive supersampling.
        opts.fixedBondLength *= scale
        opts.scalingFactor *= scale
        opts.bondLineWidth *= scale
        if opts.fixedFontSize > 0:
            opts.fixedFontSize = max(1, round(opts.fixedFontSize * scale))
    drawer.DrawMolecule(mol)
    drawer.FinishDrawing()
    return drawer.GetDrawingText()


def _acs_png_sized(mol: Any, w: int, h: int) -> bytes:
    """ACS 1996 art scaled to fill a w×h canvas.

    ACS mode draws at a fixed physical size regardless of canvas, so a plain
    fixed-canvas draw leaves a tiny molecule adrift in whitespace. Measure the
    flexi render, redraw supersampled to fit, then pad to the exact canvas.
    """
    base = _draw_acs_flexi_png(mol)
    try:
        from PIL import Image
    except ImportError:
        # Without Pillow keep the legacy fixed canvas (molecule at ACS size).
        from rdkit.Chem import Draw

        drawer = Draw.MolDraw2DCairo(w, h)
        Draw.SetACS1996Mode(drawer.drawOptions(), Draw.MeanBondLength(mol))
        drawer.DrawMolecule(mol)
        drawer.FinishDrawing()
        return drawer.GetDrawingText()

    from io import BytesIO

    nw, nh = Image.open(BytesIO(base)).size
    scale = min(w / nw, h / nh) if nw > 0 and nh > 0 else 1.0
    art = _draw_acs_flexi_png(mol, scale=scale) if scale > 0 else base
    content = Image.open(BytesIO(art)).convert("RGB")
    buf = BytesIO()
    letterbox_image(
        content, canvas_width=w, canvas_height=h, fill_ratio=1.0, padding=0
    ).save(buf, format="PNG")
    return buf.getvalue()


def render_mol_svg(
    mol: Any,
    svg_path: Path,
    *,
    width: int | None = None,
    height: int | None = None,
    image_size: int | None = None,
    style: PreviewStyle = "acs1996",
    mono: bool = True,
) -> None:
    """Render an RDKit molecule to SVG (ACS1996 by default)."""
    from rdkit.Chem import Draw

    svg_path = Path(svg_path)
    svg_path.parent.mkdir(parents=True, exist_ok=True)

    if image_size is not None:
        width = height = image_size

    use_acs = style == "acs1996" and _acs1996_supported(mol)

    if width is None and height is None:
        if use_acs:
            svg_path.write_text(Draw.MolToACS1996SVG(mol), encoding="utf-8")
            return
        drawer = Draw.MolDraw2DSVG(-1, -1)
        opts = drawer.drawOptions()
        if mono:
            opts.useBWAtomPalette()
        drawer.DrawMolecule(mol)
        drawer.FinishDrawing()
        svg_path.write_text(drawer.GetDrawingText(), encoding="utf-8")
        return

    w = int(width or height or 300)
    h = int(height or width or 300)
    drawer = Draw.MolDraw2DSVG(w, h)
    opts = drawer.drawOptions()
    if use_acs:
        Draw.SetACS1996Mode(opts, Draw.MeanBondLength(mol))
    elif mono:
        opts.useBWAtomPalette()
    drawer.DrawMolecule(mol)
    drawer.FinishDrawing()
    svg_path.write_text(drawer.GetDrawingText(), encoding="utf-8")


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
    style: PreviewStyle = "acs1996",
    backend: PreviewBackend = "auto",
    crop: bool = False,
    crop_padding: int = 2,
    keep_size: bool = True,
    target_aspect: float | None = None,
    fill_ratio: float = 0.90,
) -> str:
    """Render CDX to PNG. Returns the backend used (``rdkit`` or ``obabel``)."""
    chosen = resolve_preview_backend(backend)
    png_path = Path(png_path)
    png_path.parent.mkdir(parents=True, exist_ok=True)

    if chosen == "rdkit":
        try:
            mol = mol_from_cdx(Path(cdx_path))
            # Explicit sizes set the canvas resolution; target_aspect only
            # letterboxes below (strategy B). Flexi-draw when no size given.
            render_mol_png(
                mol,
                png_path,
                width=width,
                height=height,
                image_size=image_size,
                style=style,
                mono=mono,
            )
            if crop or target_aspect is not None:
                crop_png_whitespace(
                    png_path,
                    padding=crop_padding,
                    keep_size=keep_size if target_aspect is None else True,
                    target_aspect=target_aspect,
                    fill_ratio=fill_ratio,
                )
            return "rdkit"
        except SystemExit:
            raise
        except Exception as exc:
            if backend == "rdkit":
                raise SystemExit(f"RDKit preview failed: {exc}") from exc
            # auto: fall through to Open Babel
            chosen = "obabel"
            if not obabel_available():
                raise SystemExit(
                    f"RDKit preview failed ({exc}) and Open Babel is not available."
                ) from exc

    cmd = _obabel_cmd(
        cdx_path,
        png_path,
        output_format="png",
        width=width,
        height=height,
        image_size=image_size,
        thick_lines=thick_lines,
        no_margin=no_margin,
        no_title=no_title,
        mono=mono,
    )
    _run(cmd, tool="obabel")
    if crop or target_aspect is not None:
        crop_png_whitespace(
            png_path,
            padding=crop_padding,
            keep_size=keep_size if target_aspect is None else True,
            target_aspect=target_aspect,
            fill_ratio=fill_ratio,
        )
    return "obabel"


def render_svg(
    cdx_path: Path,
    svg_path: Path,
    *,
    width: int | None = None,
    height: int | None = None,
    image_size: int | None = None,
    thick_lines: bool = False,
    no_margin: bool = False,
    no_title: bool = True,
    mono: bool = True,
    style: PreviewStyle = "acs1996",
    backend: PreviewBackend = "auto",
) -> str:
    """Render CDX to SVG. Returns the backend used."""
    chosen = resolve_preview_backend(backend)
    svg_path = Path(svg_path)
    svg_path.parent.mkdir(parents=True, exist_ok=True)

    if chosen == "rdkit":
        try:
            mol = mol_from_cdx(Path(cdx_path))
            render_mol_svg(
                mol,
                svg_path,
                width=width,
                height=height,
                image_size=image_size,
                style=style,
                mono=mono,
            )
            return "rdkit"
        except SystemExit:
            raise
        except Exception as exc:
            if backend == "rdkit":
                raise SystemExit(f"RDKit SVG preview failed: {exc}") from exc
            chosen = "obabel"
            if not obabel_available():
                raise SystemExit(
                    f"RDKit SVG preview failed ({exc}) and Open Babel is not available."
                ) from exc

    cmd = _obabel_cmd(
        cdx_path,
        svg_path,
        output_format="svg",
        width=width,
        height=height,
        image_size=image_size,
        thick_lines=thick_lines,
        no_margin=no_margin,
        no_title=no_title,
        mono=mono,
    )
    _run(cmd, tool="obabel")
    return "obabel"


def _resolve_emf_aspect(
    *,
    target_aspect: float | None,
    width: int | None,
    height: int | None,
    image_size: int | None,
) -> float | None:
    if target_aspect is not None and target_aspect > 0:
        return target_aspect
    if width and height and height > 0:
        return width / height
    if image_size:
        return 1.0
    return None


def _inkscape_action_path(path: Path) -> str:
    """Absolute path safe for Inkscape ``--shell`` action values.

    Action lists are ``;``-separated; reject paths that would break parsing.
    """
    resolved = str(Path(path).resolve())
    if ";" in resolved or "\n" in resolved:
        raise ValueError(f"Path not safe for inkscape --shell actions: {resolved!r}")
    return resolved


def svg_to_emf_batch(
    pairs: Sequence[tuple[Path, Path]],
    *,
    text_to_path: bool = True,
    chunk_size: int = _INKSCAPE_SHELL_CHUNK,
    companion_wmf: bool = False,
) -> None:
    """Convert many SVG files to EMF with one or more ``inkscape --shell`` sessions.

    Avoids cold-starting the full Inkscape.app per file (often ~0.5s on macOS).

    *companion_wmf* also writes ``<emf>.with_suffix(".wmf")`` from the same open
    document — PowerPoint's OLE presentation cache needs both (the EMF rides
    inside the WMF; see `retain_cdx/ole/presentation.py`). Exporting twice per
    open costs one extra Inkscape action, not a second SVG render.
    """
    if not pairs:
        return
    if not inkscape_available():
        raise SystemExit(
            "Inkscape is required for EMF preview. "
            "Install it (e.g. `brew install inkscape`)."
        )
    if chunk_size < 1:
        raise ValueError("chunk_size must be >= 1")

    items = [(Path(svg), Path(emf)) for svg, emf in pairs]
    for svg, emf in items:
        if not svg.is_file():
            raise SystemExit(f"SVG not found for EMF export: {svg}")
        emf.parent.mkdir(parents=True, exist_ok=True)

    for start in range(0, len(items), chunk_size):
        chunk = items[start : start + chunk_size]
        lines: list[str] = []
        for svg, emf in chunk:
            actions = [
                f"file-open:{_inkscape_action_path(svg)}",
                f"export-filename:{_inkscape_action_path(emf)}",
            ]
            if text_to_path:
                # Match CLI ``--export-text-to-path`` so Excel does not sub fonts.
                actions.append("export-text-to-path:true")
            actions.append("export-do")
            if companion_wmf:
                actions.append(
                    f"export-filename:{_inkscape_action_path(emf.with_suffix('.wmf'))}"
                )
                actions.append("export-do")
            actions.append("file-close")
            lines.append(";".join(actions))
        lines.append("quit")
        script = "\n".join(lines) + "\n"

        proc = subprocess.run(
            ["inkscape", "--shell"],
            input=script,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        if proc.returncode != 0:
            detail = (proc.stderr or proc.stdout or "").strip()
            raise SystemExit(f"inkscape --shell failed: {detail}")

        expected = [emf for _, emf in chunk]
        if companion_wmf:
            expected += [emf.with_suffix(".wmf") for _, emf in chunk]
        missing = [
            str(path)
            for path in expected
            if not path.exists() or path.stat().st_size == 0
        ]
        if missing:
            detail = (proc.stderr or proc.stdout or "").strip()
            hint = f" ({detail})" if detail else ""
            raise SystemExit(
                "Inkscape --shell did not produce EMF for: "
                + ", ".join(missing[:5])
                + ("…" if len(missing) > 5 else "")
                + hint
            )


def render_emf(
    cdx_path: Path,
    emf_path: Path,
    *,
    width: int | None = None,
    height: int | None = None,
    image_size: int | None = None,
    thick_lines: bool = False,
    no_margin: bool = False,
    no_title: bool = True,
    mono: bool = True,
    style: PreviewStyle = "acs1996",
    backend: PreviewBackend = "auto",
    target_aspect: float | None = None,
    crop_padding: int = 2,
    fill_ratio: float = 0.92,
    companion_wmf: bool = False,
) -> str:
    """Render CDX to a **vector** EMF via SVG + Inkscape.

    Pipeline (sharp at any zoom; matches shell box aspect):
      1. flexi ACS SVG (vector, tight to molecule)
      2. letterbox SVG viewBox to shell aspect / fill ratio
      3. Inkscape SVG → EMF via ``--shell`` (keeps paths; no PNG raster)
    """
    used = render_emf_many(
        [(Path(cdx_path), Path(emf_path))],
        width=width,
        height=height,
        image_size=image_size,
        thick_lines=thick_lines,
        no_margin=no_margin,
        no_title=no_title,
        mono=mono,
        style=style,
        backend=backend,
        target_aspect=target_aspect,
        crop_padding=crop_padding,
        fill_ratio=fill_ratio,
        companion_wmf=companion_wmf,
    )
    return used[0]


def render_emf_many(
    jobs: Sequence[tuple[Path, Path]],
    *,
    width: int | None = None,
    height: int | None = None,
    image_size: int | None = None,
    thick_lines: bool = False,
    no_margin: bool = False,
    no_title: bool = True,
    mono: bool = True,
    style: PreviewStyle = "acs1996",
    backend: PreviewBackend = "auto",
    target_aspect: float | None = None,
    crop_padding: int = 2,
    fill_ratio: float = 0.92,
    companion_wmf: bool = False,
) -> list[str]:
    """Render many CDX→EMF jobs with batched ``inkscape --shell`` exports.

    *jobs* is a sequence of ``(cdx_path, emf_path)``. SVG generation is per
    molecule (fast); Inkscape is invoked in chunks so the App is not
    cold-started once per molecule.
    """
    if not jobs:
        return []
    if not inkscape_available():
        raise SystemExit(
            "Inkscape is required for EMF preview. "
            "Install it (e.g. `brew install inkscape`)."
        )

    aspect = _resolve_emf_aspect(
        target_aspect=target_aspect,
        width=width,
        height=height,
        image_size=image_size,
    )

    used_backends: list[str] = []
    with tempfile.TemporaryDirectory(prefix="retain_cdx_emf_batch_") as tmp:
        tmpdir = Path(tmp)
        pairs: list[tuple[Path, Path]] = []
        for index, (cdx_path, emf_path) in enumerate(jobs):
            cdx_path = Path(cdx_path)
            emf_path = Path(emf_path)
            emf_path.parent.mkdir(parents=True, exist_ok=True)
            svg_path = tmpdir / f"{index:05d}_{cdx_path.stem}.svg"
            used = render_svg(
                cdx_path,
                svg_path,
                # Always flexi for ACS tightness; aspect is applied via viewBox.
                width=None,
                height=None,
                image_size=None,
                thick_lines=thick_lines,
                no_margin=no_margin,
                no_title=no_title,
                mono=mono,
                style=style,
                backend=backend,
            )
            if aspect is not None and aspect > 0:
                letterbox_svg_to_aspect(
                    svg_path,
                    target_aspect=aspect,
                    fill_ratio=fill_ratio,
                    padding_px=float(crop_padding),
                )
            pairs.append((svg_path, emf_path))
            used_backends.append(used)

        svg_to_emf_batch(pairs, companion_wmf=companion_wmf)

    return used_backends


def letterbox_svg_to_aspect(
    svg_path: Path,
    *,
    target_aspect: float,
    fill_ratio: float = 0.95,
    padding_px: float = 2.0,
) -> None:
    """Pad an SVG viewBox to *target_aspect* and scale content by *fill_ratio*.

    Content drawn in the original user space stays put; the viewBox expands
    (letterbox) so the final frame matches the OLE/shell display box.
    """
    import re

    if target_aspect <= 0:
        return
    fill_ratio = min(max(fill_ratio, 0.1), 1.0)
    text = Path(svg_path).read_text(encoding="utf-8", errors="replace")

    vb_match = re.search(
        r'viewBox=(["\'])\s*([-\d.eE+]+)\s+([-\d.eE+]+)\s+([-\d.eE+]+)\s+([-\d.eE+]+)\s*\1',
        text,
    )
    if vb_match:
        x0, y0 = float(vb_match.group(2)), float(vb_match.group(3))
        content_w, content_h = float(vb_match.group(4)), float(vb_match.group(5))
    else:
        wh = re.search(
            r'width=(["\'])([-\d.eE+]+)(?:px)?\1[^>]*height=(["\'])([-\d.eE+]+)(?:px)?\3',
            text,
            flags=re.I | re.S,
        )
        if not wh:
            wh = re.search(
                r'height=(["\'])([-\d.eE+]+)(?:px)?\1[^>]*width=(["\'])([-\d.eE+]+)(?:px)?\3',
                text,
                flags=re.I | re.S,
            )
            if not wh:
                return
            content_w, content_h = float(wh.group(4)), float(wh.group(2))
        else:
            content_w, content_h = float(wh.group(2)), float(wh.group(4))
        x0, y0 = 0.0, 0.0

    if content_w <= 0 or content_h <= 0:
        return

    # Inflate content box slightly so strokes/labels are not clipped.
    pad = max(0.0, padding_px)
    content_w += 2 * pad
    content_h += 2 * pad
    x0 -= pad
    y0 -= pad

    # Letterbox content into target aspect, then expand so content uses fill_ratio.
    content_aspect = content_w / content_h
    if content_aspect > target_aspect:
        box_w = content_w
        box_h = content_w / target_aspect
    else:
        box_h = content_h
        box_w = content_h * target_aspect

    box_w /= fill_ratio
    box_h /= fill_ratio
    # Keep exact aspect after float ops.
    box_h = box_w / target_aspect

    off_x = (box_w - content_w) / 2.0
    off_y = (box_h - content_h) / 2.0
    vb_x = x0 - off_x
    vb_y = y0 - off_y

    def _replace_viewbox(match: re.Match[str]) -> str:
        quote = match.group(1)
        return f'viewBox={quote}{vb_x:g} {vb_y:g} {box_w:g} {box_h:g}{quote}'

    if vb_match:
        text = re.sub(
            r'viewBox=(["\'])\s*[-\d.eE+]+\s+[-\d.eE+]+\s+[-\d.eE+]+\s+[-\d.eE+]+\s*\1',
            _replace_viewbox,
            text,
            count=1,
        )
    else:
        # Insert viewBox after <svg ...>
        text = re.sub(
            r"(<svg\b[^>]*)(>)",
            rf'\1 viewBox="{vb_x:g} {vb_y:g} {box_w:g} {box_h:g}"\2',
            text,
            count=1,
            flags=re.I,
        )

    # Keep width/height attributes consistent with the new frame aspect.
    out_w = 400.0
    out_h = out_w / target_aspect
    text = re.sub(
        r'\bwidth=(["\'])[^"\']*\1',
        f'width="{out_w:g}px"',
        text,
        count=1,
        flags=re.I,
    )
    text = re.sub(
        r'\bheight=(["\'])[^"\']*\1',
        f'height="{out_h:g}px"',
        text,
        count=1,
        flags=re.I,
    )

    Path(svg_path).write_text(text, encoding="utf-8")


def render_preview(
    cdx_path: Path,
    output_path: Path,
    *,
    format: PreviewFormat | str = "png",
    width: int | None = None,
    height: int | None = None,
    image_size: int | None = None,
    thick_lines: bool = False,
    no_margin: bool = False,
    no_title: bool = True,
    mono: bool = True,
    style: PreviewStyle = "acs1996",
    backend: PreviewBackend = "auto",
    crop: bool = False,
    crop_padding: int = 2,
    keep_size: bool = True,
    target_aspect: float | None = None,
    fill_ratio: float = 0.92,
    companion_wmf: bool = False,
) -> ResolvedPreviewFormat:
    """Render a preview image.

    Defaults:
      - format: PNG
      - backend: RDKit ACS1996 when available, else Open Babel
      - mono: True (ACS1996 is monochrome; Open Babel uses ``-xu``/``-xU``)
    """
    resolved = resolve_preview_format(format)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if resolved == "emf":
        render_emf(
            cdx_path,
            output_path,
            width=width,
            height=height,
            image_size=image_size,
            thick_lines=thick_lines,
            no_margin=no_margin,
            no_title=no_title,
            mono=mono,
            style=style,
            backend=backend,
            target_aspect=target_aspect,
            crop_padding=crop_padding,
            fill_ratio=fill_ratio,
            companion_wmf=companion_wmf,
        )
        return "emf"

    render_png(
        cdx_path,
        output_path,
        width=width,
        height=height,
        image_size=image_size,
        thick_lines=thick_lines,
        no_margin=no_margin,
        no_title=no_title,
        mono=mono,
        style=style,
        backend=backend,
        crop=crop,
        crop_padding=crop_padding,
        keep_size=keep_size,
        target_aspect=target_aspect,
        fill_ratio=fill_ratio,
    )
    return "png"


def preview_extension(format: PreviewFormat | str) -> str:
    """Return the file extension (without dot) for a resolved preview format."""
    return resolve_preview_format(format)


def crop_png_whitespace(
    png_path: Path,
    *,
    padding: int = 2,
    background: tuple[int, int, int] = (255, 255, 255),
    threshold: int = 245,
    keep_size: bool = True,
    target_aspect: float | None = None,
    fill_ratio: float = 1.0,
) -> None:
    from PIL import Image

    image = Image.open(png_path).convert("RGB")
    cropped = _crop_non_background(image, padding=padding, background=background, threshold=threshold)
    if cropped is None:
        return
    if not keep_size:
        cropped.save(png_path)
        return

    width, height = image.size
    if target_aspect is not None and target_aspect > 0:
        if width / height > target_aspect:
            canvas_height = height
            canvas_width = max(1, round(canvas_height * target_aspect))
        else:
            canvas_width = width
            canvas_height = max(1, round(canvas_width / target_aspect))
    else:
        canvas_width, canvas_height = width, height

    letterbox_image(
        cropped,
        canvas_width=canvas_width,
        canvas_height=canvas_height,
        fill_ratio=fill_ratio,
        padding=padding,
        background=background,
    ).save(png_path)


def letterbox_png(
    png_path: Path,
    *,
    canvas_width: int,
    canvas_height: int,
    fill_ratio: float = 0.92,
    padding: int = 2,
    background: tuple[int, int, int] = (255, 255, 255),
    threshold: int = 245,
    recrop: bool = True,
) -> None:
    """Fit the (optionally re-cropped) PNG content into a fixed canvas."""
    from PIL import Image

    image = Image.open(png_path).convert("RGB")
    content = image
    if recrop:
        cropped = _crop_non_background(
            image, padding=padding, background=background, threshold=threshold
        )
        if cropped is not None:
            content = cropped
    letterbox_image(
        content,
        canvas_width=canvas_width,
        canvas_height=canvas_height,
        fill_ratio=fill_ratio,
        padding=padding,
        background=background,
    ).save(png_path)


def letterbox_image(
    image,
    *,
    canvas_width: int,
    canvas_height: int,
    fill_ratio: float = 0.92,
    padding: int = 2,
    background: tuple[int, int, int] = (255, 255, 255),
):
    """Scale *image* to occupy *fill_ratio* of the canvas and center it."""
    from PIL import Image

    fill_ratio = min(max(fill_ratio, 0.1), 1.0)
    max_width = max(1, int((canvas_width - 2 * padding) * fill_ratio))
    max_height = max(1, int((canvas_height - 2 * padding) * fill_ratio))
    scale = min(max_width / max(image.width, 1), max_height / max(image.height, 1))
    resized_width = max(1, round(image.width * scale))
    resized_height = max(1, round(image.height * scale))
    resized = image.resize((resized_width, resized_height), Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", (canvas_width, canvas_height), background)
    offset = (
        (canvas_width - resized_width) // 2,
        (canvas_height - resized_height) // 2,
    )
    canvas.paste(resized, offset)
    return canvas


def _crop_non_background(
    image,
    *,
    padding: int,
    background: tuple[int, int, int],
    threshold: int,
):
    from PIL import Image, ImageChops

    width, height = image.size
    # 每个通道与背景的差值超过容差就算前景。原来是逐像素的 Python 双重循环，
    # 580×400 就是 23 万次解释器迭代；这些是 PIL 的 C 实现。
    tolerance = 255 - threshold
    difference = ImageChops.difference(image, Image.new("RGB", image.size, background))
    # 「任一通道超差」= 三通道取最大后比较；convert("L") 是亮度加权，不等价
    red, green, blue = difference.split()
    channel_max = ImageChops.lighter(ImageChops.lighter(red, green), blue)
    mask = channel_max.point(lambda value: 255 if value > tolerance else 0)
    box = mask.getbbox()
    if box is None:
        return None
    left, top, right, bottom = box
    right -= 1  # getbbox 的右/下界是开区间，下面按闭区间处理
    bottom -= 1

    left = max(0, left - padding)
    top = max(0, top - padding)
    right = min(width, right + padding + 1)
    bottom = min(height, bottom + padding + 1)
    return image.crop((left, top, right, bottom))


def _resolve_letterbox_canvas(
    png_path: Path,
    *,
    width: int | None,
    height: int | None,
    image_size: int | None,
    target_aspect: float | None,
    min_scale: int = 4,
) -> tuple[int, int]:
    """Choose a high-res canvas that matches the shell display aspect."""
    from PIL import Image

    if image_size is not None:
        return int(image_size), int(image_size)
    if width is not None and height is not None:
        return max(1, int(width)), max(1, int(height))

    with Image.open(png_path) as im:
        content_w, content_h = im.size

    if target_aspect is not None and target_aspect > 0:
        # Prefer a canvas at least min_scale times a typical shell pixel box,
        # while remaining large enough for the cropped content.
        base = max(content_w, content_h, 120)
        if target_aspect >= 1.0:
            canvas_w = max(base * min_scale // 2, content_w)
            canvas_h = max(1, round(canvas_w / target_aspect))
        else:
            canvas_h = max(base * min_scale // 2, content_h)
            canvas_w = max(1, round(canvas_h * target_aspect))
        # Ensure content can fit at fill_ratio ~0.92
        scale_w = content_w / max(canvas_w * 0.92, 1)
        scale_h = content_h / max(canvas_h * 0.92, 1)
        grow = max(scale_w, scale_h, 1.0)
        if grow > 1.0:
            canvas_w = max(1, round(canvas_w * grow))
            canvas_h = max(1, round(canvas_h * grow))
        return canvas_w, canvas_h

    if width is not None:
        return max(1, int(width)), max(1, int(width))
    if height is not None:
        return max(1, int(height)), max(1, int(height))
    return max(content_w, 1), max(content_h, 1)


def _obabel_cmd(
    cdx_path: Path,
    output_path: Path,
    *,
    output_format: str,
    width: int | None,
    height: int | None,
    image_size: int | None,
    thick_lines: bool,
    no_margin: bool,
    no_title: bool,
    mono: bool = True,
) -> list[str]:
    cmd = [
        "obabel",
        "-icdx",
        str(cdx_path),
        f"-o{output_format}",
        "-O",
        str(output_path),
    ]
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
        # -xu: no CPK/element atom colors; -xU: ignore embedded/internal colors.
        cmd.extend(["-xu", "-xU"])
    return cmd


def _run(cmd: list[str], *, tool: str) -> None:
    proc = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip()
        raise SystemExit(f"{tool} failed: {detail}")
