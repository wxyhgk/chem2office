# Changelog

## 0.0.8 — 2026-07-13

Preview rendering, shell OLE geometry, and packaging improvements for server-friendly ChemDraw Excel embeds.

### Preview rendering

- Default preview backend is now **RDKit ACS 1996** (monochrome journal style) via `pycdxml` CDX→CDXML → RDKit draw.
- Fallback to **Open Babel** CLI when RDKit/pycdxml are unavailable.
- Default monochrome Open Babel path uses `-xu` / `-xU` (no element coloring).
- **EMF previews** are generated as **vectors**: RDKit ACS SVG → letterbox viewBox to display aspect → Inkscape EMF (no PNG raster path).
- Preview API: `preview_backend`, `preview_style`, `preview_mono`, `preview_format` (`png` / `emf` / `auto`).

### Shell OLE geometry (fixed distortion / layout)

- **Stretch bug fixed**: worksheet `oleObject` anchors use `<from>/<to>`; patching only updated `<xdr:from>/<xdr:to>`, so sheet anchors lagged behind VML/drawing and Excel stretched the preview. Both forms are now updated together.
- Anchors are computed from the **worksheet’s real column widths / row heights** (not hardcoded 64×19 defaults), so custom shell layouts no longer warp the object.
- Display box is **tightened to molecule aspect** so the outer frame hugs the structure (less empty padding).
- **Black OLE border removed** (`stroked="f"`, DrawingML `<a:ln><a:noFill/>`).
- **Center in target cell** when no explicit in-cell offset is given (important for wide/tall shell cells such as `shell_large` A1).
- `ShellTemplate.allow_resize` defaults to **True** so `cell=` placement is honored.

### Packaging / project layout

- Package metadata via `pyproject.toml` (`retain-cdx` 0.0.8).
- Bundled shell template under `retain_cdx/data/`; template unpack reference under `docs/template_reference/`.
- Examples moved under `examples/`.

### Tests

- Expanded shell/layout tests: API exports, EMF path, monochrome/ACS backend, sheet↔drawing anchor sync, no border, cell centering.
