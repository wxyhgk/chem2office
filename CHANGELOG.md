# Changelog

## 0.0.11 — 2026-07-31

Closes the remaining findings from `docs/excel_ooxml_research.md`; every C/F
item in that audit is now fixed.

### Target-workbook fidelity

- **Fixed**: inserting into a sheet that already had a chart or comments left
  the freshly copied drawing/VML part carrying the *source's* shape ids while
  the other layer was remapped, so the `oleObject` bound to a shape id present
  on only one of them. Both sides are remapped now.
- **Fixed**: appending picture shapes into a comment's VML left them pointing at
  an undefined `#_x0000_t75`; the missing `v:shapetype` definitions are copied.
- **Fixed**: `set_cell_text` dropped the cell's `s=` style on overwrite, and
  overwriting a formula cell left a stale `xl/calcChain.xml` entry behind —
  Excel reports "Removed Records: Formula" for that. The chain part (a cache
  Excel rebuilds) is now removed when a formula is replaced.
- **Fixed**: macro-enabled workbooks declare `Default Extension="bin"` as
  `vbaProject`, so embedded OLE parts inherited the wrong content type. They
  get an explicit per-part `Override` when the default disagrees.
- **Fixed**: `set_column_width` appended a second `<col>` inside an existing
  grouped range (`min="1" max="10"`), leaving overlapping definitions; the
  covering range is split instead.

### Markup compatibility

- **Fixed**: `mc:Ignorable` was overwritten with a fixed token list, dropping
  whatever the target declared (`xr6`, …). The lists are merged.
- **Fixed**: worksheets lost `xmlns` declarations for prefixes no element
  happens to use — exactly what `mc:Ignorable` names. `write_sheet_xml` now
  carries the source root's declarations over. (Same root cause as the
  PowerPoint `xmlns:v` bug: ElementTree only emits prefixes in actual use.)

### Anchors

- **Fixed**: the `to` corner stayed on the from-cell with an offset that could
  run several cells past its edge. Excel copes and re-normalises on save, but
  Excel-authored files never look like that and other consumers clamp. Offsets
  are now walked into the cell they land in, across all three representations
  (worksheet, drawing, VML). A no-op unless the box spans cells — an ordinary
  insert fits its cell and is untouched.

## 0.0.10 — 2026-07-31

### OLE container correctness

- **Fixed**: `build_chemdraw_ole` hardcoded a single FAT sector, so any CDX over
  63,488 bytes produced a silently corrupt container — `olefile` still opened
  it, and only reading `CONTENTS` revealed the damage. That is exactly the shape
  of the historical "Excel deleted the object" reports. FAT, mini-FAT and the
  directory now grow to as many sectors as the payload needs (verified from
  1 KB to 1 MB), and a payload beyond the header-only DIFAT raises instead of
  emitting a broken file.
- **Fixed**: shrinking a stream across the 4096-byte mini-stream cutoff was
  patched in place, leaving the start sector to be resolved against the wrong
  allocator. Crossing the cutoff now rebuilds.
- **Fixed**: `patch_direntry_size` assumed the directory fit in one sector.
- Rebuilt containers pad free mini-FAT entries with `FREESECT`, matching
  Windows-authored files byte for byte for the same content.
- `write_cdx_to_ole` no longer uses `SystemExit` message matching as control
  flow, and drops any cached `\x02OlePres000` presentation stream rather than
  shipping one that renders the *previous* molecule.

### Excel geometry and relationships

- **Fixed**: column widths assumed Calibri 11's 7px max digit width. The unit is
  the workbook's normal font (ECMA-376 §18.3.1.13) — the bundled template is
  itself a 等线-11 file where a default column is 72px, not 64. The font is now
  read from `xl/styles.xml`, with a lookup table and a Calibri fallback.
- **Fixed**: `baseColWidth` was converted as `chars * mdw + 5` without the
  round-up to a multiple of 8 the spec requires, so a plain openpyxl workbook
  resolved its default column to 61px instead of 64 — enough to leave a
  two-point placement 9px short.
- **Fixed**: rebuilding a sheet's `.rels` dropped every attribute beyond
  Id/Type/Target, so `TargetMode="External"` was lost and each hyperlink on the
  target sheet became a broken internal part reference.

### PowerPoint support (new)

- **Scope**: `PowerPointCDXWriter` places molecules on slides of a deck you
  already have. Building the presentation — slides, layouts, titles, tables —
  is left to a presentation library such as `python-pptx`; `input_pptx` is
  required. Verified in both directions: python-pptx can build the deck before
  we inject, and it round-trips our OLE objects intact afterwards (CDX bytes,
  media and relationships all survive), unlike openpyxl on the Excel side.
- `PowerPointCDXWriter` inserts editable ChemDraw OLE objects into `.pptx`.
  PresentationML positions everything in absolute EMU, so none of the Excel-side
  grid machinery applies — no column widths, no anchor triples, no VML layer.
  `SlidePlacement` takes `cm/mm/in/pt/px/emu` and derives the missing dimension
  from the molecule's aspect; `SlideGrid` tiles several structures on one slide.
- Bundled `retain_cdx/data/template.pptx` supplies the native PowerPoint+ChemDraw
  `p:graphicFrame` and OLE container that every insert clones. Its sample object
  (and the orphaned parts and relationships behind it) is cleared on load.
- `retain_cdx/ole` and `retain_cdx/render` are reused unchanged: the PPT
  container's `\x01CompObj` and `\x01Ole` streams are byte-identical to the
  Excel template's, with the same CLSID and ProgID.

### Fixed while building it

- **OLE**: `patch_direntry_size` assumed the directory stream fit in one sector
  and refused to work otherwise. PowerPoint's ChemDraw containers carry a
  `\x02OlePres000` presentation stream, which pushes the directory past one
  sector, so every in-place injection failed. The directory FAT chain is now
  walked properly. (Also corrects the earlier finding that ChemDraw OLE never
  carries a presentation stream — that holds for Excel, not PowerPoint.)
- **PresentationML**: ElementTree drops namespace declarations no element
  actually uses, which stripped the `xmlns:v` that the frame's
  `<mc:Choice Requires="v">` depends on — PowerPoint then offered to repair the
  deck. Slides are now written through `write_slide_xml`, which re-attaches it.
- **PresentationML**: the cloned fallback picture kept `cNvPr id="0"` on every
  structure, so shape ids collided within a slide. All ids in a clone are
  renumbered.

### Tests

- `tests/ooxml_checks.py` — a reusable package oracle (dangling relationship
  targets, missing content types, duplicate shape ids, and `mc:*` `Requires` /
  `Ignorable` prefixes that resolve to nothing). Both regressions above were
  confirmed to fail the suite when the fix is reverted.
- 12 PowerPoint tests covering placement units, grid layout, per-structure parts
  and relationships, template-orphan cleanup, and the bundled template itself.

## 0.0.9 — 2026-07-31

Excel internals restructure, preview sizing fixes, and cell-text public API.

### Excel internals (no public behavior change)

- `_source/template_*.py` modules consolidated into the `_source/template/` package.
- `_parts/ole.py` split into focused modules: `merge` (orchestrator), `context`, `surface`, `ole_objects`, `util`. The unused `_parts/ole.py` compat shim was removed (internal path, no import guarantee).
- New internal DTO layer `office/excel/models.py` (`PreviewSpec` / `LayoutSpec` / `NormalizeSpec` / `ShellSourceOptions` / `SourceBuildRequest`), now threaded end-to-end: `builder` → `workbook` → `generated` → `fill_parts` and the shell path take spec objects instead of re-declaring 12-18 scalar kwargs per signature; the two near-identical generated orchestrators were merged.
- `_ooxml` is now the single source of truth for OOXML plumbing: namespace URI constants and `register_namespaces` (previously triplicated in `template/namespaces.py` and `_parts/surface.py`), relationship-type URLs, and the `.rels` read/rewrite helpers (`read_rels` / `rewrite_rel_targets`) used by `_parts` and `_shell/media_rels`.
- Removed dead code: `template/xml.py` (facade of a facade) and the unused `extract_base_ole_anchor` export.

### Picture mode (optional, for large batches)

- New `ExcelCDXWriter.insert_pictures_many()` inserts structures as plain Excel pictures instead of OLE objects. A picture is one `xdr:pic` anchor plus a media part — no CFBF container, no VML shape, no activation machinery — so a single sheet can hold thousands of structures and stay responsive. Anchors are computed from the target sheet's real column widths / row heights, and the drawing part is reused when the sheet already has one. Trade-off: pictures cannot be opened in ChemDraw by double-click, so keep `insert_cdx_many` for rows that must stay editable.
- `add_smiles_ole_by_merge --picture-mode` exposes it. Measured on 50 molecules against the OLE path: no OLE containers (0 vs 361 KB), no VML (0 vs 33 KB), 2.9× less shape XML (241 vs 689 KB), and a smaller workbook (159 vs 296 KB).
- `add_smiles_ole_by_merge --no-smiles-column` drops the SMILES text column so structures take column A. This is a layout choice only — the SMILES text is well under 1% of the file and never the reason Excel struggles.

### OLE object display box / large-batch performance

- **Fixed**: the generated path stripped `autoPict` from `objectPr`, leaving the spec default (*true*) in force — "size this object from its referenced picture". Excel then measures every preview image at layout time, which is both wrong (previews are deliberately supersampled) and expensive on sheets holding thousands of objects. `normalize_object_pr` now writes `autoPict="0"` and `defaultSize="0"` explicitly, matching Excel-authored files.
- `add_smiles_ole_by_merge --split N` writes one workbook per N structures. Thousands of embedded OLE objects on a single sheet are slow in Excel on modest hardware regardless of file size — each object carries a CFBF container, activation machinery, and a shape in both the DrawingML and VML layers. Splitting keeps every file responsive with no loss of editability.

### Worksheet schema order (fixed Excel repair prompts)

- **Fixed**: `attach_ole_objects` inserted `drawing` / `legacyDrawing` / `oleObjects` immediately after `pageMargins`, which pushed any existing `pageSetup` / `headerFooter` behind them. That violates the fixed CT_Worksheet child sequence (ECMA-376 §18.3.1.99), and Excel responds by repairing the file — which drops the drawing and OLE parts. Any target workbook carrying page-setup settings (i.e. most real Excel files) was affected.
- New `insert_worksheet_child()` in `_ooxml/xml_io.py` places elements by schema rank; it also backs `dimension` creation in `_ooxml/cells.py`.
- `attach_ole_objects` now finds `<oleObjects>` nested in `mc:AlternateContent` (Excel's own serialization) and appends into it instead of creating a second, schema-illegal container.

### Preview sizing fixes

- **Fixed**: the RDKit PNG path discarded explicit `preview_width` / `preview_height` whenever a target aspect was in play (always, since strategy B) — requesting 1200×900 produced a ~44×30 px raster. Explicit sizes now set the canvas; the target aspect only letterboxes.
- **Fixed**: ACS 1996 pins the drawn bond length, so explicit sizes used to leave a tiny molecule adrift in whitespace (or a tiny raster upscaled blurrily). The molecule is now vector-scaled to fill the requested canvas, keeping ACS proportions.
- Default PNG previews (no explicit size) are supersampled to 4× the OLE display box instead of the ACS flexi canvas (~25 px for small molecules, blurry once Excel upscaled it).
- Shell previews fall back to the molecule's own aspect when the VML display box cannot be parsed.

### Preview rendering

- EMF previews are batched through a persistent `inkscape --shell` process (`render_emf_many`) instead of one Inkscape cold start per molecule.

### Public API

- `ExcelCDXWriter.set_cell_text` / `set_cell_texts` / `set_column_width` / `set_row_height` — write plain text (IDs, headers) and grid sizes next to inserted structures. Use these instead of post-processing with openpyxl, which drops OLE parts.
- `Placement` supports a second anchor: `to_cell` / `to_offset_x` / `to_offset_y`. When set, the display box spans from-anchor → to-anchor resolved against the target sheet's real column widths / row heights (Excel 64/19 px defaults otherwise), overriding `width_px`/`height_px`/`DisplayOptions`; the target grid is never reshaped in this mode. Works on both the generated and shell paths; the CLI's `--to` flag is back.
- `StructureSource.from_smiles` exposes ChemDraw style passthrough: `bond_length` / `line_width` / `bold_width` / `label_size` (CDX document properties, points). `add_smiles_ole_by_merge`'s style flags work again.

### CLI

- New `scripts.cli.smiles_id_structure_table`: build ID + structure tables from `.smiles` files (SMILES+ID per line), with `--id-header` / `--structure-header` / `--preview-format` / `--limit` options.
- `add_smiles_ole_by_merge` gains `--preview-format {png,emf,auto}` (was hardcoded to PNG). EMF is worth preferring for large batches: on a 50-molecule sample it was both smaller and faster than the 900×650 PNG default (0.3 MB vs 4.2 MB workbook, 2.9 s vs 10.9 s), since vector previews skip the supersampled raster path and Inkscape runs batched.
- The legacy `scripts/xlsx/`, `scripts/ole/`, `scripts/render/` trees were removed; all four CLIs now run on the `retain_cdx` public API. Ported CLI changes: `inject_cdx_into_xlsx` drops `--stream` and the `emu` unit (`--to` is kept, now backed by the public two-point `Placement`), and `--template` now uses the shell path (one `--cdx` per run); `add_smiles_ole_by_merge` requires an OLE-free input workbook. CLIs (not the library) additionally require `openpyxl` for plain-table preprocessing.

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
