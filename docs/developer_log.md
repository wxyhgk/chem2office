2026-03-06

> Historical record — paths are the layout of the day. `scripts/ole/`,
> `scripts/xlsx/` and `scripts/render/` were retired in 0.0.9; that logic now
> lives in `retain_cdx/ole/`, `retain_cdx/office/excel/_ooxml/` and
> `retain_cdx/render/`. The CFBF work described here was revisited in 0.0.10 —
> see `excel_ooxml_research.md` (C1/C2) for what was still wrong.

Issue
- Excel could open generated `.xlsx` files but immediately reported repair errors:
  - removed record from `/xl/worksheets/sheet1.xml`
  - embedded ChemDraw object was removed
  - only the preview image remained visible

Observed Cases
- Small CDX payloads could work after aligning default generation with the bundled `template.xlsx`.
- Large CDX payloads, such as `sdfData/GDG906.cdx` (`7589` bytes), still failed.
- `obabel` on this machine can read `cdx` but cannot write it.
- `pycdxml` + `rdkit` was able to generate a valid binary CDX from `GDG906.sdf`.

False Leads
- The issue was not caused by `CDXML` vs `CDX` confusion once a real binary `CDX` was produced.
- The issue was not primarily caused by sheet XML anchors or image generation.
- The issue was not just namespace formatting in worksheet XML. That had caused an earlier failure in another path, but it was not the final blocker for large CDX payloads.

Actual Root Cause
- The bundled Excel template stores the ChemDraw OLE `CONTENTS` stream as a mini-stream entry with limited size.
- The original template-mode injector only overwrote the existing `CONTENTS` bytes and patched the directory entry size.
- That approach works only when the new CDX is not larger than the original template stream.
- For larger CDX payloads, the project fell back to rebuilding the OLE container.
- The rebuilt OLE container had an internal CFBF/FAT constant bug:
  - `_FREE`, `_ENDOFCHAIN`, `_FATSECT` special sector IDs were wrong
  - `olefile` tolerated the result, but Excel did not
  - Excel therefore treated the OLE object as corrupt and removed the corresponding object record from `sheet1.xml`

Why The File Looked Like "Only A Picture"
- The worksheet still contained both:
  - a VML/drawing preview image
  - an `oleObject1.bin` relationship
- During Excel recovery, the invalid OLE object record was removed.
- The preview image relationship survived, so the final workbook looked like an image-only embed.

Fix Applied
- `scripts/ole/ole_builder.py`
  - fixed the CFBF special sector constants
  - kept large `CONTENTS` streams in regular FAT sectors instead of trying to squeeze them into mini-stream assumptions
- `scripts/cli/inject_cdx_into_xlsx.py`
  - retained the stable bundled-template XML/VML path
  - when template `CONTENTS` capacity is too small, swap only the OLE binary to a freshly rebuilt container
  - avoid falling back to the old template-free worksheet XML path for this case
- `scripts/xlsx/relationships.py`
  - preserved template relationship numbering (`rId3`/`rId4` for first OLE/image pair)
  - wrote `.rels` without `ns0:` prefixes
- `scripts/xlsx/content_types.py`
  - wrote `[Content_Types].xml` without `ns0:` prefixes

Result
- `sdfData/GDG906.sdf` -> `pycdxml`/`rdkit` -> `sdfData/GDG906.cdx` succeeded
- `sdfData/GDG906.cdx` -> `sdfData/GDG906.xlsx` now opens without Excel repair
- Embedded ChemDraw object is preserved instead of being downgraded to a preview image

Important Notes
- `pycdxml` is usable here as a practical `SDF -> CDX` toolchain.
- Current Open Babel installation cannot write `CDX`; it only reports `cdx` as read-only.
- The bundled template remains important because Excel is very sensitive to the worksheet/drawing/VML/OLE relationship structure.
- For large payloads, the safe approach is:
  - keep template XML structure
  - replace only the OLE binary with a valid rebuilt compound file
