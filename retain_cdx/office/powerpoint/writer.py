from __future__ import annotations

import tempfile
from collections.abc import Iterable
from pathlib import Path

from retain_cdx.ole.embed import write_cdx_to_ole
from retain_cdx.render.preview import preview_extension, render_preview
from retain_cdx.resources import bundled_template_pptx
from retain_cdx.structure import StructureSource

from retain_cdx.office.excel._ooxml import extract_xlsx, pack_xlsx
from retain_cdx.office.excel._utils import parse_emu

from ._pml import (
    FramePlacement,
    OleFrameTemplate,
    append_ole_frame,
    clear_ole_frames,
    slide_paths,
    slide_size,
)
from .options import PreviewOptions, SlideGrid, SlidePlacement

_DEFAULT_WIDTH_EMU = parse_emu("6cm")


class PowerPointCDXWriter:
    """Insert editable ChemDraw CDX OLE objects into a PPTX presentation.

    Positions are absolute EMU boxes — PresentationML has no cell grid, so none
    of the Excel-side column-width/anchor machinery applies here.
    """

    def __init__(self, input_pptx: str | Path | None = None):
        self._tmp = tempfile.TemporaryDirectory(prefix="retain_cdx_pptx_")
        self.package_dir = Path(self._tmp.name) / "package"
        self.package_dir.mkdir()

        if input_pptx is None:
            with bundled_template_pptx() as template:
                extract_xlsx(template, self.package_dir)
            # Start from an empty canvas: the bundled slide ships one sample
            # object whose only job is to be the frame/container template.
            for slide in slide_paths(self.package_dir):
                clear_ole_frames(slide)
        else:
            extract_xlsx(Path(input_pptx), self.package_dir)

        self._template = self._load_frame_template()

    def _load_frame_template(self) -> OleFrameTemplate:
        with tempfile.TemporaryDirectory(prefix="retain_cdx_pptx_tpl_") as tmp:
            template_dir = Path(tmp)
            with bundled_template_pptx() as template:
                extract_xlsx(template, template_dir)
            return OleFrameTemplate.from_package(template_dir)

    def close(self) -> None:
        self._tmp.cleanup()

    def __enter__(self) -> "PowerPointCDXWriter":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    @property
    def slide_count(self) -> int:
        return len(slide_paths(self.package_dir))

    @property
    def slide_size(self) -> tuple[int, int]:
        """(width, height) of the slides, in EMU."""
        return slide_size(self.package_dir)

    def save(self, output_pptx: str | Path) -> None:
        pack_xlsx(self.package_dir, Path(output_pptx))

    def insert_cdx(
        self,
        source: StructureSource | str | Path,
        *,
        slide: int = 1,
        placement: SlidePlacement | None = None,
        preview: PreviewOptions | None = None,
        name: str | None = None,
    ) -> None:
        """Insert one structure on *slide* (1-based).

        Without *placement* the structure is centred on the slide at a default
        6 cm width, keeping the molecule's aspect ratio.
        """
        self.insert_cdx_many(
            [source], slide=slide, placements=[placement], preview=preview, names=None
            if name is None
            else [name],
        )

    def insert_cdx_many(
        self,
        sources: Iterable[StructureSource | str | Path],
        *,
        slide: int = 1,
        grid: SlideGrid | None = None,
        placements: list[SlidePlacement | None] | None = None,
        preview: PreviewOptions | None = None,
        names: list[str] | None = None,
    ) -> None:
        """Insert several structures on one slide.

        Give *grid* to tile them, or *placements* for explicit boxes (one entry
        per structure; ``None`` entries fall back to centred defaults).
        """
        structures = [
            item if isinstance(item, StructureSource) else StructureSource.from_cdx(item)
            for item in sources
        ]
        if not structures:
            return

        slide_files = slide_paths(self.package_dir)
        if not 1 <= slide <= len(slide_files):
            raise ValueError(
                f"Slide {slide} out of range (presentation has {len(slide_files)})"
            )
        slide_path = slide_files[slide - 1]

        preview_options = preview or PreviewOptions(format="emf")
        extension = preview_extension(preview_options.format)

        with tempfile.TemporaryDirectory(prefix="retain_cdx_pptx_parts_") as tmp:
            work = Path(tmp)
            for index, structure in enumerate(structures):
                box = self._resolve_box(
                    structure,
                    index=index,
                    grid=grid,
                    placement=placements[index]
                    if placements and index < len(placements)
                    else None,
                )
                cdx_path = work / f"structure{index + 1}.cdx"
                structure.write_cdx(cdx_path)

                container_path = work / f"oleObject{index + 1}.bin"
                write_cdx_to_ole(
                    container_path,
                    cdx_path,
                    template_ole=self._template.container,
                    stream_name="CONTENTS",
                )

                image_path = work / f"image{index + 1}.{extension}"
                render_preview(
                    cdx_path,
                    image_path,
                    format=preview_options.format,
                    width=preview_options.width,
                    height=preview_options.height,
                    thick_lines=preview_options.thick_lines,
                    no_margin=preview_options.no_margin,
                    mono=preview_options.mono,
                    backend=preview_options.backend,
                    style=preview_options.style,
                    target_aspect=box.width / box.height if box.height else None,
                )

                append_ole_frame(
                    self.package_dir,
                    slide_path=slide_path,
                    template=self._template,
                    container=container_path.read_bytes(),
                    preview_bytes=image_path.read_bytes(),
                    preview_ext=extension,
                    placement=box,
                    name=(names[index] if names and index < len(names) else None)
                    or structure.name
                    or f"Structure {index + 1}",
                )

    def _resolve_box(
        self,
        structure: StructureSource,
        *,
        index: int,
        grid: SlideGrid | None,
        placement: SlidePlacement | None,
    ) -> FramePlacement:
        aspect = structure.content_aspect or 1.0
        if grid is not None:
            x, y, cell_w, cell_h = grid.box(index)
            width, height = _fit(cell_w, cell_h, aspect)
            # Centre the molecule inside its cell.
            return FramePlacement(
                x + (cell_w - width) // 2, y + (cell_h - height) // 2, width, height
            )

        if placement is None:
            width = _DEFAULT_WIDTH_EMU
            height = max(1, round(width / aspect))
            slide_w, slide_h = self.slide_size
            return FramePlacement(
                (slide_w - width) // 2, (slide_h - height) // 2, width, height
            )

        width, height = placement.width_emu, placement.height_emu
        if width is None and height is None:
            width = _DEFAULT_WIDTH_EMU
        if width is None:
            width = max(1, round(height * aspect))
        elif height is None:
            height = max(1, round(width / aspect))
        return FramePlacement(placement.x_emu, placement.y_emu, width, height)


def _fit(width: int, height: int, aspect: float) -> tuple[int, int]:
    """Largest box with *aspect* that fits inside width×height."""
    if width / height > aspect:
        return max(1, round(height * aspect)), height
    return width, max(1, round(width / aspect))


__all__ = ["PowerPointCDXWriter"]
