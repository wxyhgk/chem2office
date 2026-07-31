from __future__ import annotations

from pathlib import Path

from retain_cdx.structure import StructureSource

from retain_cdx.office.excel._utils import px_to_points

from .plan import InsertPlan


def write_cdx_sources(
    sources: list[StructureSource],
    output_dir: Path,
    plan: InsertPlan,
) -> list[Path]:
    cdx_paths: list[Path] = []
    layout = plan.layout
    normalize = plan.normalize
    for idx, source in enumerate(sources, start=1):
        if normalize.enabled:
            source = source.normalized(
                margin=normalize.margin,
                target_aspect=layout.width_px / layout.height_px,
                target_width=px_to_points(layout.width_px),
                target_height=px_to_points(layout.height_px),
                fill_ratio=normalize.fill_ratio,
            )
        path = output_dir / f"source_{idx}.cdx"
        source.write_cdx(path)
        cdx_paths.append(path)
    return cdx_paths
