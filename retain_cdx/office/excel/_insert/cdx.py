from __future__ import annotations

from pathlib import Path

from retain_cdx.structure import StructureSource

from .plan import InsertPlan


def write_cdx_sources(
    sources: list[StructureSource],
    output_dir: Path,
    plan: InsertPlan,
) -> list[Path]:
    cdx_paths: list[Path] = []
    for idx, source in enumerate(sources, start=1):
        if plan.normalize_cdx:
            source = source.normalized(
                margin=plan.normalize_margin,
                target_aspect=plan.width_px / plan.height_px,
                target_width=plan.width_px * 0.75,
                target_height=plan.height_px * 0.75,
                fill_ratio=plan.normalize_fill_ratio,
            )
        path = output_dir / f"source_{idx}.cdx"
        source.write_cdx(path)
        cdx_paths.append(path)
    return cdx_paths
