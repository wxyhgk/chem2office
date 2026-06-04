from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import tempfile

from .normalize import normalize_cdx_viewbox
from .geometry import cdx_content_aspect, cdx_page_aspect


@dataclass(frozen=True)
class StructureSource:
    """Canonical editable structure payload for Office embedding."""

    cdx_data: bytes
    name: str = "structure"

    @classmethod
    def from_cdx(cls, path: str | Path, *, name: str | None = None) -> "StructureSource":
        cdx_path = Path(path)
        data = cdx_path.read_bytes()
        if not data:
            raise ValueError(f"CDX is empty: {cdx_path}")
        return cls(cdx_data=data, name=name or cdx_path.stem)

    @classmethod
    def from_smiles(cls, smiles: str, *, name: str = "structure") -> "StructureSource":
        from .converters import smiles_to_cdx_bytes

        return cls(cdx_data=smiles_to_cdx_bytes(smiles), name=name)

    def write_cdx(self, path: str | Path) -> Path:
        output = Path(path)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_bytes(self.cdx_data)
        return output

    def normalized(
        self,
        *,
        margin: float = 18,
        target_aspect: float | None = None,
        target_width: float | None = None,
        target_height: float | None = None,
        fill_ratio: float = 0.8,
    ) -> "StructureSource":
        return StructureSource(
            cdx_data=normalize_cdx_viewbox(
                self.cdx_data,
                margin=margin,
                target_aspect=target_aspect,
                target_width=target_width,
                target_height=target_height,
                fill_ratio=fill_ratio,
            ),
            name=self.name,
        )

    def as_temp_cdx(self) -> tempfile.NamedTemporaryFile:
        tmp = tempfile.NamedTemporaryFile(suffix=".cdx", delete=False)
        tmp.write(self.cdx_data)
        tmp.flush()
        return tmp

    @property
    def page_aspect(self) -> float | None:
        return cdx_page_aspect(self.cdx_data)

    @property
    def content_aspect(self) -> float | None:
        return cdx_content_aspect(self.cdx_data)
