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
    def from_cdxml(
        cls, cdxml: str | bytes | Path, *, name: str | None = None
    ) -> "StructureSource":
        """Build from a CDXML file path, or the CDXML text itself.

        The drawing is converted to CDX as-is (via pycdxml): coordinates,
        orientation and labels are kept, nothing is re-laid out.
        """
        from .converters import cdxml_to_cdx_bytes

        if isinstance(cdxml, bytes):
            text = cdxml.decode("utf-8")
        elif isinstance(cdxml, str) and cdxml.lstrip().startswith("<"):
            text = cdxml
        else:
            path = Path(cdxml)
            text = path.read_text(encoding="utf-8")
            name = name or path.stem
        return cls(cdx_data=cdxml_to_cdx_bytes(text), name=name or "structure")

    @classmethod
    def from_smiles(
        cls,
        smiles: str,
        *,
        name: str = "structure",
        bond_length: float = 24,
        line_width: float = 1.2,
        bold_width: float = 4,
        label_size: float = 16,
    ) -> "StructureSource":
        """Build from SMILES; the ChemDraw style knobs map to CDX document
        properties (BondLength/LineWidth/BoldWidth/LabelSize, in points)."""
        from .converters import smiles_to_cdx_bytes

        return cls(
            cdx_data=smiles_to_cdx_bytes(
                smiles,
                bond_length=bond_length,
                line_width=line_width,
                bold_width=bold_width,
                label_size=label_size,
            ),
            name=name,
        )

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
