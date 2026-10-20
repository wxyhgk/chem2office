from __future__ import annotations

from pathlib import Path
import tempfile


def smiles_to_cdx_bytes(
    smiles: str,
    *,
    bond_length: float = 24,
    line_width: float = 1.2,
    bold_width: float = 4,
    label_size: float = 16,
) -> bytes:
    from rdkit import Chem
    from rdkit.Chem import AllChem
    from pycdxml.cdxml_converter.rdkit_chemdraw import mol_to_document
    from pycdxml.cdxml_converter.chemdraw_io import write_cdx_file

    mol = Chem.MolFromSmiles(smiles)
    if mol is None:
        raise ValueError(f"Invalid SMILES: {smiles}")

    AllChem.Compute2DCoords(mol)
    doc = mol_to_document(
        mol,
        chemdraw_style={
            "BondLength": bond_length,
            "LineWidth": line_width,
            "BoldWidth": bold_width,
            "LabelSize": label_size,
            "CaptionSize": label_size,
        },
    )

    with tempfile.TemporaryDirectory(prefix="retain_cdx_smiles_") as tmp:
        cdx_path = Path(tmp) / "structure.cdx"
        write_cdx_file(doc, cdx_path)
        return cdx_path.read_bytes()


def cdxml_to_cdx_bytes(cdxml: str) -> bytes:
    """CDXML text → binary CDX, keeping the drawing exactly as authored.

    Unlike the SMILES path nothing is re-laid out: coordinates, labels and
    styles carry over, which is the point of handing structures over as CDXML.
    """
    from pycdxml.cdxml_converter.chemdraw_objects import ChemDrawDocument

    return ChemDrawDocument.from_cdxml(cdxml).to_bytes()
