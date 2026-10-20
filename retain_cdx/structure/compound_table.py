"""Structura「化合物表」包：manifest.json + 每个化合物一个 CDXML。

交接格式 v1（zip，或解压后的目录）::

    manifest.json        程序读这个
    compounds.csv        给人看，不读
    cdxml/<id>.cdxml     每个化合物一个
    table.cdxml          整张表（给直接用 ChemDraw 的人），不读

manifest::

    {"format": "structura-compound-table", "version": 1, "columns": 3,
     "compounds": [{"id": "C-1", "file": "cdxml/C-1.cdxml", "smiles": "...",
                    "formula": "C13H12O", "mw": 184.24, "fields": {"R1": "OMe"}}]}

``compounds`` 的顺序就是表格顺序。文件名一律取 ``file`` 字段，不从 ``id`` 拼
（导出时 ``id`` 里不适合当文件名的字符会被换成 ``_``）。``smiles`` 可能没有。

选 CDXML 而不是 SMILES，是因为它原样保留画好的结构：走 SMILES 时 RDKit 会
重新排版，同一通式的核心朝向就乱了。
"""

from __future__ import annotations

import json
import posixpath
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any

from .source import StructureSource

FORMAT = "structura-compound-table"
SUPPORTED_VERSIONS = (1,)
MANIFEST = "manifest.json"


@dataclass(frozen=True)
class Compound:
    """One row of the manifest, with its structure loaded."""

    id: str
    structure: StructureSource
    smiles: str | None = None
    formula: str | None = None
    mw: float | None = None
    fields: dict[str, str] = field(default_factory=dict)


@dataclass(frozen=True)
class CompoundTable:
    """A Structura compound-table package, read into memory."""

    compounds: list[Compound]
    columns: int = 3
    version: int = 1

    @classmethod
    def load(cls, path: str | Path) -> "CompoundTable":
        """Read a package from its ``.zip`` or from the unpacked directory.

        A zip may hold the files at its root or inside one top-level folder
        (what zipping a directory produces); both are accepted.
        """
        path = Path(path)
        if path.is_dir():
            return cls._from_reader(_DirectoryReader(path), origin=str(path))
        if zipfile.is_zipfile(path):
            with zipfile.ZipFile(path) as archive:
                return cls._from_reader(_ZipReader(archive), origin=str(path))
        raise ValueError(f"Not a compound-table package (zip or directory): {path}")

    @classmethod
    def _from_reader(cls, reader: "_Reader", *, origin: str) -> "CompoundTable":
        manifest = json.loads(reader.read(MANIFEST).decode("utf-8-sig"))
        if manifest.get("format") != FORMAT:
            raise ValueError(
                f"{origin}: manifest format is {manifest.get('format')!r}, "
                f"expected {FORMAT!r}"
            )
        version = manifest.get("version")
        if version not in SUPPORTED_VERSIONS:
            raise ValueError(
                f"{origin}: unsupported {FORMAT} version {version!r} "
                f"(supported: {', '.join(map(str, SUPPORTED_VERSIONS))})"
            )
        columns = manifest.get("columns", 3)
        if not isinstance(columns, int) or columns < 1:
            raise ValueError(f"{origin}: columns must be a positive integer, got {columns!r}")

        entries = manifest.get("compounds")
        if not isinstance(entries, list) or not entries:
            raise ValueError(f"{origin}: manifest lists no compounds")

        compounds = []
        seen: set[str] = set()
        for index, entry in enumerate(entries):
            compound_id = str(entry.get("id") or "").strip()
            file_name = entry.get("file")
            if not compound_id or not file_name:
                raise ValueError(f"{origin}: compound #{index + 1} needs both id and file")
            if compound_id in seen:
                raise ValueError(f"{origin}: duplicate compound id {compound_id!r}")
            seen.add(compound_id)
            try:
                text = reader.read(_safe_member(file_name)).decode("utf-8-sig")
            except KeyError:
                raise ValueError(
                    f"{origin}: {compound_id}: file {file_name!r} is not in the package"
                ) from None
            compounds.append(
                Compound(
                    id=compound_id,
                    structure=StructureSource.from_cdxml(text, name=compound_id),
                    smiles=entry.get("smiles") or None,
                    formula=entry.get("formula") or None,
                    mw=_optional_float(entry.get("mw")),
                    fields={str(k): str(v) for k, v in (entry.get("fields") or {}).items()},
                )
            )
        return cls(compounds=compounds, columns=columns, version=version)


def _optional_float(value: Any) -> float | None:
    try:
        return None if value is None else float(value)
    except (TypeError, ValueError):
        return None


def _safe_member(name: str) -> str:
    """A package-relative path that cannot step outside the package."""
    normalized = posixpath.normpath(str(name).replace("\\", "/"))
    if normalized.startswith(("/", "../")) or normalized in ("..", "."):
        raise ValueError(f"Unsafe file path in manifest: {name!r}")
    return normalized


class _Reader:
    def read(self, member: str) -> bytes:  # pragma: no cover - interface
        raise NotImplementedError


class _DirectoryReader(_Reader):
    def __init__(self, root: Path):
        self.root = root

    def read(self, member: str) -> bytes:
        path = self.root / member
        if not path.is_file():
            raise KeyError(member)
        return path.read_bytes()


class _ZipReader(_Reader):
    def __init__(self, archive: zipfile.ZipFile):
        self.archive = archive
        names = [n for n in archive.namelist() if PurePosixPath(n).name == MANIFEST]
        if not names:
            raise ValueError(f"No {MANIFEST} in {archive.filename}")
        # The shallowest manifest wins: the package root, or its one folder.
        manifest = min(names, key=lambda n: n.count("/"))
        self.prefix = manifest[: -len(MANIFEST)]

    def read(self, member: str) -> bytes:
        return self.archive.read(self.prefix + member)


__all__ = ["Compound", "CompoundTable", "FORMAT"]
