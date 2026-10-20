"""Package-level validity checks shared by the Excel, PowerPoint and Word tests.

These are the checks that caught real defects by hand: a dangling relationship
target, a part with no declared content type, duplicate shape ids, and an
``mc:*`` ``Requires`` naming a prefix nobody declared (which is what made
PowerPoint offer to repair generated decks). Keeping them as a reusable oracle
means a regression shows up as a red test instead of a repair dialog.
"""

from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def check_package(path: Path | str) -> list[str]:
    """Return a list of problems; empty means the package looks well-formed."""
    problems: list[str] = []
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        problems += _check_relationships(archive, names)
        problems += _check_content_types(archive, names)
        problems += _check_markup_compatibility(archive, names)
        problems += _check_shape_ids(archive, names)
        problems += _check_word_object_ids(archive, names)
    return problems


def _check_relationships(archive: zipfile.ZipFile, names: set[str]) -> list[str]:
    problems: list[str] = []
    for name in sorted(n for n in names if n.endswith(".rels")):
        base = posixpath.dirname(posixpath.dirname(name))
        root = ET.fromstring(archive.read(name))
        seen: list[str] = []
        for rel in root.findall(f"{{{NS_RELS}}}Relationship"):
            seen.append(rel.get("Id"))
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target") or ""
            resolved = (
                target[1:]
                if target.startswith("/")
                else posixpath.normpath(posixpath.join(base, target))
            )
            if resolved not in names:
                problems.append(f"{name}: dangling target {target!r}")
        duplicates = {rid for rid in seen if seen.count(rid) > 1}
        if duplicates:
            problems.append(f"{name}: duplicate rIds {sorted(duplicates)}")
    return problems


def _check_content_types(archive: zipfile.ZipFile, names: set[str]) -> list[str]:
    if "[Content_Types].xml" not in names:
        return ["[Content_Types].xml missing"]
    root = ET.fromstring(archive.read("[Content_Types].xml"))
    defaults = {
        (el.get("Extension") or "").lower() for el in root.findall(f"{{{NS_CT}}}Default")
    }
    overrides = {el.get("PartName") for el in root.findall(f"{{{NS_CT}}}Override")}
    problems = []
    for name in sorted(names):
        if name == "[Content_Types].xml" or name.endswith(".rels"):
            continue
        extension = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if extension not in defaults and f"/{name}" not in overrides:
            problems.append(f"no content type declared for {name}")
    return problems


def _check_markup_compatibility(archive: zipfile.ZipFile, names: set[str]) -> list[str]:
    """Every prefix named by mc:Choice/@Requires (or mc:Ignorable) must resolve.

    ElementTree drops namespace declarations no element actually uses, so a
    round-trip silently strips the ``xmlns:v`` that ``Requires="v"`` depends on.
    """
    problems = []
    for name in sorted(n for n in names if n.endswith((".xml", ".vml"))):
        text = archive.read(name).decode("utf-8", "replace")
        declared = set(re.findall(r'xmlns:([\w.-]+)\s*=', text))
        for attribute in ("Requires", "Ignorable"):
            for value in re.findall(rf'{attribute}="([^"]+)"', text):
                for prefix in value.split():
                    if prefix not in declared:
                        problems.append(
                            f'{name}: {attribute}="{prefix}" but xmlns:{prefix} '
                            "is not declared"
                        )
    return problems


def _check_shape_ids(archive: zipfile.ZipFile, names: set[str]) -> list[str]:
    """Shape ids must be unique within a slide / drawing part."""
    problems = []
    patterns = (
        re.compile(r"ppt/slides/slide\d+\.xml$"),
        re.compile(r"xl/drawings/.*\.xml$"),
    )
    for name in sorted(names):
        if not any(pattern.match(name) for pattern in patterns):
            continue
        ids = re.findall(r"<(?:p|xdr):cNvPr[^>]*\bid=\"(\d+)\"", archive.read(name).decode())
        duplicates = {i for i in ids if ids.count(i) > 1}
        if duplicates:
            problems.append(f"{name}: duplicate cNvPr ids {sorted(duplicates)}")
    return problems


def _check_word_object_ids(archive: zipfile.ZipFile, names: set[str]) -> list[str]:
    """VML shape ids, OLE ObjectIDs and shapetype definitions must not repeat.

    Word identifies an inline OLE object by both its ``v:shape`` id and its
    ``ObjectID``; a cloned object that keeps either one collides with the
    original. Shapetypes are defined once and referenced afterwards.
    """
    problems = []
    for name in sorted(n for n in names if re.match(r"word/[^/]+\.xml$", n)):
        text = archive.read(name).decode("utf-8", "replace")
        for label, pattern in (
            ("v:shape ids", r'<v:shape\b[^>]*\bid="([^"]+)"'),
            ("ObjectIDs", r'\bObjectID="([^"]+)"'),
            ("v:shapetype ids", r'<v:shapetype\b[^>]*\bid="([^"]+)"'),
        ):
            values = re.findall(pattern, text)
            duplicates = {v for v in values if values.count(v) > 1}
            if duplicates:
                problems.append(f"{name}: duplicate {label} {sorted(duplicates)}")
    return problems
