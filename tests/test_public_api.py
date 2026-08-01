"""Public API surface: names carry their format, old ones still resolve."""

from __future__ import annotations

import pytest

import retain_cdx


def test_exported_names_are_format_explicit():
    """Every format-specific symbol says which format it belongs to.

    The Excel names predate PowerPoint support and carried no qualifier, which
    left `Placement` and `DisplayOptions` looking format-neutral.
    """
    shared = {"StructureSource", "PreviewOptions"}
    for name in set(retain_cdx.__all__) - {"__version__"} - shared:
        assert name.startswith(("Excel", "PowerPoint")), (
            f"{name} does not say which format it belongs to"
        )


def test_shared_names_carry_no_format_prefix():
    for name in ("StructureSource", "PreviewOptions"):
        assert not name.startswith(("Excel", "PowerPoint"))
        assert hasattr(retain_cdx, name)


@pytest.mark.parametrize(
    "old,new",
    [
        ("Placement", "ExcelPlacement"),
        ("DisplayOptions", "ExcelDisplayOptions"),
        ("ShellTemplate", "ExcelShellTemplate"),
        ("SlidePlacement", "PowerPointPlacement"),
        ("SlideGrid", "PowerPointGrid"),
    ],
)
def test_old_names_still_resolve_with_a_warning(old, new):
    with pytest.warns(DeprecationWarning, match=f"renamed to {new}"):
        assert getattr(retain_cdx, old) is getattr(retain_cdx, new)


def test_unknown_attribute_still_raises():
    with pytest.raises(AttributeError, match="NoSuchThing"):
        retain_cdx.NoSuchThing


def test_renamed_classes_are_not_in_all():
    """Old names resolve but must not be advertised."""
    for old in ("Placement", "DisplayOptions", "ShellTemplate", "SlidePlacement"):
        assert old not in retain_cdx.__all__
