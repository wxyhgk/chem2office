from __future__ import annotations

import io
import zipfile
from pathlib import Path

import olefile
import pytest

from retain_cdx.ole.chemdraw import build_chemdraw_ole
from retain_cdx.ole.embed import write_cdx_to_ole

CDX = Path("tests/fixtures/test.cdx")


def _payload(size: int) -> bytes:
    block = bytes(range(256))
    body = block * (size // len(block) + 1)
    return (b"VjCD0100" + body)[:size]


def _streams(blob: bytes) -> dict[str, bytes]:
    ole = olefile.OleFileIO(io.BytesIO(blob))
    return {entry[0]: ole.openstream(entry).read() for entry in ole.listdir()}


@pytest.mark.parametrize(
    "size",
    [
        1_000,
        4_095,
        4_096,  # mini-stream cutoff
        63_488,  # last payload a single FAT sector can map
        63_489,  # first payload that needs a second one
        200_000,
        1_000_000,
    ],
)
def test_container_roundtrips_at_every_size(size):
    """Regression: one hardcoded FAT sector silently corrupted CDX > ~62 KB.

    olefile could still open such a file; only reading CONTENTS revealed it —
    the same shape as the historical "Excel deleted the object" reports.
    """
    payload = _payload(size)
    streams = _streams(build_chemdraw_ole(payload))
    assert streams["CONTENTS"] == payload
    assert set(streams) == {"\x01Ole", "\x01CompObj", "CONTENTS"}


def test_extra_streams_are_carried_verbatim():
    presentation = bytes(range(256)) * 24
    streams = _streams(
        build_chemdraw_ole(_payload(1_400), extra_streams={"\x02OlePres000": presentation})
    )
    assert streams["\x02OlePres000"] == presentation
    assert streams["CONTENTS"] == _payload(1_400)


def test_minifat_free_entries_use_freesect():
    """Windows-authored containers pad with FREESECT; 0 is a valid sector id."""
    blob = build_chemdraw_ole(_payload(1_400))
    ole = olefile.OleFileIO(io.BytesIO(blob))
    minifat_sector = ole.minifatsect
    start = (minifat_sector + 1) * 512
    tail = blob[start + 256 : start + 512]
    assert tail == b"\xff\xff\xff\xff" * (len(tail) // 4)


def test_oversized_payload_raises_instead_of_corrupting():
    with pytest.raises(ValueError, match="too large"):
        build_chemdraw_ole(_payload(109 * 128 * 512 + 1))


def test_streams_match_the_excel_template_byte_for_byte():
    """\\x01Ole and \\x01CompObj must be what Windows Excel+ChemDraw writes."""
    with zipfile.ZipFile("retain_cdx/data/template.xlsx") as archive:
        reference = _streams(archive.read("xl/embeddings/oleObject1.bin"))
    ours = _streams(build_chemdraw_ole(_payload(1_400)))
    assert ours["\x01Ole"] == reference["\x01Ole"]
    assert ours["\x01CompObj"] == reference["\x01CompObj"]


def test_in_place_patch_keeps_template_layout(tmp_path):
    with zipfile.ZipFile("retain_cdx/data/template.xlsx") as archive:
        template = archive.read("xl/embeddings/oleObject1.bin")
    target = tmp_path / "ole.bin"
    target.write_bytes(template)

    write_cdx_to_ole(target, CDX)
    assert target.stat().st_size == len(template), "should patch, not rebuild"
    assert _streams(target.read_bytes())["CONTENTS"] == CDX.read_bytes()


def test_rebuild_when_payload_outgrows_the_template(tmp_path):
    target = tmp_path / "ole.bin"
    target.write_bytes(build_chemdraw_ole(_payload(200)))
    big = tmp_path / "big.cdx"
    big.write_bytes(_payload(50_000))

    write_cdx_to_ole(target, big)
    assert _streams(target.read_bytes())["CONTENTS"] == _payload(50_000)


def test_rebuild_when_crossing_the_mini_stream_cutoff(tmp_path):
    """A shrink across the cutoff moves the stream between allocators."""
    target = tmp_path / "ole.bin"
    target.write_bytes(build_chemdraw_ole(_payload(8_000)))
    small = tmp_path / "small.cdx"
    small.write_bytes(_payload(1_000))

    write_cdx_to_ole(target, small)
    assert _streams(target.read_bytes())["CONTENTS"] == _payload(1_000)


def test_stale_presentation_stream_is_dropped(tmp_path):
    """A cached rendering of the previous molecule must not survive."""
    target = tmp_path / "ole.bin"
    target.write_bytes(
        build_chemdraw_ole(_payload(1_400), extra_streams={"\x02OlePres000": b"\xff" * 6_270})
    )
    write_cdx_to_ole(target, CDX)

    streams = _streams(target.read_bytes())
    assert "\x02OlePres000" not in streams
    assert streams["CONTENTS"] == CDX.read_bytes()


def test_multi_sector_directory_is_patchable(tmp_path):
    """PowerPoint containers push the directory past one sector."""
    target = tmp_path / "ole.bin"
    extras = {f"\x02OlePres{index:03d}": b"\x00" * 100 for index in range(4)}
    target.write_bytes(build_chemdraw_ole(_payload(1_400), extra_streams=extras))
    ole = olefile.OleFileIO(str(target))
    assert len(ole.direntries) > 4, "fixture should span more than one sector"
    ole.close()

    write_cdx_to_ole(target, CDX)
    assert _streams(target.read_bytes())["CONTENTS"] == CDX.read_bytes()


def test_missing_stream_raises_a_clear_error(tmp_path):
    target = tmp_path / "ole.bin"
    target.write_bytes(build_chemdraw_ole(_payload(1_400)))
    with pytest.raises(ValueError, match="OLE stream not found"):
        write_cdx_to_ole(target, CDX, stream_name="NOPE")
