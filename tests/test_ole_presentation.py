"""OLE 演示流：让对象能被复制出去而不退化成图片。

没有 `\\x02OlePres000` 时容器仍是合法嵌入对象、双击照常进 ChemDraw，但从
PowerPoint 复制到别处只会得到一张图片——目标程序拿不到缓存渲染。
Excel 侧的容器不带这个流也正常，所以只在 PPT 路径挂上。
"""

from __future__ import annotations

import io
import struct
import zipfile
from pathlib import Path

import olefile
import pytest


from retain_cdx import (
    PowerPointCDXWriter,
    PowerPointPlacement,
    PreviewOptions,
)
from retain_cdx.ole.presentation import (
    MAX_EMBEDDED_EMF,
    PRESENTATION_STREAM,
    build_enhmetafile_presentation,
    emu_to_himetric,
    wrap_emf_in_wmf,
)

CDX = "tests/fixtures/test.cdx"
BLANK_PPTX = "tests/fixtures/blank.pptx"
_MINIMAL_EMF = bytes.fromhex("01000000") + b"\x00" * 36 + b" EMF" + b"\x00" * 24
# 最小 WMF：18 字节头 + META_EOF
_MINIMAL_WMF = struct.pack("<HHHIHIH", 1, 9, 0x0300, 12, 0, 3, 0) + struct.pack("<IH", 3, 0)


def test_header_matches_the_reference_layout():
    """字段布局照 Windows PowerPoint + ChemDraw 的真实产物。"""
    stream = build_enhmetafile_presentation(
        _MINIMAL_EMF, _MINIMAL_WMF, width_himetric=1016, height_himetric=1164
    )
    marker, fmt, target_device = struct.unpack_from("<III", stream, 0)
    assert marker == 0xFFFFFFFF
    assert fmt == 14, "必须声明 CF_ENHMETAFILE"
    assert target_device == 4, "无目标设备时该字段为 4"

    aspect, lindex, advf, reserved = struct.unpack_from("<iiii", stream, 12)
    assert (aspect, lindex, reserved) == (1, -1, 0)

    width, height, size = struct.unpack_from("<iii", stream, 28)
    assert (width, height) == (1016, 1164)
    assert stream[40 : 40 + size] == wrap_emf_in_wmf(_MINIMAL_EMF, _MINIMAL_WMF)
    assert stream[40 + size :] == struct.pack("<II", 0x494E414E, 0), "缺 TOC 尾"


def test_payload_must_be_an_emf():
    with pytest.raises(ValueError, match="EMR_HEADER"):
        build_enhmetafile_presentation(
            b"not an emf", _MINIMAL_WMF, width_himetric=10, height_himetric=10
        )


def test_emf_rides_in_a_single_comment_record():
    """整份 EMF 必须放在一条 WMFC 注释记录里。

    实测：切成两条时 PowerPoint 不崩，但取不出 EMF——复制出去只剩图片。
    """
    emf = _MINIMAL_EMF * 200
    wmf = wrap_emf_in_wmf(emf, _MINIMAL_WMF)
    assert wmf.count(b"WMFC") == 1

    size_words = struct.unpack_from("<I", wmf, 18)[0]
    function, escape, byte_count = struct.unpack_from("<HHH", wmf, 22)
    assert function == 0x0626, "首条记录应是 META_ESCAPE"
    assert escape == 0x000F, "转义类型应是 MFCOMMENT"
    assert byte_count == 34 + len(emf), "ByteCount 应覆盖注释头加整份 EMF"
    assert wmf[28 + 34 : 28 + 34 + len(emf)] == emf

    count, current, remaining, total = struct.unpack_from("<IIII", wmf, 46)
    assert (count, current, remaining, total) == (1, len(emf), 0, len(emf))

    # 记录链要走得通，且后备绘图记录还在
    offset, records = 18, 0
    while offset + 6 <= len(wmf):
        rsize, func = struct.unpack_from("<IH", wmf, offset)
        assert rsize > 0 and offset + rsize * 2 <= len(wmf), f"记录链断在 {offset}"
        records += 1
        offset += rsize * 2
        if func == 0:
            break
    assert offset == len(wmf), "记录链未走到文件末尾"
    assert records >= 2, "后备的 WMF 绘图记录被丢掉了"
    assert struct.unpack_from("<I", wmf, 6)[0] * 2 == len(wmf), "mtSize 与实际长度不符"


def test_checksum_is_not_zero():
    """校验和确实被检验：实测填 0 会让复制粘贴退化成图片。"""
    wmf = wrap_emf_in_wmf(_MINIMAL_EMF, _MINIMAL_WMF)
    checksum = struct.unpack_from("<H", wmf, 40)[0]
    assert checksum != 0

    fields = bytearray(wmf[28:62])
    fields[12:14] = b"\x00\x00"
    assert checksum == -sum(struct.unpack("<17H", bytes(fields))) & 0xFFFF


def test_oversized_emf_is_refused_rather_than_split():
    """ByteCount 是 16 位，装不下就报错——分块已实测不可行。"""
    with pytest.raises(ValueError, match="single WMFC comment record"):
        wrap_emf_in_wmf(_MINIMAL_EMF + b"\x00" * MAX_EMBEDDED_EMF, _MINIMAL_WMF)


def test_placeable_header_is_stripped():
    """Inkscape 直出的 WMF 带 22 字节 placeable 头，OLE 里不能有。"""
    placeable = bytes.fromhex("d7cdc69a") + b"\x00" * 18
    assert wrap_emf_in_wmf(_MINIMAL_EMF, placeable + _MINIMAL_WMF) == wrap_emf_in_wmf(
        _MINIMAL_EMF, _MINIMAL_WMF
    )


@pytest.mark.parametrize(
    "emu,expected",
    [(360, 1), (36000, 100), (2160000, 6000), (1, 1)],  # 1 HIMETRIC = 0.01mm = 360 EMU
)
def test_himetric_conversion(emu, expected):
    assert emu_to_himetric(emu) == expected


def _container(archive: zipfile.ZipFile) -> olefile.OleFileIO:
    name = [n for n in archive.namelist() if "/embeddings/" in n][0]
    return olefile.OleFileIO(io.BytesIO(archive.read(name)))


def test_emf_preview_attaches_a_presentation(tmp_path):
    output = tmp_path / "with_cache.pptx"
    with PowerPointCDXWriter(BLANK_PPTX) as writer:
        writer.insert_cdx(
            CDX,
            placement=PowerPointPlacement(x="1cm", y="1cm", width="4cm"),
            preview=PreviewOptions(format="emf"),
        )
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        ole = _container(archive)
        names = {entry[0] for entry in ole.listdir()}
        assert PRESENTATION_STREAM in names, "EMF 预览时应挂上缓存渲染"
        assert names == {"\x01Ole", "\x01CompObj", "CONTENTS", PRESENTATION_STREAM}

        stream = ole.openstream(PRESENTATION_STREAM).read()
        assert struct.unpack_from("<I", stream, 4)[0] == 14
        size = struct.unpack_from("<i", stream, 36)[0]
        assert len(stream) == 40 + size + 8, "声明长度与实际数据不符（含 TOC 尾）"
        assert stream[40:42] == b"\x01\x00", "载荷不是 WMF"
        assert stream.count(b"WMFC") == 1, "EMF 应整份放在一条注释记录里"


def test_presentation_size_tracks_the_display_box(tmp_path):
    """缓存尺寸应与显示框一致（HIMETRIC）。"""
    output = tmp_path / "sized.pptx"
    with PowerPointCDXWriter(BLANK_PPTX) as writer:
        writer.insert_cdx(
            CDX,
            placement=PowerPointPlacement(x="1cm", y="1cm", width="6cm", height="4cm"),
            preview=PreviewOptions(format="emf"),
        )
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        stream = _container(archive).openstream(PRESENTATION_STREAM).read()
    width, height = struct.unpack_from("<ii", stream, 28)
    assert width == emu_to_himetric(6 * 360000)  # 6cm
    assert height == emu_to_himetric(4 * 360000)


def test_png_preview_skips_the_presentation(tmp_path):
    """缓存要的是矢量图；PNG 预览时不挂。"""
    output = tmp_path / "png.pptx"
    with PowerPointCDXWriter(BLANK_PPTX) as writer:
        writer.insert_cdx(
            CDX,
            placement=PowerPointPlacement(x="1cm", y="1cm", width="4cm"),
            preview=PreviewOptions(format="png"),
        )
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        names = {entry[0] for entry in _container(archive).listdir()}
    assert PRESENTATION_STREAM not in names
    assert "CONTENTS" in names


def test_cache_matches_the_structure_it_ships_with(tmp_path):
    """缓存与 CONTENTS 必须同批产生，不能是上一个分子的渲染。"""
    output = tmp_path / "two.pptx"
    with PowerPointCDXWriter(BLANK_PPTX) as writer:
        writer.insert_cdx_many(
            ["tests/fixtures/1.cdx", "tests/fixtures/2.cdx"],
            preview=PreviewOptions(format="emf"),
        )
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        names = sorted(n for n in archive.namelist() if "/embeddings/" in n)
        assert len(names) == 2
        payloads = []
        for name in names:
            ole = olefile.OleFileIO(io.BytesIO(archive.read(name)))
            assert PRESENTATION_STREAM in {e[0] for e in ole.listdir()}
            payloads.append(
                (
                    ole.openstream("CONTENTS").read(),
                    ole.openstream(PRESENTATION_STREAM).read(),
                )
            )
    assert payloads[0][0] != payloads[1][0], "两个分子的 CDX 应不同"
    assert payloads[0][1] != payloads[1][1], "两个分子的缓存渲染也应不同"


# --- Excel 侧 -----------------------------------------------------------------


def _excel_streams(path: Path) -> list[set[str]]:
    with zipfile.ZipFile(path) as archive:
        return [
            {entry[0] for entry in olefile.OleFileIO(io.BytesIO(archive.read(name))).listdir()}
            for name in sorted(n for n in archive.namelist() if "/embeddings/" in n)
        ]


def test_excel_cache_follows_the_preview_format(tmp_path):
    """默认三态：矢量预览挂缓存，位图预览不挂。

    缺了缓存，把结构复制到 PowerPoint 会唤起 ChemDraw 现场出图（软件自己弹
    出来）。位图预览没有可用的载荷——CF_ENHMETAFILE 只吃 EMF——单为缓存再跑
    一遍矢量管线不值当。
    """
    from retain_cdx import ExcelCDXWriter

    for fmt, expected in (("emf", True), ("png", False)):
        output = tmp_path / f"{fmt}.xlsx"
        with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
            writer.insert_cdx(CDX, sheet="Sheet1", cell="B2", preview_format=fmt)
            writer.save(output)
        (streams,) = _excel_streams(output)
        assert (PRESENTATION_STREAM in streams) is expected, f"preview_format={fmt}"


def test_excel_cache_can_be_switched_off(tmp_path):
    """几千个结构的表可以关掉——每个对象多 10~15 KB。"""
    from retain_cdx import ExcelCDXWriter

    output = tmp_path / "no_cache.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx(
            CDX, sheet="Sheet1", cell="B2", preview_format="emf",
            presentation_cache=False,
        )
        writer.save(output)

    (streams,) = _excel_streams(output)
    assert PRESENTATION_STREAM not in streams


def test_excel_cache_can_be_switched_on(tmp_path):
    """开了之后复制到 PowerPoint 不会再唤起 ChemDraw 现场出图。"""
    from retain_cdx import ExcelCDXWriter

    output = tmp_path / "cached.xlsx"
    with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
        writer.insert_cdx_many(
            ["tests/fixtures/1.cdx", "tests/fixtures/2.cdx"],
            sheet="Sheet1",
            start_cell="B2",
            preview_format="emf",
            presentation_cache=True,
        )
        writer.save(output)

    for streams in _excel_streams(output):
        assert PRESENTATION_STREAM in streams

    with zipfile.ZipFile(output) as archive:
        assert not [n for n in archive.namelist() if n.endswith(".wmf")], (
            "伴生 WMF 是缓存的原料，落进包里就是一个没有内容类型声明的部件"
        )


def test_excel_cache_needs_a_vector_preview(tmp_path):
    """PNG 不能当 CF_ENHMETAFILE 载荷——明确报错，不静默降级。"""
    from retain_cdx import ExcelCDXWriter

    with pytest.raises(ValueError, match="vector preview"):
        with ExcelCDXWriter("tests/fixtures/blank.xlsx") as writer:
            writer.insert_cdx(
                CDX, sheet="Sheet1", cell="B2", preview_format="png",
                presentation_cache=True,
            )
            writer.save(tmp_path / "never.xlsx")


def test_powerpoint_cache_can_be_turned_off(tmp_path):
    """PPT 默认开，但几百页的文件可以关掉省体积。"""
    output = tmp_path / "no_cache.pptx"
    with PowerPointCDXWriter(BLANK_PPTX) as writer:
        writer.insert_cdx(
            CDX,
            placement=PowerPointPlacement(x="1cm", y="1cm", width="4cm"),
            preview=PreviewOptions(format="emf", presentation_cache=False),
        )
        writer.save(output)

    with zipfile.ZipFile(output) as archive:
        name = [n for n in archive.namelist() if "/embeddings/" in n][0]
        streams = {e[0] for e in olefile.OleFileIO(io.BytesIO(archive.read(name))).listdir()}
    assert PRESENTATION_STREAM not in streams
