"""OLE 演示流（`\\x02OlePres000`）——对象的缓存渲染。

没有它，容器仍是合法的嵌入对象、双击照样能进 ChemDraw；但**复制粘贴会退化
成图片**：目标程序拿不到缓存又不能现场唤起 OLE 服务器，只能取剪贴板里的位图。
Excel 侧的 ChemDraw 容器不带这个流也工作正常，PowerPoint 侧则需要。

流的头部照 Windows PowerPoint + ChemDraw 的真实产物（MS-OLEDS
OLEPresentationStream）::

    0   ffffffff              标记：后面是标准剪贴板格式
    4   0e000000              CF_ENHMETAFILE
    8   04000000              TargetDeviceSize = 4（无目标设备）
    12  Aspect / Lindex / Advf / Reserved1
    28  Width / Height        HIMETRIC（0.01mm）
    36  Size                  数据字节数
    40  数据
    末尾 TOCSIGNATURE("NANI") + TOCCOUNT=0

数据段**必须是 WMF**，尽管格式声明的是 CF_ENHMETAFILE。EMF 藏在 WMF 里，
靠一条 `MFCOMMENT` 转义记录（签名 `WMFC`）携带；WMF 自己的绘图记录留作
只认 WMF 的程序的后备。这不是从规范推出来的，是试出来的——四个变体在
Windows PowerPoint 上实测（2026-08-02）：

===========================  ================================================
变体                         结果
===========================  ================================================
裸 EMF 当数据段               双击即报「可能导致 PowerPoint 不稳定」
EMF 切成 2 条 WMFC 注释记录    不崩，但复制出去只是图片（EMF 没被取出）
WMFC 校验和填 0               同上——校验和确实被检验
整份 EMF 放在 1 条注释记录里    双击、复制到新 PPT、复制进 Excel 全部正常
===========================  ================================================

所以 EMF 必须整份放进**一条**注释记录，而记录的 ByteCount 是 16 位字段，
上限 65535 —— 见 `MAX_EMBEDDED_EMF`。超限时不挂这个流（宁可复制退化，
也不要产出会让 PowerPoint 崩的文件）。

校验和用的是 Wine 的算法（34 字节注释头的 16 位和取二补）。它算不出
ChemDraw 那份的 `0xe8c4`，但 PowerPoint 接受我们算的值，而填 0 会被拒——
两者都实测过。
"""

from __future__ import annotations

import struct

PRESENTATION_STREAM = "\x02OlePres000"

_MARKER = 0xFFFFFFFF
_CF_ENHMETAFILE = 14
_NO_TARGET_DEVICE = 4
_DVASPECT_CONTENT = 1
_LINDEX_NONE = -1
_ADVF = 2  # ADVF_PRIMEFIRST，与参照文件一致
_TOC = struct.pack("<II", 0x494E414E, 0)  # TOCSIGNATURE("NANI") + TOCCOUNT

_PLACEABLE = bytes.fromhex("d7cdc69a")
_WMF_HEADER = "<HHHIHIH"  # type, headerSize(词), version, size(词), objects, maxRecord(词), params
_WMF_HEADER_WORDS = 9
_META_ESCAPE = 0x0626
_MFCOMMENT = 0x000F
_COMMENT_HEADER_SIZE = 34  # WMFC 起，到 EnhancedMetafileDataSize 止

#: 单条注释记录能携带的 EMF 上限：ByteCount 是 16 位，还要减去注释头。
MAX_EMBEDDED_EMF = 0xFFFF - _COMMENT_HEADER_SIZE

EMU_PER_HIMETRIC = 360  # 1 HIMETRIC = 0.01mm = 360 EMU


def emu_to_himetric(emu: int) -> int:
    return max(1, round(emu / EMU_PER_HIMETRIC))


def wrap_emf_in_wmf(emf: bytes, wmf: bytes) -> bytes:
    """把 *emf* 塞进 *wmf* 的记录流最前面，作为一条 `WMFC` 注释记录。

    *wmf* 提供后备的绘图记录（Inkscape 直出的即可，带不带 placeable 头都行）。
    """
    if not emf.startswith(b"\x01\x00\x00\x00"):
        raise ValueError("payload is not an EMF (missing EMR_HEADER)")
    if len(emf) > MAX_EMBEDDED_EMF:
        raise ValueError(
            f"EMF is {len(emf)} bytes; a single WMFC comment record holds at "
            f"most {MAX_EMBEDDED_EMF} (splitting across records is not accepted)"
        )

    body = wmf[22:] if wmf[:4] == _PLACEABLE else wmf
    mtype, header_words, version, _, objects, max_record, params = struct.unpack_from(
        _WMF_HEADER, body, 0
    )
    records = body[header_words * 2 :]

    comment = _comment_record(emf)
    out = comment + records
    header = struct.pack(
        _WMF_HEADER,
        mtype,
        _WMF_HEADER_WORDS,
        version,
        (_WMF_HEADER_WORDS * 2 + len(out)) // 2,
        objects,
        max(max_record, len(comment) // 2),
        params,
    )
    return header + out


def _comment_record(emf: bytes) -> bytes:
    """一条 META_ESCAPE / MFCOMMENT 记录，整份携带 *emf*。"""
    fields = bytearray()
    fields += b"WMFC"
    fields += struct.pack("<II", 1, 0x00010000)  # CommentType, Version
    fields += b"\x00\x00"  # Checksum 占位
    fields += struct.pack(
        "<IIIII",
        0,  # Flags
        1,  # CommentRecordCount：只能是 1，多条会让 PowerPoint 取不出 EMF
        len(emf),  # CurrentRecordSize
        0,  # RemainingBytes
        len(emf),  # EnhancedMetafileDataSize
    )
    checksum = -sum(struct.unpack("<17H", bytes(fields))) & 0xFFFF
    fields[12:14] = struct.pack("<H", checksum)

    body = bytes(fields) + emf
    payload = body + (b"\x00" if len(body) % 2 else b"")
    size_words = (10 + len(payload)) // 2
    return (
        struct.pack("<IH", size_words, _META_ESCAPE)
        + struct.pack("<HH", _MFCOMMENT, len(body))
        + payload
    )


def build_enhmetafile_presentation(
    emf: bytes, wmf: bytes, *, width_himetric: int, height_himetric: int
) -> bytes:
    """组一份 `\\x02OlePres000` 流：*emf* 是真正的缓存，*wmf* 是后备绘图记录。"""
    data = wrap_emf_in_wmf(emf, wmf)
    header = struct.pack(
        "<IIIiiiiiii",
        _MARKER,
        _CF_ENHMETAFILE,
        _NO_TARGET_DEVICE,
        _DVASPECT_CONTENT,
        _LINDEX_NONE,
        _ADVF,
        0,  # Reserved1
        width_himetric,
        height_himetric,
        len(data),
    )
    return header + data + _TOC


__all__ = [
    "MAX_EMBEDDED_EMF",
    "PRESENTATION_STREAM",
    "build_enhmetafile_presentation",
    "emu_to_himetric",
    "wrap_emf_in_wmf",
]
