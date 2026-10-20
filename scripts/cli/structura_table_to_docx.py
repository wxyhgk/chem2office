#!/usr/bin/env python3
"""Structura 化合物表包 → 专利式 Word 表格。

每组两行：上面一行结构，下面一行编号（C-1、C-2…）。整张表用同一个缩放
系数，所有结构键长相同；默认系数让最宽的结构刚好放进一列。

Usage::

    python3 -m scripts.cli.structura_table_to_docx \\
        --input "某专利 化合物表.zip" --output 化合物表.docx

    # 指定每行个数与缩放；放进已有文档里 {{compounds}} 所在的段落
    python3 -m scripts.cli.structura_table_to_docx \\
        --input 化合物表/ --output out.docx --columns 4 --scale 0.8 \\
        --document report.docx --placeholder "{{compounds}}"
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from retain_cdx import CompoundTable, PreviewOptions, WordCDXWriter


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Structura compound-table package (zip or directory) → Word table"
    )
    parser.add_argument("--input", required=True, type=Path, help="化合物表 .zip 或解压后的目录")
    parser.add_argument("--output", required=True, type=Path, help="输出 .docx")
    parser.add_argument("--columns", type=int, help="每行几个结构（默认取包里的设置）")
    parser.add_argument(
        "--scale",
        type=float,
        help="统一缩放系数，1.0 = ACS 1996 原大（键长 14.4 pt）；默认让最宽的结构刚好放进一列",
    )
    parser.add_argument(
        "--document", type=Path, help="放进这个已有 .docx（默认新建空白 A4 文档）"
    )
    parser.add_argument(
        "--placeholder", help="用表格替换文档中正文恰好是这段文字的段落（默认追加到文末）"
    )
    parser.add_argument(
        "--valign",
        choices=("bottom", "center", "top"),
        default="bottom",
        help="同一行结构高度不同时的竖直对齐（默认 bottom，编号紧贴结构）",
    )
    parser.add_argument(
        "--preview-format",
        choices=("auto", "emf", "png"),
        default="auto",
        help="预览图格式（默认 auto：有 Inkscape 用矢量 EMF）",
    )
    parser.add_argument(
        "--no-cache",
        action="store_true",
        help="不在 OLE 容器里放渲染缓存（\\x03EPRINT），每个对象小一些",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.placeholder and not args.document:
        raise SystemExit("--placeholder needs --document")

    started = time.perf_counter()
    table = CompoundTable.load(args.input)
    preview = PreviewOptions(
        format=args.preview_format,
        presentation_cache=False if args.no_cache else None,
    )
    with WordCDXWriter(args.document) as writer:
        scale = writer.insert_compound_table(
            table,
            columns=args.columns,
            scale=args.scale,
            placeholder=args.placeholder,
            vertical_align=args.valign,
            preview=preview,
        )
        writer.save(args.output)

    columns = args.columns or table.columns
    print(
        f"{len(table.compounds)} compounds, {columns} per row, scale {scale:.3f} "
        f"→ {args.output} ({time.perf_counter() - started:.1f}s)"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
