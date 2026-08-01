# retain_cdx 包结构说明

把可编辑的 ChemDraw CDX 写成 Office 文档里的 OLE 对象。

本文说的是**包内部怎么分工**。面向用户的安装与用法看仓库根目录的
`README.md`；Excel 实现细节看 `office/excel/README.md`。

## 一句话分工

```
结构从哪来        →  structure/     SMILES / CDX / CDXML，归一化与几何
装进什么容器      →  ole/           CFBF（.bin），ChemDraw 的 CONTENTS 流
预览图怎么画      →  render/        RDKit ACS1996 / Open Babel，PNG 与矢量 EMF
往哪个文档里放    →  office/        excel/ 与 powerpoint/
```

前三个与 Office 格式**完全无关**，可以单独使用；`office/` 才关心
worksheet、slide、锚点这些东西。

## 公共 API

只有这些是稳定的，其余（含 `_` 前缀模块）随时可能重构。
**带格式前缀的归该格式，不带前缀的是两边共享的**：

```python
from retain_cdx import (
    ExcelCDXWriter,        # 往 .xlsx 写；也能写单元格文本、行高列宽
    ExcelPlacement,        # 定位：单元格 + 格内偏移，或两点锚（to_cell）
    ExcelDisplayOptions,   # 显示框、行高列宽、归一化、预览
    ExcelShellTemplate,    # 复用原生 Excel+ChemDraw 模板的路径

    PowerPointCDXWriter,   # 往已有 .pptx 的某页某位置放分子
    PowerPointPlacement,   # 绝对 EMU 框，支持 cm/mm/in/pt/px/emu
    PowerPointGrid,        # 网格平铺

    StructureSource,       # from_cdx / from_smiles，可 normalized()
    PreviewOptions,        # 预览格式/后端/尺寸（两个 writer 共用）
)
```

## 目录

### `structure/` — 分子这一侧

| 模块 | 职责 |
|---|---|
| `source.py` | `StructureSource`：CDX 字节 + 名称，`from_smiles` 可传 ChemDraw 样式 |
| `converters.py` | SMILES → CDX（经 pycdxml/RDKit） |
| `normalize.py` | 改写 CDX 的 viewbox，让分子按目标比例/填充率排布 |
| `geometry.py` | 从 CDX 读内容宽高比，供显示框贴合使用 |

### `ole/` — CFBF 容器这一侧

| 模块 | 职责 |
|---|---|
| `embed.py` | 唯一入口 `write_cdx_to_ole`：能就地补丁就补丁，否则重建 |
| `inject.py` | 写入既有流、按 FAT 链定位并修补目录项长度 |
| `cfbf.py` | 从零构造合法容器（`\x01Ole` / `\x01CompObj` / `CONTENTS` + 可选额外流） |

三个流与 Windows 产物逐字节一致，CLSID 与 ProgID 跨 Excel / PowerPoint 通用。
容器会随载荷增长到所需的 FAT / mini-FAT / 目录扇区数——早期写死单个 FAT 扇区，
导致超过 63,488 字节的 CDX 静默损坏（见 `docs/excel_ooxml_research.md` C1）。

**何时重建而非补丁**：载荷变大、跨过 4096 字节的 mini-stream 分界、
或容器里带着会过期的 `\x02OlePres000` 演示流缓存。

### `render/` — 预览图这一侧

`preview.py` 一个模块，对外提供 `render_preview` / `render_emf_many` 等。
默认 RDKit ACS 1996（单色期刊风格），不可用时回退 Open Babel。
EMF 走矢量路径（ACS SVG → letterbox → Inkscape），批量渲染时用常驻
`inkscape --shell` 进程，比逐个冷启动快得多。

预览图**不进 OLE 容器**，它是 OOXML 层的部件——这是 Excel 侧的事实；
PowerPoint 的原生容器确实带演示流，但我们不复制陈旧的那份。

### `office/` — 文档这一侧

```
office/
  _package.py   extract_ooxml / pack_ooxml —— 任何 OOXML 都是 ZIP
  _units.py     EMU / px / pt 换算
  options.py    PreviewOptions（与格式无关，两个 writer 共用）
  excel/        SpreadsheetML：网格、三套锚点、VML
  powerpoint/   PresentationML：绝对 EMU，无网格无 VML
```

**硬规则：`excel/` 与 `powerpoint/` 互不相识。** 两者共用的东西放 `office/` 层，
不允许跨格式 import 对方的 `_` 私有模块。这条规则是补出来的——PowerPoint 最初
伸手拿了 Excel 的 `_ooxml`、`_utils` 私有模块，而拿的三样东西根本不属于 Excel。

两个格式的难度差别很大：

| | Excel | PowerPoint |
|---|---|---|
| 定位 | 三套锚点（worksheet / drawing / VML）必须一致 | 一处 `p:xfrm`，绝对 EMU |
| 网格 | 列宽像素依赖工作簿字体的 max digit width | 不看行列 |
| VML | 必需 | 不存在 |
| 元素顺序 | CT_Worksheet 固定序列，插错即触发修复 | `p:spTree` 追加即可 |

Excel 的内部分层另见 `office/excel/README.md`。PowerPoint 只有三个模块：
`writer.py`（门面）、`_pml.py`（幻灯片手术）、`options.py`（定位与网格）。

### `data/`

`template.xlsx` 与 `template.pptx`：由 Windows Excel / PowerPoint + ChemDraw
真实插入产出的模板，提供 OLE 容器与骨架 XML。**它们是零件供体，不是你的起点**
——PowerPoint 侧必须传入自己的 `.pptx`。

## 依赖方向

```
       structure/   ole/   render/        与格式无关，互不依赖
              ↘      ↓      ↙
                  office/                 _package / _units / options
                  ↙      ↘
             excel/       powerpoint/     互不相识
```

- 下层不 import 上层；`office/` 不 import 具体 writer。
- `structure` / `ole` / `render` 不知道 Office 的存在。
- 跨格式共用的一律上提到 `office/`。

## 踩过的坑（都有回归测试）

这些结论来自对真实 Office 文件的解剖，不是推测：

- **位置启发式在 OOXML 里一律不可靠。** 曾用「插在 `pageMargins` 后面」放
  `drawing`，真实文件带 `pageSetup` 时就把它挤到后面，Excel 报修复并删掉 OLE。
  必须按 schema 序列排位。
- **ElementTree 会丢掉「声明了但没被元素使用」的命名空间**，而 MCE 恰恰依赖
  这种声明。同一个根因造成过 PowerPoint 报修复（`xmlns:v`）和 Excel 丢
  `mc:Ignorable` 前缀。凡碰 `mc:` 都要手动兜。
- **列宽的单位是工作簿字体的 max digit width，不是固定 7px。** 内置模板自己
  就是等线（8px）工作簿。
- **openpyxl 会丢 OLE 对象，python-pptx 不会。** 所以 Excel 侧我们必须收尾，
  PPT 侧两个库可以任意穿插。

完整审计见 `docs/excel_ooxml_research.md` 与 `docs/powerpoint_ole_research.md`；
`tests/ooxml_checks.py` 把其中几条固化成了包级校验（关系悬空、内容类型缺失、
shape id 重复、`Requires`/`Ignorable` 前缀未声明）。
