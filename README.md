# ChemDraw OLE 注入（Excel / PowerPoint / Word）

这个仓库实现了 ChemDraw 的 “OLE 模板注入” 流程，在 Linux/macOS 下批量生成
可编辑的 ChemDraw OLE。默认模式会直接复用包内内置的模板
（`retain_cdx/data/template.xlsx`）作为稳定骨架；如果显式传入
`--template`，则使用你指定的模板文件。

## 为什么默认仍然依赖模板

Excel 的 OLE 对象存储在微软 CFBF 容器里，ChemDraw 使用的是二进制 CDX 流。
虽然仓库里保留了“从零生成”的代码路径用于自定义定位，但默认输出为了兼容
Excel，仍然采用模板注入方式：

1) 用 Windows Excel 插入 ChemDraw 对象，保存模板。
2) 替换 OLE 的 `CONTENTS` 流为新的 CDX。
3) 更新预览图，让 Excel 显示正确结构。

## 项目定位

开源生态里，和这个问题相关的能力通常分成两层：

1. 生成或转换 ChemDraw 数据
   - 例如把分子、`SDF`、`CDXML` 转成二进制 `CDX`
   - 这类能力常见于化学信息学工具链，例如 `pycdxml`、`OpenEye`

2. 把 `CDX` 写成 Excel 里的可编辑 ChemDraw OLE 对象
   - 这一步不仅要处理 `CDX`
   - 还要处理 Excel 的 `xlsx` 结构、`oleObject*.bin`、预览图、
     worksheet/drawing/VML 锚点、relationships、content types

这个仓库主要解决的是第 2 层。

也就是说：
- `pycdxml` / `OpenEye` 更像是上游 `CDX` 生产器
- 这个项目负责把已有的 `CDX` 落进 Excel，生成可编辑 OLE 对象

这也是为什么这个仓库不只是“格式转换脚本”，而更接近一个面向 ChemDraw 的
Excel OLE 生成器。

## 文件说明

- `retain_cdx/` - 库主包（`ExcelCDXWriter` 等 API）。
  包内分工见 `retain_cdx/README.md`。
  - `retain_cdx/office/excel/` - Excel OLE 写入实现
    （结构说明见 `retain_cdx/office/excel/README.md`）。
  - `retain_cdx/office/powerpoint/` - PowerPoint 幻灯片 OLE 写入
    （只负责摆放分子；做 PPT 交给 python-pptx 之类，见
    `docs/powerpoint_ole_research.md`）。
  - `retain_cdx/office/word/` - Word 文档 OLE 写入（随文字排版的嵌入对象；
    文档本身交给 python-docx 之类，见 `docs/word_ole_research.md`）。
  - `retain_cdx/render/` - 预览图渲染（RDKit / Open Babel / EMF）。
  - `retain_cdx/structure/` - 结构来源、归一化、几何。
  - `retain_cdx/ole/` - CFBF/OLE 容器读写。
  - `retain_cdx/data/template.xlsx`、`template.pptx`、`template.docx` - 内置模板
    （由 Windows Excel/PowerPoint/Word + ChemDraw 产出，提供 OLE 容器与骨架）。
- `scripts/` - 脚本目录。
  - `scripts/cli/` - CLI 入口（`python3 -m scripts.cli.<名称>` 运行），
    全部基于 `retain_cdx` 公共 API 实现：
    - `inject_cdx_into_xlsx` - 把 CDX 文件批量注入 Excel（每行一个 OLE）
    - `smiles_id_structure_table` - 从 `.smiles` 文件生成「编号 + 结构」表
    - `add_smiles_ole_by_merge` - 给已有数据表插入一列 ChemDraw 结构
    - `make_state_pair_layout` - 生成「OLE + 基态/激发态标签」双行版式
    - `structura_table_to_docx` - Structura 化合物表包 → 专利式 Word 表格
  - 旧版 `scripts/ole/`、`scripts/render/`、`scripts/xlsx/` 已删除
    （逻辑与 `retain_cdx` 包内实现重复）；新代码请依赖 `retain_cdx`
    公共 API
- `tests/fixtures/` - 示例输入文件（`san.cdx`、`1.cdx`、`2.cdx` 等）。
- `docs/` - 开发记录与研究报告。
  - `excel_ooxml_research.md` / `powerpoint_ole_research.md` /
    `word_ole_research.md` - 对真实
    Office 文件的实证解剖，记录了本项目大部分修复的依据。
  - `developer_log.md` - 历史调试记录（路径为当时的目录结构）。
  - `template_reference/` - 内置模板解包后的静态参考文件。
- `tests/` - 测试、夹具与本地数据。
  - `tests/fixtures/` - 已跟踪的样例输入（`san.cdx`、`blank.xlsx` 等）。
  - `tests/data/` - 本地输入数据（不入库）。
  - `tests/output/` - 运行产物（不入库）。

## 依赖

- Python 3.10+
- `olefile`（`pip install olefile`）
- 预览图后端（二选一，默认优先 RDKit）：
  - RDKit + `pycdxml`（`.[chem]` 会安装）— 默认，ACS 1996 期刊风格。
    注意：`pycdxml` 目前未发布到 PyPI，需从 GitHub 安装：
    `pip install git+https://github.com/kienerj/pycdxml`
  - Open Babel 命令行（`obabel`）— 回退方案
- 可选：`Pillow`（用于预览图裁剪/letterbox，`.[dev]` 会安装）
- 可选但建议装：Inkscape。预览默认 `preview_format="auto"`——装了就走矢量 EMF
  （更清晰、文件更小、渲染更快），没装则退回 PNG。只有矢量预览才能挂渲染缓存，
  没有缓存时把结构复制到 PowerPoint 会自动唤起 ChemDraw 现场出图。
- 仅 `scripts/cli/` 脚本需要：`openpyxl`（用来准备/预处理普通表格部分；
  库本体不依赖它）

## 开发安装

建议在虚拟环境中用 editable 模式安装：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -U pip
python -m pip install -e ".[dev]"
```

如果需要从 `SMILES` / `SDF` 生成 `CDX`，或使用 CDX 归一化/比例分析能力，
还需要安装化学工具链。`pycdxml` 目前未发布到 PyPI，需先从 GitHub 安装：

```bash
python -m pip install git+https://github.com/kienerj/pycdxml
python -m pip install -e ".[chem]"
```

Open Babel 可通过 Homebrew 或 conda 安装：

```bash
brew install open-babel
# 或
conda install -c conda-forge openbabel
```

安装后可用下面的命令确认系统里能找到 `obabel`：

```bash
obabel -V
```

## 测试

```bash
python -m pytest
```

测试夹具位于 `tests/fixtures/`。测试会写入临时目录，不需要使用仓库根目录下的
`tests/output/` 或 `tests/data/`。

## 库 API

库入口导出在 `retain_cdx` 包级别。最小用法如下：

```python
from retain_cdx import ExcelCDXWriter

with ExcelCDXWriter("input.xlsx") as writer:
    writer.insert_cdx("structure.cdx", sheet="Sheet1", cell="A1")
    writer.save("output.xlsx")
```

批量插入时先构造 `StructureSource`，再调用 `insert_cdx_many`：

```python
from retain_cdx import ExcelCDXWriter, StructureSource

sources = [
    StructureSource.from_cdx("first.cdx"),
    StructureSource.from_cdx("second.cdx"),
]

with ExcelCDXWriter("input.xlsx") as writer:
    writer.insert_cdx_many(sources, sheet="Sheet1", start_cell="A1")
    writer.save("output.xlsx")
```

### PowerPoint

只负责把分子放到指定幻灯片的指定位置；幻灯片本身用 `python-pptx` 之类去做：

```python
from retain_cdx import PowerPointCDXWriter, PowerPointPlacement, PowerPointGrid

with PowerPointCDXWriter("deck.pptx") as writer:
    writer.insert_cdx(source, slide=1,
                      placement=PowerPointPlacement(x="2cm", y="5cm", width="6cm"))
    writer.save("deck_with_structures.pptx")
```

长度支持 `cm/mm/in/pt/px/emu`；宽高只给一个时按分子比例补另一个。
多个结构可用 `PowerPointGrid(columns=3, cell_width="5cm", cell_height="4cm")` 平铺。
两个包可以任意先后使用——python-pptx 往返不会丢 OLE 对象。

### Word

对象随文字排版（像一个字符），所以不给坐标，只说「放在文字的哪里」和尺寸。
文档本身用 `python-docx` 之类去做，最顺手的交接方式是文字占位符：

```python
from docx import Document
from retain_cdx import WordCDXWriter

doc = Document()
table = doc.add_table(rows=1, cols=2)
table.cell(0, 0).text = "咖啡因"
table.cell(0, 1).text = "{{caffeine}}"
doc.save("report.docx")

with WordCDXWriter("report.docx") as writer:
    writer.insert_cdx(caffeine, placeholder="{{caffeine}}", height="2cm")
    writer.insert_cdx(benzene, align="center")   # 文末新段落
    writer.insert_cdx(san, paragraph=0)          # 第 0 个顶层段落末尾
    writer.save("report_with_structures.docx")
```

- 占位符跨多个 run（Word 手打文字常被切开）也能匹配；表格单元格里的也算。
  批量用 `insert_cdx_many(sources, placeholders=[...])`。
- 宽高只给一个时按分子比例补另一个；两个都给时分子按比例塞进框里，不拉伸。默认宽 4 cm。
- 矢量预览时容器带 `\x03EPRINT` 渲染缓存（就是预览 EMF 本身），
  `PreviewOptions(presentation_cache=False)` 关掉。
- python-docx 往返不会丢 OLE 对象，两个包可以任意先后使用。
- `WordCDXWriter()` 不传文件时从一份空白 A4 文档开始。

### Structura 化合物表 → 专利式 Word 表格

Structura 导出的「化合物表」包（zip 或解压后的目录：`manifest.json` + 每个化合物
一个 CDXML）可以直接排成专利里常见的表：每行若干个结构，结构下面一行是编号。

```bash
python3 -m scripts.cli.structura_table_to_docx \
    --input "某专利 化合物表.zip" --output 化合物表.docx [--columns 4] [--scale 0.8]
```

```python
from retain_cdx import CompoundTable, WordCDXWriter

table = CompoundTable.load("某专利 化合物表.zip")
with WordCDXWriter() as writer:                  # 或已有文档 + placeholder="{{compounds}}"
    scale = writer.insert_compound_table(table)  # columns= / scale= 可覆盖
    writer.save("化合物表.docx")
```

- **整张表一个缩放系数**：每个对象 = 结构的自然尺寸 × 同一个系数，所以键长处处
  相同。自然尺寸是 ACS 1996 画法（键长 14.4 pt）下的大小，`scale=1` 即原大；
  默认系数让最宽的结构刚好放进一列。
- 结构走 CDXML 而不是 SMILES：坐标原样保留，同一通式的核心朝向一致。
  单个文件也能用：`StructureSource.from_cdxml(路径或文本)`。
- 同一组的结构行和编号行不会被分页拆开；结构高度不同时默认底对齐
  （`vertical_align=`）。

## 用法

推荐用法（默认，直接使用内置模板）：

```bash
python3 -m scripts.cli.inject_cdx_into_xlsx --cdx tests/fixtures/san.cdx --output output.xlsx
```

显式指定模板：

```bash
python3 -m scripts.cli.inject_cdx_into_xlsx --template my_template.xlsx --cdx tests/fixtures/san.cdx --output output.xlsx
```

多行（每行一个 OLE，默认也走内置模板）：

```bash
python3 -m scripts.cli.inject_cdx_into_xlsx --cdx tests/fixtures/1.cdx tests/fixtures/2.cdx --output output_2rows.xlsx
```

脚本会自动：
- 把文件名（去后缀）写入 A 列。
- 在 B 列放入对应的 OLE 对象。
- 生成预览 PNG（默认 RDKit ACS 1996 风格，回退 Open Babel）并绑定到对应 OLE。

注意：`--template` 模式现在复用「原生 Excel+ChemDraw shell」路径
（`ExcelShellTemplate`），一次只支持一个 `--cdx` 文件。

### 位置与尺寸控制

`scripts.cli.inject_cdx_into_xlsx` 支持两类定位方式：

1. 简单定位
   - `--col`
   - `--start-row`
   - `--margin-left`
   - `--margin-top`
   - `--width`
   - `--height`

2. 精确锚点
   - `--from`
   - `--to`

锚点语法（`--from` 与 `--to` 相同）：

```text
A1
A1+5mm,3mm
B2+0pt,40pt
```

含义：
- `A1` 表示左上角单元格。
- `+5mm,3mm` 表示在该单元格内再偏移 `5mm` 和 `3mm`。
- 支持单位：`mm`、`pt`、`in`、`px`。

### `--from` / `--to` 的行为

- `--from` 只平移对象（覆盖 `--col`/`--start-row`/`--margin-*`）。
- 不带 `--to` 时，显示框大小由 `--width`/`--height` 决定。
- `--to` 用第二个锚点定义显示框：框从 `--from` 一直伸到
  `--to` 指定单元格的**左上角**（加偏移），跨度按目标表的真实
  列宽/行高解析（未设置的列/行按 Excel 默认 64px/19px）。
  设置 `--to` 时会覆盖 `--width`/`--height`，且不再调整行高/列宽。
  对应公共 API 的 `ExcelPlacement(to_cell=..., to_offset_x=..., to_offset_y=...)`。

例子：

```bash
# 只移动到 A1 附近，保持默认尺寸
python3 -m scripts.cli.inject_cdx_into_xlsx \
  --cdx tests/fixtures/san.cdx \
  --from A1+2mm,45pt \
  -o moved_only.xlsx

# 两点锚：框从 A1（内偏 2mm,10pt）一直铺到 C4 左上角
python3 -m scripts.cli.inject_cdx_into_xlsx \
  --cdx tests/fixtures/san.cdx \
  --from A1+2mm,10pt \
  --to C4 \
  -o two_point.xlsx
```

### 关于长宽比

- OLE 对象本身并不会自动“聪明缩放”。
- `--width`/`--height` 定义的矩形就是显示框。
- 如果这个框比原始比例更高、更窄，视觉上就会出现“被拉长”或“被压扁”。

### 方形双行布局脚本

仓库另外提供了一个专门脚本，用于生成这种版式：

```text
OLE | 基态
OLE | 激发态 s1
```

入口：

```bash
python3 -m scripts.cli.make_state_pair_layout --cdx tests/fixtures/san.cdx -o state_pair.xlsx
```

默认行为：
- `A` 列宽约 `180 px`
- `A1`、`A2` 行高各约 `90 px`
- 左侧 `A1 + A2` 组成约 `180 x 180 px` 的方形区域
- OLE 放在这个方形区域中间
- `B1=基态`
- `B2=激发态 s1`

可用参数：
- `--top-label`
- `--bottom-label`
- `--a-column-pixels`
- `--row-pixels`
- `--fill-ratio`

`--fill-ratio` 的规则：
- 默认 `0.9`
- 保持 OLE 原始长宽比
- 先判断宽和高中哪一边先触到容器限制
- 把受限边拉到容器的 `90%`
- 另一边按比例缩放
- 最终整体居中
- 因为缩放取的是宽高限制中的较小值，所以不会溢出单元格

例子：

```bash
python3 -m scripts.cli.make_state_pair_layout \
  --cdx tests/fixtures/san.cdx \
  --top-label 基态 \
  --bottom-label 激发态\ s1 \
  --a-column-pixels 180 \
  --row-pixels 90 \
  --fill-ratio 0.9 \
  -o state_pair_fit90.xlsx
```

## 工作原理（默认路径）

1) 解压 `.xlsx`（本质是 ZIP）。
2) 对每个 CDX：
   - 复制模板里的 OLE 文件（`xl/embeddings/oleObject*.bin`）。
   - 替换 `CONTENTS` 流为新的 CDX 数据。
   - 若新 CDX 仍能放进模板原有流，则直接覆盖并修补目录项长度。
   - 若新 CDX 大于模板原有流，则重建合法的 ChemDraw OLE 容器，只替换
     `oleObject*.bin`，其余 worksheet/drawing/VML 结构继续复用模板。
3) 生成 PNG 预览图并更新关联关系：
   - `xl/media/image*.png`
   - `xl/worksheets/_rels/sheet1.xml.rels`
   - `xl/drawings/_rels/vmlDrawing1.vml.rels`
4) 在 `sheet1.xml` 和 `vmlDrawing1.vml` 中复制/更新锚点。
5) 重新打包为 `.xlsx`。

说明：

- 上述流程现在全部由 `retain_cdx` 库内部实现（`ExcelCDXWriter`）；
  CLI 只负责解析参数并调用公共 API。
- 定位参数（`--col`、`--start-row`、`--width`、`--height`、`--from`）
  会被翻译成 `ExcelPlacement` / `ExcelDisplayOptions` 传给库。

## 限制

- 模板必须由 Windows Excel + ChemDraw 插入生成。
- 当前脚本假设模板结构类似内置的 `retain_cdx/data/template.xlsx`
  （简单一行一对象布局）。
- 当前项目默认只接受二进制 `CDX` 作为注入输入；`CDXML` 不能直接替代
  `--cdx` 参数。

## 备注

- 之前默认生成模式和模板模式不是同一路径，可能出现 Excel 打开后修复并删除
  OLE 对象的情况。
- 当前默认模式已经改为复用内置模板，以避免这类兼容性问题。
- 对于大 `CDX`，之前的问题不在 worksheet XML，而在重建出来的 `oleObject.bin`
  不够严格合法；修正 CFBF/FAT 常量后，Excel 不再删除对象记录。

## 常见问题（FAQ）

**1) Excel 提示 “被替换的部件：/xl/worksheets/sheet1.xml 有 XML 错误”**

- 旧版本默认生成模式和模板模式不一致，可能导致 Excel 删除对象记录。
- 当前默认模式已经回到内置模板工作流；请直接使用：

```bash
python3 -m scripts.cli.inject_cdx_into_xlsx --cdx tests/fixtures/san.cdx --output output.xlsx
```

**2) OLE 对象显示了，但双击打不开**

- 最常见原因是 CDX 不是 ChemDraw 原生导出的格式。
- 建议用 ChemDraw 导出 `.cdx` 再注入，或用 Excel 中插入的 OLE 作为模板。

**3) 报错 “CDX is larger than existing stream”**

- 旧版本模板模式只会覆盖现有 `CONTENTS` 流，所以会报这个错误。
- 当前版本在默认路径下会自动回退到“重建 OLE 容器”的方式处理大 `CDX`，
  不需要手工换更大的模板。
- 如果你显式传了 `--template` 且模板结构比较特殊，仍建议优先用内置模板
  （不传 `--template`）先验证。

**4) 预览图没有更新**

- 检查 `xl/media/` 下是否生成了 `image*.png`。
- 确认 `sheet1.xml.rels` 和 `vmlDrawing1.vml.rels` 里指向的是 PNG。

**5) 对象位置对了，但看起来被拉长或压扁**

- 这是显示框比例问题，不是 OLE 本身损坏。
- `--width`/`--height` 指定的矩形越偏离原始比例，视觉变形越明显。
- 如果只是改位置，优先只传 `--from`。
- 如果需要自动保持比例并尽量填满方形区域，优先用
  `python3 -m scripts.cli.make_state_pair_layout --fill-ratio 0.9`。

**6) 手头只有 `.sdf` 或 `.cdxml`，没有 `.cdx`**

- 当前机器上的 Open Babel 只能读取 `cdx`，不能写 `cdx`。
- 一条可行路线是使用 `pycdxml` + `rdkit`：
  - `SDF -> RDKit molecule`（`Chem.MolFromMolFile`）
  - `RDKit molecule -> binary CDX`（`pycdxml` 的 `mol_to_document` + `write_cdx_file`）
  - 再把生成的 `CDX` 注入 Excel
- `CDXML` 本身不能直接传给当前 CLI 的 `--cdx` 参数。

SDF → CDX 最小示例：

```python
import tempfile
from pathlib import Path
from rdkit import Chem
from rdkit.Chem import AllChem
from pycdxml.cdxml_converter.rdkit_chemdraw import mol_to_document
from pycdxml.cdxml_converter.chemdraw_io import write_cdx_file

mol = Chem.MolFromMolFile("input.sdf")          # 或 Chem.MolFromSmiles("c1ccccc1")
AllChem.Compute2DCoords(mol)                    # SDF 已含 2D 坐标时可省略
doc = mol_to_document(mol, chemdraw_style={"BondLength": 24, "LineWidth": 1.2})
with tempfile.TemporaryDirectory() as tmp:
    out = Path(tmp) / "structure.cdx"
    write_cdx_file(doc, str(out))
    cdx_bytes = out.read_bytes()                # 交给 StructureSource(cdx_data=...)
```

**7) 同一进程内先用过 RDKit，再解析 Office 骨架会段错误**

- 在同一 Python 进程里先调用过 RDKit 的 C 扩展（`MolFromMolFile`、
  `Compute2DCoords`、`mol_to_document` 等），之后 `xml.etree.ElementTree.parse`
  解析 xlsx/docx 骨架时会触发段错误（C 层符号冲突，栈顶是
  `xml.etree.ElementTree.parse`）。
- 规避方式：把「RDKit 处理分子」和「解析/写入 Office 文件」拆到不同进程
  （`multiprocessing` / `subprocess`），保证解析 Office 骨架的进程在
  `xml.etree.parse` 之前没有加载并使用过 RDKit。
- 直接调用 `insert_cdx_many`（从 CDX 文件出发）不受影响——预览渲染对
  RDKit 是懒加载，发生在骨架解析之后；只有「同进程内先用 RDKit 读
  SDF/SMILES，再做 Office 注入」这种组合需要分进程。
