# ChemDraw OLE 注入（Excel）

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
  - `retain_cdx/data/template.xlsx` - 默认使用的内置模板
    （单行、已包含一个 ChemDraw OLE 对象）。
- `scripts/` - 模块化脚本目录。
  - `scripts/cli/` - CLI 入口（`python3 -m scripts.cli.<名称>` 运行）
  - `scripts/ole/` - OLE 注入逻辑
  - `scripts/render/` - 预览图生成
  - `scripts/xlsx/` - XLSX 结构处理
- `examples/` - 示例输入文件（`san.cdx`、`1.cdx`、`2.cdx` 等）。
- `docs/` - 开发记录与参考资料。
  - `docs/template_reference/` - 内置模板解包后的静态参考文件。
- `tests/` - 测试与夹具。

## 依赖

- Python 3.10+
- `olefile`（`pip install olefile`）
- Open Babel 命令行（`obabel`，用于生成预览 PNG）
- 可选：`Pillow`（用于预览图裁剪，`.[dev]` 会安装）

## 开发安装

建议在虚拟环境中用 editable 模式安装：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -U pip
python -m pip install -e ".[dev]"
```

如果需要从 `SMILES` 生成 `CDX`，或使用 CDX 归一化/比例分析能力，还需要安装化学工具链：

```bash
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
`output/`、`Data/` 或 `sdfData/`。

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

## 用法

推荐用法（默认，直接使用内置模板）：

```bash
python3 -m scripts.cli.inject_cdx_into_xlsx --cdx examples/san.cdx --output output.xlsx
```

显式指定模板：

```bash
python3 -m scripts.cli.inject_cdx_into_xlsx --template my_template.xlsx --cdx examples/san.cdx --output output.xlsx
```

多行（每行一个 OLE，默认也走内置模板）：

```bash
python3 -m scripts.cli.inject_cdx_into_xlsx --cdx examples/1.cdx examples/2.cdx --output output_2rows.xlsx
```

脚本会自动：
- 把文件名（去后缀）写入 A 列。
- 在 B 列放入对应的 OLE 对象。
- 用 Open Babel 生成预览 PNG 并绑定到对应 OLE。

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

锚点语法：

```text
A1
A1+5mm,3mm
B2+0pt,40pt
```

含义：
- `A1` 表示左上角单元格。
- `+5mm,3mm` 表示在该单元格内再偏移 `5mm` 和 `3mm`。
- 支持单位：`mm`、`pt`、`in`、`emu`。

### `--from` 和 `--to` 的行为

这个规则很重要：

- 只传 `--from`
  - 只平移对象
  - 保持模板原始长宽比和尺寸不变
  - 适合“换位置，不想变形”

- 同时传 `--from` 和 `--to`
  - 重新定义对象显示框
  - 宽高会按新的锚点框变化
  - 适合明确指定对象跨哪些单元格

例子：

```bash
# 只移动到 A1/A2 附近，保持原始尺寸
python3 -m scripts.cli.inject_cdx_into_xlsx \
  --cdx sdfData/GDG906.cdx \
  --from A1+2mm,45pt \
  -o moved_only.xlsx
```

```bash
# 强制定义一个新的显示框，会改变宽高
python3 -m scripts.cli.inject_cdx_into_xlsx \
  --cdx sdfData/GDG906.cdx \
  --from A1+5mm,20pt \
  --to A2+20mm,10pt \
  -o resized_box.xlsx
```

### 关于长宽比

- OLE 对象本身并不会自动“聪明缩放”。
- 如果你显式给了新的 `--from` 和 `--to`，程序会把这个矩形当成新的显示框。
- 如果这个框比原始比例更高、更窄，视觉上就会出现“被拉长”或“被压扁”。
- 如果只想改位置，优先只传 `--from`。

### 方形双行布局脚本

仓库另外提供了一个专门脚本，用于生成这种版式：

```text
OLE | 基态
OLE | 激发态 s1
```

入口：

```bash
python3 -m scripts.cli.make_state_pair_layout --cdx sdfData/GDG906.cdx -o state_pair.xlsx
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
  --cdx sdfData/GDG906.cdx \
  --top-label 基态 \
  --bottom-label 激发态\ s1 \
  --a-column-pixels 180 \
  --row-pixels 90 \
  --fill-ratio 0.9 \
  -o sdfData/GDG906_state_pair_fit90.xlsx
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

- 默认布局下，CLI 现在内部直接复用内置模板的稳定工作流。
- 只有在显式传入定位参数（如 `--col`、`--start-row`、`--width`、`--height`、
  `--from`、`--to`）时，才会走从零构造 XLSX 骨架的路径。

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
python3 -m scripts.cli.inject_cdx_into_xlsx --cdx examples/san.cdx --output output.xlsx
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
- 如果同时传了 `--from` 和 `--to`，对象会被放进你指定的新矩形里。
- 新矩形越偏离原始比例，视觉变形越明显。
- 如果只是改位置，优先只传 `--from`。
- 如果需要自动保持比例并尽量填满方形区域，优先用
  `python3 -m scripts.cli.make_state_pair_layout --fill-ratio 0.9`。

**6) 手头只有 `.sdf` 或 `.cdxml`，没有 `.cdx`**

- 当前机器上的 Open Babel 只能读取 `cdx`，不能写 `cdx`。
- 一条可行路线是使用 `pycdxml` + `rdkit`：
  - `SDF -> RDKit molecule`
  - `RDKit molecule -> binary CDX`
  - 再把生成的 `CDX` 注入 Excel
- `CDXML` 本身不能直接传给当前 CLI 的 `--cdx` 参数。
