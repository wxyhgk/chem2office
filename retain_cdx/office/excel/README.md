# Excel 模块结构说明

这个目录负责把可编辑的 ChemDraw CDX OLE 对象写入 `.xlsx` 文件。

目前最稳定的路径是 **shell 模板路径**：

1. 复用一个由 Excel + ChemDraw 原生创建的 OLE shell workbook。
2. 替换 OLE 容器里的 CDX `CONTENTS` 流。
3. 同步 Excel 表面显示用的预览图。
4. 把这个 source workbook 里的 OLE / drawing / VML / media parts 合并到目标 workbook。

## 对外 API

### `__init__.py`

Excel 模块的对外出口。这里应该只暴露稳定 public API，不要把内部工具都挂出去。

当前主要导出：

- `ExcelCDXWriter`
- `Placement`
- `PreviewOptions`
- `DisplayOptions`
- `ShellTemplate`

### `writer.py`

用户真正调用的 workbook writer。

它应该保持轻量，只负责串联流程：

1. 生成插入计划。
2. 准备临时 CDX 文件。
3. 构建 source workbook。
4. 合并 source workbook 到目标 workbook。
5. 修正 layout / content types。

不要在这里继续堆预览图、VML、relationship、OLE 二进制细节。

### `options.py`

用户可见的配置 dataclass：

- `Placement`
- `PreviewOptions`
- `DisplayOptions`
- `ShellTemplate`

这些类型属于 public API，后续改动要谨慎。

## 内部子包

带 `_` 前缀的目录都是内部实现，不建议用户直接 import。内部重构可以改这些路径，但不要影响 public API。

### `_insert/`

负责把用户传入的参数转成明确的插入输入。

- `plan.py`
  构建 `InsertPlan`，统一处理：
  - placement
  - display
  - shell
  - preview
  - normalize
  - width / height / row height / column width

- `cdx.py`
  把 `StructureSource` 写成临时 `.cdx` 文件。
  如果用户要求 normalize，也在这里处理。

### `_source/`

负责创建临时 source workbook。

source workbook 是一个中间 `.xlsx`，里面已经包含要合并的 OLE 对象。

- `builder.py`
  根据是否使用 shell 选择构建路径：
  - shell 模式
  - 非 shell/generated 模式

- `workbook.py`
  具体创建 source workbook。
  shell 模式会：
  - 解包 shell workbook
  - 替换 OLE `CONTENTS`
  - 更新预览图
  - 必要时 patch shell 几何
  - 重新打包成 source workbook

### `_shell/`

负责 Excel/ChemDraw 原生 shell 相关细节。

- `geometry.py`
  处理 shell 的显示框和 anchor：
  - 读取 VML 里的宽高
  - 在允许 resize 时 patch worksheet / drawing / VML geometry

- `preview.py`
  处理 shell 的预览图同步：
  - 默认预览后端：`RDKit ACS1996`（黑白期刊风格；CDX 经 pycdxml→CDXML 解析）
  - 回退：Open Babel CLI（`-xu`/`-xU` 黑白）
  - 可选 `preview_format="emf"`：SVG + Inkscape → EMF
  - PNG 路径会裁剪白边，并按 OLE 显示框比例生成“不变形优先”的预览图
  - 同步 worksheet relationship 和 VML relationship

注意：RDKit 路径需 `.[chem]`（rdkit + pycdxml）。EMF 仍需 Inkscape。

### `_parts/`

负责把 source workbook 里的 parts 合并到目标 workbook。

- `ole.py`
  处理：
  - drawing part 复制
  - VML part 复制
  - OLE embedding 复制
  - media 复制
  - relationship 重写
  - `oleObjects` 挂到目标 worksheet
  - shell layout 复制

这个模块只负责合并 parts，不负责生成 CDX，也不负责渲染预览图。

### `_ooxml/`

负责底层 Open XML / XLSX 操作。

- `package.py`
  解包和重新打包 `.xlsx`。

- `content_types.py`
  修正 `[Content_Types].xml`，确保 PNG、OLE bin、VML drawing、drawing XML 等类型存在。

- `relationships.py`
  写 `.rels` 文件。

- `workbook.py`
  解析 workbook，找到 sheet 名对应的 worksheet XML。

- `worksheet.py`
  修改 worksheet XML：
  - 添加 drawing / legacyDrawing / oleObjects
  - 复制 shell 行高列宽布局
  - 非 shell 模式下设置行高列宽

### `_utils/`

小型工具函数。

- `address.py`
  解析 Excel 单元格地址，处理列号转换。

- `units.py`
  处理 px / pt / EMU 等单位转换。

## 推荐依赖方向

尽量保持单向依赖：

```text
writer
  -> _insert
  -> _source
  -> _shell
  -> _parts
  -> _ooxml
  -> _utils
```

基本原则：

- 内部模块不要 import `writer.py`。
- `_utils` 不要反向 import 上层模块。
- `_ooxml` 只做 XML / package 基础操作，不关心 ChemDraw。
- `_shell` 只关心 shell，不关心目标 workbook 合并。
- `_parts` 只关心合并 parts，不关心 CDX 从哪里来。
- `writer.py` 只编排流程，不写底层 XML 细节。

## import 管理原则

每个内部子包的 `__init__.py` 是门面。

例如 `writer.py` 应优先这样导入：

```python
from retain_cdx.office.excel._insert import build_insert_plan, write_cdx_sources
from retain_cdx.office.excel._source import build_source_workbook
from retain_cdx.office.excel._parts import merge_source_workbook
from retain_cdx.office.excel._ooxml import extract_xlsx, pack_xlsx
```

不要在 `writer.py` 里到处导入具体文件：

```python
# 尽量避免
from retain_cdx.office.excel._ooxml.package import extract_xlsx
from retain_cdx.office.excel._ooxml.workbook import parse_workbook_sheets
```

这样做的目的是让内部文件可以继续重构，但上层 import 不需要频繁变化。

## 当前限制

- shell 模式目前一次只支持插入一个 CDX 对象。
- 稳定路径依赖原生 Excel/ChemDraw OLE shell。
- 预览默认 PNG + RDKit ACS1996；`preview_backend="obabel"` 可强制 Open Babel。
- `preview_format="emf"` 需要 Inkscape；`preview_format="auto"` 有 Inkscape 时用 EMF。
- 双击 OLE 打开 ChemDraw 后，Excel/ChemDraw 可能会重新生成预览。
  CDX 内容不会变，但预览风格可能与 RDKit ACS 示意仍有差异。
