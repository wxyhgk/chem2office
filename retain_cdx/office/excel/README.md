# Excel 模块结构说明

这个目录负责把可编辑的 ChemDraw CDX OLE 对象写入 `.xlsx` 文件。

概念上两步：**造零件**（bin + 预览）→ **装进 Excel**（本目录主责）。

三条路径，默认走第一条：

| 路径 | 怎么落地 | 何时用 |
|------|----------|--------|
| generated | 克隆模板片段，**直写目标包** | 默认 |
| shell | 改造一整份原生 shell 包，再 merge | 传了 `ExcelShellTemplate` |
| picture | 只放图片，不产生 OLE 对象 | `insert_pictures_many` |

## Public API vs `_` 内部（契约对照）

带 `_` 的包是**工业包常见的内部实现标记**：可随时重构路径与文件，**不保证** import 稳定。  
请只依赖下表「大门」一侧。

| 层级 | 可依赖？ | 符号 / 路径 | 说明 |
|------|----------|-------------|------|
| **大门 Public** | ✅ 稳定 | `retain_cdx.ExcelCDXWriter` | 写 workbook 的唯一推荐入口 |
| | ✅ | `ExcelPlacement`, `PreviewOptions`, `ExcelDisplayOptions`, `ExcelShellTemplate` | 用户配置 |
| | ✅ | `retain_cdx.StructureSource` | 化学结构（excel 外的领域对象） |
| **员工通道 Internal** | ⚠️ 勿当 API | `office/excel/models.py` | 内部 DTO（`PreviewSpec` / `LayoutSpec` / `ShellSourceOptions`） |
| | ⚠️ | `_insert/` | 编排：plan → CDX → 按路径写入 |
| | ⚠️ | `_ole_template.py` | 从内置模板抽出可克隆的 OLE 零件 |
| | ⚠️ | `_insert_ole.py` | generated 路径：把零件直接写进目标包 |
| | ⚠️ | `_shell/` | 仅 shell 路径：改造一份现成 shell 包 |
| | ⚠️ | `_parts/` | source → 目标 package 的 merge |
| | ⚠️ | `_shell/` | 原生 ChemDraw shell 几何/预览 |
| | ⚠️ | `_grid.py` / `_model.py` | Excel 网格度量与锚点值对象（跨子包共用） |
| | ⚠️ | `_picture.py` | 图片模式（不放 OLE） |
| | ⚠️ | `_ooxml/` / `_utils/` | XML / A1 地址底层 |
| **office 层共享** | ⚠️ | `office/_package.py` | OOXML 通用解包/打包（Excel 与 PPT 共用） |
| | ⚠️ | `office/_units.py` | 长度单位（EMU/px/pt） |
| | ✅ | `office/options.py` | `PreviewOptions`（与格式无关） |

```text
用户代码
   │  只应 import 大门
   ▼
writer + options          ← public
   │
   ▼
_insert ─┬─ generated → _ole_template + _insert_ole   （直写目标包）
         ├─ shell     → _shell/staging → _parts       （改造现成 shell 包）
         └─ picture   → _picture                      （直写，无 OLE）
```

**semver 含义**：改 public 签名要谨慎；改 `_` 包 / `models` 只要测试绿、public 行为不变即可。

### `__init__.py` / `writer.py` / `options.py`

- `__init__.py` — 只 re-export public 符号  
- `writer.py` — 薄门面：计划 → 执行 → save；另提供
  `set_cell_text(s)` / `set_column_width` / `set_row_height`
  （写结构旁边的 ID/表头等文本、调格子尺寸，
  **不要**用 openpyxl 后处理 —— 它会丢 OLE parts）  
- `insert_pictures_many()` — **图片模式**：只放图不放 OLE（见下）  
- `options.py` — 用户可见 dataclass  

不要在 `writer` 里堆预览、VML、OLE 二进制细节。

### 图片模式 `_picture.py`

`insert_pictures_many()` 与 OLE 路径并列：媒体 + drawing 的 `xdr:pic` 锚点，
不产生 embeddings / VML / oleObjects。上千个 OLE 对象会拖垮 Excel（每个都带
CFBF 容器、VML 形状和激活机制），图片则很轻，单表可装几千个结构。
代价：不能双击进 ChemDraw 编辑。锚点按目标 sheet 真实网格计算
（复用 `_grid.py` 的 `ExcelGridMetrics`），目标已有 drawing 时追加进去。

## 内部子包

### 依赖方向（硬规则）

```text
writer → _insert (pipeline) → _insert_ole / _shell / _picture / _parts
                       ↘ models ↙
写入侧的包（_insert_ole / _shell / _picture / _parts）不得 import _insert。
excel 与 powerpoint 互不相识；两者共用的东西放 `office/` 层，
**不允许**跨格式 import 对方的 `_` 私有模块。
```

`models.py`（无 `_` 前缀，**仍是内部**）：

- `PreviewSpec` / `LayoutSpec` / `NormalizeSpec` / `ShellSourceOptions` / `SourceBuildRequest`

`InsertPlan`（`_insert/plan.py`）组合上述 spec，executor 按路径取用。

### `_insert/`

编排层：把用户参数变成计划并执行插入。

- `plan.py` — 构建 `InsertPlan`
- `cdx.py` — `StructureSource` → 临时 `.cdx`（含 optional normalize）
- `executor.py` — 写 CDX → 建 source → merge → 修正 layout / content types

### generated 路径 — `_ole_template.py` + `_insert_ole.py`

默认路径。模板是**零件供体**，不是工作目录：`ExcelOleTemplate` 一次加载取出
容器字节、`<oleObject>` 声明、drawing 锚点、VML shape，以及新建 drawing/VML
部件用的空外壳；`insert_ole_objects` 把它们克隆进目标包。

**为什么克隆而不是拼**：早期版本自己拼这些 XML，Excel 打开后修复文件并删掉
OLE 对象（`docs/developer_log.md` 2026-03-06）。有模板就不该手写任何 OOXML。

**顺序要求**：锚点由 `_grid.ExcelGridMetrics` 按目标表真实行高列宽算出，
所以 `touch_sheet_layout` 必须在插入**之前**跑（见 `_insert/executor.py`）。

### `_shell/` — 原生 ChemDraw OLE shell（单对象 source）

与 generated 的 B–D 对齐，但改的是「真·Excel+ChemDraw」模板：

| 子步骤 | 做什么 | 模块 |
|--------|--------|------|
| B | 渲染预览写入 `xl/media` | `preview.py` |
| C | 更新 vml/sheet image rels | `media_rels.py` |
| D | 框贴合分子 / 挪格 / 改大小 | `geometry.py` |
| chrome | 去黑边框 | `chrome.py` |

- `preview.py` — `replace_shell_preview`（**仅** B+C：画预览 + 改 image rels）  
- `geometry.py` — `tighten_shell_display_box`、`patch_shell_geometry`（`read_vml_size_px` 定义在 `vml.py`）  
- `chrome.py` — 仅 `remove_shell_ole_border`  
- `media_rels.py` — image rId / Target  
- `staging.py` — 编排入口 `build_shell_staging` 与 `_fit_cdx_to_shell`  
- 底层：`vml` / `ooxml_patch` / `media_rels`（`_grid` / `_model` 在 excel 顶层，三条路径共用）  

**Shell 路径的编排在 `_shell/staging.py`**（顺序显式）：  
B 写 bin → D' tighten → chrome → B+C 预览 → D 挪格 → chrome → pack。

注意：RDKit 路径需 `.[chem]`（rdkit + pycdxml）。EMF 仍需 Inkscape。

### `_parts/` — 把 shell 的零件搬进目标表

只服务 shell 路径：staging 目录里已是一套自洽的 OLE 零件，这里负责搬运并
解决与目标表已有内容的冲突（rId 分配、shapeId 重映射、drawing/VML 合并）。

```text
staging 目录（已含 bin + 图 + drawing/vml/oleObjects）
        │
        ▼  merge_source_workbook
目标 package_dir
```

两个模块**三条路径共用**，不要跟着 shell 一起挪：

- `util.py` — `next_rid` / `unique_part_name` / `write_xml`
- `anchor_overflow.py` — executor 第 7 步，把越界偏移走进真实格子
  （generated 直写路径的锚点本来就在格内，对它是 no-op）

子步骤与文件：

| 子步骤 | 做什么 | 模块 |
|--------|--------|------|
| B 落盘零件 | 拷 `embeddings/*.bin`、`media/*` | `context.py` |
| C 接线 | sheet/vml `.rels` 分配 rId | `context.py` |
| D 钉框 | 合并 drawing.xml / vmlDrawing.vml | `surface.py` |
| E 挂 OLE | 改 shapeId；`attach_ole_objects`（`_ooxml/worksheet.py`） | `ole_objects.py` + `merge.py` |
| 编排 | 串起上面顺序 | `merge.py` |

- `merge.py` — 唯一入口 `merge_source_workbook`
- `context.py` — 目标目录、拷贝 bin/media、写 rels
- `surface.py` — drawing / VML / shapeId
- `ole_objects.py` — worksheet `<oleObjects>` 辅助
- `util.py` — rid / 文件名 / xml 小工具

### `_ooxml/`

负责底层 Open XML / XLSX 操作。


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
  解析 Excel 单元格地址，处理列号转换（A1 记法是表格专属，所以留在这里）。

单位换算已上移到 `office/_units.py`（EMU/px/pt 是 OOXML 通用的，PPT 也要用）；
`_utils` 仍 re-export，内部 import 不受影响。

## 推荐依赖方向

尽量保持单向依赖：

```text
writer
  -> _insert
  -> _insert_ole / _shell / _picture / _parts
  -> _shell
  -> _grid / _model        （网格度量与锚点值对象）
  -> _ooxml / _utils
  -> office/_package, office/_units   （跨格式共享）
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
from retain_cdx.office.excel._shell import build_shell_staging
from retain_cdx.office.excel._parts import merge_source_workbook
from retain_cdx.office._package import extract_ooxml, pack_ooxml
```

不要在 `writer.py` 里到处导入具体文件：

```python
# 尽量避免
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
