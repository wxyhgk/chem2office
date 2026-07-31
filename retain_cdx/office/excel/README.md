# Excel 模块结构说明

这个目录负责把可编辑的 ChemDraw CDX OLE 对象写入 `.xlsx` 文件。

概念上两步：**造零件**（bin + 预览，可外置）→ **装进 Excel**（本目录主责）。  
实现上仍是：计划 → 临时 source 小表 → merge 进目标 package。

## Public API vs `_` 内部（契约对照）

带 `_` 的包是**工业包常见的内部实现标记**：可随时重构路径与文件，**不保证** import 稳定。  
请只依赖下表「大门」一侧。

| 层级 | 可依赖？ | 符号 / 路径 | 说明 |
|------|----------|-------------|------|
| **大门 Public** | ✅ 稳定 | `retain_cdx.ExcelCDXWriter` | 写 workbook 的唯一推荐入口 |
| | ✅ | `Placement`, `PreviewOptions`, `DisplayOptions`, `ShellTemplate` | 用户配置 |
| | ✅ | `retain_cdx.StructureSource` | 化学结构（excel 外的领域对象） |
| **员工通道 Internal** | ⚠️ 勿当 API | `office/excel/models.py` | 内部 DTO（`PreviewSpec` / `SourceBuildRequest`…） |
| | ⚠️ | `_insert/` | 编排：plan → CDX → source → merge |
| | ⚠️ | `_source/` | 临时 source.xlsx 组装（generated / shell） |
| | ⚠️ | `_source/template/` | generated 模板的 XML 细节抽屉 |
| | ⚠️ | `_parts/` | source → 目标 package 的 merge |
| | ⚠️ | `_shell/` | 原生 ChemDraw shell 几何/预览 |
| | ⚠️ | `_ooxml/` / `_utils/` | ZIP/XML/单位底层 |

```text
用户代码
   │  只应 import 大门
   ▼
writer + options          ← public
   │
   ▼
_insert  →  _source  →  _parts   ← _ 内部（装进 Excel）
              │            │
         template/      surface/rels
         (XML 细节)     (merge 细节)
```

**semver 含义**：改 public 签名要谨慎；改 `_` 包 / `models` 只要测试绿、public 行为不变即可。

### `__init__.py` / `writer.py` / `options.py`

- `__init__.py` — 只 re-export public 符号  
- `writer.py` — 薄门面：计划 → 执行 → save；另提供
  `set_cell_text(s)` / `set_column_width` / `set_row_height`
  （写结构旁边的 ID/表头等文本、调格子尺寸，
  **不要**用 openpyxl 后处理 —— 它会丢 OLE parts）  
- `insert_pictures_many()` — **图片模式**：只放图不放 OLE（见下）  

### 图片模式 `_picture.py`

`insert_pictures_many()` 与 OLE 路径并列：媒体 + drawing 的 `xdr:pic` 锚点，
不产生 embeddings / VML / oleObjects。上千个 OLE 对象会拖垮 Excel（每个都带
CFBF 容器、VML 形状和激活机制），图片则很轻，单表可装几千个结构。
代价：不能双击进 ChemDraw 编辑。锚点按目标 sheet 真实网格计算
（复用 `_shell/grid.py` 的 `ExcelGridMetrics`），目标已有 drawing 时追加进去。
- `options.py` — 用户可见 dataclass  

不要在 `writer` 里堆预览、VML、OLE 二进制细节。

## 内部子包

### 依赖方向（硬规则）

```text
writer → _insert (pipeline) → _source / _parts
                ↘ models ↙
_source 与 _parts 不得 import _insert。
```

`models.py`（无 `_` 前缀，**仍是内部**）：

- `PreviewSpec` / `LayoutSpec` / `NormalizeSpec` / `ShellSourceOptions` / `SourceBuildRequest`

`InsertPlan`（`_insert/plan.py`）组合上述 spec，并提供 `to_source_request()`。

### `_insert/`

编排层：把用户参数变成计划并执行插入。

- `plan.py` — 构建 `InsertPlan`
- `cdx.py` — `StructureSource` → 临时 `.cdx`（含 optional normalize）
- `executor.py` — 写 CDX → 建 source → merge → 修正 layout / content types

### `_source/` — 在临时小表里先装齐 OLE（merge 之前）

**角色**：产出 `source.xlsx`，里面已经自洽地含有 `.bin` + 预览 + drawing/vml/oleObjects。  
随后由 `_parts` merge 进用户目标表。  
（化学格式转换在更外层；这里从 CDX 路径开始。）

```text
CDX 列表 + SourceBuildRequest
        │
        ▼  builder（shell? → shell 路径 : generated 路径）
source.xlsx  ──merge──►  目标 package
```

#### Generated 路径子步骤

| 子步骤 | 做什么 | 模块 |
|--------|--------|------|
| 几何 | 每行固定框 anchors | `anchors.py` |
| B 落盘 | 克隆 OLE 壳、写 CDX、渲染预览 | `fill_parts.py` |
| C–F | rels + drawing/vml + oleObjects + content types | `wire_surface.py` + `template/` |
| 编排 | 解包模板 → B → C–F → 打包 | `generated.py` |

- `builder.py` — 只认 `SourceBuildRequest`，分发 shell / generated  
- `workbook.py` — shell 组装；generated 转调 `generated.py`  
- `fill_parts.py` — **B** embeddings + media  
- `wire_surface.py` — **C–F** 接线与表面 XML  
- `template/` — generated 模板 XML 细节（common / namespaces / objects / sheet / vml / xml）  
- `anchors.py` — 两格锚点计算  

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
- 底层：`model` / `grid` / `vml` / `ooxml_patch`  

**Shell 路径的编排在 `_source/workbook.py`**（顺序显式）：  
B 写 bin → D' tighten → chrome → B+C 预览 → D 挪格 → chrome → pack。

注意：RDKit 路径需 `.[chem]`（rdkit + pycdxml）。EMF 仍需 Inkscape。

### `_parts/` — 「装进 Excel」（当前实现）

**概念上**这一步只做一件事：把已有的 `.bin` + 预览图装进目标 xlsx。  
**实现上**零件先在临时 source 工作簿里装齐，再 merge 进目标 package（不负责 CDX/预览生成）。

```text
source.xlsx（已含 bin + 图 + drawing/vml/oleObjects）
        │
        ▼  merge_source_workbook
目标 package_dir
```

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
