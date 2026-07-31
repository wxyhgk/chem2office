# Excel/OOXML 深度研究报告

> 2026-07-26 · 三路并行实证审计：锚点/几何模型、OLE/CFBF 容器、真实工作簿鲁棒性。
> 所有结论均有可复现证据（解剖内置模板/夹具、构造对抗性工作簿、逐字节 diff
> Windows Excel 原生产物）。复现脚本与样本存于当次会话 scratchpad。

## 一、确立的基础事实（ground truth）

### 1. 合法的 ChemDraw OLE 容器长什么样

内置模板与 `shell_large.xlsx` 的 `oleObject1.bin` 结构一致：4096 字节 CFBF v3，
**三个流全部在 mini-stream**：

| 流 | 大小 | 作用 |
|---|---|---|
| `\x01Ole` | 20 B | OLE 2.0 版本标记 + flags=0（嵌入非链接），无 Moniker |
| `\x01CompObj` | 126 B | AnsiUserType `"CS ChemDraw Drawing"`、剪贴板格式、ProgID `"ChemDraw.Document.6.0"`、Unicode 标记 `0x71B239F4` |
| `CONTENTS` | ~1.4 KB | 裸二进制 CDX（`VjCD0100` 魔数） |

- 根条目 CLSID `41BA6D21-A02E-11CE-8FD9-0020AFD1F20C`（跨 ChemDraw 版本稳定）。
- ProgID 里的 "6.0" **不是应用版本号**——v6 以来所有 ChemDraw 都注册这个固定名。
- **Excel 的容器内没有 `\x02OlePres000` 演示流**：激活前的视觉完全由
  OOXML 层承担（`xl/media` + VML imagedata + `objectPr r:id`），我们的
  预览替换模型是正确的。
  **注意**：这条只对 Excel 成立——PowerPoint 的 ChemDraw 容器带演示流，
  见 `powerpoint_ole_research.md`。
- 目录树红黑着色、header 常量与 Windows 产物逐位一致（根节点为红，虽违背
  spec 的 SHOULD，但以样本为准）。

### 2. 列宽像素模型的真相（MDW）

ECMA-376 列宽以工作簿 normal 字体的 **max digit width（MDW）** 为单位：

- Calibri 11 → MDW = 7px；**等线/DengXian 11 → MDW = 8px**。
- 内置模板本身就是等线工作簿：`width="21.375"` = (176−5)/8 整除、
  VML `margin-left:88.5pt` 反推默认列 = 72px = roundup8(8×8+5)。
- `defaultColWidth` / `<col width>` 已含 +5 padding：px = width×MDW+5。
- `baseColWidth` **不含** padding：px = baseColWidth×MDW+5 再
  **向上取整到 8 的倍数**（8 字符 → 64px，不是 61px）。

### 3. 锚点三元组语义

worksheet `<oleObject><anchor>` / drawing `xdr:twoCellAnchor` / VML `x:Anchor`：

- 我们的输出三处逐 EMU 一致（比 Excel 自己的文件还严格——Excel 原生文件
  三处间有 ≤0.4px 漂移，且 VML 与 sheet anchor 对齐）。
- 分工：VML 负责渲染/打印预览（xdr 的 `sp` 是 `hidden="1"` 占位），
  objectPr anchor 是激活/重排时的权威框，xdr 喂现代绘图层/选择。
- 语义组合（自 Excel 原生模板复制，勿"修正"）：`editAs="oneCell"`（随格移动
  不随格缩放）+ objectPr `moveWithCells="1"` + VML **仅存在 `<x:SizeWithCells/>`**。
  VML 标志极性与元素名反直觉：move-yes/size-no 编码为只写 SizeWithCells。
  网上文档互相矛盾，以夹具证据为准。

## 二、发现清单（按严重程度）

### 🔴 会产生损坏/丢数据（应尽快修）

| # | 问题 | 位置 | 证据 |
|---|------|------|------|
| ~~C1~~ **已修 0.0.10** | CDX > 63,488 B 时 `build_chemdraw_ole` 静默产出损坏容器（FAT 只写 1 扇区，csectfat 硬编码 1） | `ole/cfbf.py` L192-297 | 63,488 B 恰好可过，63,489 B → olefile "FAT sector index out of range"；strict open 仍成功，读 CONTENTS 才炸——复现历史"Excel 静默删对象"故障类 |
| ~~C2~~ **已修 0.0.10** | 就地 patch 跨 4096 mini-cutoff 损坏流（≥4096 模板注入 <4096 CDX 后 size 落入 mini 区间，start_sect 被按 mini-sector 解析） | `ole/embed.py` | 实测 8000B 容器注入 1342B → 读回空/垃圾，真实数据孤儿化在 regular sector 4。当前模板不可达；修 C1 后变可达 |
| ~~C3~~ **已修 0.0.10** | `baseColWidth` 换算 61px 应为 64px（缺 roundup-to-8） | `_shell/grid.py` from_worksheet | 双锚点 B2→E8 实测 to colOff=183px，应 192px，框短 9px。openpyxl/Calibri 工作簿今天就错 |
| ~~C4~~ **已修 0.0.10** | MDW 硬编码 7（Calibri），等线等 CJK 默认工作簿每默认列漂 8px、跨列框宽 ~8% 拉伸 | `_shell/grid.py`、`_source/template/vml.py` | 内置模板自己就是 MDW=8 的等线工作簿（见基础事实 2） |
| ~~C5~~ **已修 2026-07-30** | 真实 Excel 表插入时 CT_Worksheet 子元素顺序错（`pageMargins+1` 启发式）；worksheet 级 `mc:AlternateContent` 包裹的 `<oleObjects>` 找不到 → 造出第二个 `<oleObjects>`（maxOccurs=1 违规） | `_ooxml/worksheet.py` attach_ole_objects | 含 pageSetup/headerFooter/tableParts 的表 → drawing 插到 pageSetup 前；outer-AC 表 → 双 oleObjects + 顺序错乱 |

> **C5 后记**：这是本报告中唯一在真实生产数据上兑现的条目。2026-07-30 转换
> 两个 2812 行 SMILES 工作簿（原文件带 `pageSetup`）时，产物 Excel 打开即触发
> 修复提示。修复：新增 `_ooxml/xml_io.insert_worksheet_child()` 持有完整
> CT_Worksheet 序列并按秩位插入，替换 `drawing`/`legacyDrawing`/`oleObjects`
> 三处位置启发式；同时支持在 `mc:AlternateContent/mc:Choice` 内查找已有
> `<oleObjects>`。回归测试见 `tests/test_worksheet_schema_order.py`。
> 教训：位置启发式（"插在某元素后面"）在 OOXML 里一律不可靠，必须按 schema 序列。
| ~~C6~~ **已修 0.0.10** | 目标表的外部关系（超链接等）被重建时丢 `TargetMode="External"` → 链接损坏/Excel 修复 | `_ooxml/relationships.py` read_rels/build_rels 只传 (Id,Type,Target) | `https://example.com/x` 被解析为内部部件路径 |
| ~~C7~~ **已修 0.0.11** | 目标表已有图表/批注时，新拷贝的 drawing/VML 部件 shapeId 不重映射 → OLE↔VML 绑定悬空；追加进批注 VML 时缺 `_x0000_t75` shapetype 定义 | `_parts/surface.py` merge_drawing_and_vml（created_* 分支跳过 rewrite） | chart 表：oleObject shapeId=1026 但 VML 只有 _x0000_s1025；comments 表：t75 未定义 |
| ~~C8~~ **已修 0.0.11** | `set_cell_texts` 覆盖单元格时丢 `s=` 样式属性；覆盖公式格后 calcChain.xml 残留 → "Removed Records" 修复 | `_ooxml/cells.py` | 粗体/填充实测消失；calcChain 仍引用 B2 |
| ~~C9~~ **已修 0.0.11** | `.xlsm`：`Default bin → vbaProject` 已存在时嵌入的 oleObject*.bin 被判为 VBA 工程（content type 冲突未用 Override 解决） | `_ooxml/content_types.py` | 实测输出无嵌入 Override。（好消息：vbaProject/customXml 字节级保留） |

### 🟠 高风险/易碎

| # | 问题 | 位置 |
|---|------|------|
| ~~F1~~ **已修 0.0.10** | generated 路径删除 `autoPict="0"`——ECMA-376 默认 true=按预览图原始尺寸调整对象，而预览图被我们 4× 超采样，激活/重排时框有 4× 放大风险 | `_parts/ole_objects.py` normalize_object_pr |
| ~~F2~~ **已修 0.0.11** | 同格溢出 `to` 标记（col/rowOff 超过格子尺寸）：Excel 按 origin+offset 容忍并在保存时归一化，但第三方消费者可能 clamp；Excel 原生文件从不写这种标记。**注意**：需先修 C3/C4 再做归一化，否则现在的字体无关字面偏移反而是保护 | `_source/anchors.py` compute_anchors |
| ~~F3~~ **已修 0.0.11** | `mc:Ignorable` 被整体覆盖为硬编码 5 个 token，目标已有的（如 xr6）不再声明 → 严格 MCE 消费者必须拒绝 | `_ooxml/xml_io.py` ensure_sheet_compat_namespaces |
| ~~F4~~ **已修 0.0.11** | `set_column_width` 遇 Excel 分组列区间（`min=1 max=10`）产生重叠 `<col>` 定义，哪个生效由消费者决定 | `_ooxml/sheet_layout.py` _find_exact_col |
| ~~F5~~ **已修 0.0.10** | `embed.py` 用 `except SystemExit` + 报错字符串匹配做重建回退的控制流 | `ole/embed.py` L29 |

### 🟡 表面性/低危

- generated 路径 VML 绝对 `style`（margin-left/top）用源模板常量+均匀行高的"幻想网格"计算（实测 margin-top:110pt vs 真实 15pt）；Excel 用 `x:Anchor` 所以无感，第三方按 style 摆放会错位。shell 路径已正确（用 grid metrics）。
- generated VML 强制 `stroked="t"`（shell 路径与 Excel 模板都是 `stroked="f"`）。
- ~~重建 miniFAT 空闲项填 0x00 而非 FREESECT~~ **已修 0.0.10**。
- `grid.py` 忽略 hidden 行列；`int()` 截断 vs Excel 四舍五入 ±1px。
- `read_vml_size_px` 正则要求 pt 单位且 width 在 height 前。
- 单 token `<dimension ref="A1"/>` 扩张时被忽略（Excel 视 dimension 为提示性，低危）。
- `pack_xlsx` 的 `[Content_Types].xml` 恰好是第一个 zip 条目只是 `os.walk` 的巧合；"必须第一"对 Excel/LO 是误传（都读 central directory），但确定性排序是廉价保险。

### ✅ 验证过没问题的

- 就地 patch 路径：无扇区孤儿、无上一分子数据泄漏（zero-pad 到旧长度）。
- `\x01Ole`/`\x01CompObj`/CLSID/header 与 Windows 产物逐字节一致。
- 锚点三元组逐 EMU 一致；行高（pt 基）换算字体无关且正确。
- 插入自己产出的文件、Excel inner-AC 模板、多 sheet + definedNames +
  非连续 sheet 部件名、Override 风格 content types、mergeCells 共存——全部干净。
- LibreOffice headless 对所有输出可开（但它连硬 schema 违规都容忍，
  只能作必要非充分 oracle；且 LO 会丢 ChemDraw OLE，不能做保真检验）。

## 三、修复路线图（全部完成）

> C1–C9 与 F1–F5 已全部修复（0.0.10 / 0.0.11）。以下为当初的建议顺序，保留作记录。


1. **OLE 容器双修**：C1 多 FAT 扇区支持（DIFAT ≤109 扇区可覆盖 ~6.8MB）+
   C2 cutoff 穿越守卫（`(旧≥4096) != (新≥4096)` 即重建）+ 顺手 F5、FREESECT。
2. **网格度量双修**：C3 baseColWidth roundup + C4 从 `styles.xml` normal
   字体解析 MDW（Calibri→7、等线→8，未知回退 7）。两处（grid.py、
   template/vml.py）同步。
3. **worksheet 手术**：C5 schema-index 插入器（持有 CT_Worksheet 全序列，
   MCE-aware）+ AC 内查找 oleObjects。
4. **rels 保真**：C6 read_rels/build_rels 携带全部属性（尤其 TargetMode）。
5. **merge 对称性**：C7 created_* 分支同样跑 shapeId 重映射 + shapetype 补拷。
6. **cells/CT 收尾**：C8 保留 `s=` + 删 calcChain；C9 `.xlsm` 用 Override；
   F1 保留 `autoPict="0"`；F3 Ignorable 求并集；F4 列区间拆分。
7. **最后**：F2 溢出锚点归一化（必须在 2 之后做）。
