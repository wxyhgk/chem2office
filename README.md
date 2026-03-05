# ChemDraw OLE 注入（Excel）

这个仓库实现了 ChemDraw 的 “OLE 模板注入” 流程，在 Linux/macOS 下批量生成
可编辑的 ChemDraw OLE。默认模式会直接复用仓库内置的 `template.xlsx`
作为稳定骨架；如果显式传入 `--template`，则使用你指定的模板文件。

## 为什么默认仍然依赖模板

Excel 的 OLE 对象存储在微软 CFBF 容器里，ChemDraw 使用的是二进制 CDX 流。
虽然仓库里保留了“从零生成”的代码路径用于自定义定位，但默认输出为了兼容
Excel，仍然采用模板注入方式：

1) 用 Windows Excel 插入 ChemDraw 对象，保存模板。
2) 替换 OLE 的 `CONTENTS` 流为新的 CDX。
3) 更新预览图，让 Excel 显示正确结构。

## 文件说明

- `inject_cdx_into_xlsx.py` - 入口脚本（兼容旧用法，内部调用 `scripts/`）。
- `scripts/` - 拆分后的模块化脚本目录。
  - `scripts/cli/` - CLI 入口
  - `scripts/ole/` - OLE 注入逻辑
  - `scripts/render/` - 预览图生成
  - `scripts/xlsx/` - XLSX 结构处理
- `template.xlsx` - 默认使用的内置模板（单行、已包含一个 ChemDraw OLE 对象）。
- `template/` - `template.xlsx` 解包后的静态参考文件。

## 依赖

- Python 3
- `olefile`（`pip install olefile`）
- Open Babel 命令行（`obabel`，用于生成预览 PNG）

## 用法

推荐用法（默认，直接使用内置 `template.xlsx`）：

```bash
python3 inject_cdx_into_xlsx.py --cdx san.cdx --output output.xlsx
```

也可以直接运行模块入口：

```bash
python3 -m scripts.cli.inject_cdx_into_xlsx --cdx san.cdx --output output.xlsx
```

显式指定模板：

```bash
python3 inject_cdx_into_xlsx.py --template template.xlsx --cdx san.cdx --output output.xlsx
```

多行（每行一个 OLE，默认也走内置模板）：

```bash
python3 inject_cdx_into_xlsx.py --cdx 1.cdx 2.cdx --output output_2rows.xlsx
```

脚本会自动：
- 把文件名（去后缀）写入 A 列。
- 在 B 列放入对应的 OLE 对象。
- 用 Open Babel 生成预览 PNG 并绑定到对应 OLE。

## 工作原理（默认路径）

1) 解压 `.xlsx`（本质是 ZIP）。
2) 对每个 CDX：
   - 复制模板里的 OLE 文件（`xl/embeddings/oleObject*.bin`）。
   - 替换 `CONTENTS` 流为新的 CDX 数据。
   - 修补 OLE 目录项的流长度（`olefile` 不支持变长写入）。
3) 生成 PNG 预览图并更新关联关系：
   - `xl/media/image*.png`
   - `xl/worksheets/_rels/sheet1.xml.rels`
   - `xl/drawings/_rels/vmlDrawing1.vml.rels`
4) 在 `sheet1.xml` 和 `vmlDrawing1.vml` 中复制/更新锚点。
5) 重新打包为 `.xlsx`。

说明：

- 默认布局下，CLI 现在内部直接复用 `template.xlsx` 的稳定工作流。
- 只有在显式传入定位参数（如 `--col`、`--from`、`--to`）时，才会走从零构造
  XLSX 骨架的路径。

## 限制

- CDX 大小必须 <= 模板中 OLE 的 `CONTENTS` 流长度（脚本会补零并修正长度）。
- 模板必须由 Windows Excel + ChemDraw 插入生成。
- 当前脚本假设模板结构类似 `template.xlsx`（简单一行一对象布局）。

## 备注

- 之前默认生成模式和模板模式不是同一路径，可能出现 Excel 打开后修复并删除
  OLE 对象的情况。
- 当前默认模式已经改为复用内置模板，以避免这类兼容性问题。

## 常见问题（FAQ）

**1) Excel 提示 “被替换的部件：/xl/worksheets/sheet1.xml 有 XML 错误”**

- 旧版本默认生成模式和模板模式不一致，可能导致 Excel 删除对象记录。
- 当前默认模式已经回到内置模板工作流；请直接使用：

```bash
python3 inject_cdx_into_xlsx.py --cdx san.cdx --output output.xlsx
```

**2) OLE 对象显示了，但双击打不开**

- 最常见原因是 CDX 不是 ChemDraw 原生导出的格式。
- 建议用 ChemDraw 导出 `.cdx` 再注入，或用 Excel 中插入的 OLE 作为模板。

**3) 报错 “CDX is larger than existing stream”**

- OLE 的 `CONTENTS` 流不能变长写入。
- 解决方法：用更“大”的模板（OLE 中 CDX 内容更复杂时流更大），
  或选择更小的 CDX。

**4) 预览图没有更新**

- 检查 `xl/media/` 下是否生成了 `image*.png`。
- 确认 `sheet1.xml.rels` 和 `vmlDrawing1.vml.rels` 里指向的是 PNG。
