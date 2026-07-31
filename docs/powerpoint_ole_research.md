# PowerPoint ChemDraw OLE 研究报告

> 2026-07-31 · 解剖 `ben-test.pptx`（Windows PowerPoint + ChemDraw 插入的苯环
> OLE 对象），确立 PPT 侧的落地模型，并与已经吃透的 Excel 侧对照。

## 一、样本结构

37.5 KB，与 OLE 相关的只有三处（**没有 VML**）：

```
ppt/slides/slide1.xml            2349 B   graphicFrame + oleObj
ppt/slides/_rels/slide1.xml.rels  594 B   oleObject + image + slideLayout
ppt/embeddings/oleObject1.bin   12288 B   CFBF 容器
ppt/media/image1.emf             5396 B   矢量预览
```

`[Content_Types].xml`：`Default bin → oleObject`、`Default emf → image/x-emf`。

## 二、幻灯片里的 OLE 标记

```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="4" name="对象 3"><a:extLst>…creationId…</a:extLst></p:cNvPr>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr>
    <p:nvPr><p:extLst>…p14:modId…</p:extLst></p:nvPr>
  </p:nvGraphicFramePr>
  <p:xfrm><a:off x="5910263" y="3217863"/><a:ext cx="369887" cy="419100"/></p:xfrm>
  <a:graphic><a:graphicData uri=".../presentationml/2006/ole">
    <mc:AlternateContent>
      <mc:Choice Requires="v">
        <p:oleObj name="CS ChemDraw Drawing" r:id="rId2"
                  imgW="370294" imgH="419073" progId="ChemDraw.Document.6.0">
          <p:embed/>
        </p:oleObj>
      </mc:Choice>
      <mc:Fallback>
        <p:oleObj …同上…>
          <p:embed/>
          <p:pic>…<a:blip r:embed="rId3"/>…<a:xfrm>同 p:xfrm</a:xfrm>…</p:pic>
        </p:oleObj>
      </mc:Fallback>
    </mc:AlternateContent>
  </a:graphicData></a:graphic>
</p:graphicFrame>
```

**与 Excel 的关键差异（几乎全是好消息）：**

| | Excel | PowerPoint |
|---|---|---|
| 定位 | 三套锚点（worksheet/drawing/VML）必须一致 | **绝对 EMU 坐标**，`p:xfrm` 一处（Fallback 的 `p:pic` 再重复一次） |
| 网格依赖 | 列宽像素公式依赖 normal 字体 MDW，是一整类 bug 的来源 | **完全没有**——不看行列 |
| VML | 必需，且标志极性反直觉 | **不存在** |
| 预览引用 | VML imagedata + `objectPr r:id` | `p:pic/a:blip r:embed`（在 mc:Fallback 里） |
| 元素顺序 | CT_Worksheet 固定序列，插错就触发修复 | `p:spTree` 里追加即可，无顺序陷阱 |

也就是说 Excel 侧最痛的几类问题（MDW 字体依赖、baseColWidth 取整、锚点三元组一致性、CT_Worksheet 顺序、VML 层）在 PPT 侧**全部不存在**。

`imgW`/`imgH` 是对象固有尺寸（EMU），`p:xfrm/a:ext` 是显示尺寸；样本里二者接近但不等
（370294×419073 vs 369887×419100），说明 PowerPoint 允许缩放显示。

## 三、OLE 容器：与 Excel 同源，但多一个演示流

| 流 | PPT | Excel 模板 | 结论 |
|---|---|---|---|
| `\x01CompObj` | 126 B | 126 B | **逐字节相同** |
| `\x01Ole` | 20 B | 20 B | **逐字节相同** |
| `\x02OlePres000` | 6270 B | 无 | **仅 PPT 有** |
| `CONTENTS` | 1489 B | 1489 B | 同长不同分子 |

根 CLSID `41BA6D21-A02E-11CE-8FD9-0020AFD1F20C`、ProgID `ChemDraw.Document.6.0`
与 Excel 完全一致 → **`retain_cdx/ole` 的容器构建逻辑对 PPT 直接可用**。

`\x02OlePres000` 头部：剪贴板格式 = **14 (CF_ENHMETAFILE)**，偏移 142 处是 ` EMF`
签名，即内嵌一份 EMF 渲染。注意它与 `ppt/media/image1.emf`（5396 B）**不是同一份**
（6168 B），是两套独立渲染。

> 修正之前的 Excel 研究结论：「ChemDraw OLE 不带演示流」只对 Excel 成立。

## 四、已发现并修复的缺陷

> 实现 `PowerPointCDXWriter` 时又踩到两个，都已修 + 有回归测试（见 §七）。

**多扇区目录流**（`ole/inject.py`）— PPT 容器有 4 个流、12288 B，目录流跨扇区，
而 `patch_direntry_size` 硬算 `(first_dir_sector+1)*sector_size + sid*128`，直接报
"Directory stream spans multiple sectors; size patch not supported" 拒绝工作。
已改为沿 FAT 走目录链定位条目。实测注入 `examples/1.cdx`、`san.cdx` 后：
CONTENTS 正确替换、`OlePres000`/`CompObj` 完好、olefile strict 可读，Excel 侧 47 个测试不受影响。

## 五、待解决

1. **`OlePres000` 会变陈旧**（severity: 未知，需真 PowerPoint 验证）。注入新分子后演示流
   仍是旧分子（苯）的 EMF。三个选项：(a) 原样保留；(b) 用我们生成的 EMF 重写演示流
   （需按 MS-OLEDS 组装头部）；(c) 整个删掉（Excel 侧根本没有这个流且工作正常）。
   PowerPoint 显示用的是 `p:pic` 的 media EMF，所以日常视图大概率不受影响，
   但复制粘贴/旧版本路径可能用到。**这条必须在真 PowerPoint 上验证。**
2. **重建路径会丢演示流** — `build_chemdraw_ole` 只写三个流；一旦 CDX 超出原
   CONTENTS 触发重建，`OlePres000` 就没了（与选项 c 等价，但是意外而非有意）。
3. Excel 研究里的 C1（CDX > 63,488 B 时 FAT 只写一扇区导致静默损坏）、
   C2（跨 4096 mini-cutoff）同样适用于 PPT 容器，且 PPT 容器更大更容易触发。

## 六、实现路线

复用现成的：`retain_cdx/ole`（容器）、`retain_cdx/render`（EMF/PNG 预览）、
`retain_cdx/structure`（SMILES/CDX）。**只有 PresentationML 那层是新的**，
而且比 Excel 简单得多。

建议与 Excel 侧同构，走模板注入：

1. 把 `ben-test.pptx` 清理后作为内置模板（`retain_cdx/data/template.pptx`）——
   它提供 Windows PowerPoint + ChemDraw 原生产出的 `graphicFrame` 骨架和容器。
2. `retain_cdx/office/powerpoint/` 下实现 `PowerPointCDXWriter`：
   - `insert_cdx(source, slide=1, position=..., size=...)` — 克隆 graphicFrame、
     换 `p:xfrm`、注入 CDX、换 media EMF、分配 rId
   - 多对象：网格/列表布局辅助
3. 定位 API 用绝对长度（EMU/mm/pt/in），不需要 Placement 那套单元格语义。

**以上均已实现**（`retain_cdx/office/powerpoint/`）。

## 七、实现期间踩到的两个坑

### 1. `mc:Choice Requires="v"` 的前缀声明被 ElementTree 丢弃 —— PowerPoint 报修复

模板里是 `<mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v">`。
`v` 前缀**没有任何元素使用**（PPT 侧本来就没有 VML 内容），那条声明纯粹是给
`Requires` 查的「前向声明」。ElementTree 只为实际用到的前缀输出 `xmlns`，
于是序列化后只剩 `Requires="v"` 指向一个未声明前缀 —— MCE 硬违规。

五行可复现：

```python
src = '<root xmlns:mc="urn:mc"><mc:Choice xmlns:v="urn:vml" Requires="v"/></root>'
ET.tostring(ET.fromstring(src))   # → <mc:Choice Requires="v" />，xmlns:v 没了
```

修：`_pml.write_slide_xml()` 序列化后把 `xmlns:v` 补回根元素
（与 Excel 侧 `write_sheet_xml` 补 `x14ac`/`xr` 同一手法）。

**普遍教训**：OOXML 经 ElementTree 往返时，「声明了但没被元素使用」的命名空间会
静默消失，而 MCE 恰恰依赖这类声明。Excel 研究里的 F3（`mc:Ignorable` 被覆盖）
是同一失效模式。凡是碰 `mc:` 的地方都要手动兜。

### 2. 克隆帧里的 `cNvPr id="0"` 重复

`mc:Fallback` 里的预览图 `p:pic` 在模板中是 `id="0"`，逐个克隆后一张幻灯片上
出现 N 个 id=0。已改为克隆时给帧内所有 `cNvPr` 重新编号。

### 3. 模板样例对象的孤儿部件

最初只从 `p:spTree` 删掉 graphicFrame，模板自带的苯环 `.bin`/`.emf` 和对应
relationships 仍留在包里 —— 产物里多出一个「幽灵结构」。`clear_ole_frames()`
现在连带清理 rels 与部件。

## 八、回归防线

`tests/ooxml_checks.py` 提供包级校验（Excel/PPT 通用）：relationship 悬空、
内容类型缺失、shape id 重复、以及**所有 `Requires`/`Ignorable` 前缀是否真的声明**。
上述 1、2 两条经验证——回退任一修复，测试即变红。
