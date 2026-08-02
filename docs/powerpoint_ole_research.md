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

1. ~~**`OlePres000` 会变陈旧**~~ / ~~**重建路径会丢演示流**~~ —— 已解决，见 §九。
   当时列的三个选项里选了 (b)：用我们自己的渲染重写演示流。丢失确实有后果：
   缺了它，对象双击正常，但**从 PPT 复制出去只会得到一张图片**，不能双击、
   也不能再粘进 Excel。日常视图不受影响的判断是对的，复制路径的猜测也应验了。
2. Excel 研究里的 C1（CDX > 63,488 B 时 FAT 只写一扇区导致静默损坏）、
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

## 九、演示流：载荷形式是试出来的，不是推出来的

§五第 1 条的悬案。缺 `\x02OlePres000` 的后果已确认：对象合法、双击进 ChemDraw
正常，但复制到另一个 PPT 只得到一张图片——目标程序拿不到缓存渲染，又不会现场
唤起 OLE 服务器，只能取剪贴板里的位图。

头一版按 `CF_ENHMETAFILE`（=14）的字面含义放了裸 EMF，**PowerPoint 双击直接报
「可能导致 PowerPoint 不稳定」**。回头拆参照文件才发现：格式字段虽是 14，数据段
却是一份 **WMF**，EMF 藏在它的一条 `MFCOMMENT` 转义记录里（签名 `WMFC`），
WMF 自己的绘图记录是给只认 WMF 的程序的后备。末尾还有 8 字节
`TOCSIGNATURE`（`NANI`）+ `TOCCOUNT=0`。

在 Windows PowerPoint 上逐变体实测（2026-08-02），每个变体只改一个量：

| # | 数据段 | 双击 | 复制到新 PPT |
|---|---|---|---|
| A | 参照文件的演示流原字节 | 正常 | 正常 |
| B | 裸 EMF | **崩** | — |
| D | EMF 切成 2 条 WMFC 注释记录 | 正常 | 只是图片 |
| F | 1 条注释记录，校验和填 0 | 正常 | 只是图片 |
| G | 我们的 EMF + 参照的后备绘图记录 | 正常 | 正常 |
| E | 1 条注释记录，校验和照算 | 正常 | **正常，可再粘进 Excel** |

A 是关键的对照组：把参照的演示流原字节挂到**我们自己造的容器**上一切正常，
说明容器构建、CDX 注入、幻灯片标记全都无罪，变量只剩流内容。G 则证明我们的
EMF 和 Inkscape 的后备 WMF 都没问题。

由此定下三条硬约束（`retain_cdx/ole/presentation.py`）：

- **整份 EMF 必须放在一条注释记录里。** 记录的 `ByteCount` 是 16 位字段，所以
  EMF 上限 `MAX_EMBEDDED_EMF` = 65501 字节。超限时不挂这个流——复制退化好过
  产出会让 PowerPoint 崩的文件。
- **校验和真的被检验**（F 那一行就是证据）。用 Wine 的算法：34 字节注释头的
  16 位和取二补。它算不出 ChemDraw 那份的 `0xe8c4`——穷举过和 / 异或 / 二补 /
  DWORD 折叠 / CRC16 与全部连续区间，都对不上——但 PowerPoint 接受我们算的值。
- **Inkscape 直出的 WMF 带 22 字节 placeable 头**，进 OLE 前必须剥掉。

EMF 与 WMF 从同一份（letterbox 之后的）SVG 导出，在同一次 `inkscape --shell`
打开里连导两次，只多一次导出动作，不重渲 SVG。

**教训**：参照文件从一开始就在手上，规范条文和实物冲突时我选了条文，还把
「比套一层 WMF 更贴合规范」写进注释当理由。逐字节照抄实物本该是第一选择。
