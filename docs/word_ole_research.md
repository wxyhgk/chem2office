# Word ChemDraw OLE 研究报告

> 2026-10-13 · 解剖 `word.docx`（Windows Word 16 + ChemDraw 插入的单个分子），
> 确立 Word 侧的落地模型，并与 Excel / PowerPoint 对照。该样例原样入库为
> `retain_cdx/data/template.docx`。

## 一、样本结构

18.5 KB，与 OLE 相关的只有三处：

```
word/document.xml                4040 B   一个段落 → 一个 run → w:object
word/_rels/document.xml.rels     1094 B   rId4 → image，rId5 → oleObject
word/embeddings/oleObject1.bin  12288 B   CFBF 容器
word/media/image1.emf            7624 B   矢量预览
```

`[Content_Types].xml`：`Default bin → oleObject`、`Default emf → image/x-emf`。

## 二、文档里的 OLE 标记

```xml
<w:p>
  <w:r>
    <w:object w:dxaOrig="929" w:dyaOrig="842" w14:anchorId="550DBFC2">
      <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" …>…</v:shapetype>
      <v:shape id="_x0000_i1025" type="#_x0000_t75"
               style="width:46.6pt;height:42pt" o:ole="">
        <v:imagedata r:id="rId4" o:title=""/>
      </v:shape>
      <o:OLEObject Type="Embed" ProgID="ChemDraw.Document.6.0"
                   ShapeID="_x0000_i1025" DrawAspect="Content"
                   ObjectID="_1853394678" r:id="rId5"/>
    </w:object>
  </w:r>
</w:p>
```

- **嵌入式（随文字排版）**：对象就是 run 里的一个「字符」，没有锚点、没有坐标。
  位置 = 它在文字流里的位置。
- **只有 VML，没有 DrawingML**。`v:shape` 的 style 给显示宽高（pt），
  `v:imagedata` 指预览图。
- `w:dxaOrig` / `w:dyaOrig` 是原始尺寸，单位 twips（1/20 pt）。929 × 842 twips
  = 46.45 × 42.1 pt，与 style 的 46.6 × 42 pt 基本一致——即 100% 缩放。
- 三个**文档内唯一**的标识：VML shape id（`_x0000_i1025`）、`ObjectID`
  （`_1853394678`，旧 `.doc` 里 ObjectPool 的存储名）、以及 `w14:anchorId`。
- `_x0000_t75` 是标准「图片框」shapetype。Word 只在第一次用到时定义，后面的
  对象只用 `type="#_x0000_t75"` 引用。

## 三、与 Excel / PowerPoint 对照

| | Excel | PowerPoint | **Word** |
|---|---|---|---|
| 定位 | 三套锚点（worksheet / drawing / VML）必须一致 | `p:xfrm` 绝对 EMU | **文字流位置**；只给宽高 |
| VML | 必需 | 不存在 | **必需，且是唯一的显示层** |
| DrawingML | 有 | 有 | **无** |
| 元素顺序 | CT_Worksheet 固定序列 | `p:spTree` 追加 | `w:sectPr` 必须是 body 最后一个子元素 |

Word 的几何比 Excel 简单得多：没有网格、没有字体相关的列宽、没有锚点同步。

## 四、OLE 容器：同源，但流不一样

| 流 | Excel | PowerPoint | **Word** |
|---|---|---|---|
| `\x01CompObj`（126 B） | ✓ | ✓ | ✓ |
| `\x01Ole`（20 B） | ✓ | ✓ | ✓ |
| `CONTENTS`（CDX） | ✓ | ✓ | ✓ |
| `\x02OlePres000` | — | ✓ | **—** |
| `\x03EPRINT` | — | — | **✓ 7624 B** |
| `\x03ObjInfo` | — | — | **✓ 6 B** |

- 根 CLSID、`\x01CompObj`、`\x01Ole` 三种格式**逐字节相同**，
  `ole/chemdraw.py` 直接复用。
- **`\x03EPRINT` 与 `word/media/image1.emf` 逐字节相同**。Word 的渲染缓存就是
  预览 EMF 原样放进容器——不必像 PowerPoint 的 `\x02OlePres000` 那样把 EMF 藏进
  WMF 注释记录、算校验和、受 64 KB 上限约束。
- `\x03ObjInfo` = `00 00 03 00 0d 00`（MS-DOC 的 ODT 结构）。照抄。

## 五、实现

1. 样例原样入库为 `retain_cdx/data/template.docx`，作为 `w:object` 骨架的零件供体。
2. `retain_cdx/ole/eprint.py`：两个流的名字与 ObjInfo 字节；`write_cdx_to_ole`
   新增 `extra_streams=`。`\x03EPRINT` 也被列为「陈旧缓存」——就地补丁 Word
   容器时会重建，不会留下旧分子的渲染。
3. `retain_cdx/office/word/`：
   - `_wml.py`：document.xml 手术。分配 shape id / ObjectID（扫描全部
     `word/*.xml`，页眉页脚里也有 VML）、只定义一次 shapetype、去掉克隆来的
     `w14:anchorId`、写 rels 与部件（按 Word 的方式编号：`oleObject2.bin`）。
   - `writer.py`：`WordCDXWriter`，三种落点——占位符、段落末尾、新段落。
4. 公共的 OPC 记账（rels、内容类型、部件名）从 PowerPoint 私有模块上提到
   `office/_opc.py`，Word 与 PowerPoint 共用。

### 尺寸

只给宽或高，另一边按分子比例；两个都给，分子按比例**塞进**这个框。永远不写
与 CDX 比例不一致的框——Excel 侧刚修过这个坑：比例不符的对象在 Excel 里看着
对（预览已 letterbox），被别的程序按 CDX 重渲时就会拉伸。

### 占位符

Word 会因格式、拼写检查、编辑历史把手打的文字切成多个 run，所以 `{{mol}}`
常常跨 run。匹配在段落拼接后的文本上做，再映射回各 run：前半段留在第一个
run，后半段留在它**原来的** run（保留原格式），中间的 `w:t` 删掉，对象 run
插在中间。被改过的 `w:t` 加 `xml:space="preserve"`，否则对象两侧的空格会被吞。

## 六、实现期间的坑

**`mc:Ignorable` 前缀被 ElementTree 丢掉。** Word 的根元素声明了二十多个前缀
（`w15`、`w16cid`、`w16du`…），大多只为让 `mc:Ignorable` 能列出它们，没有元素
使用。ElementTree 只输出用到的声明，往返一次就留下 `Ignorable` 指向未声明的
前缀——与 PowerPoint 的 `xmlns:v`、Excel 的 `xr6` 是同一个根因。写回时把原根
元素的声明全部补回；`test_mc_ignorable_prefixes_stay_declared` 守着（回退修复即变红）。

## 七、回归防线

`tests/ooxml_checks.py` 新增 Word 检查：`word/*.xml` 里 `v:shape` id、
`ObjectID`、`v:shapetype` id 不得重复。`tests/test_word_writer.py` 覆盖容器
流组成、与原生容器的逐字节一致、占位符的各种拆分、尺寸比例。

## 八、待验证（需要 Windows Word + ChemDraw）

以下只能在真机上确认，目前是按样本推断：

1. 生成的文档在 Word 中打开不报修复，双击能进入 ChemDraw 编辑。
2. 带 `\x03EPRINT` 时，从 Word 复制到别处是否不再唤起 ChemDraw；不带时是否会
   （即 Excel 侧「不带就不能用」的结论在 Word 是否同样成立）。
3. 只有一份样本；如果 ChemDraw / Word 版本不同时 `\x03ObjInfo` 有别的取值，
   需要再采样。
