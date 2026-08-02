# ole — ChemDraw 的 OLE 容器

## 这一层解决什么

Excel/PPT 解压后，每个 ChemDraw 对象对应一个 `xl/embeddings/oleObject1.bin`。
它不是图片也不是 CDX，而是一个**微软复合文件**（CFBF / MS-CFB）——
一个文件里的迷你文件系统，有扇区、分配表和目录树。

真实容器长这样：

```
oleObject1.bin  4096 字节，魔数 D0CF11E0
根 CLSID 41BA6D21-A02E-11CE-8FD9-0020AFD1F20C   ← ChemDraw.Document
  /\x01Ole          20 字节   声明「嵌入对象，不是链接」
  /\x01CompObj     126 字节   双击时启动谁：ProgID = ChemDraw.Document.6.0
  /CONTENTS       1489 字节   .cdx 文件的原始字节，一字不差
```

本层要做的事一句话：**把 CONTENTS 里的旧分子换成新分子**，并且换完之后
Excel/PowerPoint 仍然认这个容器。

最后半句才是难点。历史上因为 FAT 常量写错，Excel 打开时会「修复」文件、
把 OLE 对象整个删掉，只剩一张预览图。

## 五个文件

```
embed.py     决策层：该补丁还是该重建         ← 唯一对外入口
├─ inject.py         就地补丁（改现成容器）
├─ presentation.py   演示流：对象的缓存渲染（PPT 侧才要）
└─ chemdraw.py       ChemDraw 身份：哪三个流、CLSID、ProgID
   └─ cfbf/          通用复合文件写入器（不认识化学）
        const.py  layout.py  alloc.py  directory.py  header.py
```

### `embed.py` — 只回答一个问题

对外只有 `write_cdx_to_ole()`。拿到新 CDX 后判断**能不能就地改**，
三种情况必须重建：

| 情况 | 原因 |
|---|---|
| 新载荷更大 | 原坑位塞不下 |
| 跨过 4096 字节分界 | 小于 4096 的流走 mini 流，跨界后起始扇区号会被按错误的分配表解析 |
| 容器带 `\x02OlePres000` | 那是上一个分子的渲染缓存，留着会显示错的分子（PowerPoint 的容器有） |

否则走补丁——快，而且模板里其它一切原样保留。

### `inject.py` — 就地补丁

- `inject_cdx_stream`：写入新字节，**不足处补零**，防止上个分子的残留泄漏。
- `patch_direntry_size`：olefile 改不了流长度，所以手改目录项里那个 8 字节字段。

难点是**算出那个字段在文件里的偏移**。目录本身也是一条扇区链，条目不一定在
第一个扇区——PowerPoint 的容器因为多了演示流就超出了一个扇区，
`_direntry_offset` 沿 FAT 链走才修好。

### `presentation.py` — 「别人看到的样子」

只有 PowerPoint 用。少了这个流，对象照样合法、双击照样进 ChemDraw，但
**复制粘贴会退化成一张图片**：目标程序拿不到缓存渲染，又不会现场唤起 OLE
服务器，只能取剪贴板里的位图——粘出来不能双击、也不能再复制进 Excel。

流是 40 字节头 + 数据 + 8 字节 TOC 尾（MS-OLEDS OLEPresentationStream）：

```
0   ffffffff   标记：后面是标准剪贴板格式
4   0000000e   CF_ENHMETAFILE
8   00000004   TargetDeviceSize（无目标设备）
12  Aspect / Lindex / Advf / Reserved1
28  Width / Height    HIMETRIC（0.01mm，1 单位 = 360 EMU）
36  Size              数据字节数
40  数据
末尾 "NANI" + 0      TOCSIGNATURE + TOCCOUNT
```

**数据段必须是 WMF**，尽管格式声明的是 `CF_ENHMETAFILE`。EMF 藏在 WMF 里，
靠一条 `MFCOMMENT` 转义记录（签名 `WMFC`）携带；WMF 自己的绘图记录是给只认
WMF 的程序的后备。这不是从规范推出来的——规范的字面意思正好相反。是在
Windows PowerPoint 上一个变体一个变体试出来的（2026-08-02）：

| 数据段 | 双击 | 复制到新 PPT |
|---|---|---|
| 裸 EMF | **崩**（「可能导致 PowerPoint 不稳定」） | — |
| EMF 切成 2 条 WMFC 注释记录 | 正常 | 只是图片 |
| 1 条注释记录，校验和填 0 | 正常 | 只是图片 |
| **1 条注释记录，校验和照算** | 正常 | **正常，可再复制进 Excel** |

三条硬约束由此而来：

- **整份 EMF 必须放在一条注释记录里。** 记录的 `ByteCount` 是 16 位字段，
  所以 EMF 上限 `MAX_EMBEDDED_EMF` = 65501 字节。超了就不挂这个流——复制退化
  好过产出会崩的文件。
- **校验和真的被检验**（填 0 那一栏就是证据）。用的是 Wine 的算法：34 字节
  注释头的 16 位和取二补。它算不出 ChemDraw 那份的 `0xe8c4`，但 PowerPoint
  接受我们算的值。
- **Inkscape 直出的 WMF 带 22 字节 placeable 头**，进 OLE 前要剥掉。

EMF 和 WMF 从同一份 SVG 导出，在同一次 `inkscape --shell` 打开里连着导两次
（`render_preview(..., companion_wmf=True)`），只多一次导出动作，不多渲染一遍。

### `chemdraw.py` — 「ChemDraw 是什么」

全部化学侧知识都在这 70 行里：CLSID、ProgID、`\x01Ole` 的 20 字节、
`\x01CompObj` 的构造。这三个流与 Windows 产出的模板**逐字节一致**
（有测试盯着），CLSID 与 ProgID 在 Excel / PowerPoint 两侧通用。

`build_chemdraw_ole()` 只是组好流表，交给下面的通用写入器。

### `cfbf/` — 「复合文件怎么写」

跟 ChemDraw 毫无关系，**能装任何东西**。之所以自己写：olefile 只读不能建。

模块划分对应格式自身的结构，懂 MS-CFB 的人可以直接定位：

| 模块 | 做什么 | 行数 |
|---|---|---|
| `const.py` | 格式词汇：扇区大小、特殊扇区号、签名 | 30 |
| `layout.py` | 规划：谁住 mini 流、各段几个扇区、起点在哪（只算数，不出字节） | 125 |
| `alloc.py` | FAT 与 mini-FAT | 39 |
| `directory.py` | 目录树与 128 字节条目 | 120 |
| `header.py` | 512 字节文件头与 DIFAT | 41 |
| `__init__.py` | `build_compound_file`：把上面五步串起来 | 69 |

入口读起来就是这五步：

```python
layout = plan(streams)
out.write(build_header(layout))
out.write(build_fat(layout))
out.write(build_directory(layout, root_clsid=...))
out.write(build_minifat(layout))
out.write(layout.mini_stream)
```

布局：

```
header(512B)  FAT×n  目录×n  mini-FAT×n  mini流×n  大流…
```

**几个容易写错、都有测试守着的地方**：

- **FAT 要能映射自己。** FAT 本身占扇区，所以扇区数要迭代到收敛。早期写死
  一个 FAT 扇区，导致载荷超过 63,488 字节就静默产出损坏容器——olefile 还能
  打开它，读 CONTENTS 才炸。
- **mini-FAT 空闲位填 `FREESECT`（0xFFFFFFFF）而不是 0**，因为 0 是合法的
  扇区号。填 0 曾是我们与 Windows 产物之间仅剩的 404 字节差异。
- **目录序不是字典序**：先比名字长度，再比大写形式。
- 超出头部 DIFAT 能放下的 109 个 FAT 扇区时**直接报错**，而不是产出坏文件。

## 与 Office 层的关系

本层不知道 Excel / PowerPoint 的存在，只吐 `.bin` 字节。
谁把它放进包里、怎么写 rels 和锚点，是 `office/` 的事。

延伸阅读：`docs/excel_ooxml_research.md`（C1/C2 容器缺陷）、
`docs/powerpoint_ole_research.md`（PPT 容器多出的演示流）。
