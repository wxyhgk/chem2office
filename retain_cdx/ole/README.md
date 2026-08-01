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

## 四个文件

```
embed.py     决策层：该补丁还是该重建         ← 唯一对外入口
├─ inject.py     就地补丁（改现成容器）
└─ chemdraw.py   ChemDraw 身份：哪三个流、CLSID、ProgID
   └─ cfbf.py    通用复合文件写入器（不认识化学）
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

### `chemdraw.py` — 「ChemDraw 是什么」

全部化学侧知识都在这 70 行里：CLSID、ProgID、`\x01Ole` 的 20 字节、
`\x01CompObj` 的构造。这三个流与 Windows 产出的模板**逐字节一致**
（有测试盯着），CLSID 与 ProgID 在 Excel / PowerPoint 两侧通用。

`build_chemdraw_ole()` 只是组好流表，交给下面的通用写入器。

### `cfbf.py` — 「复合文件怎么写」

跟 ChemDraw 毫无关系，**能装任何东西**。之所以自己写：olefile 只读不能建。

主流程是四个阶段，每个阶段一节：

| 阶段 | 做什么 |
|---|---|
| `_plan` | 每个流放 mini 流还是常规扇区；各区段占几个扇区；起始位置 |
| `_build_fat` | 扇区分配表，把每条链串起来 |
| `_build_directory` | 根条目 + 每流一条，排成按 CFB 序的平衡二叉查找树 |
| `_build_header` | 512 字节头部与 DIFAT |

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
