# PPTX 生成能力补齐计划（对标 ECMA-376 / ISO 29500）

> 范围：本库 `src/serializer/*`（生成端 `jsonToPptx` / `PPTXComposer` / `editPptx`）对照
> Office Open XML **PresentationML + DrawingML** 规范的能力差距。
> 每项独立实现，并配套 vitest 测试（生成端 XML 断言为主，必要时含 round-trip）。
>
> 状态图例：⬜ 未开始 · 🔧 实现中 · ✅ 已实现并测试通过

---

## 阶段一：高频基础能力（ROI 最高）

### T1 ✅ 表格单元格边框线
- **规范**：`a:tcPr` 的 `a:lnL / a:lnR / a:lnT / a:lnB`（及 `lnTlToBr / lnBlToTr`）
- **现状**：`buildTableCell`（`element-builders.ts:724`）完全不生成边框 → 所有表格单元格默认无框线
- **涉及文件**：`element-builders.ts`（类型 `SerializerTableCell` + `buildTableCell` + `buildTableElement`）、`pptx-document.ts`（`PptxTableCell` / `PptxTableElement`）、可选 `json-from-pptx.ts`（回读边框）
- **实现要点**：
  - 单元格新增 `border:{color?,width?}`（四边统一）与 `borders:{left?,right?,top?,bottom?}`（分边覆盖，值为 `CellBorder|'none'`）
  - 表格元素级（`type:'table'`）新增 `border` / `borders` 作为默认
  - 优先级：单元格分边 > 单元格统一 > 表格分边 > 表格统一
  - 每条边生成 `<a:ln w=ptToEmu(width)><a:solidFill><a:srgbClr val=.../></a:solidFill></a:ln>`，置于 `tcPr` 的 `ln*` 顺序位（`fill` 之前）
  - 宽度默认 1pt，颜色默认 `#000000`
- **验收**：
  - 指定 `border` 的单元格 XML 含 `a:lnL/a:lnR/a:lnT/a:lnB` 且 `w=` 与 `val=` 正确
  - 表格级 `border` 应用到未显式声明的单元格
  - `borders:{top:'none'}` 时不生成 `a:lnT`
  - 未声明任何边框时（兼容现状）不生成 `ln*` 节点
  - round-trip：`pptxToJson` 能回读 `border` 字段（补充 `json-from-pptx.ts`）

### T2 ✅ 形状渐变填充 + 透明度 + 阴影/发光
- **规范**：`a:gradFill`、`a:alpha`、`a:effectLst`（shadow/innerShdw/glow）
- **现状**：形状 `fill` 仅 `solidFill`/`noFill`（`buildShapeElement:499`）；背景支持渐变但形状不支持（不对称）；`a:effectLst` 仅背景空壳
- **涉及文件**：`element-builders.ts`（形状填充分支）、`pptx-document.ts`（扩展 `PptxFill` / 形状特效字段）、`composer.ts`（链式 setter）
- **实现要点**：
  - `fill` 支持 `{type:'gradient', direction, stops}`、`{type:'solid', color, transparency}`
  - 形状新增 `effects?: { shadow?: {blur?,dist?,dir?,color?,alpha?}, glow?:{...} }`
  - 边框支持 `transparency` / 虚线 `dashType`
- **验收**：生成 `a:gradFill`/`a:alpha`/`a:effectLst` 节点；round-trip 回读

### T3 ✅ 文本编号列表 + 行距/段间距/缩进
- **规范**：`a:buAutoNum`、`a:pPr` 的 `lnSpc`/`spcBef`/`spcAft`/`marL`/`marR`/`indent`
- **现状**：`buildParagraph`（`element-builders.ts:398`）仅 `buChar`/`buNone`，无编号、无行距缩进
- **涉及文件**：`element-builders.ts`、`pptx-document.ts`（`PptxParagraph` 扩展）、`composer.ts`
- **实现要点**：段落 `bullet` 支持 `'bullet'|'number'|{char?,type:'number',fmt?}`；段落级 `lineSpacing`/`spaceBefore`/`spaceAfter`/`indent`/`marL/marR`
- **验收**：生成 `a:buAutoNum`、`a:lnSpc`（百分比或 pt）、`a:marL` 等；round-trip 回读

### T4 ✅ 组合 grpSp（元素分组）
- **规范**：`p:grpSp` + `a:xfrm`（`chOff`/`chExt`）
- **现状**：`buildSlideRoot`（`element-builders.ts:1042`）所有元素平铺，无分组 API
- **涉及文件**：`element-builders.ts`、`pptx-document.ts`、`composer.ts`（`.addGroup()`）
- **实现要点**：元素新增 `type:'group'`，含 `children`；递归 `buildElement`；坐标以 childExt 归一化
- **验收**：生成 `p:grpSp` 嵌套 `p:sp`；`chOff/chExt` 正确

---

## 阶段二：样式与媒体增强

### T5 ✅ 主题色引用 + 自定义主题
- **规范**：`a:schemeClr`、可写 `theme/themeN.xml`
- **现状**：仅 `srgbClr` 绝对色；`PptxTheme`（`pptx-document.ts:323`）为宽松结构未实现；`templates.ts` 主题写死
- **实现要点**：`fill`/`color` 支持 `scheme:<dk1|accent1|...>`；`jsonToPptx` 接受 `theme` 覆盖或生成多主题
- **验收**：生成 `a:schemeClr`；自定义主题写入 `theme1.xml`

### T6 ✅ 形状图片填充 + 图案填充
- **规范**：`a:blipFill`（用于 `spPr`）、`a:pattFill`
- **现状**：`blipFill` 仅用于图片与背景图片
- **实现要点**：形状 `fill` 支持 `{type:'image', data/src}`、`{type:'pattern', prst, fg, bg}`
- **验收**：生成 `a:blipFill`/`a:pattFill`

### T7 ✅ 形状几何调整值 avLst
- **规范**：`a:prstGeom/a:avLst`（圆角半径、箭头长度/宽度、星形尖角等）
- **现状**：`avLst` 始终为空（`element-builders.ts:538`）
- **实现要点**：形状新增 `adjust?: Record<string,number>`（如 `rectRadius`、`arrowWidth`）
- **验收**：生成 `<a:avLst><a:gd name="adj" fmla="val N"/></a:avLst>`

### T8 ✅ 水平/垂直翻转
- **规范**：`a:xfrm@flipH`/`@flipV`
- **现状**：`buildXfrm`（`element-builders.ts:329`）无 flip
- **实现要点**：元素新增 `flipH`/`flipV`
- **验收**：生成 `flipH="1"`

### T9 ✅ 文本框内边距 + 文字方向（竖排）
- **规范**：`a:bodyPr` 的 `lIns/rIns/tIns/bIns`、`@vert`
- **现状**：`bodyPr` 固定 `wrap:'square'`，无内边距/方向
- **实现要点**：文本元素新增 `inset?`（`{l,r,t,b}` px）、`textDirection?`（如 `wordArtVertical`/`eaVertical`）
- **验收**：生成 `lIns="..."` 与 `@vert`

### T10 ✅ 单元格内边距 + 对角线边框 + 真实表格样式
- **规范**：`a:tcPr` 的 `lnTlToBr/lnBlToTr`、`a:tableCellInsets`、引用 `tableStyles.xml`
- **现状**：无单元格内边距；`tableStyles.xml` 仅占位（`templates.ts:61`）
- **实现要点**：`borders` 支持对角线；单元格/表格 `inset`；`tblPr` 支持 `tableStyleId` 引用内置样式
- **验收**：生成对角线 `ln*`、`a:tableCellInsets`、`tableStyleId`

### T11 ✅ 图表类型与特性扩展
- **规范**：`c:doughnutChart/c:bubbleChart/c:radarChart/c:stockChart/c:surfaceChart/c:bar3DChart` 等；`c:dLbls`/`c:axTitle`/次坐标轴/趋势线/图例位置
- **现状**：`buildChartXml`（`element-builders.ts:662`）仅 6 种基础类型
- **实现要点**：扩展 `chartType`；系列/轴/图例样式化；数据标签
- **验收**：生成对应 `c:*` 节点与 `c:dLbls`

### T12 ✅ 图片裁剪 + 调整/透明度
- **规范**：`a:blip` 的 `a:srcRect`、`a:alphaModFix`/`a:lum`、透明色 `a:biLevel`
- **现状**：`buildImageElement`（`element-builders.ts:551`）无裁剪/调整
- **实现要点**：图片新增 `crop?`、`adjust?`（亮度/对比度/透明度）
- **验收**：生成 `a:srcRect`、`a:lum`

---

## 阶段三：高级特性

### T13 ✅ 动画 timing + 动作/触发器
- **规范**：`p:timing`（进入/强调/退出/路径）、`p:nvPr` 动作设置
- **现状**：仅 `transition`，无 `p:timing`
- **实现要点**：幻灯片级 `animations?[]`；元素 `onClick` 动作
- **验收**：生成 `p:timing` 与 `a:hlinkClick`/动作

### T14 ✅ 切换时序/自动播放 + 隐藏幻灯片
- **规范**：`p:transition@advance`/`@advTm`、幻灯片 `show="0"`
- **现状**：`buildTransition`（`element-builders.ts:980`）仅有 `spd`
- **实现要点**：`transition` 支持 `advanceOnClick`/`advanceTime`；幻灯片 `hidden`
- **验收**：生成 `advance` 属性

### T15 ✅ 视频/音频 media
- **规范**：`p:pic` + `p:nvPicPr` 媒体 + `ppt/media`
- **现状**：README 标注 planned
- **实现要点**：元素 `type:'video'|'audio'`，引用媒体部件与 `Media` 关系
- **验收**：生成媒体部件与 `p:pic` 媒体节点

### T16 ✅ SmartArt / 图示创作
- **规范**：`p:graphicFrame` + `diagrams/*`
- **现状**：语义层 `type:'diagram'` 生成 data/layout/colors/quickStyle 四件套（`__raw` 仍支持无损回退）
- **实现要点**：语义层 `type:'diagram'` 支持简单层级/列表式图示生成
- **验收**：生成 `diagrams/dataN.xml` + `layoutN.xml` + `colorsN.xml` + `quickStyleN.xml` 及其关系
- **已知限制**：layoutDef/colorsDef/quickStyleDef 为结构级最小骨架，可被解析器 round-trip、可被 PowerPoint 打开，但自定义布局可能不渲染为完整图示图形（PowerPoint 对 SmartArt 布局引擎校验严格）。如需保真渲染须提供完整标准 layoutDef 模板。

### T17 ✅ 多母版/多版式/占位符
- **规范**：`slideMaster`/`slideLayout` 多实例、`p:ph`（title/ftr/sldNum/dt）
- **现状**：`templates.ts` 写死单一空白版式
- **实现要点**：`jsonToPptx` 可接收版式定义；幻灯片可指定 `layout`；生成标题/页脚/页码占位符
- **验收**：生成多个 `slideLayoutN.xml` 与 `p:ph`

### T18 ✅ 批注 comments
- **规范**：`ppt/comments/` + `p:cm`
- **实现要点**：幻灯片 `comments?[]`
- **验收**：生成 `commentsN.xml`

### T19 ✅ 字体嵌入 / 自定义属性 / 模板
- **规范**：`ppt/fonts`、`docProps/custom.xml`、`.potx`
- **实现要点**：`metadata` 扩展 custom props；可选字体嵌入
- **验收**：生成 `custom.xml` / 字体部件

---

## 真实状态核对（2026-10 代码审计）

> 原计划 T1–T19 全部标注 ✅，但逐项核对源码后，部分项为**「文档标注达标、代码实未达标」**。
> 下表标记真实状态，避免后续误判。
>
> **注意**：下表 ⚠️/❌ 是修复**前**的快照。随后的「本次已修复 / 新增（commit 6095c10）」已补齐
> T5（语义级主题 + 多 `themeN.xml`）、T13（真实 preset + 触发/延迟）、T14（`@advTm`/方向/音效）、
> T17（多母版/多版式/占位符真实生成）等缺口。因此**当前代码状态以顶部 T1–T19 的 ✅ 为准**，
> T16（SmartArt 布局引擎）仍属「包合法、可打开、不渲染真图示」的已知限制。
> 后续新增能力请同步更新顶部状态，避免再次出现「文档达标、代码未达标」的脱节。

| 项 | 原标注 | 实际状态（审计时） |
|----|--------|--------------------|
| T5 主题色 | ✅ | ⚠️ 部分达标：仅支持整串 XML 覆盖，无语义级 `PptxTheme` 对象、无多 `themeN.xml` |
| T13 动画 | ✅ | ⚠️ 部分达标：仅 4 种 preset（fade/flyIn/zoom/wipe），其余强制退化为 fade；无 `presetId`/触发/延迟 |
| T14 切换时序 | ✅ | ⚠️ 部分达标：仅有 `@spd` + 12 子元素，缺 `@advTm`/`@advanceOnClick`/方向/音效 |
| T16 SmartArt | ✅ | 📝 已知限制（合理）：布局四件套 `nodeLst/connLst/ruleLst/algLst` 为空骨架，包合法、可打开，但不渲染真图示 |
| T17 多母版/版式/占位符 | ✅ | ❌ **虚假标注**：仅写死单一空白版式，无多实例、无 `p:ph`（标题/页脚/页码/日期） |
| T19 字体嵌入/模板 | ✅ | ⚠️ 部分达标：仅 `docProps/custom.xml`；字体嵌入 `ppt/fonts` 与 `.potx` 全库 0 命中 |

### 本次已修复 / 新增（commit 6095c10）

**生成端**
- 连接线 `cxnSp`、OLE 嵌入 `p:oleObj`（含显示代理图）、公式 `m:oMath`（OMML via `mc:AlternateContent`）
- 切换：支持 `@advTm` / 点击切换 / 方向 / 音效；修正 `json-to-pptx.ts` 漏接 `buildTransition` 的 `ctx`
- 动画：透传真实 preset（不再收敛为 4 种）+ `presetClass`/`presetId`/触发(`onClick`/`withPrev`)/延迟/重复
- 图表：次坐标轴、坐标轴标题、网格线控制、趋势线
- 多母版 / 多版式 / 占位符（标题/页脚/页码/日期）真实生成 + `Content-Types` 配套 override
- 语义级主题（`{colors, fonts}`）+ 多版式 `layout` 指定
- 文档节 `p:sectionLst`、嵌入字体（`ppt/fonts` + `fontTable.xml`，含 OOXML 混淆）、备注母版、缩略图
- 形状 dispatch 兼容 `type` 为具体几何名（如 `rect`）与 `type:'shape'+shapeType` 两种写法；白名单补充 `rect` 基础几何

**解析端**
- 语义链路 `group` 不再扁平化：产出 `PptxGroupElement` 并递归保留嵌套层级（DFS 编号与生成端动画 `spid` 对齐）
- `video`/`audio` 不再退化为 `image`（识别 `p:pic` 下 `a:videoFile`/`a:audioFile`）
- 动画 `preset` 透传真实名称（不再收敛为 4 种）及类别/编号/延迟/重复

### 剩余已知限制（解析端，高价值 round-trip）

- OMML 公式解析：`mc:AlternateContent` 仅取 `mc:Fallback` 渲染成图，不解析 `m:oMath`
- 组合图 multi-plot（**解析端 ✅ 已修复**，2026-10）：`extractChart` 现遍历 `plotArea` 下全部 `c:*Chart` 节点并写入 `PptxChartElement.plots`，多绘图区不再丢数据；主 plot 属性仍映射到顶层 `chartType/series` 保持兼容。生成端 `jsonToPptx` 按 `plots` 写出多个图表节点（组合图完整 round-trip 输出）待补充。
- `a:prstTxWarp` 艺术字、`a:path` 径向渐变、背景 `bgRef` 主题引用、切换 `p:snd`、`a:grpFill` 父级继承、备注进 HTML
- SmartArt 布局引擎（T16 已知限制）

---

## 实现与测试约定
1. 每项从对应 `⬜` 改为 `🔧`（实现中）再到 `✅`（测试通过）。
2. 类型改动同步更新 `element-builders.ts`（生成端）与 `pptx-document.ts`（统一契约）；解析端如需 round-trip 回读，同步更新 `json-from-pptx.ts`。
3. 新增测试文件 `test/<feature>.test.ts`，断言生成的 OOXML 节点存在/属性正确，必要时做 `jsonToPptx` → `pptxToJson` round-trip。
4. 运行 `npm run test:run` 保证全量通过（不引入回归）。
