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

## 实现与测试约定
1. 每项从对应 `⬜` 改为 `🔧`（实现中）再到 `✅`（测试通过）。
2. 类型改动同步更新 `element-builders.ts`（生成端）与 `pptx-document.ts`（统一契约）；解析端如需 round-trip 回读，同步更新 `json-from-pptx.ts`。
3. 新增测试文件 `test/<feature>.test.ts`，断言生成的 OOXML 节点存在/属性正确，必要时做 `jsonToPptx` → `pptxToJson` round-trip。
4. 运行 `npm run test:run` 保证全量通过（不引入回归）。
