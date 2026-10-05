# PptxDocument JSON 契约

权威定义在 `src/types/pptx-document.ts`（入口已 `export *`）。坐标单位：**px**（96 DPI）；字号/线宽/间距：**pt**；颜色：`#RRGGBB`（或 `scheme:accent1` 主题色引用，解析端会按该页真实主题解析为绝对色）。

```
PptxDocument
├── version: '1.0'
├── slideSize: { width, height }          // px，缺省 1280x720 由调用方给定
├── metadata?: { title, subject, author, keywords, description, lastModifiedBy,
│                created, modified, category, status, contentType, language, version }
├── customProps?: Record<string, string>  // 写回 docProps/custom.xml
├── media?: Record<id, { base64, mime }>  // 可选，元素可用 ref 引用避免重复内联
├── theme?: PptxTheme | string             // 语义级主题对象 或 完整 theme XML 字符串
├── masters?: PptxSlideMaster[]            // 母版/版式定义（省略时退回单一空白版式）
├── sections?: PptxSection[]               // 文档节（presentation.xml 的 p:section）
├── fonts?: PptxFontResource[]             // 嵌入字体（fntdata + fontTable.xml）
└── slides: PptxSlide[]
```

## PptxSlide

| 字段 | 类型 | 说明 |
|---|---|---|
| `layout` | `number` | 使用的版式索引（0 基，按 `masters[].layouts` 展平后顺序）；省略用第 0 个 |
| `background` | `string \| {type:'solid',color} \| {type:'gradient',direction,stops} \| {type:'image',data?,src?,extension?,srcRect?,tile?}` | 页背景；缺省继承主题 |
| `transition` | `PptxTransition` | 解析自 `p:transition`，生成端写回 |
| `notes` | `string` | 演讲者备注（写回 notesSlide 部件） |
| `comments` | `PptxComment[]` | 写回 `ppt/comments/commentsN.xml` |
| `hidden` | `boolean` | 对应 `p:sld show="0"` |
| `advanceTime` | `number` | 自动播放停留毫秒（`p:timing` afterTime） |
| `animations` | `PptxAnimation[]` | 元素动画（解析自 `p:timing` 的 `p:spTgt`） |
| `elements` | `PptxElement[]` | 见下 |

`PptxTransition`：`{ type, duration(ms), advanceOnClick?, advanceAfterTime?(ms), direction?, sound?: {name?, data?, extension?} }`
- `advanceAfterTime` 与 `advanceTime` 互为等价表达；生成时优先写 `@advTm`
- `sound` 为内嵌音频关系目标或 `{ name }` 内置音效名

`PptxAnimation`：`{ target, type, duration(秒), presetClass?, presetId?, presetSubtype?, trigger?, delay?, repeat?, direction?, path? }`
- `target`：字符串=OOXML spid（推荐，group 嵌套也正确）；数字=elements 下标（向后兼容，仅扁平布局）
- `type`：OOXML preset 名（`flyIn` / `wipe` / `bounce` / `path`…），解析端透传真实值
- `presetClass`：`'entr'` | `'exit'` | `'emph'` | `'path'` | `'mediacall'`
- `trigger`：`{type:'afterPrev'\|'withPrev'\|'onClick', delay?, target?}`

`PptxComment`：`{ author?, text, dt?, pos?: {x?, y?} }`，`pos` 单位 EMU（缺省 1 英寸处）。

## 公共字段（所有元素）

| 字段 | 说明 |
|---|---|
| `name?` | 元素名，便于编辑时区分 |
| `descr?` | 替代文本（无障碍，`p:cNvPr@descr`） |
| `decorative?` | 是否为装饰性元素 |
| `__raw?` | 原始 OOXML 载荷 `{tag, node, rels?, parts?}`，解析端产出 |
| `rawFallback?` | 强制以 `__raw` 回写（需配合 `rawDeps:'all'` 解析出的载荷） |

## text

```ts
{ type:'text', x, y, width, height, rotation?, flipH?, flipV?,
  paragraphs?, runs?, text?,                       // 优先级：paragraphs > runs > text
  fontSize?, color?, bold?, italic?, underline?, fontFace?, href?,
  align?, valign?,
  bullet?, lineSpacing?, spaceBefore?, spaceAfter?, indentLeft?, indentRight?, indent?,
  inset?: { l?, r?, t?, b? },                      // 文本框内边距 px
  textDirection?: 'horz'|'vert'|'vert270'|'wordArtVert'|'eaVert'|'mongolianVert'|'wordArtVertRtl',
  // —— 底层形状（带文字的形状，如椭圆/饼图/弧线）——
  shapeType?, fill?, line?, adjust?, effects?,
  // —— 高级排版 ——
  numCol?, spcCol?,                                // 分栏数 / 栏间距 pt
  autofit?: 'none'|'normal'|'shape',               // 自动适配（a:noAutofit / normAutofit / spAutoFit）
  fontScale?, lnSpcReduction?,                     // normAutofit 的字号缩放% / 行距缩减%
  prstTxWarp?,                                     // 艺术字变形预设（textArchUp / textWave / …）
  noWrap? }                                        // 不换行（a:bodyPr@wrap="none"）
```

`PptxParagraph`：`{ text?, runs?, align?, bullet?, lineSpacing?, spaceBefore?, spaceAfter?, indentLeft?, indentRight?, indent?, rtl? }`

`PptxTextRun`：
```ts
{ text, fontSize?, color?, bold?, italic?, underline?, fontFace?, href?,
  field?: string,         // 动态字段（a:fld@type）：'slidenum' 页码、'datetime' 日期
  break?: boolean,        // 软换行（a:br）：本 run 前强制换行，text 为空
  outline?: { color?, width? } | 'none',   // 文字描边（a:rPr/a:ln）
  shadow?: { color?, blur?, x?, y?, alpha? } }  // 文字外阴影（a:rPr/a:effectLst/a:outerShdw）
```
- `href` 支持 `'#N'` 内部跳转
- `lineSpacing`：数字=百分比（100=单倍）或 `{type:'pt'|'percent', value}`
- `bullet`：`true`=项目符号；`'number'`=自动编号；或 `{type:'number'|'bullet'|'picture', fmt?, start?, char?, font?, sizePct?, data?, rid?}`

## shape

```ts
{ type:'shape', shapeType, x, y, width, height, rotation?, flipH?, flipV?,
  fill?, line?, effects?, adjust?, custGeom?, threeD? }
```

- `shapeType`：OOXML `prstGeom` 预设名（`rect` / `roundRect` / `ellipse` / `triangle` / `arrow` / `star5` / `foldedCorner` …）
- `fill`：`string` | `{type:'solid', color, transparency?}` | `{type:'gradient', direction, stops:[{color,position}]}` | `{type:'image', data|src, extension?, tile?:{sx,sy,tx,ty}, srcRect?:{l,t,r,b}}` | `{type:'pattern', prst, fg, bg}` | `'none'` | `null`
- `line`：`{ color?, width?(pt), transparency?, dashType? }` | `'none'` | `null`
- `effects`：`{ shadow?: PptxShadow|true, glow?: PptxGlow|true }`，`PptxShadow{type?,color?,blur?,distance?,angle?,transparency?}`、`PptxGlow{color?,blur?}`（blur 单位 pt）
- `adjust`：几何调整值，key 必须是 OOXML 的 gd 名：`{ adj: 25000 }`（roundRect/snip）或 `{ adj1: 50000, adj2: 40000 }`（箭头/标注/星形）。值域通常 0–100000
- `custGeom`：自定义几何（`a:custGeom`），指定时优先于 `shapeType` 的预设几何
  ```
  { paths: [{ w?, h?, closed?, commands: PptxGeometryCommand[] }] }
  ```
  `PptxGeometryCommand`：`moveTo{x,y}` | `lnTo{x,y}` | `cubicBezTo{x1,y1,x2,y2,x,y}` | `quadBezTo{x1,y1,x,y}` | `arcTo{wR,hR,stAng,swAng}` | `close`
  坐标可传 0~1 归一化（自动放大到坐标空间）或绝对值
- `threeD`：三维属性 `{ shape?: PptxShape3D, scene?: PptxScene3D }`
  - `PptxShape3D`：`{ extrusionHeight?, contourWidth?, extrusionColor?, contourColor?, bevelTop?, bevelBottom?, material? }`（尺寸 pt）
  - `PptxScene3D`：`{ camera?, fov?, zoom?, rotX?, rotY?, rotZ?, lightRig?, lightDir? }`

## image

```ts
{ type:'image', x, y, width, height, rotation?,
  data?: string,        // dataURL（data:<mime>;base64,<b64>）或裸 base64
  src?: string,         // 远程 URL（生成端会下载为媒体）
  extension?: string,
  href?: string,        // 图片级超链接（'#N' 内部跳转）
  crop?: { l?, r?, t?, b? },                              // 百分比 0–100
  imageAdjust?: { brightness?(-100..100), contrast?(-100..100), transparency?(0..100) } }
```

> 解析端图片/媒体统一输出完整 `data:<mime>;base64,<b64>` dataURL（不再输出裸 base64）。

## chart

```ts
{ type:'chart', chartType, x, y, width, height,
  title?, legend?, varyColors?, barDir?: 'bar'|'col',
  categories?: string[], series?: PptxChartSeries[],
  grouping?: 'clustered'|'stacked'|'percentStacked'|'standard',
  holeSize?, smooth?, marker?, ofPieType?: 'pie'|'bar',
  numberFormat?, bubble3D?, showNegBubbles?, bubbleScale?, wireframe?,
  spaceFill?,                          // 图表区填充（'none' 或 #RRGGBB）
  view3D?: { rotX?, rotY?, depthPercent?, rAngAx? },
  axisTitles?: { category?, value?, secondaryValue? },
  secondaryValueAxis?: boolean,        // 启用次数值轴
  dataLabels?: boolean,                // 显示数据标签（系列级可覆盖）
  gridlines?: { major?, minor? } }
```

- `chartType`（ECMA-376 plotArea 下全部节点）：`barChart` `bar3DChart` `lineChart` `line3DChart` `areaChart` `area3DChart` `pieChart` `pie3DChart` `doughnutChart` `ofPieChart` `scatterChart` `bubbleChart` `radarChart` `stockChart` `surfaceChart` `surface3DChart`
- `PptxChartSeries`：
  ```ts
  { name?, values?: number[], x?, y?, open?, high?, low?, close?,  // 散点用 x/y，股票用 open/high/low/close
    color?, pointColors?: (string|undefined)[],  // 逐点填充色（c:dPt）
    axis?: 'primary'|'secondary',                // 绑定到哪条数值轴
    dataLabels?: boolean,                        // 系列级数据标签覆盖
    trendlines?: PptxTrendline[] }               // 趋势线
  ```
  `PptxTrendline`：`{ type?: 'linear'|'exp'|'log'|'poly'|'movingAvg'|'power', name?, order?, period?, forward?, backward?, showEquation?, showRSquared?, intercept? }`

## table

```ts
{ type:'table', x, y, width, height,
  colWidths?: number[], rowHeights?: number[],   // px，缺省均分
  border?: { color?, width? },                   // 表格级四边统一
  borders?: { left?, right?, top?, bottom?, diagonal?: 'tlBr'|'blTr'|'both' },
  inset?: { l?, r?, t?, b? },                    // 表格级单元格内边距 px
  tableStyleId?: string,                          // 引用 tableStyles.xml 的 GUID
  rows: PptxTableRow[] }
```

`PptxTableRow`：`{ height?: number, cells: PptxTableCell[] }`
`PptxTableCell`：`{ text?, paragraphs?, colSpan?, rowSpan?, hMerge?, vMerge?, fill?, border?, borders?{left,right,top,bottom}, align?, valign?: 'top'|'middle'|'bottom', fontSize?, color?, bold?, italic?, underline?, fontFace?, inset? }`
- `hMerge`/`vMerge`：被合并吞并的单元格（OOXML 仍需存在这些节点，缺失会导致表格结构错乱）
- 单元格 `borders.left` 等可为 `{color,width}` 或 `'none'`

常用 `tableStyleId`：
- `{5940675A-B579-460E-94D1-54222C63F5DA}` — Table Grid（黑色网格，库内默认）
- `{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}` — Medium Style 2 - Accent 1（强调色表头 + 白色网格线 + 交替行）

## diagram（SmartArt）

```ts
{ type:'diagram', x, y, width, height,
  diagramType?: 'list'|'hierarchy'|'process'|'cycle'|'pyramid',   // 创作端
  nodes?: PptxDiagramNode[],                                        // 创作端层级
  texts?: string[],                                                 // 解析端：数据部件文本
  shapes?: PptxDiagramShape[],                                      // 解析端：缓存绘图形状
  dataPath?: string }                                               // 数据部件路径（调试用）
```

`PptxDiagramNode`：`{ text, children?: PptxDiagramNode[] }`

`PptxDiagramShape`（解析自缓存绘图 `drawingN.xml`，坐标 px、相对图示框）：
```ts
{ x, y, width, height, prst?, adjust?, connector?,
  fill?, lineColor?, lineWidth?, text?, fontSize?, color?, bold?,
  align?, anchor?, flipH?, flipV? }
```
- `connector: true` 表示连接线（无填充文字，仅描边）
- `prst` 为 OOXML 预设几何（`roundRect` / `ellipse` / `rect` / `straightConnector1` …）

## group

```ts
{ type:'group', x, y, width, height,
  childrenCoordinates?: 'local'|'page'|'relative',   // 默认 'local'
  children: PptxElement[] }
```
- `local`（OOXML 标准）：children 的 x/y 为相对组左上角的局部坐标
- `page`：children 的 x/y 为页绝对坐标，生成时自动减 group 偏移做相对化
- `relative`：解析端产出——children 坐标已由 `chOff/chExt` 空间换算为相对组左上角的偏移

## connector / ole / math

```ts
// 连接线（p:cxnSp）
{ type:'connector', x, y, width, height,
  shapeType?: 'straightConnector1'|'bentConnector3'|'curvedConnector2'|...,  // 默认 'straightConnector1'
  start?: { x, y }, end?: { x, y },     // 精确端点（px），优先于 x/y/width/height
  rotation?, flipH?, flipV?, line?, adjust? }

// OLE 嵌入对象（p:oleObj）
{ type:'ole', x, y, width, height,
  progId?: string,            // 'Excel.Sheet.12' / 'PowerPoint.Show.12'
  target?: string,            // ppt/embeddings/*.xlsx
  data?: string, extension?: string,    // 嵌入文件 base64
  showAsIcon?: boolean, poster?: { data?, src?, extension? } }

// 公式（OMML m:oMathPara / m:oMath）
{ type:'math', x, y, width, height,
  omml?: string,              // OMML XML 字符串
  text?: string }             // 纯文本回退
```

## video / audio / raw

```ts
{ type:'video'|'audio', x, y, width, height, data?, src?, extension?,
  poster?: { data?, src?, extension? } }   // 仅 video 有 poster

{ type:'raw', x, y, width, height }        // 解析兜底，始终按 __raw 回写
```

## 文档级高级字段

### theme（PptxTheme）

```ts
{ name?: string,
  colors?: PptxThemeColorScheme,   // 12 色槽：dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink
  fonts?: PptxThemeFontScheme }    // major/minor 各含 { latin, ea, cs }
```
提供 `colors`/`fonts` 时生成端据此构造完整 `themeN.xml`；也可传整串 XML 字符串。

### masters（PptxSlideMaster[]）

```ts
[{ name?, background?, elements?: PptxElement[],
   placeholders?: PptxPlaceholder[],     // title/body/ftr/sldNum/dt 的默认位置与样式
   layouts?: PptxSlideLayout[] }]        // 该母版下的版式列表
```
`PptxSlideLayout`：`{ name?, background?, elements?, placeholders?, showMasterSp? }`
`PptxPlaceholder`：`{ type, x, y, width, height, idx?, prompt?, name?, fontSize?, color?, bold?, fontFace?, align?, valign?, bullet? }`

### sections（PptxSection[]）

```ts
[{ name?: string, slides: number[] }]   // slides 为 0 基页索引
```

### fonts（PptxFontResource[]）

```ts
[{ name: string,           // 字体族名
   data: string,           // base64（生成端会按 ECMA-376 做 XOR 混淆；解析端自动解混淆）
   panose?: string,        // 20 位十六进制分类
   bold?: boolean, italic?: boolean,
   embedType?: 'full'|'subset' }]
```

## 单位换算速查

| 场景 | 换算 |
|---|---|
| EMU → px | `px = emu × 96 / 914400`（`emu / 9525`） |
| px → EMU | `emu = px × 9525` |
| 表格单元格默认内边距 | 左右 0.1in = 91440 EMU = 9.6px；上下 0.05in = 45720 EMU = 4.8px |
| 线宽 1pt | 12700 EMU |
| OOXML 千分比属性（`a:tile@sx`、`a:srcRect@l`） | `值 = 比例 × 100000` |
