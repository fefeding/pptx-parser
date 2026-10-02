# PptxDocument JSON 契约

权威定义在 `src/types/pptx-document.ts`（入口已 `export *`）。坐标单位：**px**（96 DPI）；字号/线宽/间距：**pt**；颜色：`#RRGGBB`（或 `scheme:accent1` 主题色引用）。

```
PptxDocument
├── version: '1.0'
├── slideSize: { width, height }          // px，缺省 1280x720 由调用方给定
├── metadata?: { title, subject, author, keywords, description, lastModifiedBy,
│                created, modified, category, status, contentType, language, version }
├── customProps?: Record<string, string>  // 写回 docProps/custom.xml
├── media?: Record<id, { base64, mime }>  // 可选，元素可用 ref 引用避免重复内联
├── theme?: Record<string, unknown>
└── slides: PptxSlide[]
```

## PptxSlide

| 字段 | 类型 | 说明 |
|---|---|---|
| `background` | `string \| {type:'solid',color} \| {type:'gradient',direction,stops} \| {type:'image',data?,src?,extension?,srcRect?,tile?}` | 页背景；缺省继承主题 |
| `transition` | `{ type, duration(ms), advanceOnClick? }` | 解析自 `p:transition`，生成端写回 |
| `notes` | `string` | 演讲者备注（写回 notesSlide 部件） |
| `comments` | `PptxComment[]` | 写回 `ppt/comments/commentsN.xml` |
| `hidden` | `boolean` | 对应 `p:sld show="0"` |
| `advanceTime` | `number` | 自动播放停留毫秒（`p:timing` afterTime） |
| `animations` | `{ target, type:'fade'\|'flyIn'\|'zoom'\|'wipe', duration }[]` | `target` 为本页 `elements` 下标 |
| `elements` | `PptxElement[]` | 见下 |

`PptxComment`：`{ author?, text, dt?, pos?: {x?, y?} }`，`pos` 单位 EMU（缺省 1 英寸处）。

## 公共字段（所有元素）

| 字段 | 说明 |
|---|---|
| `name?` | 元素名，便于编辑时区分 |
| `__raw?` | 原始 OOXML 载荷 `{tag, node, rels?, parts?}`，解析端产出 |
| `rawFallback?` | 强制以 `__raw` 回写（需配合 `rawDeps:'all'` 解析出的载荷） |

## text

```ts
{ type:'text', x, y, width, height, rotation?, align?, valign?,
  paragraphs?, runs?, text?,                       // 优先级：paragraphs > runs > text
  fontSize?, color?, bold?, italic?, underline?, fontFace?, href?,
  bullet?, lineSpacing?, spaceBefore?, spaceAfter?, indentLeft?, indentRight?, indent?,
  inset?: { l?, r?, t?, b? },                      // 文本框内边距 px
  textDirection?: 'horz'|'vert'|'vert270'|'wordArtVert'|'eaVert'|'mongolianVert'|'wordArtVertRtl' }
```

`PptxParagraph`：`{ text?, runs?, align?, bullet?, lineSpacing?, spaceBefore?, spaceAfter?, indentLeft?, indentRight?, indent? }`
`PptxTextRun`：`{ text, fontSize?, color?, bold?, italic?, underline?, fontFace?, href? }`（`href` 支持 `'#N'` 内部跳转）

- `lineSpacing`：数字=百分比（100=单倍）或 `{type:'pt'|'percent', value}`
- `bullet`：`true`=项目符号；`'number'`=自动编号；或 `{type:'number'|'bullet', fmt?, start?, char?}`

## shape

```ts
{ type:'shape', shapeType, x, y, width, height, rotation?,
  fill?, line?, effects?, adjust? }
```

- `shapeType`：OOXML `prstGeom` 预设名（`rect` / `roundRect` / `ellipse` / `triangle` / `arrow` / `star5` / `foldedCorner` …）
- `fill`：`string` | `{type:'solid', color, transparency?}` | `{type:'gradient', direction, stops:[{color,position}]}` | `'none'` | `null`
  - 生成端还额外支持形状**图片填充**与**图案填充**：`{type:'image', data|src, extension?, tile?:{sx,sy,tx,ty}, srcRect?:{l,t,r,b}}`、`{type:'pattern', prst, fg, bg}`
- `line`：`{ color?, width?(pt), transparency?, dashType? }` | `'none'` | `null`
- `effects`：`{ shadow?: PptxShadow|true, glow?: PptxGlow|true }`，`PptxShadow{type?,color?,blur?,distance?,angle?,transparency?}`、`PptxGlow{color?,blur?}`（blur 单位 pt）
- `adjust`：几何调整值，key 必须是 OOXML 的 gd 名：`{ adj: 25000 }`（roundRect/snip）或 `{ adj1: 50000, adj2: 40000 }`（箭头/标注/星形）。值域通常 0–100000

## image

```ts
{ type:'image', x, y, width, height, rotation?,
  data?: string,        // dataURL 或裸 base64
  src?: string,         // 远程 URL（生成端会下载为媒体）
  extension?: string,
  href?: string,        // 图片级超链接（'#N' 内部跳转）
  crop?: { l?, r?, t?, b? },                              // 百分比 0–100
  imageAdjust?: { brightness?(-100..100), contrast?(-100..100), transparency?(0..100) } }
```

## chart

```ts
{ type:'chart', chartType, x, y, width, height,
  title?, legend?, varyColors?, barDir?: 'bar'|'col',
  categories?: string[], series?: PptxChartSeries[],
  grouping?: 'clustered'|'stacked'|'percentStacked'|'standard',
  holeSize?, smooth?, marker?, ofPieType?: 'pie'|'bar',
  numberFormat?, bubble3D?, showNegBubbles?, bubbleScale?, wireframe? }
```

- `chartType`（ECMA-376 plotArea 下全部节点）：`barChart` `bar3DChart` `lineChart` `line3DChart` `areaChart` `area3DChart` `pieChart` `pie3DChart` `doughnutChart` `ofPieChart` `scatterChart` `bubbleChart` `radarChart` `stockChart` `surfaceChart` `surface3DChart`
- `PptxChartSeries`：`{ name?, values?: number[], x?, y?, open?, high?, low?, close?, color? }`（散点用 `x/y`，股票用 `open/high/low/close`）

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
`PptxTableCell`：`{ text?, paragraphs?, colSpan?, rowSpan?, fill?, border?, borders?{left,right,top,bottom}, align?, valign?: 'top'|'middle'|'bottom', fontSize?, color?, bold?, italic?, underline?, fontFace?, inset? }`
（单元格 `borders.left` 等可为 `{color,width}` 或 `'none'`）

常用 `tableStyleId`：
- `{5940675A-B579-460E-94D1-54222C63F5DA}` — Table Grid（黑色网格，库内默认）
- `{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}` — Medium Style 2 - Accent 1（强调色表头 + 白色网格线 + 交替行）

## diagram（SmartArt）

```ts
{ type:'diagram', x, y, width, height,
  diagramType?: 'list'|'hierarchy'|'process'|'cycle'|'pyramid',   // 创作端
  nodes?: { text, children? }[],                                    // 创作端层级
  texts?: string[], dataPath?: string }                             // 解析端
```

解析端只保留可读文本 + `__raw`，还原依赖原始回写。

## group / video / audio / raw

```ts
{ type:'group', x, y, width, height,
  childrenCoordinates?: 'local'|'page',   // 默认 'local'（子坐标为组内局部坐标）
  children: PptxElement[] }

{ type:'video'|'audio', x, y, width, height, data?, src?, extension?,
  poster?: { data?, src?, extension? } }   // 仅 video 有 poster

{ type:'raw', x, y, width, height }        // 解析兜底，始终按 __raw 回写
```

## 单位换算速查

| 场景 | 换算 |
|---|---|
| EMU → px | `px = emu × 96 / 914400`（`emu / 9525`） |
| px → EMU | `emu = px × 9525` |
| 表格单元格默认内边距 | 左右 0.1in = 91440 EMU = 9.6px；上下 0.05in = 45720 EMU = 4.8px |
| 线宽 1pt | 12700 EMU |
| OOXML 千分比属性（`a:tile@sx`、`a:srcRect@l`） | `值 = 比例 × 100000` |
