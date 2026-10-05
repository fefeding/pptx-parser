# 元素配方（可直接抄）

所有片段来自 `examples/generate-test-pptx.mjs`（T1–T19，21 页），已通过 93 项自检断言。坐标单位 px，字号/线宽 pt。

## 文本

```js
// 富文本：runs 混合样式
{ type: 'text', x: 60, y: 120, width: 600, height: 60,
  runs: [
    { text: '普通 ' },
    { text: '加粗红色', bold: true, color: '#ef4444', fontSize: 20 },
    { text: ' 链接', href: 'https://example.com', underline: true, color: '#2563eb' }
  ] }

// 编号 + 行距 + 段间距 + 缩进
{ type: 'text', x: 60, y: 120, width: 1040, height: 420, fontSize: 18, color: '#334155',
  paragraphs: [
    { text: '编号第一项', bullet: { type: 'number', fmt: 'ea1ChsPeriod' } },  // 一、二、
    { text: '缩进 + 行距28pt + 段前/后10', bullet: true, indent: 40,
      lineSpacing: { type: 'pt', value: 28 }, spaceBefore: 10, spaceAfter: 10 },
    { text: '普通项（无符号）', bullet: false }
  ] }

// 文本框内边距 + 竖排（textDirection 必须是合法 ST_TextVerticalType）
{ type: 'text', x: 360, y: 120, width: 160, height: 320, text: '竖排文字',
  fontSize: 22, color: '#7c3aed', textDirection: 'eaVert', inset: { l: 20, r: 20, t: 20, b: 20 } }

// run 级描边 + 外阴影 + 动态字段
{ type: 'text', x: 60, y: 120, width: 600, height: 60,
  runs: [
    { text: '描边文字', outline: { color: '#1e293b', width: 0.5 }, fontSize: 28, color: '#fbbf24' },
    { text: ' 带阴影', shadow: { color: '#000000', blur: 4, x: 2, y: 2, alpha: 60 } },
    { text: ' 第', break: false }, { text: '', field: 'slidenum' }, { text: ' 页' }  // 动态页码
  ] }

// 分栏 + 自动适配 + 艺术字变形
{ type: 'text', x: 60, y: 120, width: 400, height: 200, text: '分栏文本…',
  numCol: 2, spcCol: 12, autofit: 'normal', fontScale: 90,
  prstTxWarp: 'textArchUp' }
```

## 形状

```js
// 渐变填充 + 边框
{ type: 'shape', shapeType: 'rect', x: 60, y: 120, width: 240, height: 150,
  fill: { type: 'gradient', direction: 'horizontal',
          stops: [{ position: 0, color: '#6366f1' }, { position: 1, color: '#ec4899' }] },
  line: { color: '#111827', width: 1 } }

// 纯色 + 透明度 40%
{ type: 'shape', shapeType: 'roundRect', x: 340, y: 120, width: 240, height: 150,
  fill: { type: 'solid', color: '#10b981', transparency: 40 }, line: { color: '#047857', width: 1.5 } }

// 阴影 / 发光
{ type: 'shape', shapeType: 'ellipse', x: 620, y: 120, width: 160, height: 160, fill: { color: '#3b82f6' },
  effects: { shadow: { blur: 10, distance: 5, angle: 90, color: '#000000', transparency: 40 } } }
{ type: 'shape', shapeType: 'diamond', x: 820, y: 120, width: 160, height: 160, fill: { color: '#f59e0b' },
  effects: { glow: { color: '#ef4444', blur: 15 } } }

// 几何调整值（key 必须是 OOXML gd 名）
{ type: 'shape', shapeType: 'roundRect', x: 60, y: 120, width: 240, height: 120,
  fill: { color: '#6366f1' }, adjust: { adj: 30000 } }
{ type: 'shape', shapeType: 'rightArrow', x: 340, y: 120, width: 240, height: 120,
  fill: { color: '#10b981' }, adjust: { adj1: 50000, adj2: 40000 } }

// 翻转
{ type: 'shape', shapeType: 'rect', x: 60, y: 120, width: 200, height: 140, fill: { color: '#3b82f6' }, flipH: true }

// 主题色引用（配合 jsonToPptx 的 options.theme 自定义主题）
{ type: 'shape', shapeType: 'rect', x: 60, y: 120, width: 240, height: 120,
  fill: { type: 'solid', color: 'scheme:accent1' }, line: { color: 'scheme:dk1', width: 1.5 } }

// 自定义几何（a:custGeom）：坐标可传 0~1 归一化或绝对值
{ type: 'shape', shapeType: 'rect', x: 60, y: 120, width: 200, height: 200, fill: { color: '#6366f1' },
  custGeom: { paths: [{ closed: true, commands: [
    { type: 'moveTo', x: 0.5, y: 0 },        // 顶点
    { type: 'lnTo', x: 1, y: 1 },            // 右下
    { type: 'lnTo', x: 0, y: 1 },            // 左下
    { type: 'close' }
  ] }] } }

// 三维挤出 + 相机视角
{ type: 'shape', shapeType: 'roundRect', x: 60, y: 120, width: 200, height: 120, fill: { color: '#3b82f6' },
  threeD: { shape: { extrusionHeight: 50, extrusionColor: '#1e40af', bevelTop: { preset: 'relaxedInset', width: 6, height: 6 } },
            scene: { camera: 'perspectiveRelaxed', rotX: 15, rotY: 20, lightRig: 'balanced' } } }
```

### 形状的图片填充 / 图案填充 / 平铺 / 裁剪

```js
// 图片铺满
{ type: 'shape', shapeType: 'rect', x: 60, y: 120, width: 240, height: 160,
  fill: { type: 'image', data: 'data:image/png;base64,…' } }

// 图案填充（prst = OOXML 预设图案名）
{ type: 'shape', shapeType: 'roundRect', x: 340, y: 120, width: 240, height: 160,
  fill: { type: 'pattern', prst: 'diagCross', fg: '#ef4444', bg: '#fef3c7' } }

// 平铺：sx/sy 为「每格占图片原始尺寸的比例」，sx=sy=1 即按原图尺寸重复
{ type: 'shape', shapeType: 'rect', x: 620, y: 120, width: 240, height: 160,
  fill: { type: 'image', data: 'data:image/png;base64,…', tile: { sx: 1, sy: 1 } } }

// 平铺 + 源图裁剪：只取图片居中 50% 作为一格
{ type: 'shape', shapeType: 'rect', x: 900, y: 120, width: 240, height: 160,
  fill: { type: 'image', data: 'data:image/png;base64,…',
          tile: { sx: 1, sy: 1 }, srcRect: { l: 0.25, t: 0.25, r: 0.25, b: 0.25 } } }
```

## 图片

```js
{ type: 'image', data: 'data:image/png;base64,…', x: 60, y: 120, width: 240, height: 240, name: 'Original' }
// 四边各裁 10%（l/r/t/b 为 0~1 比例）
{ type: 'image', data: 'data:image/png;base64,…', x: 360, y: 120, width: 240, height: 240,
  crop: { l: 0.1, r: 0.1, t: 0.1, b: 0.1 } }
// 亮度/对比度/透明度
{ type: 'image', data: 'data:image/png;base64,…', x: 660, y: 120, width: 240, height: 240,
  imageAdjust: { brightness: 20, contrast: 30, transparency: 10 } }
```

## 表格

```js
{ type: 'table', x: 60, y: 110, width: 1100, height: 280,
  colWidths: [280, 280, 280, 260],
  border: { color: '#94a3b8', width: 1 },             // 表格级四边统一
  rows: [
    { cells: [
      { text: '四边粗框', border: { color: '#ef4444', width: 2 } },
      { text: '上下右分边', borders: { top: { color: '#10b981', width: 2 },
                                     right: { color: '#10b981', width: 2 } } },
      { text: '无框', borders: { top: 'none', bottom: 'none', left: 'none', right: 'none' } },
      { text: '内边距 20px', inset: { l: 20, r: 20, t: 10, b: 10 } }
    ] },
    { cells: [
      { text: '对角线', borders: { diagonal: 'both' } },       // 'tlBr' | 'blTr' | 'both'
      { text: '垂直居中', valign: 'middle', align: 'center', fill: '#f1f5f9' },
      { text: '跨列', colSpan: 2, bold: true, fontSize: 16 }
    ] }
  ] }
```

指定样式（GUID 必须在 `ppt/tableStyles.xml` 有定义，生成端会自动补）：

```js
{ type: 'table', …, tableStyleId: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' }  // Medium Style 2 - Accent 1
{ type: 'table', …, tableStyleId: '{5940675A-B579-460E-94D1-54222C63F5DA}' }  // Table Grid（默认）
```

## 图表

```js
// 柱状（多系列 + 系列色 + 图例在右）
{ type: 'chart', chartType: 'barChart', title: '柱状（clustered+系列色）', legend: 'r',
  x: 40, y: 110, width: 380, height: 200, categories: ['Q1','Q2','Q3','Q4'],
  series: [{ name: '产品A', values: [10,20,15,25], color: '#6366f1' },
           { name: '产品B', values: [15,10,20,18], color: '#ec4899' }] }

// 折线（平滑 + 标记）
{ type: 'chart', chartType: 'lineChart', smooth: true, marker: true, … }

// 面积（堆叠）
{ type: 'chart', chartType: 'areaChart', grouping: 'stacked', … }

// 横向百分比堆叠 + 数字格式
{ type: 'chart', chartType: 'barChart', barDir: 'bar', grouping: 'percentStacked',
  numberFormat: '#,##0%', … }

// 饼图（多色 + 百分比标签）
{ type: 'chart', chartType: 'pieChart', varyColors: true, dataLabels: { showPercent: true }, … }

// 环图 / 子母饼 / 气泡 / 雷达 / 股票 / 曲面
{ type: 'chart', chartType: 'doughnutChart', holeSize: 30, dataLabels: { showValue: true }, … }
{ type: 'chart', chartType: 'ofPieChart', ofPieType: 'bar', … }
{ type: 'chart', chartType: 'bubbleChart', bubble3D: true, bubbleScale: 150,
  series: [{ name: 'b', x: [1,2,3,4], y: [2,3,1,5], values: [3,4,2,6] }] }
{ type: 'chart', chartType: 'radarChart', dataLabels: { showValue: true }, … }
{ type: 'chart', chartType: 'stockChart',
  series: [{ name: 'OHLC', open: [10,12,11,13], high: [15,16,14,17],
             low: [8,10,9,11], close: [12,11,13,15] }] }
{ type: 'chart', chartType: 'surfaceChart', wireframe: true, … }   // 3D 曲面需 ≥2 系列
```

`legend` 位置：`'r'`（右）/ `'b'`（底）/ `false`。3D 变体：`bar3DChart` `line3DChart` `area3DChart` `pie3DChart` `surface3DChart`。

```js
// 三维视角 + 坐标轴标题
{ type: 'chart', chartType: 'bar3DChart', title: '3D 柱状',
  view3D: { rotX: 30, rotY: 20, depthPercent: 150, rAngAx: false },
  axisTitles: { category: '季度', value: '金额（万）' },
  … }

// 逐点配色（c:dPt）+ 趋势线
{ type: 'chart', chartType: 'barChart', categories: ['Q1','Q2','Q3','Q4'],
  series: [{ name: '产品A', values: [10, 20, 15, 25],
             pointColors: ['#ef4444', undefined, '#10b981', undefined],  // 第 1/3 点覆盖
             trendlines: [{ type: 'linear', name: '趋势', showEquation: true }] }] }

// 次坐标轴（双轴图）
{ type: 'chart', chartType: 'barChart', secondaryValueAxis: true,
  axisTitles: { value: '销量', secondaryValue: '增长率%' },
  series: [{ name: '销量', values: [100, 120, 130] },
           { name: '增长率', values: [10, 20, 8], axis: 'secondary' }] }

// 网格线 + 图表区填充
{ type: 'chart', chartType: 'lineChart', gridlines: { major: true, minor: false },
  spaceFill: '#f8fafc', … }
```

## 组合 / 图示 / 媒体

```js
{ type: 'group', x: 80, y: 120, width: 600, height: 320, childrenCoordinates: 'local',
  children: [
    { type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 200, height: 100, fill: { color: '#60a5fa' } },
    { type: 'text', x: 20, y: 120, width: 320, height: 40, text: '组合内文本', fontSize: 18 }
  ] }

// SmartArt：diagramType = list | hierarchy | process | cycle | pyramid
{ type: 'diagram', diagramType: 'hierarchy', x: 60, y: 120, width: 820, height: 400,
  nodes: [{ text: '根', children: [{ text: '子A' }, { text: '子B', children: [{ text: '孙' }] }] }] }

{ type: 'video', data: 'data:video/mp4;base64,…', extension: 'mp4',
  x: 60, y: 120, width: 360, height: 210, poster: { data: 'data:image/png;base64,…' } }
{ type: 'audio', data: 'data:audio/m4a;base64,…', extension: 'm4a', x: 480, y: 120, width: 200, height: 200 }
```

### 连接线 / OLE / 公式

```js
// 连接线（p:cxnSp）：用 start/end 精确端点
{ type: 'connector', shapeType: 'straightConnector1',
  start: { x: 100, y: 100 }, end: { x: 300, y: 200 },
  line: { color: '#64748b', width: 1.5 } }

// 弯折连接器 + 调整值
{ type: 'connector', shapeType: 'bentConnector3',
  x: 100, y: 100, width: 200, height: 100,
  line: { color: '#3b82f6', width: 1 }, adjust: { adj1: 50000 } }

// OLE 嵌入 Excel
{ type: 'ole', x: 100, y: 100, width: 400, height: 300,
  progId: 'Excel.Sheet.12', data: 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,…',
  extension: 'xlsx', showAsIcon: false,
  poster: { data: 'data:image/png;base64,…' } }

// 公式（OMML）
{ type: 'math', x: 100, y: 100, width: 300, height: 60,
  omml: '<m:oMathPara xmlns:m="…"><m:oMath><m:r><m:t>E=mc^2</m:t></m:r></m:oMath></m:oMathPara>' }
```

## 页级能力

```js
{
  background: '#ffffff',                       // 或 {type:'gradient',…} / {type:'image', data, tile, srcRect}
  transition: { type: 'fade', duration: 1000, advanceOnClick: false,
                advanceAfterTime: 3000, direction: 'l',
                sound: { name: 'applause' } },
  advanceTime: 3000,                           // 自动播放停留毫秒
  animations: [
    { target: 1, type: 'flyIn', duration: 0.5, presetClass: 'entr',
      presetSubtype: 4, direction: 'l',        // 4='from left'
      trigger: { type: 'onClick' } },
    { target: 2, type: 'fade', duration: 0.8, presetClass: 'entr',
      trigger: { type: 'afterPrev', delay: 0.5 }, repeat: 3 }
  ],
  hidden: true,                                // 隐藏该页（p:sld show="0"）
  notes: '备注文本',
  comments: [{ author: 'Alice', text: '这是一条批注', dt: '2026-01-01T00:00:00Z' }],
  elements: [ … ]
}
```

## 文档级

```js
const data = await jsonToPptx({
  slideSize: { width: 1280, height: 720 },
  metadata: { title: '报告', author: 'me', subject: '…', keywords: 'a;b', description: '…' },
  customProps: { '部门': '研发', '版本': '1.0' },     // 写回 docProps/custom.xml
  // 语义级主题（生成端据此构造 themeN.xml）
  theme: { name: '自定义', colors: { dk1: '#1e293b', lt1: '#ffffff', accent1: '#6366f1', accent2: '#ec4899' },
           fonts: { major: { latin: 'Arial', ea: '微软雅黑' }, minor: { latin: 'Arial', ea: '微软雅黑' } } },
  // 母版/版式（幻灯片通过 layout 索引引用）
  masters: [{ name: '母版1', background: '#f8fafc',
              placeholders: [{ type: 'title', x: 60, y: 20, width: 1160, height: 60, fontSize: 36, bold: true }],
              layouts: [{ name: '标题页', placeholders: [{ type: 'title', x: 200, y: 300, width: 880, height: 80 }] }] }],
  // 文档节
  sections: [{ name: '第一部分', slides: [0, 1, 2] }, { name: '第二部分', slides: [3, 4] }],
  // 嵌入字体（生成端自动做 XOR 混淆）
  fonts: [{ name: 'Source Han Sans', data: 'base64…', panose: '02000000040000000000', embedType: 'subset' }],
  slides
}, { outputType: 'uint8array' });
```

## 修改已有 PPTX

```js
const editor = await editPptx(fileData);
await editor.deleteSlide(3);
await editor.moveSlide(1, 5);
await editor.setMetadata({ title: '新标题' });
await editor.addSlide({ background: '#fff', elements: [{ type: 'text', x: 60, y: 60, width: 600, height: 50, text: '新增页' }] });
const out = await editor.save({ outputType: 'uint8array' });
```
