# @fefeding/ppt-parser

一个轻量级的 PPTX 解析库，让处理 PowerPoint 文件变得简单。基于纯 TypeScript 编写，零框架依赖，同时支持浏览器和 Node.js 环境。

## 特性

- **简单好用** — 只需几行代码即可解析和转换 PPTX 文件
- **零依赖** — 不绑定任何框架，适用于任何 JavaScript/TypeScript 项目
- **双向转换** — 支持 PPTX 到 HTML 或 JSON 的解析，也能把 JSON（或链式 `PPTXComposer`）序列化回合法的 PPTX 文件
- **元素丰富** — 文本、形状、表格、图片、图表等全面支持
- **智能单位转换** — 自动处理 EMU 到 PX 的单位转换
- **通用模块** — 同时支持 ESM 和 CommonJS
- **浏览器 & Node.js** — 两种环境均可无缝运行

## 安装

```bash
npm install @fefeding/ppt-parser
```

或直接下载 `dist` 目录下的文件在浏览器中使用。

## 快速开始

### 解析 PPTX 为 HTML

```javascript
import { pptxToHtml } from '@fefeding/ppt-parser';

const fileInput = document.querySelector('#ppt-upload');

fileInput.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  const fileData = await file.arrayBuffer();
  const result = await pptxToHtml(fileData, {
    mediaProcess: true,
    themeProcess: true
  });

  // result.slides 包含解析后的幻灯片 HTML
  console.log('幻灯片数量:', result.slides.length);
  console.log('幻灯片尺寸:', result.slideSize);
  console.log('元数据:', result.metadata);
  console.log('图表:', result.charts);

  // 渲染所有幻灯片
  const container = document.getElementById('preview');
  result.slides.forEach(slide => {
    const div = document.createElement('div');
    div.innerHTML = slide.html;
    container.appendChild(div);
  });
});
```

### 解析 PPTX 为 JSON

```javascript
import { pptxToJson } from '@fefeding/ppt-parser';

const result = await pptxToJson(fileData);
console.log('幻灯片数量:', result.slides.length);
console.log('幻灯片尺寸:', result.slideSize);
console.log('元数据:', result.metadata);
```

### 提取文件内容

```javascript
import { pptxToFiles } from '@fefeding/ppt-parser';

const result = await pptxToFiles(fileData);
console.log('文件列表:', result.files);
console.log('内容:', result.content);
```

## 配置选项

```typescript
interface PptxParserOptions {
  // 是否处理媒体文件（图片等）
  mediaProcess?: boolean;

  // 主题处理方式
  themeProcess?: boolean | 'colorsAndImageOnly';

  // 幻灯片尺寸调整
  incSlide?: {
    width: number;
    height: number;
  };

  // 自定义样式表
  styleTable?: Record<string, { name: string; text: string; suffix?: string }>;

  // 回调函数
  callbacks?: {
    onFileStart?: () => void;
    onError?: (error: { type: string; message: string }) => void;
    onSlide?: (data: any, info: { slideNum: number; fileName: string }) => void;
    onThumbnail?: (thumbnail: string | null) => void;
    onSlideSize?: (slideSize: { width: number; height: number }) => void;
    onGlobalCSS?: (css: string) => void;
    onComplete?: (info: {
      executionTime: number;
      slideWidth: number;
      slideHeight: number;
      styleTable: any;
      settings: PptxParserOptions;
    }) => void;
  };
}
```

## 序列化回 PPTX

### 链式构建（Composer）

```javascript
import { PPTXComposer } from '@fefeding/ppt-parser';

const composer = new PPTXComposer();
composer
  .title('My Deck')
  .author('me')
  .addSlide(slide => {
    slide.background('#ffffff');
    slide.addText(t => t.value('Hello World').x(100).y(80).fontSize(28).bold());
    slide.addShape({ shapeType: 'roundRect', x: 100, y: 400, width: 200, height: 80, fill: { color: '#4f46e5' } });
    slide.addImage({ data: dataUrl, x: 500, y: 400, width: 100, height: 100 });
  })
  .addSlide(slide => {
    // 超链接：外部 URL，或 '#N' 跳转到第 N 页
    slide.addText(t => t.runs([
      { text: 'External link', options: { href: 'https://example.com' } },
      { text: ' / ', options: {} },
      { text: 'Jump to slide 1', options: { href: '#1' } }
    ]).x(100).y(100));
  });

const data = await composer.save(); // Uint8Array
```

也支持对象式配置：`slide.addText({ text: 'Hi', x: 0, y: 0 })`。

### JSON 转 PPTX

```javascript
import { jsonToPptx } from '@fefeding/ppt-parser';

const data = await jsonToPptx({
  metadata: { title: 'My Deck', author: 'me' },
  slideSize: { width: 1280, height: 720 },  // px，默认 16:9
  slides: [
    {
      background: '#ffffff',
      elements: [
        { type: 'text', x: 100, y: 80, width: 600, height: 60, text: 'Title\nSubtitle', fontSize: 24, color: '#1e293b' },
        { type: 'shape', shapeType: 'ellipse', x: 600, y: 300, width: 150, height: 150, fill: { color: '#ed7d31' }, line: { color: '#000', width: 1 } },
        { type: 'image', data: dataUrl, x: 100, y: 300, width: 200, height: 150 }
      ]
    }
  ]
});
```

支持的元素类型：`text`（多段落、run 内字号/颜色/粗体/斜体/下划线/字体、超链接、项目符号、编号列表、描边、阴影、动态字段）、预设形状（渐变/纯色/图片/图案填充、平铺/裁剪、阴影/发光、三维挤出、自定义几何）、图片（dataURL / base64 / 远程 `src`、裁剪/亮度/对比度/透明度）、表格（单元格边框/对角线/跨行跨列/合并/样式）、图表（2D/3D、多系列、趋势线、次坐标轴、逐点配色、网格线）、组合、图示（SmartArt）、连接线、视频/音频、OLE 嵌入对象、公式、批注、自定义文档属性等。

### 编辑已有 PPTX

```javascript
import { editPptx } from '@fefeding/ppt-parser';

const editor = await editPptx(fileData);
await editor.deleteSlide(2);                       // 删除第 2 页
await editor.moveSlide(1, 3);                      // 调整顺序
await editor.addSlide({ elements: [/* 同上的元素格式 */] });
await editor.setMetadata({ title: 'Updated', author: 'me' });
const updated = await editor.save();
```

生成的文件可经 `pptxToJson` / `pptxToHtml` 往返解析。

## Headless 编辑器内核

除命令式封装 `editPptx` 外，库内置一套**与 UI 完全无关**的编辑器内核（`src/editor`，已随包发布）。它把文档模型、PPTX 双向转换、编辑操作、图表渲染与元素几何从具体渲染层下沉到库本身，因此任意前端框架或自有 UI 都能直接复用同一套业务逻辑，无需重写。

```javascript
import {
  createStore, createActions, docFromPptx, docToPptx, renderChartSVG,
  elementRect, rotatedRect, effectMargin, absoluteElementRect,
  normalizeDoc, normalizeElement,
  EditorStore
} from '@fefeding/ppt-parser';

// 1) 创建无 UI 的文档状态（可选持久化：传入实现 getItem/setItem 的 storage 适配）
const store = createStore({ storage: localStorage });
const sem = await pptxToJson(fileData, { mode: 'semantic' });
store.setDoc(docFromPptx(sem.document));

// 2) 创建编辑操作（工厂函数，不依赖单例，便于多实例/测试）
const actions = createActions(store);
actions.addSlide({ elements: [/* 与上文一致的元素格式 */] });
actions.alignElements('hcenter');
actions.updateElement(id, { fill: { color: '#4f46e5' } });

// 3) 导出为 PPTX
const pptxDoc = docToPptx(store.doc);
const data = await jsonToPptx(pptxDoc);

// 4) 图表与几何（纯函数，宿主自行决定如何挂载）
const svg = renderChartSVG(chartElement); // 返回 SVG 字符串
const rect = elementRect(element);        // 元素实际矩形（组合取子元素并集）
```

核心 API：

- `createStore(options?)` → `EditorStore`：持有 `doc`，提供 `setDoc / getDoc / update`，并可选持久化（传入 `storage` 适配 `getItem/setItem`）。
- `createActions(store)` → `EditorActions`：返回 `addSlide / deleteSelected / duplicateSelected / copySelected / paste / selectAll / nudge / zOrder / alignElements / distribute / groupSelection / ungroupSelection / toggleLock / toggleHidden / addSlide / duplicateSlide / deleteSlide / moveSlide / toggleSlideHidden / applyLayout / setBackground / setSlideSize / applyTheme / setNotes / applyTextStyleSel / setBackgroundImage / setElementGeo / resizeTable / updateElement / findInDoc / setTransition / addAnimation / updateAnimation / removeAnimation / moveAnimation`。
- `docFromPptx(semanticDoc, { fileName? })` / `docToPptx(doc)`：编辑器文档 ↔ 标准 `PptxDocument` 的双向转换。`docFromPptx` 的标题优先级为 `core.xml 的 dc:title` > `fileName`（自动去扩展名）> `'导入的演示文稿'`；当原文件无标题元信息时，以文件名兜底，并在导出时一并写回 `core.xml`，保证来回一致。
- `renderChartSVG(chartEl)`：把图表元素渲染为 SVG 字符串（宿主决定如何挂载到 DOM）。
- `elementRect / rotatedRect / effectMargin / absoluteElementRect`：元素几何计算（组合并集矩形、旋转包围盒、阴影/发光外边距、相对某父级的绝对矩形）。
- `normalizeDoc(doc)` / `normalizeElement(el)`：导入/合并后补齐缺省字段（`slideSize`、`elements`、`zIndex`、元素 `id`/`name` 等），确保文档完整。
- `EditorStore`（类型）：`createStore` 返回，结构为 `{ doc, setDoc, getDoc, update, storage? }`。

## 支持的元素

- **文本** — 多段落、run 内字号/颜色/粗体/斜体/下划线/字体、超链接（外部 URL 或 `#N` 跳转到第 N 页）、项目符号与编号列表、行距、缩进、文本框内边距、竖排文字方向、分栏、自动适配、艺术字变形、run 级描边与阴影、动态字段（页码/日期）
- **形状** — 全部预设几何（`rect`、`roundRect`、`ellipse`、`triangle`、`arrow`、`star5`、`foldedCorner` 等）；渐变 / 纯色 / 图片 / 图案填充；图片填充支持平铺（`tile`）与源矩形裁剪（`srcRect`）；线型、阴影与发光效果；几何调整（`avLst`）；水平/垂直翻转；自定义几何路径（`custGeom`）；三维挤出/斜面/相机/光照（`threeD`）
- **图片** — PNG、JPEG、GIF、BMP、WEBP、SVG；base64/dataURL 或远程 `src`；裁剪、亮度/对比度/透明度
- **表格** — 完整样式：单元格边框、对角线边框（`tlBr` / `blTr` / `both`）、单元格内边距、跨行跨列（`colSpan`/`rowSpan`）、合并（`hMerge`/`vMerge`）、填充、对齐、表格样式（`tableStyleId`）
- **图表** — 柱状/条形、折线、面积、饼/环、子母饼、散点、气泡、雷达、股票、曲面（2D 与 3D）；多系列、系列颜色、逐点配色（`pointColors`）、图例、数据标签、分组、`barDir`、`holeSize`、`ofPieType`、`bubble3D`、`wireframe`、三维视角（`view3D`）、次坐标轴、趋势线、网格线、坐标轴标题
- **组合** — `group` 元素，支持 `childrenCoordinates: 'local' | 'page' | 'relative'`
- **图示** — SmartArt（list / hierarchy / process / cycle / pyramid），含缓存绘图形状与连接线
- **连接线** — `connector`（`straightConnector1` / `bentConnector3` / `curvedConnector2`），支持精确起止端点
- **媒体** — 视频（`mp4`/`m4v`）与音频（`m4a`/`mp3`），可选封面图
- **OLE** — 嵌入对象（Excel 表格等），可显示为图标
- **公式** — OMML 数学公式
- **幻灯片级** — 背景（纯色/渐变/图片）、切换（含声音/方向）、自动播放停留、动画（entr/exit/emph/path + 触发时机）、隐藏页、备注、批注、自定义文档属性
- **文档级** — 母版/版式/占位符、文档分节、嵌入字体（自动 XOR 混淆）、语义级主题（配色/字体）

## 平台使用方式

### 浏览器

```html
<script src="./dist/ppt-parser.browser.js"></script>
<script>
  const result = await pptxParser.pptxToHtml(fileData);
</script>
```

### Node.js

```javascript
const fs = require('fs');
const { pptxToHtml } = require('@fefeding/ppt-parser');

const buffer = fs.readFileSync('presentation.pptx');
const result = await pptxToHtml(buffer);
```

### Vue 示例

```vue
<script setup>
import { pptxToHtml } from '@fefeding/ppt-parser';

async function handleUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const fileData = await file.arrayBuffer();
  const result = await pptxToHtml(fileData, {
    mediaProcess: true,
    themeProcess: true
  });

  slides.value = result.slides;
}
</script>
```

> 如需在 Vue 中渲染图表，请将 `examples/chart-lib/chart-renderer.js` 复制到项目中，并将 `echarts` 设为全局依赖。完整示例请参见 `examples/vue-demo` 目录。

## 开发

```bash
# 克隆项目
git clone https://github.com/fefeding/pptx-parser.git
cd pptx-parser

# 安装依赖
npm install

# 开发模式
npm run dev

# 构建
npm run build

# 运行测试
npm test
```

## TypeScript

本包内置 TypeScript 类型定义，所有接口和类型均从包入口导出。

## 浏览器兼容性

- Chrome >= 80
- Firefox >= 75
- Edge >= 80
- Safari >= 14

## 许可证

[MIT](LICENSE)

## AI 技能

本仓库在 [`skills/pptx-parser/`](skills/pptx-parser) 下提供了一份开箱即用的 AI 编程助手技能，同时也随 npm 包一起发布（`package.json` 的 `files` 中已包含 `skills`），因此任何依赖 `@fefeding/ppt-parser` 的项目都能直接获得它。

该技能完整描述了**把 PPTX 解析为 HTML / JSON / 标准 JSON、由 JSON 生成或编辑 PPTX、以及修复 OOXML 合规问题**的精确契约——包含可直接抄的元素配方、完整的字段参考，以及一份真实的「坑位清单」（例如 `tableCellInsets` 不是合法 OOXML、`a:lnTlToBr` 本身就是一条线、`tableStyleId` 必须在 `tableStyles.xml` 中有对应定义等）。

```
skills/pptx-parser/
├── SKILL.md                 # 入口：能力矩阵、快速上手、验证工作流、已知限制
├── references/
│   ├── api.md               # 全部导出 API 的签名、选项、返回结构
│   ├── json-schema.md       # PptxDocument 字段契约 + 单位换算
│   ├── cookbook.md          # 元素配方（文本/形状/图片/表格/图表/组合/图示/媒体/……）
│   └── ooxml-pitfalls.md    # 真实「被 PowerPoint/WPS 静默忽略」案例 + 自查清单
└── scripts/
    ├── pptx-info.mjs        # 打印文件概览（页数、元素、文本、元数据、图表）
    ├── pptx-to-html.mjs     # 渲染为自包含 HTML（Node 下需 jsdom）
    └── json-to-pptx.mjs     # JSON → PPTX，并支持 --check 往返自检
```

当你需要用代码预览、提取、生成或修改 PPTX 时，让 AI 助手读取 `skills/pptx-parser/SKILL.md`（或将其配置为一个技能）即可。

## 致谢

本项目受 [pptxjs](https://github.com/meshesha/pptxjs) 项目启发，感谢原作者提供的架构基础。
