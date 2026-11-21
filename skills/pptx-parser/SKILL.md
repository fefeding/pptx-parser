---
name: pptx-parser
description: 使用 @fefeding/ppt-parser 处理 PowerPoint 文件的完整指南——解析 PPTX 为 HTML / 结构化 JSON / 标准 JSON（PptxDocument），由 JSON 生成或编辑 PPTX，把 PPTX 转换成 DOCX / PDF / Markdown / HTML（插件化转换器注册表），复用 Headless 编辑器内核，以及排查渲染差异与 OOXML 合规问题。当需要预览 PPT、提取内容与图表数据、批量生成或修改 PPT、转换文档格式、做 round-trip 保真验证时使用。
license: MIT
---

# PPTX 解析与生成（@fefeding/ppt-parser）

一个零框架依赖的纯 TypeScript PPTX 库，浏览器与 Node.js 同构。核心能力是**双向转换**：

```
PPTX ──pptxToStandard──> PptxDocument JSON ──jsonToPptx──> PPTX
PPTX ──pptxToHtml──────> HTML + CSS（浏览器预览）
PPTX ──pptxToFiles─────> 部件清单（xml / 图片 base64 / 二进制）
PPTX ──editPptx────────> 页级增删改 → 新 PPTX（不动语义层）
PPTX ──convert(target)─> DOCX / PDF / Markdown / HTML（插件化转换引擎）
```

## 何时使用

- 需要在网页里预览 PPTX（渲染成 HTML，配 `styles.global` 的 CSS）
- 需要提取 PPTX 的文本、表格、图表数据、备注、批注、自定义属性
- 需要用代码生成 PPTX（报告、批量导出、模板填充）
- 需要修改已有 PPTX（删页、移动页、追加页、改元数据）
- 需要把 PPTX 转成 DOCX / PDF / Markdown / HTML，或注册自定义转换格式
- 需要在自有 UI 里复用编辑器内核（状态、操作、几何、吸附、SmartArt 文本）
- 渲染结果与 WPS/PowerPoint 不一致，需要定位是生成端 XML 不合规还是解析端未支持

## 安装与导入

```bash
npm install @fefeding/ppt-parser
```

```js
// ESM / TypeScript
import { pptxToHtml, pptxToJson, pptxToStandard, pptxToFiles, jsonToPptx, editPptx, PPTXComposer } from '@fefeding/ppt-parser';
// 默认导出即 pptxToHtml
import pptxToHtml from '@fefeding/ppt-parser';

// CJS
const { jsonToPptx } = require('@fefeding/ppt-parser');

// 格式转换（独立子路径）
import { convert, converters, registerConverter } from '@fefeding/ppt-parser/convert';

// 浏览器 <script src="./dist/ppt-parser.browser.js"></script> → 全局变量 pptxParser
// 浏览器转换 <script src="./dist/convert.browser.js"></script>
```

**输入数据**：`ArrayBuffer | Uint8Array | Buffer | string`（底层是 JSZip，Node 下可直接 `fs.readFile`；浏览器下用 `file.arrayBuffer()`）。

> Node 下把 Buffer 转 ArrayBuffer 时注意切片：`buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)`，否则共享内存池会读到脏数据。

## 三个最小示例

### 1. 解析为 HTML 预览

```js
const result = await pptxToHtml(fileData, { mediaProcess: true, themeProcess: true });
result.slides.forEach(s => { const div = document.createElement('div'); div.innerHTML = s.html; container.appendChild(div); });
// 必须同时注入全局 CSS，否则样式丢失
const style = document.createElement('style'); style.textContent = result.styles.global;
```

### 2. 解析为标准 JSON，再生成回去（round-trip）

```js
const doc = await pptxToStandard(fileData);                 // PptxDocument
const out = await jsonToPptx(doc, { outputType: 'uint8array' });
fs.writeFileSync('out.pptx', Buffer.from(out));
```

### 3. 从零生成一页 PPT

```js
const data = await jsonToPptx({
  slideSize: { width: 1280, height: 720 },
  metadata: { title: '月报', author: 'me' },
  slides: [{
    background: '#ffffff',
    elements: [
      { type: 'text',  x: 60, y: 40, width: 1100, height: 60, text: '月度报告', fontSize: 32, bold: true },
      { type: 'shape', shapeType: 'roundRect', x: 60, y: 140, width: 520, height: 300,
        fill: { type: 'gradient', direction: 'horizontal',
                stops: [{ position: 0, color: '#6366f1' }, { position: 1, color: '#ec4899' }] },
        line: { color: '#111827', width: 1 } },
      { type: 'chart', chartType: 'barChart', x: 620, y: 140, width: 540, height: 300,
        title: '销量', categories: ['1月', '2月', '3月'],
        series: [{ name: '线上', values: [120, 200, 160] }, { name: '线下', values: [80, 90, 110] }] }
    ]
  }]
});
```

## 能力矩阵

| 需求 | API | 说明 |
|---|---|---|
| 网页预览 | `pptxToHtml` | 返回每页 HTML + 全局 CSS；图表需额外用 echarts 渲染 |
| 结构化数据 | `pptxToJson` | 返回简化 XML 树（`slide.data`），适合逐个节点精读 |
| 语义 JSON（可回写） | `pptxToStandard` | 返回 `PptxDocument`，与 `jsonToPptx` 同源，可无损往返 |
| 部件清单 | `pptxToFiles` | 所有 zip 条目；xml 给文本，图片给 base64 + dataUrl |
| 生成 PPTX | `jsonToPptx` | 输入 `PptxDocument` 或 `PPTXComposer` 实例 |
| 链式构建 | `PPTXComposer` | `.slideSize().title().addText()…` 最后交给 `jsonToPptx` |
| 编辑已有文件 | `editPptx` | 删页/移动页/追加页/改元数据，`save()` 输出 |
| 格式转换（子路径） | `convert(input, target, options)` | `@fefeding/ppt-parser/convert`；内置 `docx`/`pdf`/`md`/`html`，可 `registerConverter` 扩展 |
| 编辑器内核 | `createStore` / `createActions` / `snapBox` … | 与 UI 无关的文档状态、编辑操作与几何计算，见下文 |

完整签名、选项与返回结构见 [`references/api.md`](references/api.md)。格式转换详见 [`references/api.md`](references/api.md) 的「格式转换」一节。

## 元素类型速查

`slides[i].elements` 支持：`text`、`shape`、`image`、`chart`、`table`、`diagram`（SmartArt）、`group`、`connector`（连接线）、`video`、`audio`、`ole`（嵌入对象）、`math`（公式）、`raw`（解析兜底）。

- 坐标 `x/y/width/height` 统一为 **px**（96 DPI）；内部按 `SLIDE_FACTOR = 96/914400` 与 EMU 互转
- 字体大小统一为 **pt**，线宽为 **pt**，透明度/裁剪/平铺为 **0~1 或 0~100** 的比例（见下文各字段注释）
- 颜色用 `#RRGGBB`（也支持 `scheme:accent1` 这类主题色引用写法，解析端会按该页真实主题解析为绝对色）
- 文本 run 支持 `outline`（描边）、`shadow`（外阴影）、`field`（动态字段如页码）、`break`（软换行）
- 形状支持 `custGeom`（自定义几何路径）、`threeD`（三维挤出/相机/光照）、`effects`（阴影/发光）
- 图表支持 `view3D`（三维视角）、`secondaryValueAxis`（双轴）、`trendlines`（趋势线）、`pointColors`（逐点配色）、`gridlines`（网格线）
- 文档级支持 `masters`（母版/版式/占位符）、`sections`（分节）、`fonts`（嵌入字体）、`theme`（语义级主题对象）

逐个字段的权威定义见 [`references/json-schema.md`](references/json-schema.md)；可直接抄的片段见 [`references/cookbook.md`](references/cookbook.md)。

## 格式转换（PPTX → DOCX / PDF / MD / HTML）

独立于主入口的子路径 `@fefeding/ppt-parser/convert`，采用**可插拔转换器注册表**：内置四种，可覆盖或新增。

```js
import { convert, converters, registerConverter } from '@fefeding/ppt-parser/convert';

const docxBytes = await convert(fileData, 'docx', { title: 'Deck', autoFont: true }); // Uint8Array → 写文件
const md = await convert(fileData, 'md');                                              // string

converters.list().map(c => c.id);   // ['docx', 'md', 'pdf', 'html']
registerConverter({ id: 'my', ext: 'my', async convert(input, options) { return /* bytes|string */; } });
```

**输入**：`ArrayBuffer | Uint8Array | Buffer | Blob | base64 | data URL`。**输出**：二进制目标（docx/pdf）返回 `Uint8Array`，文本目标（md/html）返回 `string`。

**路径与保真要点**（排障时先确认走的是哪条链路）：
- `pdf` / `html`：幻灯片版式直排，一页 = 一张幻灯片，还原度最高；主题色 `schemeClr` 按该页真实主题解析。
- `docx` / `md`：走块级中间表示 `pptxToBlocks`。DOCX 每页插 `pageBreak`、页/小标题映射 `Heading 1`/`Heading 2`（保留原 run 字号/颜色）；Markdown 用 `hr` 分隔页。
- 表格单元格边框（颜色/线宽/`none`）已全链路贯通：PPTX → block IR → DOCX 单元格边框 / HTML 行内 `border-*` / PDF 四边 `drawLine`。
- 中文 PDF：`autoFont` 默认 `true`（仅 Node 自动探测系统字体）；浏览器调用方需显式传 `fonts`（TTF 字节），WOFF2/EOT 会被拒绝并报错。
- `convert(target)` 未注册会抛 `TypeError` 并列出已注册 id；其它错误经 `wrapError` 包装为 `ConvertError`。

## 编辑器内核（可选复用）

除 `editPptx` 页级编辑外，包入口还导出了与 UI 无关的编辑器内核：`createStore` / `createActions` / `docFromPptx` / `docToPptx` / `renderChartSVG` / `elementRect` / `rotatedRect` / `effectMargin` / `absoluteElementRect` / `snapBox` / `animDurationMs` / `updateDiagramTexts` / `normalizeDoc` / `normalizeElement` / `base64ToBytes` / `bytesToBase64` / `stripDataUrl`（类型 `EditorStore` / `SnapBoxRect` / `SnapGuide` / `SnapResult` / `DiagramTextUpdate`）。适合在自有 UI 中复用文档模型与 PPTX 双向转换逻辑，而不必依赖 `examples/editor`。完整签名见 [`references/api.md`](references/api.md) 的「编辑器内核」一节。

导入时建议把打开的文件名传给 `docFromPptx(pptxDoc, { fileName })`：这样没有标题元信息的 PPTX 也能以文件名作标题，且导出时会写回 `core.xml`，来回一致。

**异步对话框复用注意**：表格/图表/SmartArt 等属性对话框打开期间用户可能换页，保存时不能用当前选中页。`findInDoc(doc, id, slideIndex?)` 与 `updateElement(id, patch, { slideIndex? })` 支持显式传入打开时记录的页码；`snapBox` 为纯函数（传入缩放 `zoom` 换算屏幕容差），宿主自行决定如何画参考线。`updateDiagramTexts` 必须用来改 SmartArt 文本（语义 `el.texts` + 缓存 drawing 两处），只改前者导出会被 `__raw` 覆盖。

## 目录与可复用资源

| 路径 | 用途 |
|---|---|
| `examples/generate-test-pptx.mjs` | 生成能力全覆盖样例（T1–T19，21 页）+ 93 项自检断言，是最好的「能力清单 + 参数写法」参考 |
| `examples/round-trip-pptx.mjs` | PPTX → JSON → PPTX 往返，输出页数/元素/文本保真对比 |
| `examples/parse-pptx.mjs` | 最小解析脚本，打印页数、尺寸、每页文本 |
| `examples/index.html` | 浏览器预览页（含 echarts 图表渲染接入示例 + DOCX/PDF/MD/HTML 转换 demo） |
| `examples/cjk-font.js` + `vite.dev.config.ts` | demo 的中文字体获取：`/api/cjk-font` 代理拿真正 TTF，避免 WOFF2 导致 PDF 文字不可见 |
| `examples/chart-lib/` | `chart-renderer.js` + echarts，用于渲染 `result.charts` |
| `examples/editor/` | 在线编辑器示例（导入 PPTX → 检视/编辑形状与图表 → 导出） |
| `examples/vue-demo/` | Vue 集成示例 |
| `src/convert/` | 插件化转换引擎源码（`index.ts` 注册表 + docx/pdf/markdown/pptx 各链路） |
| `vscode/` | 官方 VSCode 插件，复用主库编辑器内核与 `@fefeding/ppt-parser/convert` |
| `test/*.test.ts` | 每项能力的回归测试，也是「预期行为」的权威说明 |
| `skills/pptx-parser/scripts/` | 本 skill 附带的三个可运行脚本（见下） |

### 附带的脚本（可直接 `node` 运行）

| 脚本 | 用途 |
|---|---|
| `scripts/pptx-info.mjs <file.pptx> [--no-text] [--json]` | 打印页数/画布/元数据/自定义属性/图表/每页元素类型与文本 |
| `scripts/pptx-to-html.mjs <file.pptx> [--out o.html] [--page 10]` | 渲染为自包含 HTML（Node 下需 `npm i -D jsdom`） |
| `scripts/json-to-pptx.mjs <doc.json> [--out o.pptx] [--check]` | JSON → PPTX，并可选做往返自检 |

## 验证工作流（务必执行）

生成或修改 PPTX 后不要凭感觉判断，按下面顺序验证：

1. **构建**（示例脚本从 `dist` 导入）：`npm run build`
2. **重新解析自检**：`node examples/generate-test-pptx.mjs`（默认会跑 93 项断言）
3. **往返保真**：`node examples/round-trip-pptx.mjs <file>` → 对比页数/元素类型/文本缺失
4. **单元测试**：`npx vitest run`；**类型**：`npm run type-check`
5. **格式转换端到端**：`cd vscode && npm run verify:convert`（`test-sample.pptx` → docx/pdf/md/html + 命令路径，31 项断言）；检查导出文件的页数（DOCX 应一页一张幻灯片）、标题样式与表格边框
6. **肉眼核对**：渲染出 HTML 后用无头浏览器截图

```bash
node skills/pptx-parser/scripts/pptx-to-html.mjs <file.pptx> --page 10 --out /tmp/p10.html
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --disable-gpu \
  --hide-scrollbars --screenshot=/tmp/p10.png --window-size=1280,720 \
  --virtual-time-budget=3000 file:///tmp/p10.html
```

7. **真机**：最终用 WPS/PowerPoint 打开确认（本库自检只能证明 XML 符合我们的预期，不能替代真机）

## 已知限制（实测，避免误判）

用 `pptxToStandard` 回读 `examples/test-sample.pptx` 的结果：

| 场景 | 实测行为 |
|---|---|
| 组合 `group` | 解析端已提取子元素（`childrenCoordinates: 'relative'`，坐标已换算到组内相对偏移），往返保真 |
| 视频 / 音频 | 解析端已识别 `p:pic` 下的 `a:videoFile` / `a:audioFile`，回读为 `video`/`audio` 元素（不再退化为 `image`）；生成端从 JSON 创作时支持 `video`/`audio` |
| SmartArt `diagram` | 解析端保留 `texts`（文本）+ `shapes`（缓存绘图形状，含坐标/连接线）+ `__raw`；创作端可用 `diagramType`+`nodes` 生成原生 diagrams 部件 |
| 连接线 `connector` | 生成端支持（`start`/`end` 精确端点）；解析端在图示缓存绘图中以 `PptxDiagramShape.connector=true` 提取，普通页面的 `p:cxnSp` 当前解析为 `shape` |
| OLE / 公式 | 类型定义完整、生成端支持；解析端目前按 `raw` 兜底（依赖 `__raw` 回写） |
| 形状 `blipFill` / `pattFill` | 生成端支持（含 tile/srcRect），但 `readSpPr` 未映射回 `PptxFill`，往返会丢失 |
| 图表 | 解析端只给数据（`result.charts`），HTML 预览需自行用 echarts 渲染 |

## 高频陷阱（详见 `references/ooxml-pitfalls.md`）

1. **JSON → PPTX 后真机"没效果"** —— 优先怀疑写出的 XML 不符合 OOXML schema。PowerPoint/WPS 对非法节点是**静默忽略**，不会报错。已踩过的坑：自造元素名、多包一层节点、枚举值不在 schema 内、属性写成了元素。
2. **表格没有边框** —— `tableStyleId` 指向的 GUID 必须在 `ppt/tableStyles.xml` 里有定义，否则退化为"无样式无网格"。本库生成端会为引用到的 GUID 自动补等价定义；手写 XML 时要注意。
3. **样式丢失** —— `pptxToHtml` 的 `styles.global` 必须一起注入，否则全部是裸 HTML。
4. **`__raw` 依赖不完整** —— 语义类型（text/shape/image/chart/table）默认不带 `__raw.rels/parts`，若想强制原始回写需 `pptxToStandard(file, { rawDeps: 'all' })`，否则 `r:embed` 会悬空。
5. **Buffer ≠ ArrayBuffer** —— Node 下直接用 `buf.buffer` 可能读到内存池其它内容，必须切片。
6. **图表在 HTML 里不显示** —— `pptxToHtml` 只输出占位 + `result.charts` 数据，需自行用 `examples/chart-lib/chart-renderer.js`（依赖全局 echarts）渲染。

## 参考文档

- [`references/api.md`](references/api.md) —— 全部导出 API、选项、返回结构
- [`references/json-schema.md`](references/json-schema.md) —— PptxDocument 完整字段契约
- [`references/cookbook.md`](references/cookbook.md) —— 元素配方（文本/形状/图片/表格/图表/组合/图示/媒体/动画/批注）
- [`references/ooxml-pitfalls.md`](references/ooxml-pitfalls.md) —— OOXML 合规要点与本项目已修复的真实案例
