---
name: pptx-parser
description: 使用 @fefeding/ppt-parser 处理 PowerPoint 文件的完整指南——解析 PPTX 为 HTML / 结构化 JSON / 标准 JSON（PptxDocument），由 JSON 生成或编辑 PPTX，以及排查渲染差异与 OOXML 合规问题。当需要预览 PPT、提取内容与图表数据、批量生成或修改 PPT、做 round-trip 保真验证时使用。
license: MIT
---

# PPTX 解析与生成（@fefeding/ppt-parser）

一个零框架依赖的纯 TypeScript PPTX 库，浏览器与 Node.js 同构。核心能力是**双向转换**：

```
PPTX ──pptxToStandard──> PptxDocument JSON ──jsonToPptx──> PPTX
PPTX ──pptxToHtml──────> HTML + CSS（浏览器预览）
PPTX ──pptxToFiles─────> 部件清单（xml / 图片 base64 / 二进制）
PPTX ──editPptx────────> 页级增删改 → 新 PPTX（不动语义层）
```

## 何时使用

- 需要在网页里预览 PPTX（渲染成 HTML，配 `styles.global` 的 CSS）
- 需要提取 PPTX 的文本、表格、图表数据、备注、批注、自定义属性
- 需要用代码生成 PPTX（报告、批量导出、模板填充）
- 需要修改已有 PPTX（删页、移动页、追加页、改元数据）
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

// 浏览器 <script src="./dist/ppt-parser.browser.js"></script> → 全局变量 pptxParser
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

完整签名、选项与返回结构见 [`references/api.md`](references/api.md)。

## 元素类型速查

`slides[i].elements` 支持：`text`、`shape`、`image`、`chart`、`table`、`diagram`（SmartArt）、`group`、`video`、`audio`、`raw`（解析兜底）。

- 坐标 `x/y/width/height` 统一为 **px**（96 DPI）；内部按 `SLIDE_FACTOR = 96/914400` 与 EMU 互转
- 字体大小统一为 **pt**，线宽为 **pt**，透明度/裁剪/平铺为 **0~1 或 0~100** 的比例（见下文各字段注释）
- 颜色用 `#RRGGBB`（也支持 `scheme:accent1` 这类主题色引用写法，见 cookbook）

逐个字段的权威定义见 [`references/json-schema.md`](references/json-schema.md)；可直接抄的片段见 [`references/cookbook.md`](references/cookbook.md)。

## 目录与可复用资源

| 路径 | 用途 |
|---|---|
| `examples/generate-test-pptx.mjs` | 生成能力全覆盖样例（T1–T19，21 页）+ 93 项自检断言，是最好的「能力清单 + 参数写法」参考 |
| `examples/round-trip-pptx.mjs` | PPTX → JSON → PPTX 往返，输出页数/元素/文本保真对比 |
| `examples/parse-pptx.mjs` | 最小解析脚本，打印页数、尺寸、每页文本 |
| `examples/index.html` | 浏览器预览页（含 echarts 图表渲染接入示例） |
| `examples/chart-lib/` | `chart-renderer.js` + echarts，用于渲染 `result.charts` |
| `examples/vue-demo/` | Vue 集成示例 |
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
5. **肉眼核对**：渲染出 HTML 后用无头浏览器截图

```bash
node skills/pptx-parser/scripts/pptx-to-html.mjs <file.pptx> --page 10 --out /tmp/p10.html
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --disable-gpu \
  --hide-scrollbars --screenshot=/tmp/p10.png --window-size=1280,720 \
  --virtual-time-budget=3000 file:///tmp/p10.html
```

6. **真机**：最终用 WPS/PowerPoint 打开确认（本库自检只能证明 XML 符合我们的预期，不能替代真机）

## 已知限制（实测，避免误判）

用 `pptxToStandard` 回读 `examples/test-sample.pptx` 的结果：

| 场景 | 实测行为 |
|---|---|
| 组合 `group` | 语义提取**丢失组合内的子元素**（T4 页只剩标题）。需要保留组合请改用 `editPptx` 直接操作 XML |
| 视频 / 音频 | 回读降级为 `image` 元素（媒体数据保留，类型信息丢失） |
| SmartArt `diagram` | 可回读为 `diagram`（保留 `texts`，连接线 `cxnSp` 现已渲染为 `straightConnector1`）；还原依赖 `__raw` |
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
