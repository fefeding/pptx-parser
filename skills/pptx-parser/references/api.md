# API 参考

## 导出清单

```ts
// 默认导出 = pptxToHtml
import pptxToHtml, {
  pptxToJson, pptxToStandard, pptxToFiles,
  jsonToPptx, editPptx, PPTXComposer
} from '@fefeding/ppt-parser';

// 类型全部从入口导出（两处）：
//   export * from './types/pptx-document'   —— PptxDocument / PptxSlide / PptxElement …
//   export * from './compatibility-types'   —— PptxHtmlResult / PptxEditor / SlideElement …
```

## 解析选项 `PptxParserOptions`

| 字段 | 类型 | 默认 | 说明 |
|---|---|---|---|
| `mediaProcess` | `boolean` | `true` | 是否处理媒体（图片等）。关闭可提速，但图片填充/图片元素会缺失 |
| `themeProcess` | `boolean \| 'colorsAndImageOnly'` | `true` | `true`=完整主题（含母版背景）；`'colorsAndImageOnly'`=只取主题色与背景图 |
| `incSlide` | `{ width: number; height: number }` | `{0,0}` | 幻灯片尺寸增量调整 |
| `styleTable` | `Record<string, {name;text;suffix?}>` | `{}` | 自定义样式表；解析过程中会被填充，最终由 `genGlobalCSS` 生成 `styles.global` |
| `callbacks` | `Callbacks` | - | 见下 |

回调：`onFileStart`、`onError({type,message})`、`onSlide(data,{slideNum,fileName})`、`onThumbnail`、`onSlideSize`、`onGlobalCSS(css)`、`onComplete({executionTime,slideWidth,slideHeight,styleTable,settings})`。

---

## `pptxToHtml(fileData, options?)`

把 PPTX 渲染为可直接在浏览器展示的 HTML。

```ts
const result = await pptxToHtml(fileData, { mediaProcess: true, themeProcess: true });
```

返回 `PptxHtmlResult`：

```ts
{
  slides: [{ html: string, data: any, slideNum: number, fileName: string, hidden: boolean }],
  slideSize: { width: number, height: number },   // px
  thumbnail: string | null,                        // base64 jpeg
  styles: { global: string },                      // ★ 必须一起注入的 CSS
  metadata: { title?, author?, created?, … },
  customProps: Record<string, string>,
  charts: ChartData[]                              // 图表数据（需自行渲染）
}
```

要点：
- 每页 `html` 根节点是 `<section class='slide' style='width:…px;height:…px'>`，带 `id="slide-N"`；隐藏页带 `data-hidden='true'`，过渡/计时带 `data-transition` / `data-timing`。
- **必须**把 `styles.global` 作为 `<style>` 注入，类名形如 `_css_1` / `_tbl_cell_css_3`。
- 图表只输出占位容器 + `result.charts`，需用 echarts 渲染（见 `examples/chart-lib/chart-renderer.js`，要求 `echarts` 为全局）。
- 批注（`ppt/comments/commentsN.xml`）会以 `.pptx-comment` 气泡渲染在页内；文档自定义属性渲染为左下角面板。

---

## `pptxToJson(fileData, options?)`

返回结构化（简化 XML 树）结果，适合逐节点精读。

```ts
const json = await pptxToJson(buffer, { mediaProcess: true, themeProcess: true });
// {
//   slides: [{ data: ProcessedSlideData, slideNum, fileName }],
//   slideSize, thumbnail, styles: { global }, metadata, customProps, charts
// }
```

`slide.data` 关键字段：`slideContent`（tXml 简化树，根 `p:sld`）、`slideLayoutContent`、`slideMasterContent`、`themeContent`、`slideResObj`（rId → `{type,target}`）、`tableStyles`、`defaultTextStyle`、`notesContent`。

文本提取：简化树里 DrawingML 文本是键 `'a:t'` 对应的字符串，递归收集即可（参考 `examples/parse-pptx.mjs`）。

`mode: 'semantic'` 时额外返回 `result.document`（等价于 `pptxToStandard` 的结果）。

---

## `pptxToStandard(fileData, options?)`

返回 `PptxDocument`（`types/pptx-document.ts`），与 `jsonToPptx` 同一契约，可直接回写。

```ts
const doc = await pptxToStandard(fileData, { rawDeps: 'all' });
const out = await jsonToPptx(doc, { outputType: 'uint8array' });
```

选项：
- `rawDeps: 'all'` — 为语义类型（text/shape/image/chart/table）也附带 `__raw.rels`/`__raw.parts`，否则只有 `{tag, node}`
- 其余解析选项同 `PptxParserOptions`（`mediaProcess`、`themeProcess` 等）

- 解析端对**语义层不支持**的类型（diagram / group / OLE / 未知标签）自动附加 `__raw`（原始节点 + 关系 + 部件），序列化时原样回写，保证不丢信息。
- 语义类型（text/shape/image/chart/table）默认 `__raw` 只有 `{tag,node}`；要强制原始回写必须先用 `rawDeps: 'all'` 解析，否则 `r:embed` 等引用会悬空。
- 元素上可设 `rawFallback: true` 强制走 `__raw` 回写。
- 解析端图片/媒体统一输出完整 `data:<mime>;base64,<b64>` dataURL。
- 主题色引用（`scheme:accent1`）解析时按该页真实主题（`themeContent`）解析为绝对色，多主题文件中不同页可绑定不同主题。

---

## `pptxToFiles(fileData)`

列出并读出 PPTX 内所有部件，便于直接检查 OOXML。

```ts
const { files, content } = await pptxToFiles(fileData);
// files: [{ name, dir, size }]
// content[name]:
//   { type:'text', content }                          // xml / rels
//   { type:'image', format, base64, dataUrl }         // png/jpg/gif/bmp/svg
//   { type:'binary', base64 }                         // 其它（mp4/m4a/emf…）
```

**排查真机差异时最有用**：`content['ppt/slides/slide10.xml']` 直接看生成端到底写了什么。

---

## `jsonToPptx(presentation, options?)`

```ts
const data = await jsonToPptx(pres, {
  outputType: 'uint8array',   // uint8array | arraybuffer | blob | nodebuffer | base64
  theme: themeXmlString       // 可选：完整 theme XML 字符串（见 examples/generate-test-pptx.mjs 的 CUSTOM_THEME）
});
```

- `pres`：`PptxDocument` 或 `PPTXComposer` 实例（有 `toJSON()` 即可）。
- `pres.theme`：可传 `PptxTheme` 语义对象（`{ name, colors, fonts }`），生成端据此构造完整 `themeN.xml`；也可传整串 XML 字符串覆盖。
- `pres.masters`：提供时生成端按此写出多个 `slideMasterN.xml` 及其版式，幻灯片通过 `PptxSlide.layout` 指定所用版式。
- `pres.fonts`：嵌入字体（生成端自动做 XOR 混淆为 `fntdata`）。
- 约束：`slides` 非空，否则抛 `jsonToPptx: 演示文稿至少需要一页幻灯片`。
- 写入的部件包含 `ppt/tableStyles.xml`（按文档引用到的 `tableStyleId` 动态补等价定义）、`docProps/custom.xml`、`ppt/comments/commentsN.xml`、`ppt/commentAuthors.xml`、主题、图表、媒体等，无需手工拼装。
- 返回值类型由 `outputType` 决定；Node 落盘用 `Buffer.from(data)`。

---

## `PPTXComposer`（链式构建）

```ts
const p = new PPTXComposer();
p.slideSize(1280, 720).title('标题').author('me');
p.addSlide((s) => {
  s.background('#ffffff');
  s.addText({ x: 60, y: 40, width: 600, height: 60, text: 'Hello', fontSize: 32, bold: true });
  s.addShape({ shapeType: 'roundRect', x: 60, y: 140, width: 300, height: 160, fill: '#3b82f6' });
  s.addImage({ x: 400, y: 140, width: 300, height: 200, data: 'data:image/png;base64,…' });
  s.addChart({ chartType: 'pieChart', x: 60, y: 340, width: 300, height: 200,
               categories: ['A','B'], series: [{ name: '占比', values: [3, 7] }] });
  s.addGroup(...); s.addDiagram(...);
});
const data = await jsonToPptx(p);
```

## `editPptx(fileData)`

对已有 PPTX 做页级编辑，不经过语义层，适合"只删一页/追加一页"这类操作。

```ts
const editor = await editPptx(fileData);
await editor.getSlideCount();            // 逻辑页数（sldIdLst 顺序）
await editor.getSlide(3);                // 该页简化 XML 树（与 pptxToJson 的 slideContent 同构）
await editor.deleteSlide(3);             // 删除（至少保留一页）
await editor.moveSlide(1, 5);            // 移动页码（会物理重编号 slide 文件并重映射内部跳转）
await editor.setMetadata({ title: '新标题' });
await editor.addSlide({
  background: '#fff',
  transition: { type: 'fade', duration: 1000 },
  animations: [{ target: 1, type: 'flyIn', duration: 0.5, presetClass: 'entr' }],
  elements: [ … ]
});
const out = await editor.save({ outputType: 'uint8array' });
```

`editor.zip` 是底层 JSZip 实例，需要改任意部件可直接操作。`addSlide` 支持 `transition` 和 `animations`（与 `PptxSlide` 同构）。

---

## 编辑器内核（Headless Editor Core）

与 UI 无关的可复用编辑器能力，全部从包入口导出（实现位于 `src/editor`，随 `dist`/`src` 一起发布）。任意前端框架或自有 UI 都能直接复用同一套文档模型与 PPTX 双向转换逻辑，而不必依赖 `examples/editor` 的具体渲染。

```ts
import {
  createStore, createActions, docFromPptx, docToPptx, renderChartSVG,
  elementRect, rotatedRect, effectMargin, absoluteElementRect,
  normalizeDoc, normalizeElement,
  EditorStore
} from '@fefeding/pptx-parser';
```

### 文档状态 `createStore`

```ts
const store = createStore({ storage? });        // storage 可选，实现 getItem/setItem 即可持久化（如 localStorage）
store.setDoc(doc);                              // doc 为 docFromPptx 得到的编辑器文档
const doc = store.getDoc();                     // 取回当前文档快照
store.update(patch);                            // 浅合并式更新
```

类型 `EditorStore`：`{ doc, setDoc, getDoc, update, storage? }`。

### 编辑操作 `createActions`

```ts
const actions = createActions(store);           // 工厂函数，不依赖单例，便于多实例/测试
actions.addSlide({ elements: [/* … */] });
actions.duplicateSlide(1);
actions.deleteSlide(2);
actions.moveSlide(1, 3);
actions.updateElement(id, { fill: { color: '#4f46e5' } });
actions.alignElements('hcenter');
actions.distribute('horizontal');
actions.groupSelection(); actions.ungroupSelection();
actions.toggleLock(id); actions.toggleHidden(id);
actions.applyLayout(0); actions.setBackground('#fff');
actions.setSlideSize(1280, 720);
actions.applyTheme(themeObj); actions.setNotes('…');
actions.addAnimation({ target, type: 'flyIn', presetClass: 'entr', duration: 0.5 });
actions.removeAnimation(idx); actions.moveAnimation(from, to);
// 文本 / 表格 / 几何 / 查找
actions.applyTextStyleSel({ bold: true });
actions.setElementGeo(id, geo); actions.resizeTable(id, rows, cols);
actions.findInDoc('关键字');
```

### 文档 ↔ PptxDocument 转换

```ts
// 解析端产物（PptxDocument）转编辑器文档；fileName 用于标题兜底
const doc = docFromPptx(pptxStandardDoc, { fileName: 'my-deck.pptx' });

// 编辑器文档转回标准 PptxDocument，再交 jsonToPptx 落盘
const pptxDoc = docToPptx(doc);
const out = await jsonToPptx(pptxDoc, { outputType: 'uint8array' });
```

`docFromPptx` 的标题优先级：**PPTX 自带 `core.xml` 的 `dc:title`** > **传入的 `fileName`（自动去扩展名）** > 固定字符串 `'导入的演示文稿'`。当原文件无标题元信息时，会以文件名兜底，并在导出时一并写回 `core.xml`，保证来回一致。

### 图表与几何（纯函数）

```ts
const svg = renderChartSVG(chartElement);       // 返回 SVG 字符串，宿主自行决定挂载方式
const rect = elementRect(element);              // 元素实际矩形（group = 子元素并集）
const box  = rotatedRect(element);              // 含旋转的外接矩形
const m    = effectMargin(element);             // 阴影/发光外边距
const abs  = absoluteElementRect(element, parent); // 相对某父级的绝对矩形
```

### 文档规范化

- `normalizeDoc(doc)`：补齐缺省字段（`slideSize`、`elements`、`zIndex` 序号等），导入/合并后调用可确保文档完整。
- `normalizeElement(el)`：单元素级规范化（补 `id`、默认 `name`、类型字段等）。

---

## 环境注意事项

- **Node**：`fs.readFile` 得到 Buffer 可直接传入；若要显式 ArrayBuffer，务必切片：
  `buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)`。
- **浏览器**：`await file.arrayBuffer()`；或 `<script src="dist/ppt-parser.browser.js">` 后用全局 `pptxParser`。
- 解析是异步的（内部 JSZip 解压 + 主题/母版读取）；大文件建议开 `callbacks.onSlide` 做进度反馈。
