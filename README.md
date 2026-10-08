# @fefeding/ppt-parser

A lightweight PPTX parsing library that makes working with PowerPoint files simple. Built with pure TypeScript, zero framework dependencies, and full support for both browser and Node.js environments.

## Features

- **Simple & Easy** — Parse and convert PPTX files with just a few lines of code
- **Zero Dependencies** — No framework lock-in, works in any JavaScript/TypeScript project
- **Dual Conversion** — Parse PPTX to HTML or JSON, and serialize JSON (or the fluent `PPTXComposer`) back to a valid PPTX file
- **Comprehensive Elements** — Text, shapes, tables, images, charts and more
- **Smart Unit Handling** — Automatic EMU to PX conversion
- **Universal Module** — Supports both ESM and CommonJS
- **Browser & Node.js** — Runs seamlessly in both environments

## Live Demo

Try it online: [https://fefeding.github.io/pptx-parser/examples/index.html](https://fefeding.github.io/pptx-parser/examples/index.html)

## Installation

```bash
npm install @fefeding/ppt-parser
```

Or use directly from the `dist` folder for browser builds.

## Quick Start

### Parse PPTX to HTML

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

  // result.slides contains the parsed slide HTML
  console.log('Slides:', result.slides.length);
  console.log('Slide size:', result.slideSize);
  console.log('Metadata:', result.metadata);
  console.log('Charts:', result.charts);

  // Render slides
  const container = document.getElementById('preview');
  result.slides.forEach(slide => {
    const div = document.createElement('div');
    div.innerHTML = slide.html;
    container.appendChild(div);
  });
});
```

### Parse PPTX to JSON

```javascript
import { pptxToJson } from '@fefeding/ppt-parser';

const result = await pptxToJson(fileData);
console.log('Slides:', result.slides.length);
console.log('Slide size:', result.slideSize);
console.log('Metadata:', result.metadata);
```

### Extract File Contents

```javascript
import { pptxToFiles } from '@fefeding/ppt-parser';

const result = await pptxToFiles(fileData);
console.log('Files:', result.files);
console.log('Content:', result.content);
```

## Serialize JSON to PPTX

### Fluent Composer

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
    // Hyperlinks: external URLs, or '#N' to jump to slide N
    slide.addText(t => t.runs([
      { text: 'External link', options: { href: 'https://example.com' } },
      { text: ' / ', options: {} },
      { text: 'Jump to slide 1', options: { href: '#1' } }
    ]).x(100).y(100));
  });

const data = await composer.save(); // Uint8Array
```

Object-style config is also accepted: `slide.addText({ text: 'Hi', x: 0, y: 0 })`.

### JSON to PPTX

```javascript
import { jsonToPptx } from '@fefeding/ppt-parser';

const data = await jsonToPptx({
  metadata: { title: 'My Deck', author: 'me' },
  slideSize: { width: 1280, height: 720 },  // px, default 16:9
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

Supported elements: text (multi-paragraph, runs with font size/color/bold/italic/underline/font face, hyperlinks, bullets, numbered lists, outline, shadow, dynamic fields), preset shapes (gradient/solid/picture/pattern fills, tiling, cropping, shadow, glow, 3D extrusion, custom geometry), images (dataURL/base64/remote URL, cropping, brightness/contrast/transparency), tables (per-cell borders, diagonal borders, spans, styles), charts (2D/3D, multi-series, trendlines, secondary axis, point colors, gridlines), groups, diagrams (SmartArt), connectors, video/audio, OLE objects, math formulas, and more.

### Edit Existing PPTX

```javascript
import { editPptx } from '@fefeding/ppt-parser';

const editor = await editPptx(fileData);
await editor.deleteSlide(2);                       // remove slide 2
await editor.moveSlide(1, 3);                      // reorder
await editor.addSlide({ elements: [/* same element format */] });
await editor.setMetadata({ title: 'Updated', author: 'me' });
const updated = await editor.save();
```

Generated files round-trip through `pptxToJson` / `pptxToHtml`.

## Options

```typescript
interface PptxParserOptions {
  // Process media files (images, etc.)
  mediaProcess?: boolean;

  // Theme processing mode
  themeProcess?: boolean | 'colorsAndImageOnly';

  // Slide size adjustment
  incSlide?: {
    width: number;
    height: number;
  };

  // Custom style table
  styleTable?: Record<string, { name: string; text: string; suffix?: string }>;

  // Callbacks
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

## Headless Editor Core

Beyond the command-style `editPptx` wrapper, the library ships a **UI-agnostic editor core** (`src/editor`, included in the package). It lifts the document model, PPTX round-trip, edit operations, chart rendering and element geometry out of any specific renderer, so any frontend framework or custom UI can reuse the same business logic without rewriting it.

```javascript
import {
  createStore, createActions, docFromPptx, docToPptx, renderChartSVG,
  elementRect, rotatedRect, effectMargin, absoluteElementRect,
  normalizeDoc, normalizeElement,
  EditorStore
} from '@fefeding/pptx-parser';

// 1) Create a UI-free document state (optional persistence via a getItem/setItem storage adapter)
const store = createStore({ storage: localStorage });
const sem = await pptxToJson(fileData, { mode: 'semantic' });
store.setDoc(docFromPptx(sem.document));

// 2) Create edit operations (factory, no singleton — easy multi-instance/testing)
const actions = createActions(store);
actions.addSlide({ elements: [/* element format as above */] });
actions.alignElements('hcenter');
actions.updateElement(id, { fill: { color: '#4f46e5' } });

// 3) Export back to PPTX
const pptxDoc = docToPptx(store.doc);
const data = await jsonToPptx(pptxDoc);

// 4) Chart & geometry (pure functions — host decides how to mount)
const svg = renderChartSVG(chartElement); // returns an SVG string
const rect = elementRect(element);        // actual rect (group = union of children)
```

Core API:

- `createStore(options?)` → `EditorStore`: holds `doc`, exposes `setDoc / getDoc / update`, with optional persistence (pass a `storage` implementing `getItem/setItem`).
- `createActions(store)` → `EditorActions`: returns `addSlide / deleteSelected / duplicateSelected / copySelected / paste / selectAll / nudge / zOrder / alignElements / distribute / groupSelection / ungroupSelection / toggleLock / toggleHidden / addSlide / duplicateSlide / deleteSlide / moveSlide / toggleSlideHidden / applyLayout / setBackground / setSlideSize / applyTheme / setNotes / applyTextStyleSel / setBackgroundImage / setElementGeo / resizeTable / updateElement / findInDoc / setTransition / addAnimation / updateAnimation / removeAnimation / moveAnimation`.
- `docFromPptx(semanticDoc, { fileName? })` / `docToPptx(doc)`: bidirectional conversion between the editor document and the standard `PptxDocument`. `docFromPptx` resolves the document title with priority `core.xml dc:title` > `fileName` (extension stripped) > `'导入的演示文稿'`; when the source file has no title metadata, the file name is used as a fallback and written back into `core.xml` on export so the title stays consistent round-trip.
- `renderChartSVG(chartEl)`: renders a chart element to an SVG string (host decides how to mount it).
- `elementRect / rotatedRect / effectMargin / absoluteElementRect`: element geometry helpers (group union rect, rotated bounding box, shadow/glow margin, absolute rect relative to a parent).
- `normalizeDoc(doc)` / `normalizeElement(el)`: fill in default fields after import/merge (`slideSize`, `elements`, `zIndex`, element `id`/`name`, …) to keep the document complete.
- `EditorStore` (type): `{ doc, setDoc, getDoc, update, storage? }` returned by `createStore`.

## Supported Elements

- **Text** — Multi-paragraph, runs with font size/color/bold/italic/underline/font face, hyperlinks (external URL or `#N` to jump to slide N), bullet & numbered lists, line spacing, indentation, text box inset, vertical text direction, multi-column, autofit, WordArt transforms, run-level outline & shadow, dynamic fields (slide number, datetime)
- **Shapes** — All preset geometries (`rect`, `roundRect`, `ellipse`, `triangle`, `arrow`, `star5`, `foldedCorner`, …); gradient / solid / picture / pattern fills; picture fill with tiling (`tile`) and source-rectangle cropping (`srcRect`); line style, shadow & glow effects; geometric adjustment (`avLst`); horizontal/vertical flip; custom geometry paths (`custGeom`); 3D extrusion, bevel, camera & lighting (`threeD`)
- **Images** — PNG, JPEG, GIF, BMP, WEBP, SVG; base64/dataURL or remote `src`; cropping, brightness/contrast/transparency
- **Tables** — Full styling: per-cell borders, diagonal borders (`tlBr` / `blTr` / `both`), cell inset, span (`colSpan`/`rowSpan`), merge (`hMerge`/`vMerge`), fill, alignment, and table styles (`tableStyleId`)
- **Charts** — Bar/column, line, area, pie/doughnut, ofPie, scatter, bubble, radar, stock, surface (2D & 3D); multi-series, series colors, per-point colors (`pointColors`), legend, data labels, grouping, `barDir`, `holeSize`, `ofPieType`, `bubble3D`, `wireframe`, 3D view (`view3D`), secondary axis, trendlines, gridlines, axis titles
- **Group** — `group` elements with `childrenCoordinates: 'local' | 'page' | 'relative'`
- **Diagram** — SmartArt (list / hierarchy / process / cycle / pyramid) with cached drawing shapes and connectors
- **Connector** — Connection lines (`straightConnector1` / `bentConnector3` / `curvedConnector2`) with precise start/end points
- **Media** — Video (`mp4`/`m4v`) and audio (`m4a`/`mp3`) with optional poster
- **OLE** — Embedded objects (Excel sheets, etc.) with icon/poster
- **Math** — OMML formulas
- **Slide-level** — Background (solid/gradient/image), transition (with sound/direction), auto-advance timing, animations (entr/exit/emph/path with triggers), hidden slides, notes, comments, custom document properties
- **Document-level** — Masters/layouts/placeholders, sections, embedded fonts (with XOR obfuscation), semantic theme (colors/fonts)

## Platform Usage

### Browser

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

### Vue Example

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

> For chart rendering in Vue, copy `examples/chart-lib/chart-renderer.js` to your project and set up `echarts` as a global dependency. See the `examples/vue-demo` directory for a complete working example.

## Development

```bash
# Clone the repo
git clone https://github.com/fefeding/pptx-parser.git
cd pptx-parser

# Install dependencies
npm install

# Development mode
npm run dev

# Build
npm run build

# Run tests
npm test
```

## TypeScript

This package includes built-in TypeScript definitions. All interfaces and types are exported from the package entry.

## Browser Support

- Chrome >= 80
- Firefox >= 75
- Edge >= 80
- Safari >= 14

## AI Skills

This repository ships a ready-to-use skill for AI coding assistants under [`skills/pptx-parser/`](skills/pptx-parser). It is also published inside the npm tarball (the `skills` entry in `package.json` `files`), so any project depending on `@fefeding/ppt-parser` gets it.

The skill documents the exact contract for **parsing PPTX to HTML / JSON / standard JSON, generating or editing PPTX from JSON, and fixing OOXML compliance issues** — with copy-paste element recipes, a full field reference, and a list of real pitfalls (e.g. `tableCellInsets` is not valid OOXML, `a:lnTlToBr` is itself a line, `tableStyleId` must have a matching definition in `tableStyles.xml`).

```
skills/pptx-parser/
├── SKILL.md                 # entry point: capability matrix, quick start, verification workflow, known limits
├── references/
│   ├── api.md               # every exported API: signatures, options, return shapes
│   ├── json-schema.md       # PptxDocument field contract + unit conversions
│   ├── cookbook.md          # element recipes (text/shape/image/table/chart/group/diagram/media/...)
│   └── ooxml-pitfalls.md    # real "silently ignored by PowerPoint/WPS" cases + self-check checklist
└── scripts/
    ├── pptx-info.mjs        # print file overview (pages, elements, text, metadata, charts)
    ├── pptx-to-html.mjs     # render a self-contained HTML (needs jsdom in Node)
    └── json-to-pptx.mjs     # JSON → PPTX, with --check round-trip self-check
```

Point your AI assistant at `skills/pptx-parser/SKILL.md` (or configure it as a skill) whenever you need to preview, extract, generate, or modify PPTX programmatically.

## License

[MIT](LICENSE)

## Acknowledgements

Inspired by the [pptxjs](https://github.com/meshesha/pptxjs) project. Thanks to the original authors for the foundational architecture.
