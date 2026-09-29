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

Supported elements: text (multi-paragraph, runs with font size/color/bold/italic/underline/font face, hyperlinks, bullets, alignment), preset shapes, and images (dataURL / base64 / remote URL via `src`).

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

## Supported Elements

- **Text** — Rich text, hyperlinks, bullet lists, numbered lists
- **Images** — PNG, JPEG, SVG, and more
- **Shapes** — Rectangles, circles, triangles, custom shapes
- **Tables** — Full table support with custom styling
- **Charts** — Bar, line, pie, and other chart types
- **Media** — Video and audio (planned)

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

## License

[MIT](LICENSE)

## Acknowledgements

Inspired by the [pptxjs](https://github.com/meshesha/pptxjs) project. Thanks to the original authors for the foundational architecture.
