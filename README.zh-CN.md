# @fefeding/ppt-parser

一个轻量级的 PPTX 解析库，让处理 PowerPoint 文件变得简单。基于纯 TypeScript 编写，零框架依赖，同时支持浏览器和 Node.js 环境。

## 特性

- **简单好用** — 只需几行代码即可解析和转换 PPTX 文件
- **零依赖** — 不绑定任何框架，适用于任何 JavaScript/TypeScript 项目
- **双向转换** — 支持 PPTX 到 HTML 或 JSON 的解析
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

## 支持的元素

- **文本** — 富文本、超链接、项目符号、编号列表
- **图片** — PNG、JPEG、SVG 等格式
- **形状** — 矩形、圆形、三角形、自定义形状
- **表格** — 完整表格支持，包含自定义样式
- **图表** — 柱状图、折线图、饼图等
- **媒体** — 视频和音频（计划中）

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

## 致谢

本项目受 [pptxjs](https://github.com/meshesha/pptxjs) 项目启发，感谢原作者提供的架构基础。
