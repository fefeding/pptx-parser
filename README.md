# PPT-Parser

一个轻量级的 PPTX 解析库，让处理 PowerPoint 文件变得简单。

## 特性

- 📦 **简单易用** - 几行代码即可完成 PPTX 文件的解析和生成
- 🔧 **纯 TypeScript** - 完整的类型定义，优秀的开发体验
- 🎯 **零框架依赖** - 可在任何 JavaScript/TypeScript 项目中使用
- 📱 **双向支持** - 支持 PPTX 文件 → HTML/JSON、HTML/JSON → PPTX 双向转换
- 🎨 **支持多种元素** - 文本、形状、表格、图片等常见元素
- 🔄 **智能转换** - 自动处理 EMU ↔ PX 单位转换
- 📦 **双格式输出** - 同时支持 ESM 和 CommonJS 模块
- 🌐 **浏览器/Node.js 双支持** - 可在浏览器环境和 Node.js 中使用

## 安装

```bash
npm install @fefeding/ppt-parser
```

或者直接下载 [`dist`](./dist) 目录下的文件使用。

## 快速开始

### 解析 PPTX 文件为 HTML（推荐）

```javascript
import { pptxToHtml } from '@fefeding/ppt-parser';

// 上传并解析 PPTX 文件为 HTML
const fileInput = document.querySelector('#ppt-upload');

fileInput.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  // 读取文件为 ArrayBuffer
  const fileData = await file.arrayBuffer();

  const result = await pptxToHtml(fileData, {
    mediaProcess: true,    // 处理媒体文件
    themeProcess: true     // 处理主题样式
  });

  // result.slides 是幻灯片数组
  console.log('幻灯片数量:', result.slides.length);
  console.log('幻灯片大小:', result.slideSize);
  console.log('样式:', result.styles);
  console.log('元数据:', result.metadata);
  console.log('图表数据:', result.charts);

  // 渲染所有幻灯片
  const container = document.getElementById('preview');
  result.slides.forEach(slide => {
    const div = document.createElement('div');
    div.innerHTML = slide.html;
    container.appendChild(div);
  });
});
```

### 解析 PPTX 文件为 JSON

```javascript
import { pptxToJson } from '@fefeding/ppt-parser';

// 解析 PPTX 文件为 JSON 数据
const fileInput = document.querySelector('#ppt-upload');

fileInput.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  // 读取文件为 ArrayBuffer
  const fileData = await file.arrayBuffer();

  const result = await pptxToJson(fileData);
  // result.slides 是幻灯片 JSON 数据数组
  console.log('幻灯片数量:', result.slides.length);
  console.log('幻灯片大小:', result.slideSize);
  console.log('元数据:', result.metadata);
});
```

### 解析 PPTX 文件获取所有文件索引和内容

```javascript
import { pptxToFiles } from '@fefeding/ppt-parser';

// 解析 PPTX 文件获取所有文件的索引和内容
const fileInput = document.querySelector('#ppt-upload');

fileInput.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  // 读取文件为 ArrayBuffer
  const fileData = await file.arrayBuffer();

  const result = await pptxToFiles(fileData);

  // 查看文件索引
  console.log('文件列表:', result.files);
  // [
  //   { name: 'ppt/slides/slide1.xml', dir: false, size: 12345 },
  //   { name: 'ppt/media/image1.png', dir: false, size: 6789 },
  //   ...
  // ]

  // 获取特定文件内容
  const slide1Content = result.content['ppt/slides/slide1.xml'];
  // slide1Content.type === 'text' 时，content 字段为文本内容
  console.log('Slide1 内容:', slide1Content.content);

  // 获取图片
  const image1 = result.content['ppt/media/image1.png'];
  // image1.type === 'image' 时，有 base64 和 dataUrl 字段
  console.log('图片 Data URL:', image1.dataUrl);
});
```

`pptxToFiles` 返回值结构：
```javascript
{
  files: [
    {
      name: "ppt/slides/slide1.xml",    // 文件路径
      dir: false,                         // 是否为目录
      size: 12345                         // 解压后大小
    }
  ],
  content: {
    "ppt/slides/slide1.xml": {
      type: "text",
      content: "<?xml version=\"1.0\"..."  // XML 文件内容
    },
    "ppt/media/image1.png": {
      type: "image",
      format: "png",
      base64: "iVBORw0KGgoAAAANSUhEUg...",   // Base64 编码
      dataUrl: "data:image/png;base64,iVBORw0KGgo..."  // Data URL
    }
  }
}
```

### 导出 PPTX 文件

> 注意：当前版本导出功能正在完善中，主要支持解析功能

## 输出格式

### pptxToHtml 返回值结构

```typescript
interface PptxHtmlResult {
  // 幻灯片 HTML 结果数组
  slides: Array<{
    html: string;       // 幻灯片 HTML
    slideNum: number;   // 幻灯片编号
    fileName: string;   // 幻灯片文件名
  }>;

  // 幻灯片大小信息
  slideSize: {
    width: number;
    height: number;
    defaultTextStyle?: any;
  };

  // 缩略图（Base64 Data URL）
  thumbnail: string | null;

  // 样式信息
  styles: {
    global: string;     // 全局 CSS
  };

  // 元数据
  metadata: {
    title?: string;
    subject?: string;
    author?: string;
    keywords?: string;
    description?: string;
    lastModifiedBy?: string;
    created?: string;
    modified?: string;
    category?: string;
    status?: string;
    contentType?: string;
    language?: string;
    version?: string;
  };

  // 图表数据
  charts: Array<{
    chartId: string;    // 图表容器 ID
    type: string;       // 图表类型
    data: Array<{       // 图表系列数据
      key: string;
      values: Array<{ x: string; y: number }>;
      xlabels: { [key: string]: string };
    }>;
  }>;
}
```

### pptxToJson 返回值结构

与 `pptxToHtml` 结构相同，但 `slides` 数组中的每个元素为：

```typescript
{
  data: any;          // 幻灯片 JSON 数据
  slideNum: number;   // 幻灯片编号
  fileName: string;   // 幻灯片文件名
}
```

## 功能特性

本库提供完整的PPTX解析能力，支持标准PPTX文件的所有元素类型。

### 支持的元素类型

- 📝 **文本** - 富文本、超链接、项目符号、编号列表
- 🖼️ **图片** - JPG、PNG、SVG 等格式
- 🔷 **形状** - 矩形、圆形、三角形、自定义形状等
- 📊 **表格** - 自定义表格样式
- 📈 **图表** - 柱状图、折线图、饼图等
- 🎬 **媒体** - 视频和音频支持（计划中）
- 🎨 **效果** - 阴影、渐变、3D 效果（计划中）

### 解析选项

```typescript
interface PptxParserOptions {
  // 是否处理媒体文件
  mediaProcess?: boolean;

  // 主题处理方式
  themeProcess?: boolean | 'colorsAndImageOnly';

  // 幻灯片尺寸调整
  incSlide?: {
    width: number;
    height: number;
  };

  // 样式表
  styleTable?: {
    [key: string]: {
      name: string;
      text: string;
      suffix?: string;
    };
  };

  // 回调函数
  callbacks?: {
    onFileStart?: () => void;
    onError?: (error: { type: string; message: string }) => void;
    onSlide?: (data: any, info: { slideNum: number; fileName: string }) => void;
    onThumbnail?: (thumbnail: string | null) => void;
    onSlideSize?: (slideSize: SlideSize) => void;
    onGlobalCSS?: (css: string) => void;
    onComplete?: (info: {
      executionTime: number;
      slideWidth: number;
      slideHeight: number;
      styleTable: StyleTable;
      settings: PptxParserOptions;
    }) => void;
  };
}
```

使用示例：

```javascript
const result = await pptxToHtml(fileData, {
  mediaProcess: true,
  themeProcess: true,
  callbacks: {
    onSlide: (data, info) => {
      console.log(`处理幻灯片 ${info.slideNum}`);
    },
    onComplete: (info) => {
      console.log(`解析完成，耗时 ${info.executionTime}ms`);
    }
  }
});
```

### 浏览器中使用

```html
<script src="./dist/ppt-parser.browser.js"></script>
<script>
  const fileInput = document.querySelector('#ppt-upload');

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // 读取文件为 ArrayBuffer
    const fileData = await file.arrayBuffer();

    const result = await pptxParser.pptxToHtml(fileData);

    // 渲染所有幻灯片
    const container = document.getElementById('preview');
    result.slides.forEach(slide => {
      const div = document.createElement('div');
      div.innerHTML = slide.html;
      container.appendChild(div);
    });

    // 注入全局样式
    if (result.styles?.global) {
      const styleEl = document.createElement('style');
      styleEl.innerHTML = result.styles.global;
      document.head.appendChild(styleEl);
    }
  });
</script>
```

### Node.js 中使用

```javascript
const fs = require('fs');
const { pptxToHtml } = require('@fefeding/ppt-parser');

async function parsePptx() {
  const buffer = fs.readFileSync('presentation.pptx');
  const result = await pptxToHtml(buffer);

  // 输出幻灯片信息
  console.log('幻灯片数量:', result.slides.length);
  console.log('幻灯片大小:', result.slideSize);
  console.log('元数据:', result.metadata);

  // 输出每张幻灯片的 HTML
  result.slides.forEach(slide => {
    console.log(`幻灯片 ${slide.slideNum}:`, slide.html);
  });
}

parsePptx();
```

### Vue 中使用

在 Vue 项目中使用 PPTX Parser，包括图表渲染功能。

#### 1. 安装依赖

```bash
npm install @fefeding/ppt-parser echarts jszip
```

#### 2. 配置组件

```vue
<template>
  <div class="pptx-viewer">
    <!-- 文件上传 -->
    <input type="file" accept=".pptx" @change="handleFileUpload" :disabled="loading" />
    
    <!-- PPT 预览区域 -->
    <div v-if="slides.length > 0">
      <div v-for="(slide, index) in slides" :key="index">
        <div v-html="slide.html"></div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, nextTick, onMounted } from 'vue'
import { pptxToHtml } from '@fefeding/ppt-parser'
import JSZip from 'jszip'
import * as echarts from 'echarts'
import { chartRenderer } from './chart-renderer'  // 需要从示例中复制到项目

// 初始化全局依赖（必须在解析前设置）
onMounted(() => {
  ;(window as any).JSZip = JSZip
  ;(window as any).echarts = echarts
  ;(window as any).chartRenderer = chartRenderer
})

const loading = ref(false)
const slides = ref([])

async function handleFileUpload(event: Event) {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  if (!file) return

  loading.value = true

  try {
    // 读取文件为 ArrayBuffer
    const fileData = await file.arrayBuffer()

    // 解析 PPTX 文件
    const result = await pptxToHtml(fileData, {
      mediaProcess: true,      // 处理媒体文件
      themeProcess: true,      // 处理主题样式
      callbacks: {
        onSlide: (data, info) => {
          console.log(`处理幻灯片 ${info.slideNum}`)
        }
      }
    })

    // 保存解析结果
    slides.value = result.slides || []

    // 等待 DOM 更新后注入全局样式
    await nextTick()
    if (result.styles?.global) {
      applyGlobalStyles(result.styles.global)
    }

    // 渲染图表（关键步骤）
    if (result.charts && result.charts.length > 0) {
      await nextTick()
      console.log('检测到图表:', result.charts.length, '个')
      chartRenderer.renderCharts(result.charts)
    }

  } catch (error) {
    console.error('PPTX 解析失败:', error)
  } finally {
    loading.value = false
  }
}

function applyGlobalStyles(css: string) {
  let styleEl = document.getElementById('pptx-global-styles')
  if (!styleEl) {
    styleEl = document.createElement('style')
    styleEl.id = 'pptx-global-styles'
    document.head.appendChild(styleEl)
  }
  styleEl.innerHTML = css
}
</script>

<style>
/* 引入 PPTX 样式文件 */
@import '@fefeding/ppt-parser/src/css/pptxjs.css';
</style>
```

#### 3. 图表渲染说明

PPTX Parser 支持解析 PPTX 中的图表，并提供两种渲染方式：

**方式一：使用内置图表渲染器（推荐）**

需要从 [`examples/chart-lib/chart-renderer.js`](./examples/chart-lib/chart-renderer.js) 复制该文件到你的项目中。

```typescript
import { chartRenderer } from './chart-renderer'  // 从示例复制到你的项目

// 确保已设置全局 echarts
;(window as any).echarts = echarts

// 解析完成后渲染图表
if (result.charts && result.charts.length > 0) {
  await nextTick() // 等待 DOM 更新
  chartRenderer.renderCharts(result.charts)
}
```

**方式二：自定义图表渲染**

```typescript
// 解析结果中的图表数据结构
interface ChartData {
  chartId: string      // 图表容器 ID
  type: string         // 图表类型（bar, line, pie 等）
  data: Array<any>     // 图表数据
  style: object        // 图表样式
}

// 自定义渲染逻辑
result.charts.forEach((chart: ChartData) => {
  const element = document.getElementById(chart.chartId)
  if (element) {
    const myChart = echarts.init(element)
    const option = convertToEChartsOption(chart) // 自定义转换函数
    myChart.setOption(option)
  }
})
```

#### 4. 完整示例

参考 [`examples/vue-demo`](./examples/vue-demo) 目录查看完整的使用示例，包括：
- 文件上传处理
- 进度显示
- 全屏预览
- 样式注入
- 图表渲染

**重要**：图表渲染器 `chart-renderer.js` 位于 [`examples/chart-lib/chart-renderer.js`](./examples/chart-lib/chart-renderer.js)，需要将其复制到你的 Vue 项目中使用。

#### 5. 注意事项

- **全局依赖设置**：必须在解析前设置 `window.JSZip`、`window.echarts` 和 `window.chartRenderer`
- **图表渲染器**：需要从 [`examples/chart-lib/chart-renderer.js`](./examples/chart-lib/chart-renderer.js) 复制到你的项目中
- **样式加载**：需要引入 PPTX Parser 的样式文件 `pptxjs.css`
- **DOM 更新**：渲染图表前必须使用 `nextTick()` 等待 DOM 更新完成
- **图表容器**：确保图表容器已正确挂载到 DOM 中
- **响应式处理**：ECharts 实例会自动监听窗口大小变化并调整

## 使用场景

- 📊 在线 PPT 编辑器
- 📑 PPT 文件内容提取
- 🔄 PPT 格式转换
- 📤 PPT 报表导出
- 🎨 PPT 模板生成
- 📱 移动端 PPT 查看

## 路线图

查看 [docs/FEATURES.md](./docs/FEATURES.md) 了解功能规划和实现进度。

## 浏览器兼容性

- Chrome ≥ 80
- Firefox ≥ 75
- Edge ≥ 80
- Safari ≥ 14

## 开发

```bash
# 克隆项目
git clone https://github.com/fefeding/pptx-parser.git

# 安装依赖
npm install

# 开发模式
npm run dev

# 构建
npm run build

# 运行测试
npm test
```

## 文档

- [API 文档](./docs/API.md) - 完整的 API 参考
- [功能规划](./docs/FEATURES.md) - 功能开发和路线图

## 贡献

欢迎提交 Issue 和 Pull Request！

## 许可证

[MIT License](LICENSE)

## 致谢

本库在开发过程中参考和借鉴了 [pptxjs](https://github.com/meshesha/pptxjs) 项目的部分实现思路，特此表示感谢。pptxjs 是一个优秀的客户端PPTX解析库，为本项目的架构设计提供了重要参考。

---

**Made with ❤️ for developers**