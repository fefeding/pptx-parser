# 增强版功能使用指南

## 目录

1. [快速开始](#快速开始)
2. [核心概念](#核心概念)
3. [API详解](#api详解)
4. [实战示例](#实战示例)
5. [最佳实践](#最佳实践)
6. [故障排除](#故障排除)

## 快速开始

### 安装

```bash
npm install @fefeding/pptx-parser
# 或
pnpm add @fefeding/pptx-parser
# 或
yarn add @fefeding/pptx-parser
```

### 最简单的示例

```typescript
import { parsePptxEnhanced } from '@fefeding/pptx-parser';

async function parseFile(file: File) {
  const result = await parsePptxEnhanced(file);
  console.log('PPT标题:', result.title);
  console.log('幻灯片数量:', result.slides.length);
}
```

## 核心概念

### PPTX文件结构

PPTX本质是一个ZIP压缩包，包含以下核心文件：

```
presentation.pptx
├── [Content_Types].xml        # 内容类型定义
├── _rels/                    # 根目录关系文件
│   └── .rels
├── docProps/                 # 元数据
│   ├── core.xml             # 核心属性（标题、作者等）
│   └── app.xml             # 应用属性
├── ppt/
│   ├── presentation.xml      # 演示文稿主文件
│   ├── slides/             # 幻灯片目录
│   │   ├── slide1.xml
│   │   ├── slide2.xml
│   │   └── _rels/         # 幻灯片关系文件
│   │       ├── slide1.xml.rels
│   │       └── slide2.xml.rels
│   ├── media/             # 媒体资源
│   │   ├── image1.png
│   │   └── image2.jpg
│   └── theme/             # 主题文件
└── README.md
```

### 命名空间

PPTX使用XML命名空间区分不同类型的标签：

```typescript
import { NS } from 'pptx-parser';

// PresentationML - 幻灯片相关标签
NS.p  // 'http://schemas.openxmlformats.org/presentationml/2006/main'

// DrawingML - 绘图相关标签
NS.a  // 'http://schemas.openxmlformats.org/drawingml/2006/main'

// Relationships - 关联关系标签
NS.r  // 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
```

### EMU单位

PPTX内部使用EMU（English Metric Unit）作为坐标单位：

- 1 英寸 = 914,400 EMU
- 1 像素（96 DPI）= 9,525 EMU

```typescript
import { emu2px, px2emu } from 'pptx-parser';

// EMU 转 像素
const px = emu2px('914400'); // = 96

// 像素 转 EMU
const emu = px2emu(96); // = 914400
```

## API详解

### parsePptxEnhanced

#### 参数说明

| 参数 | 类型 | 必填 | 说明 |
|-----|------|------|------|
| file | File \| Blob \| ArrayBuffer | 是 | PPTX文件 |
| options | ParseOptions | 否 | 解析选项 |

#### ParseOptions选项

```typescript
interface ParseOptions {
  // 是否解析图片为Base64（默认: true）
  parseImages?: boolean;

  // 是否保留原始XML字符串（默认: false）
  keepRawXml?: boolean;

  // 是否输出详细日志（默认: false）
  verbose?: boolean;

  // 自定义命名空间映射（可选）
  customNS?: Record<string, string>;
}
```

#### 返回值说明

```typescript
interface PptxParseResult {
  // 基本信息
  id: string;                    // 文档ID
  title: string;                  // PPT标题
  author?: string;               // 作者
  subject?: string;              // 主题
  keywords?: string;             // 关键词
  description?: string;          // 描述

  // 时间信息
  created?: string;              // 创建时间（ISO 8601）
  modified?: string;             // 修改时间（ISO 8601）

  // 幻灯片数据
  slides: SlideParseResult[];     // 幻灯片数组

  // 页面属性
  props: {
    width: number;              // 宽度（像素）
    height: number;             // 高度（像素）
    ratio: number;              // 宽高比
    pageSize?: '4:3' | '16:9' | '16:10' | 'custom';
  };

  // 全局关系映射
  globalRelsMap?: RelsMap;      // relId -> Relation映射
}
```

### SlideParseResult

```typescript
interface SlideParseResult {
  id: string;                    // 幻灯片ID
  title: string;                 // 幻灯片标题
  background: string;            // 背景颜色（十六进制）

  // 元素列表
  elements: ParsedSlideElement[];

  // 关联关系映射
  relsMap: RelsMap;            // 当前幻灯片的rels映射

  // 原始XML（可选）
  rawXml?: string;
}
```

### ParsedSlideElement

所有元素共有的基础属性：

```typescript
interface ParsedSlideElement {
  // 基本属性
  id: string;                   // 元素唯一ID
  type: ElementType;            // 元素类型
  name?: string;                // 元素名称
  hidden?: boolean;             // 是否隐藏

  // 位置和尺寸（像素）
  rect: PptRect {
    x: number;
    y: number;
    width: number;
    height: number;
  };

  // 样式
  style: PptStyle {
    fontSize?: number;
    color?: string;
    backgroundColor?: string;
    // ... 更多样式属性
  };

  // 内容
  content: any;                 // 类型特定内容

  // 附加属性
  props: Record<string, unknown>;
  relId?: string;              // 关联ID（用于引用图片等）

  // 调试信息
  text?: string;               // 纯文本内容（用于搜索）
  attrs?: Record<string, string>; // 原始XML属性
  rawNode?: Element;           // 原始DOM节点
}
```

### 元素类型详解

#### 文本/形状元素 (type: 'text' | 'shape')

```typescript
interface ParsedShapeElement extends ParsedSlideElement {
  type: 'text' | 'shape';

  // 形状类型（矩形、圆形等）
  shapeType?: 'rectangle' | 'ellipse' | 'triangle' | 'diamond' | ...;

  // 文本内容
  text?: string;

  // 文本样式数组（支持多段不同样式）
  textStyle?: Array<{
    text: string;
    style: {
      fontSize?: number;
      fontFamily?: string;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      strike?: boolean;
      color?: string;
    };
  }>;

  // 占位符信息
  isPlaceholder?: boolean;
  placeholderType?: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer';
}
```

#### 图片元素 (type: 'image')

```typescript
interface ParsedImageElement extends ParsedSlideElement {
  type: 'image';

  // 图片数据
  src: string;                 // Base64或URL
  relId: string;               // 关系ID
  mimeType?: string;            // MIME类型（image/png等）
  altText?: string;            // 替代文本
}
```

#### OLE对象元素 (type: 'ole')

```typescript
interface ParsedOleElement extends ParsedSlideElement {
  type: 'ole';

  // OLE对象信息
  progId?: string;             // 程序ID（如 "Excel.Sheet.8"）
  relId: string;              // 关系ID
  name?: string;               // 对象名称
  hasFallback?: boolean;        // 是否有降级图片
}
```

#### 图表元素 (type: 'chart')

```typescript
interface ParsedChartElement extends ParsedSlideElement {
  type: 'chart';

  // 图表信息
  chartType?: string;          // 图表类型（bar、line、pie等）
  relId: string;              // 关系ID
}
```

#### 分组元素 (type: 'group')

```typescript
interface ParsedGroupElement extends ParsedSlideElement {
  type: 'group';

  // 分组内的子元素
  children: ParsedSlideElement[];
}
```

## 实战示例

### 示例1：构建PPT查看器

```typescript
<template>
  <div class="ppt-viewer">
    <div class="toolbar">
      <button @click="prevSlide" :disabled="currentSlide === 0">上一页</button>
      <span>{{ currentSlide + 1 }} / {{ result?.slides.length }}</span>
      <button @click="nextSlide" :disabled="currentSlide >= (result?.slides.length || 0) - 1">下一页</button>
    </div>

    <div class="slide-container" v-if="currentSlideData">
      <div class="slide" :style="slideStyle">
        <div v-for="element in currentSlideData.elements" :key="element.id"
             class="ppt-element" :style="elementStyle(element)">
          <span v-if="element.text">{{ element.text }}</span>
          <img v-else-if="element.type === 'image'" :src="(element as any).src" />
          <div v-else>元素类型: {{ element.type }}</div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue';
import { parsePptxEnhanced } from 'pptx-parser';

const currentSlide = ref(0);
const result = ref<any>(null);

const currentSlideData = computed(() =>
  result.value?.slides[currentSlide.value]
);

const slideStyle = computed(() => ({
  width: `${result.value?.props.width}px`,
  height: `${result.value?.props.height}px`
}));

function elementStyle(element: any) {
  return {
    position: 'absolute',
    left: `${element.rect.x}px`,
    top: `${element.rect.y}px`,
    width: `${element.rect.width}px`,
    height: `${element.rect.height}px`
  };
}

function prevSlide() {
  currentSlide.value--;
}

function nextSlide() {
  currentSlide.value++;
}

async function loadFile(file: File) {
  result.value = await parsePptxEnhanced(file, {
    parseImages: true
  });
}
</script>
```

### 示例2：文本提取和搜索

```typescript
import { parsePptxEnhanced } from 'pptx-parser';

class PPTSearcher {
  private result: any = null;

  async load(file: File) {
    this.result = await parsePptxEnhanced(file);
  }

  /**
   * 搜索包含特定文本的所有元素
   */
  search(query: string): Array<{
    slideIndex: number;
    slideTitle: string;
    element: any;
    matchText: string;
  }> {
    const matches: any[] = [];

    this.result.slides.forEach((slide: any, slideIndex: number) => {
      slide.elements.forEach((element: any) => {
        if (element.text && element.text.includes(query)) {
          matches.push({
            slideIndex,
            slideTitle: slide.title,
            element,
            matchText: element.text
          });
        }
      });
    });

    return matches;
  }

  /**
   * 提取所有文本
   */
  extractAllText(): string {
    const texts: string[] = [];

    this.result.slides.forEach((slide: any) => {
      slide.elements.forEach((element: any) => {
        if (element.text) {
          texts.push(element.text);
        }
      });
    });

    return texts.join('\n');
  }
}

// 使用示例
const searcher = new PPTSearcher();
await searcher.load(file);

const matches = searcher.search('重要');
console.log('找到', matches.length, '个匹配');

const allText = searcher.extractAllText();
console.log('所有文本:', allText);
```

### 示例3：图片提取和下载

```typescript
import { parsePptxEnhanced } from 'pptx-parser';

async function extractImages(file: File) {
  const result = await parsePptxEnhanced(file, {
    parseImages: true
  });

  const images: Array<{
    slideIndex: number;
    elementId: string;
    src: string;
    relId: string;
    mimeType: string;
  }> = [];

  result.slides.forEach((slide, slideIndex) => {
    slide.elements.forEach((element) => {
      if (element.type === 'image') {
        const img = element as any;
        images.push({
          slideIndex,
          elementId: img.id,
          src: img.src,
          relId: img.relId,
          mimeType: img.mimeType
        });
      }
    });
  });

  console.log(`提取到 ${images.length} 个图片`);

  // 下载所有图片
  images.forEach((img, index) => {
    downloadImage(img.src, `slide_${img.slideIndex + 1}_image_${index + 1}.png`);
  });
}

function downloadImage(base64Data: string, filename: string) {
  const link = document.createElement('a');
  link.href = base64Data;
  link.download = filename;
  link.click();
}

// 使用示例
<input type="file" @change="e => extractImages(e.target.files[0])" />
```

### 示例4：元素统计

```typescript
import { parsePptxEnhanced } from 'pptx-parser';

interface PPTStats {
  slideCount: number;
  textElements: number;
  images: number;
  shapes: number;
  oleObjects: number;
  charts: number;
  groups: number;
  totalCharacters: number;
}

async function analyzePPT(file: File): Promise<PPTStats> {
  const result = await parsePptxEnhanced(file);

  const stats: PPTStats = {
    slideCount: result.slides.length,
    textElements: 0,
    images: 0,
    shapes: 0,
    oleObjects: 0,
    charts: 0,
    groups: 0,
    totalCharacters: 0
  };

  result.slides.forEach((slide) => {
    slide.elements.forEach((element) => {
      switch (element.type) {
        case 'text':
          stats.textElements++;
          stats.totalCharacters += (element.text || '').length;
          break;
        case 'image':
          stats.images++;
          break;
        case 'shape':
          stats.shapes++;
          break;
        case 'ole':
          stats.oleObjects++;
          break;
        case 'chart':
          stats.charts++;
          break;
        case 'group':
          stats.groups++;
          break;
      }
    });
  });

  return stats;
}

// 使用示例
const stats = await analyzePPT(file);
console.table({
  '幻灯片数量': stats.slideCount,
  '文本元素': stats.textElements,
  '图片数量': stats.images,
  '形状数量': stats.shapes,
  'OLE对象': stats.oleObjects,
  '图表数量': stats.charts,
  '分组数量': stats.groups,
  '总字符数': stats.totalCharacters
});
```

## 最佳实践

### 1. 性能优化

**按需加载图片：**
```typescript
// 不解析图片（快速解析）
const result = await parsePptxEnhanced(file, {
  parseImages: false
});

// 按需加载图片
async function loadImage(relId: string): Promise<string> {
  // 实现懒加载逻辑
}
```

**使用Web Worker：**
```typescript
// pptx.worker.ts
import { parsePptxEnhanced } from 'pptx-parser';

self.onmessage = async (e) => {
  const result = await parsePptxEnhanced(e.data);
  self.postMessage(result);
};

// 主线程
const worker = new Worker('pptx.worker.js');
worker.postMessage(file);
worker.onmessage = (e) => {
  console.log('解析结果:', e.data);
};
```

### 2. 错误处理

```typescript
import { parsePptxEnhanced } from 'pptx-parser';

async function safeParse(file: File) {
  try {
    const result = await parsePptxEnhanced(file);

    // 验证结果
    if (!result || result.slides.length === 0) {
      console.warn('PPT文件为空');
      return null;
    }

    return result;
  } catch (error) {
    console.error('PPT解析失败:', error);

    // 分类错误
    if (error instanceof Error) {
      if (error.message.includes('ZIP')) {
        alert('文件不是有效的PPTX格式');
      } else if (error.message.includes('XML')) {
        alert('PPT文件包含损坏的数据');
      } else {
        alert('解析失败，请重试');
      }
    }

    return null;
  }
}
```

### 3. 内存管理

```typescript
// 对于大文件，分片处理
async function processLargeFile(file: File) {
  const chunkSize = 5; // 每次处理5页
  const result = await parsePptxEnhanced(file);

  for (let i = 0; i < result.slides.length; i += chunkSize) {
    const chunk = result.slides.slice(i, i + chunkSize);
    processChunk(chunk); // 处理分片

    // 释放内存
    await new Promise(resolve => setTimeout(resolve, 0));
  }
}
```

## 故障排除

### 问题：解析失败

**可能原因：**
1. 文件不是有效的PPTX格式
2. 文件已损坏
3. 缺少必要的XML文件

**解决方案：**
```typescript
try {
  const result = await parsePptxEnhanced(file);
} catch (error) {
  console.error('错误详情:', error);

  // 检查文件类型
  if (!file.name.endsWith('.pptx')) {
    console.error('文件扩展名不是.pptx');
  }

  // 检查文件大小
  if (file.size === 0) {
    console.error('文件为空');
  }
}
```

### 问题：图片无法显示

**可能原因：**
1. 关联关系文件解析失败
2. 图片文件路径不正确

**解决方案：**
```typescript
const result = await parsePptxEnhanced(file, {
  parseImages: true,
  verbose: true  // 查看详细日志
});

// 检查元素
result.slides.forEach(slide => {
  slide.elements.forEach(element => {
    if (element.type === 'image') {
      const img = element as any;
      if (!img.src || !img.src.startsWith('data:')) {
        console.warn('图片未正确解析:', img.relId);
      }
    }
  });
});
```

### 问题：文本乱码

**可能原因：**
1. XML编码问题
2. 字体不兼容

**解决方案：**
```typescript
// 增强版已自动处理编码
// 如果仍有问题，检查浏览器控制台的警告信息

const result = await parsePptxEnhanced(file, {
  verbose: true
});

// 输出原始文本检查
result.slides.forEach(slide => {
  slide.elements.forEach(element => {
    if (element.text) {
      console.log('文本:', element.text);
      console.log('原始属性:', element.attrs);
    }
  });
});
```

## 总结

本指南涵盖了PPTX Parser增强版的核心概念、API详解、实战示例和最佳实践。如需更多帮助，请参考：

- API文档：`docs/API.md`
- 示例代码：`examples/usage-enhanced.ts`
- 原版文档：`README.md`
