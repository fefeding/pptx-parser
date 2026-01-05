# PPTX Parser 增强版文档

## 概述

PPTX Parser 增强版基于原库进行增量式扩展，完全兼容原有API，同时大幅提升解析能力。增强版遵循 ECMA-376 OpenXML 标准，能够完整解析标准PPTX文件的所有内容。

## 核心特性

### ✅ 完全兼容原库
- 保留所有原有的API调用方式
- `parsePptx()` 函数签名不变
- 类型定义完全向后兼容
- 零破坏性修改，老代码无缝迁移

### ✅ 完整解析幻灯片元素
增强版能够解析 `<p:spTree>` 下的4类核心节点：

| 元素类型 | XML标签 | 说明 |
|---------|---------|------|
| 形状/文本框 | `<p:sp>` | 普通文本框、自定义形状、占位符 |
| 图片 | `<p:pic>` | 幻灯片中的图片资源 |
| OLE嵌入对象 | `<p:graphicFrame>` | Excel表格、think-cell等第三方组件 |
| 分组 | `<p:grpSp>` | 元素分组，支持嵌套 |

### ✅ 命名空间标准处理
- 使用标准的命名空间查询（`getElementsByTagNameNS`）
- 支持完整命名空间路径（如 `p:cSld`）和简化路径（如 `cSld`）
- 完全遵循 ECMA-376 OpenXML 标准

### ✅ 图片资源解析
- 自动解析图片为Base64格式
- 支持PNG、JPG、GIF等常见格式
- 从关联关系文件（rels）正确映射图片路径

### ✅ 文本样式解析
- 字体大小、颜色、字体名称
- 加粗、斜体、下划线、删除线
- 支持多段落、多文本运行拼接
- 中英文混合文本自动合并

### ✅ 元数据提取
- 标题、作者、主题
- 创建时间、修改时间
- 关键词、描述信息

### ✅ 完善的容错处理
- 节点不存在时返回默认值
- 解析失败时返回空数组而非抛出异常
- 支持非标准扩展标签（如微软的 a16、p14）

## 快速开始

### 基础使用

```typescript
import { parsePptxEnhanced } from 'pptx-parser';

// 最简单的调用
const result = await parsePptxEnhanced(file);

console.log('PPT标题:', result.title);
console.log('幻灯片数量:', result.slides.length);
```

### 带选项解析

```typescript
import { parsePptxEnhanced, type ParseOptions } from 'pptx-parser';

const options: ParseOptions = {
  parseImages: true,    // 解析图片为Base64
  keepRawXml: false,    // 保留原始XML
  verbose: true          // 详细日志
};

const result = await parsePptxEnhanced(file, options);
```

### 遍历幻灯片和元素

```typescript
const result = await parsePptxEnhanced(file);

result.slides.forEach((slide, slideIndex) => {
  console.log(`幻灯片 ${slideIndex + 1}: ${slide.title}`);

  slide.elements.forEach(element => {
    console.log(`  ${element.type}: ID=${element.id}`);

    if (element.type === 'text' || element.type === 'shape') {
      console.log(`    文本: ${element.text}`);
    } else if (element.type === 'image') {
      console.log(`    图片: ${element.src}`);
    } else if (element.type === 'ole') {
      console.log(`    OLE对象: ${element.progId}`);
    }
  });
});
```

## API文档

### parsePptxEnhanced(file, options?)

主解析函数，解析PPTX文件为结构化数据。

**参数：**
- `file`: File | Blob | ArrayBuffer - PPTX文件
- `options?`: ParseOptions - 解析选项

**返回值：** `Promise<PptxParseResult>`

**ParseOptions 选项：**
| 选项 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| parseImages | boolean | true | 是否解析图片为Base64 |
| keepRawXml | boolean | false | 是否保留原始XML字符串 |
| verbose | boolean | false | 是否输出详细日志 |

### PptxParseResult

解析结果的顶层结构。

```typescript
interface PptxParseResult {
  id: string;                    // PPT文档ID
  title: string;                  // PPT标题
  author?: string;               // 作者
  subject?: string;              // 主题
  keywords?: string;             // 关键词
  description?: string;          // 描述
  created?: string;              // 创建时间（ISO格式）
  modified?: string;             // 修改时间（ISO格式）
  slides: SlideParseResult[];     // 幻灯片列表
  props: {
    width: number;              // 页面宽度（像素）
    height: number;             // 页面高度（像素）
    ratio: number;              // 宽高比
    pageSize?: '4:3' | '16:9' | '16:10' | 'custom';
  };
  globalRelsMap?: RelsMap;      // 全局关联关系映射
}
```

### SlideParseResult

单个幻灯片的解析结果。

```typescript
interface SlideParseResult {
  id: string;                  // 幻灯片ID
  title: string;               // 幻灯片标题
  background: string;          // 背景颜色（十六进制）
  elements: ParsedSlideElement[]; // 元素列表
  relsMap: RelsMap;          // 关联关系映射
  rawXml?: string;            // 原始XML（如果keepRawXml为true）
}
```

### ParsedSlideElement

解析后的幻灯片元素（基础接口）。

```typescript
interface ParsedSlideElement {
  id: string;                 // 元素ID
  type: 'shape' | 'image' | 'ole' | 'chart' | 'group';
  rect: PptRect;            // 位置和尺寸（像素）
  style: PptStyle;          // 样式
  content: any;              // 内容
  props: Record<string, unknown>; // 附加属性
  name?: string;             // 元素名称
  hidden?: boolean;          // 是否隐藏
  text?: string;             // 纯文本内容
  relId?: string;           // 关联ID
  attrs?: Record<string, string>; // 原始属性
  rawNode?: Element;        // 原始XML节点
}
```

### 元素类型特定接口

#### ParsedShapeElement（形状/文本框）

```typescript
interface ParsedShapeElement extends ParsedSlideElement {
  type: 'shape' | 'text';
  shapeType?: string;            // 形状类型
  text?: string;                // 文本内容
  textStyle?: Array<{          // 文本样式数组
    text: string;
    style: Partial<PptTextStyle>;
  }>;
  isPlaceholder?: boolean;       // 是否占位符
  placeholderType?: string;     // 占位符类型
}
```

#### ParsedImageElement（图片）

```typescript
interface ParsedImageElement extends ParsedSlideElement {
  type: 'image';
  src: string;                 // 图片URL或Base64
  relId: string;               // 关联ID
  mimeType?: string;            // MIME类型
  altText?: string;            // 替代文本
}
```

#### ParsedOleElement（OLE对象）

```typescript
interface ParsedOleElement extends ParsedSlideElement {
  type: 'ole';
  progId?: string;             // OLE类型标识符
  relId: string;              // 关联ID
  name?: string;               // 对象名称
  hasFallback?: boolean;        // 是否有降级图片
}
```

#### ParsedChartElement（图表）

```typescript
interface ParsedChartElement extends ParsedSlideElement {
  type: 'chart';
  chartType?: string;          // 图表类型
  relId: string;              // 关联ID
}
```

#### ParsedGroupElement（分组）

```typescript
interface ParsedGroupElement extends ParsedSlideElement {
  type: 'group';
  children: ParsedSlideElement[]; // 分组内的子元素
}
```

## 工具函数

### emu2px(emu)

EMU单位转换为像素。

```typescript
import { emu2px } from 'pptx-parser';

const pixels = emu2px('914400'); // = 96px
```

### px2emu(px)

像素转换为EMU单位。

```typescript
import { px2emu } from 'pptx-parser';

const emu = px2emu(96); // = 914400
```

### getAttrs(node)

提取XML节点的所有属性。

```typescript
import { getAttrs } from 'pptx-parser';

const attrs = getAttrs(xmlElement);
console.log(attrs.id, attrs.name, attrs.hidden);
```

### parseRels(relsXml)

解析关联关系文件。

```typescript
import { parseRels } from 'pptx-parser';

const relsMap = parseRels(relsXmlString);
console.log(relsMap['rId1'].target);
```

### parseMetadata(coreXml)

解析元数据文件。

```typescript
import { parseMetadata } from 'pptx-parser';

const metadata = parseMetadata(coreXmlString);
console.log(metadata.title, metadata.author);
```

## 使用场景

### 场景1：提取所有文本

```typescript
const result = await parsePptxEnhanced(file);

const allText: string[] = [];

result.slides.forEach(slide => {
  slide.elements.forEach(element => {
    if (element.text) {
      allText.push(element.text);
    }
  });
});

console.log('所有文本:', allText.join('\n'));
```

### 场景2：提取所有图片

```typescript
const result = await parsePptxEnhanced(file, {
  parseImages: true
});

const images: Array<{
  slideIndex: number;
  src: string;
  relId: string;
}> = [];

result.slides.forEach((slide, slideIndex) => {
  slide.elements.forEach(element => {
    if (element.type === 'image') {
      images.push({
        slideIndex,
        src: (element as any).src,
        relId: (element as any).relId
      });
    }
  });
});

console.log(`找到 ${images.length} 个图片`);
```

### 场景3：搜索文本

```typescript
const result = await parsePptxEnhanced(file);
const searchText = '重要';

result.slides.forEach((slide, slideIndex) => {
  slide.elements.forEach(element => {
    if (element.text && element.text.includes(searchText)) {
      console.log(`在幻灯片 ${slideIndex + 1} 找到匹配: ${element.text}`);
    }
  });
});
```

### 场景4：统计元素类型

```typescript
const result = await parsePptxEnhanced(file);

const stats = {
  text: 0,
  image: 0,
  shape: 0,
  ole: 0,
  chart: 0,
  group: 0
};

result.slides.forEach(slide => {
  slide.elements.forEach(element => {
    const type = element.type as string;
    if (stats[type as keyof typeof stats] !== undefined) {
      (stats[type as keyof typeof stats] as number)++;
    }
  });
});

console.log('元素统计:', stats);
```

## 迁移指南

### 从原版迁移到增强版

原版代码：
```typescript
import { parsePptx } from 'pptx-parser';
const result = await parsePptx(file);
```

迁移到增强版（只需修改导入）：
```typescript
import { parsePptxEnhanced as parsePptx } from 'pptx-parser';
const result = await parsePptx(file);
```

### 访问新的增强功能

```typescript
import { parsePptxEnhanced } from 'pptx-parser';

const result = await parsePptxEnhanced(file, {
  parseImages: true
});

// 访问新的属性
console.log('作者:', result.author);        // 新增
console.log('元数据:', result.created);     // 新增

// 访问增强的元素属性
result.slides.forEach(slide => {
  slide.elements.forEach(element => {
    console.log('文本:', element.text);      // 新增
    console.log('关联ID:', element.relId);   // 新增
    console.log('原始属性:', element.attrs); // 新增
  });
});
```

## 故障排除

### 问题1：解析不出元素

**原因：** 命名空间匹配错误

**解决：** 确保使用增强版的 `parsePptxEnhanced` 函数，它会自动处理命名空间。

### 问题2：图片无法显示

**原因：** 关联关系文件解析失败

**解决：** 检查 `parseImages` 选项是否为 `true`，并查看控制台日志。

### 问题3：中文乱码

**原因：** XML编码问题

**解决：** 增强版已修复编码问题，使用 `DOMParser` 自动处理。

## 性能优化

### 大文件处理

对于大文件（>10MB），建议：

```typescript
const result = await parsePptxEnhanced(file, {
  parseImages: false  // 延迟加载图片
});

// 按需加载图片
async function loadImage(relId: string, slideIndex: number) {
  // 单独解析图片
}
```

### 批量处理

```typescript
// 使用 Web Worker 避免阻塞UI线程
const worker = new Worker('pptx-worker.js');
worker.postMessage(file);
worker.onmessage = (e) => {
  const result = e.data;
  // 处理结果
};
```

## 常见问题

**Q: 增强版和原版有什么区别？**

A: 增强版在完全兼容原版的基础上，新增了：
1. 完整的元素解析（支持OLE对象、分组等）
2. 图片Base64解析
3. 文本样式解析
4. 元数据提取
5. 完善的错误处理

**Q: 可以混用原版和增强版API吗？**

A: 可以。原版API保持不变，可以继续使用 `parsePptx`，新功能通过 `parsePptxEnhanced` 访问。

**Q: 性能如何？**

A: 增强版的解析速度与原版相当，图片解析会略微增加内存占用，可以通过 `parseImages: false` 优化。

**Q: 支持哪些PPTX版本？**

A: 支持所有标准PPTX格式（.pptx），基于 ECMA-376 标准。

## 技术支持

- 文档：见项目 `docs/` 目录
- 示例：见 `examples/usage-enhanced.ts`
- 问题反馈：提交 Issue 到项目仓库

## 更新日志

### v1.0.0 (2025-01-01)
- ✅ 完整解析幻灯片元素（形状、图片、OLE、分组）
- ✅ 命名空间标准处理
- ✅ 图片资源Base64解析
- ✅ 文本样式解析
- ✅ 元数据提取
- ✅ 完善的容错处理
- ✅ 向后兼容原版API
