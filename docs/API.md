# API 设计文档

## 概述

本文档详细说明了 PPT-Parser 的 API 设计，包括配置选项、方法签名和类型定义。

---

## 核心接口

### parsePptx

解析 PPTX 文件为结构化 JSON 数据。

```typescript
function parsePptx(
  file: File | Blob | string | ArrayBuffer,
  options?: ParseOptions
): Promise<PptDocument>;
```

#### 参数

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| file | `File \| Blob \| string \| ArrayBuffer` | 是 | PPTX 文件，支持多种输入格式 |
| options | `ParseOptions` | 否 | 解析配置选项 |

#### ParseOptions

```typescript
interface ParseOptions {
  /** 是否提取图片二进制数据（默认：false） */
  extractImages?: boolean;
  /** 是否解析媒体文件（视频、音频）（默认：true） */
  parseMedia?: boolean;
  /** 是否解析主题（默认：false） */
  parseTheme?: boolean;
  /** 是否解析母版（默认：false） */
  parseMaster?: boolean;
  /** 是否解析演讲者备注（默认：false） */
  parseNotes?: boolean;
  /** 自定义单位转换比例（默认：96 / 914400） */
  emuToPxRatio?: number;
  /** 进度回调函数 */
  onProgress?: (progress: number, message: string) => void;
}
```

#### 返回值

返回 `Promise<PptDocument>`，详见类型定义。

---

### serializePptx

将结构化 JSON 数据序列化为 PPTX 文件。

```typescript
function serializePptx(
  document: PptDocument,
  options?: SerializeOptions
): Promise<Blob>;
```

#### 参数

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| document | `PptDocument` | 是 | PPT 文档结构化数据 |
| options | `SerializeOptions` | 否 | 序列化配置选项 |

#### SerializeOptions

```typescript
interface SerializeOptions {
  /** 是否包含备注（默认：false） */
  includeNotes?: boolean;
  /** 是否压缩输出（默认：true） */
  compress?: boolean;
  /** 压缩级别（0-9，默认：6） */
  compressionLevel?: number;
  /** 进度回调函数 */
  onProgress?: (progress: number, message: string) => void;
}
```

#### 返回值

返回 `Promise<Blob>`，可直接下载或保存。

---

## 工具函数

### PptParseUtils

```typescript
const PptParseUtils = {
  // ID 生成
  generateId(prefix?: string): string;

  // XML 解析
  parseXmlText(text: string): string;
  parseXmlAttrs(attrs: NamedNodeMap): Record<string, string>;
  parseXmlToTree(xmlStr: string): XmlElement;

  // 坐标解析
  parseXmlRect(attrs: Record<string, string>): PptRect;

  // 样式解析
  parseXmlStyle(attrs: Record<string, string>): PptStyle;

  // 单位转换
  px2emu(px: number): number;
  emu2px(emu: number): number;

  // 颜色转换
  hexToRgb(hex: string): { r: number; g: number; b: number };
  rgbToHex(r: number, g: number, b: number): string;
  parseColor(color: string): string;

  // 形状处理
  createShapePath(type: PptShapeType, rect: PptRect): string;
  calculateShapeBounds(path: string): PptRect;

  // 文本处理
  parseTextRuns(textBody: string): PptTextParagraph[];
  escapeHtml(text: string): string;
};
```

---

## 类型定义详解

### PptDocument

```typescript
interface PptDocument {
  /** 文档唯一 ID */
  id: string;
  /** 文档标题 */
  title: string;
  /** 作者 */
  author?: string;
  /** 主题 */
  subject?: string;
  /** 关键词 */
  keywords?: string;
  /** 描述 */
  description?: string;
  /** 创建时间 */
  created?: string;
  /** 修改时间 */
  modified?: string;
  /** 幻灯片数组 */
  slides: PptSlide[];
  /** 主题定义 */
  theme?: PptTheme;
  /** 文档属性 */
  props: {
    width: number;
    height: number;
    ratio: number;
    pageSize?: '4:3' | '16:9' | '16:10' | 'custom';
  };
}
```

### PptSlide

```typescript
interface PptSlide {
  /** 幻灯片唯一 ID */
  id: string;
  /** 幻灯片标题 */
  title: string;
  /** 背景颜色或填充 */
  bgColor: string | PptFill;
  /** 背景图片 */
  backgroundImage?: string;
  /** 元素数组 */
  elements: PptElement[];
  /** 幻灯片属性 */
  props: {
    width: number;
    height: number;
    slideLayout: PptSlideLayout;
    transition?: PptSlideTransition;
    notes?: string;
    slideNumber?: number;
  };
}
```

### PptElement

```typescript
interface PptElement {
  /** 元素唯一 ID */
  id: string;
  /** 元素类型 */
  type: PptNodeType;
  /** 坐标和尺寸 */
  rect: PptRect;
  /** 变换效果 */
  transform?: PptTransform;
  /** 样式 */
  style: PptStyle;
  /** 内容 */
  content: PptElementContent;
  /** 扩展属性 */
  props: Record<string, unknown>;
  /** 子元素 */
  children?: PptElement[];
  /** 父元素 ID */
  parentId?: string;
}
```

### PptStyle

```typescript
interface PptStyle {
  // 文本样式
  fontSize?: number;
  fontFamily?: string;
  fontStyle?: 'normal' | 'italic';
  fontWeight?: 'normal' | 'bold';
  textDecoration?: 'none' | 'underline' | 'line-through';
  color?: string;
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  textVerticalAlign?: 'top' | 'middle' | 'bottom';
  lineHeight?: number;
  letterSpacing?: number;
  textShadow?: string;

  // 填充
  backgroundColor?: string | PptFill;
  fill?: PptFill;

  // 边框
  borderColor?: string;
  borderWidth?: number;
  borderStyle?: 'solid' | 'dashed' | 'dotted' | 'double';
  border?: PptBorder;

  // 效果
  shadow?: PptShadow;
  reflection?: PptReflection;
  glow?: PptGlow;
  effect3d?: PptEffect3D;

  // 其他
  opacity?: number;
  zIndex?: number;
}
```

---

## 使用示例

### 基础解析

```typescript
import PptParserCore from 'ppt-parser';

// 解析文件
const file = document.querySelector('#ppt-upload').files[0];
const pptDoc = await PptParserCore.parse(file);

console.log(pptDoc.title);
console.log(pptDoc.slides.length);
```

### 高级解析

```typescript
const pptDoc = await PptParserCore.parse(file, {
  extractImages: true,
  parseMedia: true,
  parseTheme: true,
  onProgress: (progress, message) => {
    console.log(`${progress}%: ${message}`);
  }
});
```

### 导出 PPTX

```typescript
const pptDoc = {
  id: 'doc-1',
  title: '我的演示文稿',
  slides: [
    {
      id: 'slide-1',
      title: '第一页',
      bgColor: '#ffffff',
      elements: [
        {
          id: 'text-1',
          type: 'text',
          rect: { x: 100, y: 100, width: 400, height: 50 },
          style: { fontSize: 32, color: '#333333' },
          content: '欢迎使用 PPT-Parser',
          props: {}
        }
      ],
      props: { width: 1280, height: 720, slideLayout: 'blank' }
    }
  ],
  props: { width: 1280, height: 720, ratio: 1.78 }
};

const blob = await PptParserCore.serialize(pptDoc);

// 下载文件
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'presentation.pptx';
a.click();
URL.revokeObjectURL(url);
```

### 创建复杂幻灯片

```typescript
const slide = {
  id: 'slide-1',
  title: '产品介绍',
  bgColor: '#ffffff',
  elements: [
    // 标题
    {
      id: 'title-1',
      type: 'text',
      rect: { x: 100, y: 50, width: 1080, height: 80 },
      style: {
        fontSize: 48,
        fontWeight: 'bold',
        textAlign: 'center',
        color: '#1a73e8'
      },
      content: '产品介绍',
      props: {}
    },
    // 图片
    {
      id: 'image-1',
      type: 'image',
      rect: { x: 100, y: 180, width: 500, height: 350 },
      style: {},
      content: {
        src: 'data:image/png;base64,...',
        alt: '产品图片'
      },
      props: {}
    },
    // 文本框
    {
      id: 'text-2',
      type: 'text',
      rect: { x: 650, y: 180, width: 530, height: 350 },
      style: {
        fontSize: 18,
        lineHeight: 1.6,
        color: '#5f6368'
      },
      content: [
        { text: '• 产品特点1', bullet: { type: 'bullet', level: 0 } },
        { text: '• 产品特点2', bullet: { type: 'bullet', level: 0 } },
        { text: '• 产品特点3', bullet: { type: 'bullet', level: 0 } }
      ],
      props: {}
    }
  ],
  props: {
    width: 1280,
    height: 720,
    slideLayout: 'contentWithCaption'
  }
};
```

### 使用工具函数

```typescript
import { PptParseUtils } from 'ppt-parser';

// 生成唯一 ID
const slideId = PptParseUtils.generateId('slide');
// "slide-1703284800000-1234"

// 单位转换
const emu = PptParseUtils.px2emu(100);
// 952500

const px = PptParseUtils.emu2px(914400);
// 96

// 解析 XML
const xmlTree = PptParseUtils.parseXmlToTree(xmlString);
```

---

## 错误处理

### 错误类型

```typescript
class PptParseError extends Error {
  constructor(
    message: string,
    public code: string,
    public details?: any
  ) {
    super(message);
  }
}

class PptSerializeError extends Error {
  constructor(
    message: string,
    public code: string,
    public details?: any
  ) {
    super(message);
  }
}
```

### 错误代码

| 代码 | 说明 |
|------|------|
| `INVALID_FILE` | 文件格式无效 |
| `PARSE_ERROR` | XML 解析失败 |
| `MISSING_RESOURCE` | 缺少必需资源（图片、媒体等） |
| `SERIALIZE_ERROR` | 序列化失败 |
| `UNSUPPORTED_FEATURE` | 不支持的功能 |

---

## 性能优化

### 大文件处理

```typescript
// 使用流式处理（计划支持）
const pptDoc = await PptParserCore.parse(file, {
  stream: true,
  chunkSize: 1024 * 1024 // 1MB chunks
});
```

### 内存优化

```typescript
// 不提取图片，仅保留引用
const pptDoc = await PptParserCore.parse(file, {
  extractImages: false
});
```

---

## 浏览器兼容性

| 特性 | Chrome | Firefox | Safari | Edge | IE11 |
|------|--------|---------|--------|------|------|
| 基础解析 | ✅ | ✅ | ✅ | ✅ | ✅ |
| 文本样式 | ✅ | ✅ | ✅ | ✅ | ⚠️ |
| 图片处理 | ✅ | ✅ | ✅ | ✅ | ⚠️ |
| 媒体支持 | ✅ | ✅ | ✅ | ✅ | ❌ |
| 渐变填充 | ✅ | ✅ | ✅ | ✅ | ⚠️ |
| 3D 效果 | ✅ | ✅ | ⚠️ | ✅ | ❌ |

⚠️ 部分支持
