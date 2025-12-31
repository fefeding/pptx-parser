# 功能实现迁移指南

本文档说明如何从基础功能迁移到扩展功能。

## 概述

PPT-Parser 现在分为两层架构：

1. **core.ts** - 基础功能（文本、形状、图片、表格、图表）
2. **core-extended.ts** - 扩展功能（渐变、项目符号、超链接、阴影、变换等）

---

## 使用扩展功能

### 1. 导入扩展模块

```typescript
// 基础功能
import { parsePptx, serializePptx, PptParseUtils } from 'ppt-parser';

// 扩展功能
import { PptParseUtilsExtended } from 'ppt-parser/core-extended';
```

### 2. 解析扩展功能

解析 PPTX 文件后，使用扩展工具函数处理高级特性：

```typescript
async function parseWithExtended(file: File) {
  // 1. 使用基础解析
  const pptDoc = await parsePptx(file);

  // 2. 解析幻灯片 XML 以获取详细信息
  const JSZip = await import('jszip');
  const zip = await JSZip.default.loadAsync(file);

  // 3. 读取幻灯片 XML
  const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(slideXml, 'application/xml');

  // 4. 使用扩展功能解析
  const shapeElements = xmlDoc.querySelectorAll('p\\:sp');
  shapeElements.forEach(shape => {
    const spPr = shape.querySelector('p\\:spPr');

    // 解析渐变填充
    const fillNode = PptParseUtilsExtended.parseXmlFill(spPr);
    if (fillNode.type === 'gradient') {
      console.log('发现渐变:', fillNode.gradientStops);
    }

    // 解析阴影效果
    const effectLst = spPr.querySelector('a\\:effectLst');
    const shadow = PptParseUtilsExtended.parseXmlShadow(effectLst);
    if (shadow) {
      console.log('发现阴影:', shadow);
    }

    // 解析变换效果
    const xfrm = spPr.querySelector('a\\:xfrm');
    const transform = PptParseUtilsExtended.parseXmlTransform(xfrm);
    if (transform.rotate || transform.flipH || transform.flipV) {
      console.log('发现变换:', transform);
    }
  });
}
```

### 3. 序列化扩展功能

创建包含扩展特性的文档：

```typescript
import type { PptDocument } from 'ppt-parser';

function createAdvancedDocument(): PptDocument {
  return {
    id: 'doc-1',
    title: '高级特性示例',
    slides: [
      {
        id: 'slide-1',
        title: '渐变和阴影',
        bgColor: '#ffffff',
        elements: [
          {
            id: 'shape-1',
            type: 'shape',
            rect: { x: 100, y: 100, width: 400, height: 300 },
            style: {
              fontSize: 16,
              fill: {
                type: 'gradient',
                gradientStops: [
                  { position: 0, color: '#ff6b6b' },
                  { position: 1, color: '#4ecdc4' },
                ],
                gradientDirection: 45,
              },
              shadow: {
                color: '#000000',
                blur: 10,
                offsetX: 5,
                offsetY: 5,
                opacity: 0.3,
              },
            },
            content: { shapeType: 'rectangle' },
            props: {},
          },
        ],
        props: {
          width: 1280,
          height: 720,
          slideLayout: 'blank',
        },
      },
    ],
    props: {
      width: 1280,
      height: 720,
      ratio: 1.78,
    },
  };
}
```

---

## 功能对比

| 功能 | 基础功能 (core.ts) | 扩展功能 (core-extended.ts) |
|------|---------------------|----------------------------|
| 文本解析 | ✅ 基础文本 | ✅ 项目符号、编号列表、超链接 |
| 文本样式 | ✅ 字体、颜色、对齐 | ✅ 下划线、删除线、行高、字间距 |
| 形状类型 | ✅ 矩形、圆形等基础形状 | ✅ 180+ Office 预设形状 |
| 填充效果 | ✅ 纯色填充 | ✅ 渐变、图片、图案填充 |
| 边框样式 | ✅ 基础边框 | ✅ 虚线、点线、双线、自定义样式 |
| 阴影效果 | ❌ | ✅ 阴影（颜色、模糊、偏移、透明度） |
| 变换效果 | ❌ | ✅ 旋转、水平翻转、垂直翻转 |
| 主题支持 | ❌ | ✅ 主题颜色映射 |
| 关系映射 | ❌ | ✅ 图片、媒体资源映射 |

---

## API 差异

### 基础样式 vs 扩展样式

#### 基础样式（core.ts）

```typescript
interface PptStyle {
  fontSize?: number;
  color?: string;
  fontWeight?: 'normal' | 'bold';
  textAlign?: 'left' | 'center' | 'right';
  backgroundColor?: string;
  borderColor?: string;
  borderWidth?: number;
}
```

#### 扩展样式（core-extended.ts）

```typescript
interface PptStyle {
  // 基础样式
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

  // 扩展样式
  fill?: PptFill;
  border?: PptBorder;
  shadow?: PptShadow;
  opacity?: number;
}
```

### 填充效果

```typescript
// 基础：仅支持纯色
backgroundColor: '#ff0000';

// 扩展：支持多种填充类型
fill: {
  type: 'gradient',
  gradientStops: [
    { position: 0, color: '#ff0000' },
    { position: 1, color: '#0000ff' },
  ],
  gradientDirection: 45,
}
```

### 文本内容

```typescript
// 基础：简单字符串
content: '这是一段文本';

// 扩展：支持段落级别的详细控制
content: [
  {
    text: '• 项目符号文本',
    bullet: { type: 'bullet', level: 0 },
  },
  {
    text: '  • 二级项目符号',
    bullet: { type: 'bullet', level: 1 },
  },
  {
    text: '1. 编号列表',
    bullet: { type: 'numbered', level: 0 },
  },
  {
    text: '访问链接',
    hyperlink: { url: 'https://example.com', tooltip: '提示' },
  },
];
```

---

## 迁移步骤

### 从基础功能迁移到扩展功能

#### 步骤 1: 更新导入

```typescript
// 之前
import { PptParseUtils } from 'ppt-parser';

// 之后
import { PptParseUtils } from 'ppt-parser';
import { PptParseUtilsExtended } from 'ppt-parser/core-extended';
```

#### 步骤 2: 更新样式定义

```typescript
// 之前
const element = {
  style: {
    fontSize: 16,
    color: '#333333',
    backgroundColor: '#ffffff',
  },
};

// 之后
const element = {
  style: {
    fontSize: 16,
    color: '#333333',
    fill: {
      type: 'gradient',
      gradientStops: [
        { position: 0, color: '#ffffff' },
        { position: 1, color: '#f0f0f0' },
      ],
    },
    shadow: {
      color: '#000000',
      blur: 5,
      offsetX: 3,
      offsetY: 3,
      opacity: 0.2,
    },
  },
};
```

#### 步骤 3: 更新文本内容

```typescript
// 之前
const textElement = {
  content: '这是一段文本',
};

// 之后
const textElement = {
  content: [
    {
      text: '• 第一项',
      bullet: { type: 'bullet', level: 0 },
    },
    {
      text: '• 第二项',
      bullet: { type: 'bullet', level: 0 },
    },
  ],
};
```

#### 步骤 4: 添加变换效果

```typescript
const element = {
  rect: { x: 100, y: 100, width: 200, height: 150 },
  transform: {
    rotate: 45,        // 旋转45度
    flipH: false,      // 不水平翻转
    flipV: false,      // 不垂直翻转
  },
};
```

---

## 常见问题

### Q1: 如何同时使用基础功能和扩展功能？

A: 可以同时导入并使用。基础功能用于解析和序列化，扩展功能用于处理高级特性。

### Q2: 扩展功能是否会增加文件体积？

A: 不会显著增加。core-extended.ts 是独立的模块，按需导入使用。

### Q3: 向后兼容性如何保证？

A: 基础 API 保持不变，扩展功能是可选的。现有代码无需修改即可继续工作。

### Q4: 如何判断 PPTX 文件包含扩展特性？

A: 解析文件后，检查元素的 style、content、transform 等属性是否包含扩展特性。

```typescript
function hasExtendedFeatures(element: PptElement): boolean {
  return !!(
    element.style.fill ||
    element.style.shadow ||
    element.style.border ||
    element.transform
  );
}
```

---

## 示例项目

查看 `examples/extended-features.ts` 了解完整的使用示例。

---

## 下一步

1. 阅读 [API 文档](./API.md) 了解完整的 API 参考
2. 查看 [功能规划](./FEATURES.md) 了解即将推出的功能
3. 运行示例代码体验扩展功能
