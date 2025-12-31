# 快速开始

5 分钟上手 PPT-Parser！

---

## 基础使用

### 1. 安装

```bash
npm install ppt-parser
```

### 2. 解析 PPTX 文件

```typescript
import PptParserCore from 'ppt-parser';

const file = document.querySelector('#ppt-upload').files[0];
const pptDoc = await PptParserCore.parse(file);

console.log('幻灯片数量:', pptDoc.slides.length);
console.log('标题:', pptDoc.title);
```

### 3. 导出 PPTX 文件

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
          rect: { x: 100, y: 100, width: 1080, height: 100 },
          style: {
            fontSize: 48,
            color: '#333333',
            textAlign: 'center',
          },
          content: '欢迎使用 PPT-Parser',
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

const blob = await PptParserCore.serialize(pptDoc);

// 下载
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'presentation.pptx';
a.click();
URL.revokeObjectURL(url);
```

---

## 使用扩展功能

### 渐变填充

```typescript
import PptParserCore from 'ppt-parser';
const { utilsExtended } = PptParserCore;

// 创建带渐变的形状
const slide = {
  id: 'slide-1',
  title: '渐变填充',
  bgColor: '#ffffff',
  elements: [
    {
      id: 'shape-1',
      type: 'shape',
      rect: { x: 100, y: 100, width: 400, height: 300 },
      style: {
        fill: {
          type: 'gradient',
          gradientStops: [
            { position: 0, color: '#ff6b6b' },
            { position: 1, color: '#4ecdc4' },
          ],
          gradientDirection: 45,
        },
      },
      content: { shapeType: 'rectangle' },
      props: {},
    },
  ],
  props: { width: 1280, height: 720, slideLayout: 'blank' },
};
```

### 项目符号和编号

```typescript
// 创建带项目符号的文本
const textElement = {
  id: 'text-1',
  type: 'text',
  rect: { x: 100, y: 100, width: 1080, height: 520 },
  style: {
    fontSize: 18,
    lineHeight: 1.8,
    color: '#333333',
  },
  content: [
    { text: '• 一级项目符号', bullet: { type: 'bullet', level: 0 } },
    { text: '  • 二级项目符号', bullet: { type: 'bullet', level: 1 } },
    { text: '1. 编号列表项 1', bullet: { type: 'numbered', level: 0 } },
    { text: '2. 编号列表项 2', bullet: { type: 'numbered', level: 0 } },
  ],
  props: {},
};
```

### 阴影效果

```typescript
// 创建带阴影的形状
const shapeElement = {
  id: 'shape-1',
  type: 'shape',
  rect: { x: 200, y: 150, width: 300, height: 200 },
  style: {
    backgroundColor: '#ffffff',
    fill: { type: 'solid', color: '#ffffff' },
    shadow: {
      color: '#000000',
      blur: 15,
      offsetX: 8,
      offsetY: 8,
      opacity: 0.4,
    },
  },
  content: { shapeType: 'rectangle' },
  props: {},
};
```

### 旋转和翻转

```typescript
// 创建带变换效果的形状
const element = {
  id: 'shape-1',
  type: 'shape',
  rect: { x: 200, y: 150, width: 200, height: 200 },
  transform: {
    rotate: 45,        // 旋转45度
    flipH: false,      // 不水平翻转
    flipV: false,      // 不垂直翻转
  },
  style: {
    backgroundColor: '#ff6b6b',
  },
  content: { shapeType: 'rectangle' },
  props: {},
};
```

---

## 完整示例

查看 `examples/extended-features.ts` 了解更多示例，包括：

- 渐变填充示例
- 项目符号示例
- 超链接示例
- 阴影效果示例
- 变换效果示例
- 边框样式示例

---

## 下一步

1. 📖 阅读 [API 文档](./API.md) 了解完整的 API 参考
2. 🚀 查看 [功能规划](./FEATURES.md) 了解即将推出的功能
3. 🔄 查看 [迁移指南](./MIGRATION.md) 从基础功能迁移到扩展功能
4. 💻 运行示例代码：`npm run dev` 然后 `node examples/extended-features.ts`

---

## 常见问题

### Q: 如何处理大文件？

A: 使用 `onProgress` 回调跟踪解析进度：

```typescript
const pptDoc = await PptParserCore.parse(file, {
  onProgress: (progress, message) => {
    console.log(`${progress}%: ${message}`);
  },
});
```

### Q: 如何提取图片？

A: 使用 `extractImages` 选项：

```typescript
const pptDoc = await PptParserCore.parse(file, {
  extractImages: true,
});

// 图片会作为 Base64 数据嵌入到元素中
pptDoc.slides.forEach(slide => {
  slide.elements.forEach(element => {
    if (element.type === 'image') {
      console.log('图片:', element.content);
    }
  });
});
```

### Q: 如何自定义输出？

A: 使用序列化选项：

```typescript
const blob = await PptParserCore.serialize(pptDoc, {
  includeNotes: true,    // 包含演讲者备注
  compress: true,         // 压缩输出
  compressionLevel: 6     // 压缩级别 (0-9)
});
```

---

## 获取帮助

- 📧 提交 Issue: [GitHub Issues](https://github.com/fefeding/pptx-parser/issues)
- 💬 讨论: [GitHub Discussions](https://github.com/fefeding/pptx-parser/discussions)
- 📧 邮件: support@example.com

---

开始使用 PPT-Parser，轻松处理 PowerPoint 文件！
