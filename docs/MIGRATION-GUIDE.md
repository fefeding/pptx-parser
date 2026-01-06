# PPTX Parser 架构重构迁移指南

## 概述

从版本 0.x 升级到 1.0+，PPTX Parser 进行了重大架构重构。主要的改进包括：

1. **统一的元素架构**：所有文档结构（包括文档本身）都继承自 `BaseElement`
2. **移除独立的 renderer 目录**：HTML 渲染功能整合到各个 Element 类中
3. **引入 DocumentElement**：作为整个文档的主对象，封装所有文档信息和渲染逻辑

## 主要变化

### 1. DocumentElement 替代 PptxDocument

**旧版本：**
```typescript
import { parsePptx, ppt2HTMLDocument } from 'pptx-parser';

const result = await parsePptx(file);
const html = ppt2HTMLDocument(result);
```

**新版本：**
```typescript
import { parsePptx, createDocument } from 'pptx-parser';

const result = await parsePptx(file);
const doc = createDocument(result); // 创建 DocumentElement
const html = doc.toHTML(); // 调用 toHTML 方法
```

### 2. 统一的 toHTML() API

**旧版本（使用独立的渲染器函数）：**
```typescript
import { slide2HTML, ppt2HTML } from 'pptx-parser';

const slideHTML = slide2HTML(slide, options);
const allHTML = ppt2HTML(result, options);
```

**新版本（使用元素的 toHTML 方法）：**
```typescript
import { createDocument } from 'pptx-parser';

const doc = createDocument(result);

// 转换单个幻灯片
const slideHTML = doc.getSlide(0)?.toHTML();

// 转换整个文档
const html = doc.toHTML();
```

### 3. 文档元素的属性访问

**DocumentElement 提供了丰富的属性和方法：**

```typescript
const doc = createDocument(result);

// 文档信息
console.log(doc.title);      // 标题
console.log(doc.author);     // 作者
console.log(doc.width);      // 宽度
console.log(doc.height);     // 高度
console.log(doc.ratio);      // 宽高比

// 访问内容
const firstSlide = doc.getSlide(0);
const layout = doc.getLayout('layout1');
const master = doc.getMaster('master1');

// 幻灯片数量
console.log(doc.slides.length);

// 布局数量
console.log(Object.keys(doc.layouts).length);

// 母版数量
console.log(doc.masters.length);

// 标签数量
console.log(doc.tags.length);

// 备注数量
console.log(doc.notesSlides.length);
```

### 4. 自定义渲染选项

```typescript
import type { HtmlRenderOptions } from 'pptx-parser';

const options: HtmlRenderOptions = {
  includeStyles: true,          // 包含样式
  includeScripts: true,          // 包含脚本（导航功能）
  includeLayoutElements: true,   // 包含布局和母版元素
  withNavigation: false,         // 不带导航（静态展示）
  customCss: `
    .ppt-slide {
      border-radius: 8px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
  `
};

const html = doc.toHTML(options);
```

### 5. 两种渲染模式

**静态模式（默认）：**
```typescript
const html = doc.toHTML({ withNavigation: false });
// 或者直接调用
const html = doc.toHTMLDocument();
```

**交互模式（带导航）：**
```typescript
const html = doc.toHTML({ withNavigation: true });
// 或者直接调用
const html = doc.toHTMLWithNavigation();
```

### 6. 标签和备注访问

```typescript
const doc = createDocument(result);

// 访问标签
if (doc.tags.length > 0) {
  const firstTag = doc.tags[0];
  const tagValue = firstTag.getTag('tagName');
  const propValue = firstTag.getProperty('propName');
}

// 访问备注
if (doc.notesSlides.length > 0) {
  const firstNote = doc.notesSlides[0];
  console.log(firstNote.text);       // 备注文本
  console.log(firstNote.slideId);     // 关联的幻灯片ID
}
```

## 向后兼容性

为了平滑迁移，旧版本的 API 仍然可用，但已标记为 `@deprecated`：

```typescript
// 这些函数仍然可用，但不推荐使用新项目
import { slide2HTML, ppt2HTML, ppt2HTMLDocument } from 'pptx-parser';

// 内部实现已改为使用 DocumentElement
const html = ppt2HTMLDocument(result);
```

## 元素类层次结构

所有文档元素都继承自 `BaseElement`：

```
BaseElement (基类)
├── DocumentElement (文档元素 - 新增)
├── SlideElement (幻灯片元素)
├── LayoutElement (布局元素)
│   └── PlaceholderElement (占位符元素)
├── MasterElement (母版元素)
├── TagsElement (标签元素)
├── NotesMasterElement (备注母版元素)
├── NotesSlideElement (备注页元素)
├── ShapeElement (形状元素)
├── ImageElement (图片元素)
├── ChartElement (图表元素)
├── TableElement (表格元素)
├── DiagramElement (图解元素)
├── OleElement (OLE对象元素)
└── GroupElement (组元素)
```

每个元素都有 `toHTML()` 方法用于 HTML 渲染。

## Vue 示例更新

**旧版本：**
```vue
<script setup lang="ts">
import { parsePptx, slide2HTML } from 'pptx-parser';

const currentSlideHTML = computed(() => {
  if (!parsedData.value || !currentSlide.value) return '';
  return slide2HTML(currentSlide.value, {
    includeLayoutElements: true
  });
});
</script>
```

**新版本：**
```vue
<script setup lang="ts">
import { parsePptx, createDocument } from 'pptx-parser';

const documentElement = ref(null);

async function handleFileUpload(event) {
  const buffer = await file.arrayBuffer();
  const result = await parsePptx(buffer);
  documentElement.value = createDocument(result);
}

const currentSlideHTML = computed(() => {
  if (!documentElement.value) return '';
  const slide = documentElement.value.getSlide(currentSlideIndex.value);
  return slide ? slide.toHTML() : '';
});
</script>
```

## 迁移检查清单

- [ ] 将 `ppt2HTMLDocument()` 替换为 `createDocument(result).toHTML()`
- [ ] 将 `slide2HTML()` 替换为 `documentElement.getSlide(index)?.toHTML()`
- [ ] 将 `ppt2HTML()` 替换为 `documentElement.slides.map(s => s.toHTML())`
- [ ] 检查自定义渲染选项是否正确传递给 `toHTML(options)`
- [ ] 更新类型导入：`import type { HtmlRenderOptions } from 'pptx-parser'`
- [ ] 验证 Vue 组件中的计算属性是否正确更新

## 常见问题

### Q: 旧版本的代码还能运行吗？
A: 可以，但建议尽快迁移到新 API。旧 API 已标记为 `@deprecated`。

### Q: renderer 目录在哪里？
A: 已移除。所有 HTML 渲染逻辑已整合到各个 Element 类的 `toHTML()` 方法中。

### Q: 如何获取文档属性？
A: 使用 `DocumentElement` 的属性：
```typescript
const doc = createDocument(result);
console.log(doc.title, doc.author, doc.width, doc.height);
```

### Q: 如何访问布局和母版？
A: 使用提供的方法：
```typescript
const layout = doc.getLayout('layoutId');
const master = doc.getMaster('masterId');
```

### Q: 如何自定义样式？
A: 通过 `HtmlRenderOptions.customCss` 传递：
```typescript
const html = doc.toHTML({
  customCss: '.ppt-slide { border-radius: 8px; }'
});
```

## 支持与帮助

如有问题，请查看：
- 示例代码：`examples/document-element-usage.ts`
- Vue 示例：`examples/vue-demo/src/App.vue`
- API 文档：`docs/API.md`
