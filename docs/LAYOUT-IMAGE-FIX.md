# 布局图片渲染修复

## 问题描述

当幻灯片使用某个布局（如 slide1 使用 layout1）时，转换为 HTML 后，布局中的图片（如 rId4 对应的 `../media/image4.png`）没有正确显示。

## 根本原因

1. **`LayoutElement.toHTML()` 方法只渲染占位符**
   - 原始实现只渲染了 `placeholders`，没有渲染实际元素（`elements`）
   - 布局中的图片、形状等元素没有被包含在 HTML 输出中

2. **`mediaMap` 没有正确传递到布局和母版元素**
   - `parseImages()` 函数已经处理了布局和母版中的图片，并将它们转换为 base64 URL 存储在 `result.mediaMap` 中
   - 但是在创建 `LayoutElement` 和 `MasterElement` 时，没有传递 `mediaMap`
   - 导致图片元素的 `src` 仍然使用相对路径（如 `../media/image4.png`），而不是 base64 URL

## 解决方案

### 1. 更新 `LayoutElement` 类

**添加 `elements` 属性和渲染逻辑：**

```typescript
export class LayoutElement extends BaseElement {
  type: 'layout' = 'layout';
  
  name?: string;
  placeholders: PlaceholderElement[];
  elements: BaseElement[];  // 新增
  textStyles?: any;
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string };
  relsMap: Record<string, any>;
  mediaMap?: Map<string, string>;  // 新增
  
  // ...
  
  toHTML(): string {
    const background = this.getBackgroundStyle();
    const style = [
      this.getContainerStyle(),
      background
    ].join('; ');
    
    const placeholdersHTML = this.placeholders
      .map(ph => ph.toHTML())
      .join('\n');
    
    // 渲染实际元素（图片、形状等）
    const elementsHTML = this.elements
      .map(el => el.toHTML())
      .join('\n');
    
    return `<div class="ppt-layout" style="${style}" data-layout-id="${this.id}" data-layout-name="${this.name || ''}">
${placeholdersHTML}
${elementsHTML}
</div>`;
  }
}
```

**更新 `fromResult()` 方法传递 `mediaMap`：**

```typescript
static fromResult(result: SlideLayoutResult, mediaMap?: Map<string, string>): LayoutElement {
  const placeholders = (result.placeholders || []).map(ph => {
    return new PlaceholderElement(ph.id, ph.type, ph.rect, { idx: ph.idx, name: ph.name, rawNode: ph.rawNode });
  });
  
  // 将解析的元素数据转换为 BaseElement 实例，并传递 mediaMap
  const elements = (result.elements || []).map((el: any) => {
    if (el instanceof BaseElement) {
      return el;
    }
    return createElementFromData(el, result.relsMap || {}, mediaMap);
  }).filter((el: any) => el !== null) as BaseElement[];
  
  return new LayoutElement(
    result.id,
    result.name,
    placeholders,
    elements,
    {
      textStyles: result.textStyles,
      background: result.background,
      relsMap: result.relsMap,
      colorMap: result.colorMap,
      mediaMap  // 传递 mediaMap
    }
  );
}
```

### 2. 更新 `MasterElement` 类

**添加 `mediaMap` 属性并更新 `fromResult()` 方法：**

```typescript
export class MasterElement extends BaseElement {
  type: 'master' = 'master';
  
  masterId?: string;
  elements: BaseElement[];
  placeholders: PlaceholderElement[];
  textStyles?: any;
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string };
  colorMap: Record<string, string>;
  mediaMap?: Map<string, string>;  // 新增
  
  static fromResult(result: MasterSlideResult, mediaMap?: Map<string, string>): MasterElement {
    const placeholders = (result.placeholders || []).map((ph: any) => {
      const phEl = new PlaceholderElement(
        ph.id,
        ph.type || 'other',
        ph.rect || { x: 0, y: 0, width: 100, height: 50 },
        { idx: ph.idx, name: ph.name }
      );
      return phEl;
    });
    
    // 将 result.elements 转换为 BaseElement 实例，并传递 mediaMap
    const elements: BaseElement[] = (result.elements || []).map(elementData => 
      createElementFromData(elementData, result.relsMap, mediaMap)
    ).filter((el): el is BaseElement => el !== null);
    
    return new MasterElement(
      result.id,
      elements,
      placeholders,
      {
        masterId: result.masterId,
        textStyles: result.textStyles,
        background: result.background,
        colorMap: result.colorMap,
        relsMap: result.relsMap,
        mediaMap  // 传递 mediaMap
      }
    );
  }
}
```

### 3. 更新 `DocumentElement.fromParseResult()` 方法

**传递 `mediaMap` 给布局和母版：**

```typescript
// 解析母版
if (result.masterSlides && result.masterSlides.length > 0) {
  doc.masters = result.masterSlides.map(master => MasterElement.fromResult(master, doc.mediaMap));
}

// 解析布局
if (result.slideLayouts) {
  Object.entries(result.slideLayouts).forEach(([layoutId, layout]) => {
    doc.layouts[layoutId] = LayoutElement.fromResult(layout, doc.mediaMap);
  });
}

// 解析备注母版
if (result.notesMasters && result.notesMasters.length > 0) {
  doc.notesMasters = result.notesMasters.map(nm => NotesMasterElement.fromResult(nm, doc.mediaMap));
}
```

## 工作流程

### 图片解析流程

1. **`parseImages()` 函数**（`image-parser.ts`）
   - 遍历所有幻灯片、布局和母版
   - 查找图片元素并获取其 `relId`
   - 使用 `resolveImageRelId()` 从 ZIP 中提取图片文件
   - 将图片转换为 base64 URL
   - 更新元素的 `src` 属性为 base64 URL
   - 将 relId -> base64 URL 的映射保存到 `result.mediaMap`

2. **创建 `DocumentElement`**
   - 从解析结果创建 `DocumentElement`
   - `result.mediaMap` 已经包含了所有图片的 base64 URL

3. **创建 `LayoutElement`**
   - 从 `SlideLayoutResult` 创建 `LayoutElement`
   - 调用 `createElementFromData(el, result.relsMap, mediaMap)`
   - 如果 `mediaMap` 存在且元素的 `relId` 在其中，则使用 base64 URL 作为 `src`

4. **`ImageElement.toHTML()`**
   - 使用 `src` 属性生成 HTML
   - 如果 `src` 是 base64 URL，图片可以直接显示
   - 如果 `src` 是相对路径，浏览器无法显示（这是之前的问题）

## 测试验证

使用以下代码验证修复：

```typescript
import { parsePptx, createDocument } from 'pptx-parser';

const result = await parsePptx(file);
const doc = createDocument(result);

// 检查布局元素
const layout = doc.getLayout('slideLayout1');
if (layout && layout.elements.length > 0) {
  console.log(`布局 ${layout.name} 包含 ${layout.elements.length} 个元素`);
  
  // 查找图片元素
  layout.elements.forEach(el => {
    if (el.type === 'image') {
      console.log(`图片 src: ${el.src}`);
      // 应该是 base64 URL，而不是 ../media/image4.png
    }
  });
}

// 生成 HTML
const html = doc.toHTML();
console.log(html);
```

## 相关文件

- `src/core/image-parser.ts` - 图片解析逻辑
- `src/elements/LayoutElement.ts` - 布局元素类
- `src/elements/MasterElement.ts` - 母版元素类
- `src/elements/DocumentElement.ts` - 文档元素类
- `src/elements/element-factory.ts` - 元素工厂函数

## 注意事项

1. **`mediaMap` 的作用**
   - `mediaMap` 是一个 `Map<string, string>`，键是 relId，值是 base64 URL
   - 它在 `parseImages()` 函数中构建
   - 它存储在 `result.mediaMap` 中，然后在创建元素时传递

2. **`createElementFromData()` 的行为**
   - 函数签名：`createElementFromData(data, relsMap, mediaMap?)`
   - 如果 `mediaMap` 存在且图片元素有 `relId`，则优先使用 `mediaMap.get(relId)` 作为 `src`
   - 否则使用 `data.src`（可能是相对路径）

3. **布局元素的渲染顺序**
   - 在 `SlideElement.toHTML()` 中，渲染顺序是：
     1. 母版元素（`renderMasterElements()`）
     2. 布局元素（`renderLayoutElements()`）
     3. 幻灯片元素（`renderSlideElementsWithLayout()`）

## 总结

修复后，布局中的图片（如 rId4: `../media/image4.png`）将会正确显示为 base64 URL 格式的图片，而不是使用相对路径。这样在浏览器中打开生成的 HTML 时，图片可以正常显示。
