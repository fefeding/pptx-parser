# 主题解析模块使用指南

## 概述

主题解析模块（ThemeElement）用于解析和管理 PPTX 文件的主题信息，包括颜色方案、字体方案和效果方案。主题样式会自动生成带有主题名称前缀的 CSS 类，方便在 HTML 输出中应用一致的主题样式。

## 功能特性

1. **自动解析主题信息**
   - 从 PPTX 文件的 theme*.xml 中提取主题名称
   - 解析完整的颜色方案（12种预设颜色）
   - 支持字体方案和效果方案（可选）

2. **生成带前缀的 CSS 类**
   - 使用主题名称作为 CSS 类前缀（如 `theme-office`）
   - 生成 CSS 变量（CSS Custom Properties）
   - 提供颜色、背景、文字、边框工具类
   - 自动生成预设样式类（标题、正文、链接等）

3. **自动集成到 HTML 输出**
   - 在 `DocumentElement.toHTML()` 中自动生成 `<style>` 元素
   - 所有生成的 HTML 元素可以方便地应用主题样式类

## API 文档

### ThemeElement 类

#### 构造函数

```typescript
constructor(
  id: string,
  name: string,
  colors: ThemeColors,
  themeId?: string,
  fonts?: FontScheme,
  effects?: EffectScheme
)
```

#### 主要方法

##### `getThemeClassPrefix(): string`
获取主题的 CSS 类前缀。

```typescript
const prefix = theme.getThemeClassPrefix();
// 例如: "theme-office"
```

##### `getThemeClass(suffix: string): string`
生成完整的主题 CSS 类名。

```typescript
const className = theme.getThemeClass('accent1');
// 例如: "theme-office-accent1"
```

##### `generateThemeCSS(): string`
生成完整的主题 CSS 样式。

```typescript
const css = theme.generateThemeCSS();
// 返回包含所有主题样式的 CSS 字符串
```

##### `getColor(colorKey: keyof ThemeColors): string | undefined`
获取指定主题颜色值。

```typescript
const accent1 = theme.getColor('accent1');
// 例如: "#4472C4"
```

##### `getMajorFont(): string | undefined`
获取主要字体（标题等）。

##### `getMinorFont(): string | undefined`
获取次要字体（正文等）。

### 静态方法

##### `fromResult(result: ThemeResult, themeName: string = 'theme1'): ThemeElement`
从解析结果创建 ThemeElement。

```typescript
const themeElement = ThemeElement.fromResult(result.theme, result.theme.name);
```

## 主题颜色方案

PPTX 主题包含以下12种预设颜色：

| 颜色名称 | 说明 |
|---------|------|
| `bg1` | 背景色1（通常为白色） |
| `tx1` | 文本色1（通常为黑色） |
| `bg2` | 背景色2（浅灰色） |
| `tx2` | 文本色2（深灰色） |
| `accent1` | 强调色1（蓝色） |
| `accent2` | 强调色2（橙色） |
| `accent3` | 强调色3（灰色） |
| `accent4` | 强调色4（黄色） |
| `accent5` | 强调色5（浅蓝色） |
| `accent6` | 强调色6（绿色） |
| `hlink` | 超链接颜色 |
| `folHlink` | 已访问链接颜色 |

## 生成的 CSS 类

### CSS 变量

主题会生成 CSS 变量，格式为 `--{prefix}-{color-name}`：

```css
.theme-office {
  --theme-office-bg1: #FFFFFF;
  --theme-office-tx1: #000000;
  --theme-office-accent1: #4472C4;
  /* ... 更多变量 */
}
```

### 颜色工具类

每种颜色都会生成三种工具类：

#### 背景色类
```css
.theme-office-bg-accent1 {
  background-color: var(--theme-office-accent1);
}
```

#### 文本色类
```css
.theme-office-text-accent1 {
  color: var(--theme-office-accent1);
}
```

#### 边框色类
```css
.theme-office-border-accent1 {
  border-color: var(--theme-office-accent1);
}
```

### 预设样式类

#### 标题样式
```css
.theme-office-title {
  color: var(--theme-office-tx1);
  font-weight: bold;
}
```

#### 正文样式
```css
.theme-office-body {
  color: var(--theme-office-tx2);
}
```

#### 链接样式
```css
.theme-office-link {
  color: var(--theme-office-hlink);
  text-decoration: none;
}
```

#### 强调色样式
```css
.theme-office-accent-1 { color: var(--theme-office-accent1); }
.theme-office-accent-2 { color: var(--theme-office-accent2); }
/* ... */
```

## 使用示例

### 基本使用

```typescript
import { parsePptx, DocumentElement } from 'pptx-parser';

async function main() {
  // 解析 PPTX 文件
  const result = await parsePptx(file);

  // 创建文档元素
  const doc = DocumentElement.fromParseResult(result);

  // 访问主题
  const theme = doc.theme;
  if (theme) {
    console.log('主题名称:', theme.name);
    console.log('主题前缀:', theme.getThemeClassPrefix());

    // 获取主题颜色
    const accent1 = theme.getColor('accent1');
    console.log('Accent 1 颜色:', accent1);
  }

  // 生成 HTML（自动包含主题样式）
  const html = doc.toHTML({ includeStyles: true });

  // 保存 HTML
  fs.writeFileSync('output.html', html);
}
```

### 在 HTML 中使用主题类

```html
<!DOCTYPE html>
<html>
<head>
  <title>演示文稿</title>
  <!-- 主题样式会自动包含在生成的 HTML 中 -->
</head>
<body>
  <div class="theme-office">
    <!-- 标题使用主题文字色 -->
    <h1 class="theme-office-title">演示文稿标题</h1>

    <!-- 正文使用次要文字色 -->
    <p class="theme-office-body">这是正文内容</p>

    <!-- 使用强调色 -->
    <p class="theme-office-text-accent1">这是强调文本</p>

    <!-- 使用主题背景色 -->
    <div class="theme-office-bg-accent2">
      <p>带有主题背景的区域</p>
    </div>

    <!-- 使用 CSS 变量 -->
    <p style="color: var(--theme-office-accent3);">
      使用 CSS 变量的文本
    </p>
  </div>
</body>
</html>
```

### 自定义主题样式

如果需要自定义主题样式，可以通过 HTML 渲染选项添加：

```typescript
const html = doc.toHTML({
  includeStyles: true,
  customCss: `
    .custom-highlight {
      background: var(--theme-office-accent1);
      color: white;
      padding: 10px;
      border-radius: 4px;
    }
  `
});
```

### 动态生成主题类

```typescript
// 为元素动态添加主题类
function addThemeClass(element: HTMLElement, theme: ThemeElement, suffix: string) {
  element.classList.add(theme.getThemeClass(suffix));
}

// 示例：为所有标题添加主题标题类
document.querySelectorAll('h1, h2, h3').forEach(el => {
  if (doc.theme) {
    el.classList.add(doc.theme.getThemeClass('title'));
  }
});
```

## 多主题支持

如果 PPTX 文件包含多个主题（虽然 PPTX 标准通常只使用一个主题），可以通过主题 ID 区分：

```typescript
// 假设有多个主题
if (result.theme1) {
  const theme1 = ThemeElement.fromResult(result.theme1, 'theme1');
  const prefix1 = theme1.getThemeClassPrefix();
  // 生成 theme1-* 前缀的类
}

if (result.theme2) {
  const theme2 = ThemeElement.fromResult(result.theme2, 'theme2');
  const prefix2 = theme2.getThemeClassPrefix();
  // 生成 theme2-* 前缀的类
}
```

## 注意事项

1. **主题命名**：主题名称会自动转换为 CSS 友好的格式（小写、连字符分隔）
2. **CSS 变量**：使用 CSS 自定义属性可以轻松实现主题切换
3. **浏览器兼容性**：CSS 变量需要现代浏览器支持（IE11 不支持）
4. **字体方案**：当前版本支持解析字体方案，但 CSS 生成功能相对简单
5. **效果方案**：效果方案的解析和 CSS 生成功能可以进一步扩展

## 扩展开发

### 添加新的颜色解析

如果需要支持更多颜色类型，可以在 `theme-parser.ts` 中扩展：

```typescript
function parseColorValue(colorEl: Element): string {
  // 添加新的颜色类型支持
  const schemeClr = getFirstChildByTagNS(colorEl, 'schemeClr', NS.a);
  if (schemeClr) {
    const val = schemeClr.getAttribute('val');
    // 处理方案颜色引用
  }
  // ... 现有代码
}
```

### 添加字体方案 CSS 生成

可以在 `ThemeElement.generateThemeCSS()` 中扩展字体相关样式：

```typescript
if (this.fonts) {
  // 添加更多字体相关样式
  css.push(`.${prefix}-font-heading {`);
  css.push(`  font-family: "${this.fonts.majorFont.latin}", sans-serif;`);
  css.push('}');
}
```

## 相关文件

- `src/elements/ThemeElement.ts` - 主题元素类
- `src/core/theme-parser.ts` - 主题解析器
- `src/elements/DocumentElement.ts` - 文档元素（集成主题输出）
- `src/core/types.ts` - 主题类型定义
- `examples/theme-usage-example.html` - 主题样式使用示例

## 参考资料

- [Office Open XML - Themes](https://docs.microsoft.com/en-us/openspecs/office_standards/)
- [PPTXjs Documentation](https://github.com/meshesha/PPTXjs)
