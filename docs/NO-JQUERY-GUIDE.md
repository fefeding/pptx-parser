# PPTX Parser - 无 jQuery 版本指南

## 概述

本项目提供两个版本的 PPTX 解析器：

1. **原始版本** (`src/index.html`) - 使用 PPTXjs 原始库，依赖 jQuery
2. **无 jQuery 版本** (`src/index-nojquery.html`) - 使用 TypeScript 重写的 PPTXjs，无需 jQuery

## 无 jQuery 版本

### 特点

- ✅ 完全移除 jQuery 依赖
- ✅ 使用原生 JavaScript (ES6+)
- ✅ 使用 TypeScript 编写的 PPTXjs 替代方案
- ✅ 更轻量，性能更好
- ✅ 完整的类型定义

### 运行方式

#### 开发模式
```bash
pnpm run static:nojquery
```
这将在 `http://localhost:3002` 启动开发服务器。

#### 生产构建
```bash
pnpm run static:nojquery:build
```
这将在 `dist-nojquery/` 目录生成静态文件。

### 功能说明

1. **文件上传**
   - 点击 "选择文件" 按钮
   - 支持 `.pptx` 格式
   - 自动解析并渲染幻灯片

2. **全屏模式**
   - 点击 "Fullscreen" 按钮进入全屏
   - 使用原生 Fullscreen API
   - 支持 Chrome、Firefox、Edge 等现代浏览器

3. **支持的元素**
   - 幻灯片
   - 图片
   - 表格
   - 更多元素正在开发中...

### 代码结构

```javascript
// 导入无 jQuery 版本的 PPTX 解析器
import { parsePptx } from '@fefeding/ppt-parser';

// 解析 PPTX 文件
const result = await parsePptx(arrayBuffer, {
  processFullTheme: true,
  slideMode: false,
  slideType: 'div'
});

// result 包含：
// - slides: 幻灯片数组
// - size: 幻灯片尺寸
// - globalCSS: 全局样式
```

### API 接口

#### PptxjsParserOptions

```typescript
interface PptxjsParserOptions {
  processFullTheme?: boolean;    // 是否完整处理主题
  incSlideWidth?: number;       // 增加幻灯片宽度
  incSlideHeight?: number;      // 增加幻灯片高度
  slideMode?: boolean;          // 是否使用幻灯片模式
  slideType?: 'div' | 'section' | 'divs2slidesjs' | 'revealjs';
  slidesScale?: string;          // 幻灯片缩放百分比
}
```

#### SlideData

```typescript
interface SlideData {
  id: number;                 // 幻灯片 ID
  fileName: string;           // 文件名
  width: number;             // 宽度 (px)
  height: number;            // 高度 (px)
  bgColor?: string;          // 背景颜色
  bgFill?: any;            // 背景填充
  shapes: any[];            // 形状数组
  images: any[];            // 图片数组
  tables: any[];            // 表格数组
  charts: any[];            // 图表数组
}
```

## 原始版本（使用 jQuery）

如果需要使用原始的 PPTXjs（功能更完整），可以运行：

```bash
pnpm run static
```

### 运行方式

```bash
pnpm run static
```
端口：3001

### 注意事项

- 依赖 jQuery 1.11.3
- 依赖 PPTXjs 原始库
- 功能更完整（支持更多元素类型）

## 对比

| 特性 | 无 jQuery 版本 | 原始版本 |
|------|--------------|----------|
| jQuery 依赖 | ❌ 无 | ✅ 需要 |
| 文件大小 | ~40KB | ~800KB |
| 性能 | 更好 | 一般 |
| 功能覆盖 | 基础功能 | 完整功能 |
| TypeScript 支持 | ✅ 完整 | ❌ 无 |
| 开发体验 | 更好 | 一般 |

## 开发路线图

- [x] 基础幻灯片渲染
- [x] 图片支持
- [x] 表格支持
- [ ] 文本和段落渲染
- [ ] 形状渲染
- [ ] 图表渲染
- [ ] 超链接支持
- [ ] 主题和样式继承
- [ ] 完整的 HTML 生成

## 贡献

欢迎贡献代码来完善无 jQuery 版本！请查看 `src/pptxjs/` 目录下的 TypeScript 实现。
