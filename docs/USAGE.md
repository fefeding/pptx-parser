# 使用指南

## 文件说明

构建后会在 `dist` 目录下生成以下文件：

### Node.js / 模块化环境

- **dist/ppt-parser.esm.js** (716KB) - ES Module 格式，未压缩
  - 适用于支持 ES6 modules 的环境
  - 推荐用于 Vue/React/Angular 等现代前端框架
  - **不包含**第三方依赖

- **dist/ppt-parser.cjs.js** (716KB) - CommonJS 格式，未压缩
  - 适用于 Node.js 环境
  - 或需要 CommonJS 格式的打包工具
  - **不包含**第三方依赖

### 浏览器环境

- **dist/ppt-parser.browser.js** (846KB) - ES Module 格式，未压缩
  - 包含所有依赖（jszip, tinycolor2, txml）
  - 使用 ES6 module 导出
  - 适用于开发调试
  - 同时导出到 `window.pptxToHtml` 以兼容旧用法

- **dist/ppt-parser.browser.min.js** (342KB) - ES Module 格式，已压缩
  - 包含所有依赖（jszip, tinycolor2, txml）
  - 使用 ES6 module 导出
  - 适用于生产环境
  - 同时导出到 `window.pptxToHtml` 以兼容旧用法

### 类型定义

- **dist/types/index.d.ts** - TypeScript 类型声明文件

## 使用方式

### 1. Node.js / ES Module

```javascript
// 使用 npm 安装后
import { pptxToHtml } from '@fefeding/ppt-parser';

// 或者直接使用构建文件
import { pptxToHtml } from './dist/ppt-parser.esm.js';

pptxToHtml(document.getElementById('result'), {
  fileInputId: "uploadFileInput",
  slideMode: false
});
```

### 2. CommonJS (Node.js)

```javascript
const { pptxToHtml } = require('@fefeding/ppt-parser');
// 或者
const { pptxToHtml } = require('./dist/ppt-parser.cjs.js');

pptxToHtml(document.getElementById('result'), {
  fileInputId: "uploadFileInput",
  slideMode: false
});
```

### 3. 浏览器 - ES Module (推荐)

```html
<!DOCTYPE html>
<html>
<body>
  <div id="result"></div>
  <script type="module">
    import { pptxToHtml } from './dist/ppt-parser.browser.js';

    pptxToHtml(document.getElementById('result'), {
      fileInputId: "uploadFileInput",
      slideMode: false
    });
  </script>
</body>
</html>
```

### 4. 浏览器 - Script 标签 + 全局变量 (兼容旧浏览器)

```html
<!DOCTYPE html>
<html>
<head>
  <!-- 生产环境使用压缩版本 -->
  <script src="./dist/ppt-parser.browser.min.js"></script>
</head>
<body>
  <div id="result"></div>
  <script>
    // 使用全局变量 window.pptxToHtml
    window.pptxToHtml(document.getElementById('result'), {
      fileInputId: "uploadFileInput",
      slideMode: false
    });
  </script>
</body>
</html>
```

## 依赖说明

### Node.js / ES Module 版本

这些版本**不包含**第三方依赖，需要在项目中单独安装：

```bash
npm install @fefeding/ppt-parser jszip tinycolor2 txml
```

### 浏览器版本

浏览器版本（.browser.js 和 .browser.min.js）**已包含**所有依赖，无需额外安装。

## 浏览器兼容性

- **ES Module 方式**: 需要支持 ES6 modules 的现代浏览器
  - Chrome 61+
  - Firefox 60+
  - Safari 11+
  - Edge 16+

- **全局变量方式**: 支持所有现代浏览器，包括不支持 ES6 modules 的旧浏览器

## 构建命令

```bash
# 构建所有版本
npm run build

# 开发模式（监听文件变化）
npm run dev
```

## 文件大小对比

| 文件 | 大小 | 格式 | 依赖 | 说明 |
|------|------|------|------|------|
| ppt-parser.esm.js | 716KB | ESM | ❌ | Node.js/现代浏览器，未压缩 |
| ppt-parser.cjs.js | 716KB | CJS | ❌ | Node.js，未压缩 |
| ppt-parser.browser.js | 846KB | ESM | ✅ | 浏览器，未压缩，含依赖 |
| ppt-parser.browser.min.js | 342KB | ESM | ✅ | 浏览器，已压缩，含依赖 |

## 推荐使用场景

| 场景 | 推荐版本 | 原因 |
|------|----------|------|
| Vue/React/Angular 项目 | ppt-parser.esm.js | 支持 tree-shaking，体积小 |
| Node.js 后端 | ppt-parser.cjs.js | Node.js 原生支持 |
| 纯 HTML/JS 项目（生产） | ppt-parser.browser.min.js | 单文件，体积小，含依赖 |
| 纯 HTML/JS 项目（开发） | ppt-parser.browser.js | 未压缩，便于调试 |
| 纯 HTML/JS 项目（旧浏览器） | ppt-parser.browser.min.js + 全局变量 | 兼容性最好 |
