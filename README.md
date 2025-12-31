# 🔥 PPT-Parser

PPTX 文件解析与序列化核心库，纯 TypeScript 编写，零框架依赖。支持将 `.pptx` 文件解析为结构化 JSON 数据，也可将 JSON 数据逆向序列化为标准可打开的 `.pptx` 文件，开箱即用。

## ✨ 核心特性

- ✅ 纯 TypeScript 开发，严格类型约束，完整的类型声明文件，TS/JS 项目友好
- ✅ 双向解析能力：PPTX 文件 → JSON 结构化数据、JSON 数据 → PPTX 文件
- ✅ 零框架依赖，可无缝集成到 Vue/React/Angular/原生 JS/Node.js 项目
- ✅ 支持解析 PPTX 中的文本、形状、表格、图片、幻灯片基础样式等核心内容
- ✅ 支持 EMU ↔ PX 单位自动转换（PPTX 原生单位为 EMU，自动转为前端常用的 PX）
- ✅ 打包产物双格式：ESM + CommonJS，适配所有模块化规范
- ✅ 生产级别代码，无冗余、无报错、零警告，编译压缩后体积轻量化

## 📦 安装方式

### 方式 1：本地集成（推荐，直接使用打包产物）

将项目 `dist` 目录复制到你的项目中，直接导入使用即可。

### 方式 2：npm 本地安装（推荐，项目内使用）

```bash
# 进入你的项目根目录
npm install ./path-to/ppt-parser --save
```

### 方式 3：开发调试安装

```bash
# 克隆本库后安装依赖
npm install

# 开发热更新（实时监听文件修改，自动编译）
npm run dev

# 生产打包（生成压缩后的 ESM/CJS 产物 + 类型声明）
npm run build

# TS 类型校验（无报错校验）
npm run type-check
```

## 📖 快速上手

### 环境要求

- **Node.js**: >= 16.0.0
- **npm**: >= 8.0.0
- **浏览器**: 支持 ES2020+ 语法的现代浏览器（Chrome/Firefox/Edge/Safari）

### 基础导入

```typescript
// ES Module 导入（推荐，前端项目/ES模块规范）
import PptParserCore from './dist/ppt-parser-core.esm.js';
// 按需解构导入
import { parsePptx, serializePptx, PptParseUtils } from './dist/ppt-parser-core.esm.js';

// CommonJS 导入（Node.js 项目/CommonJS规范）
const PptParserCore = require('./dist/ppt-parser-core.cjs.js');
```
## 🚀 核心 API 使用示例

### ✅ 1. 解析本地 PPTX 文件为 JSON 结构

适用于前端文件上传场景，解析用户上传的 `.pptx` 文件为结构化 JSON 数据，方便前端渲染/处理。

```typescript
import PptParserCore from './dist/ppt-parser-core.esm.js';

// 获取文件上传DOM
const fileInput = document.querySelector('#ppt-upload') as HTMLInputElement;

// 监听文件上传事件
fileInput.addEventListener('change', async (e) => {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (!file || !file.name.endsWith('.pptx')) {
    alert('请选择有效的 .pptx 文件！');
    return;
  }

  try {
    // 核心解析方法：PPTX 文件 → JSON 结构化数据
    const pptJson = await PptParserCore.parse(file);
    console.log('✅ PPTX解析成功，结构化数据：', pptJson);
    // 可在此处处理解析后的JSON数据，如前端渲染幻灯片、提取文本等
  } catch (error) {
    console.error('❌ PPTX解析失败：', error);
  }
});
```

### ✅ 2. 将 JSON 结构序列化为 PPTX 文件并下载

适用于前端根据结构化数据，生成并下载标准的 `.pptx` 文件，生成的文件可直接用 Office/WPS 打开编辑。

```typescript
import PptParserCore from './dist/ppt-parser-core.esm.js';

/**
 * 导出PPTX文件
 * @param pptJson 解析后的PPT结构化JSON数据
 */
async function exportPptxFile(pptJson: PptDocument) {
  if (!pptJson) return;

  try {
    // 核心序列化方法：JSON 数据 → PPTX Blob 文件流
    const pptBlob = await PptParserCore.serialize(pptJson);

    // 生成下载链接并触发下载
    const downloadUrl = URL.createObjectURL(pptBlob);
    const a = document.createElement('a');
    a.href = downloadUrl;
    a.download = `${pptJson.title || '我的PPT'}.pptx`;
    a.click();

    // 释放临时URL资源
    URL.revokeObjectURL(downloadUrl);
    console.log('✅ PPTX导出成功！');
  } catch (error) {
    console.error('❌ PPTX导出失败：', error);
  }
}
```

### ✅ 3. 工具函数使用（单位转换 / 唯一 ID 生成）

内置常用工具函数，满足开发中的基础需求，无需额外封装。

```typescript
import PptParserCore from './dist/ppt-parser-core.esm.js';
const { utils } = PptParserCore;

// 1. PX 转 PPTX 原生单位 EMU
const emu = utils.px2emu(100); // 输入：像素值，输出：EMU值

// 2. EMU 转 前端常用单位 PX
const px = utils.emu2px(914400); // 输入：EMU值，输出：像素值

// 3. 生成唯一ID（用于幻灯片/元素ID标识）
const uniqueId = utils.generateId('slide'); // 可选前缀，默认：ppt-node
```
## 📋 数据结构说明（完整 TS 类型）

所有数据结构均有严格的 TypeScript 类型约束，以下是核心结构的简化说明，完整类型见项目 `src/types.ts`。

### PptDocument（完整 PPT 文档结构）

```typescript
interface PptDocument {
  id: string; // 文档唯一ID
  title: string; // 文档标题
  slides: PptSlide[]; // 幻灯片数组
  props: {
    width: number; // 幻灯片宽度(px)
    height: number; // 幻灯片高度(px)
    ratio: number; // 宽高比
  };
}
```

### PptSlide（单张幻灯片结构）

```typescript
interface PptSlide {
  id: string; // 幻灯片唯一ID
  title: string; // 幻灯片标题
  bgColor: string; // 幻灯片背景色
  elements: PptElement[]; // 幻灯片内元素（文本/形状/表格/图片）
  props: {
    width: number;
    height: number;
    slideLayout: string; // 幻灯片布局类型
  };
}
```

### PptElement（幻灯片元素结构）

```typescript
type PptNodeType = 'text' | 'image' | 'shape' | 'table' | 'chart' | 'container' | 'media';

interface PptElement {
  id: string; // 元素唯一ID
  type: PptNodeType; // 元素类型
  rect: { x: number; y: number; width: number; height: number }; // 元素坐标和尺寸(px)
  style: { // 元素样式
    fontSize?: number;
    color?: string;
    fontWeight?: 'normal' | 'bold';
    textAlign?: 'left' | 'center' | 'right';
    backgroundColor?: string;
    borderColor?: string;
    borderWidth?: number;
  };
  content: string | string[][] | Record<string, any>; // 元素内容，不同类型对应不同格式
  props: Record<string, unknown>; // 扩展属性
}
```
## 🛠 脚本命令说明

项目内置完整的开发/构建/校验脚本，在项目根目录执行对应命令即可：

```bash
# 开发模式：实时监听 src 目录文件修改，自动重新编译，生成未压缩的产物
npm run dev

# 生产构建：清空旧的dist目录 → 编译TS → 生成ESM/CJS双格式产物 → 代码压缩 → 生成类型声明文件
npm run build

# TS类型校验：仅校验TypeScript语法和类型约束，不生成编译产物，快速排查语法错误
npm run type-check

# 发布预检：发布前自动执行 build + type-check，确保产物无问题
npm run prepublishOnly
```

## 📁 项目目录结构

标准的 TypeScript + Rollup 工程化结构，清晰易懂，便于维护和扩展：

```
ppt-parser/
├── src/                # 源码目录（核心代码）
│   ├── index.ts        # 库的统一导出入口
│   ├── core.ts         # 核心解析/序列化算法 + 工具函数
│   └── types.ts        # 完整TS类型定义文件
├── dist/               # 打包产物目录（npm run build 生成）
│   ├── ppt-parser-core.esm.js    # ESM模块（前端项目推荐）
│   ├── ppt-parser-core.cjs.js    # CommonJS模块（Node.js项目推荐）
│   ├── *.js.map        # 源码映射文件（调试用）
│   └── types/          # 自动生成的类型声明文件目录
├── tsconfig.json       # TypeScript编译配置
├── rollup.config.mjs   # Rollup打包配置
├── package.json        # 依赖/脚本/包信息配置
└── README.md           # 项目说明文档（当前文件）
```

## ❗ 注意事项

- 支持解析/序列化 `.pptx` 格式文件，不支持 `.ppt`（97-03 版）格式，如需兼容可先将 ppt 转为 pptx。
- 解析的图片资源目前返回 ID 占位符，如需解析图片二进制内容可基于源码扩展。
- 生成的 PPTX 文件为标准 Office 格式，可直接用 WPS/Microsoft PowerPoint 打开和编辑。
- 浏览器环境下仅支持通过 File 对象解析，Node.js 环境下可传入 Blob/Buffer 解析。

## 🧩 兼容性说明

- **Node.js**: >= 16.0.0（LTS 版本推荐 16.x/18.x）
- **浏览器**: Chrome ≥ 80、Firefox ≥ 75、Edge ≥ 80、Safari ≥ 14
- **模块化**: 支持 ESM / CommonJS 双规范，无模块化兼容问题
- **打包工具**: 兼容 Vite/Rollup/Webpack/Parcel 等主流前端打包工具

## 📄 License

MIT License

---

## ✅ 最后说明

本库为 PPTX 文件的轻量级解析与序列化解决方案，无多余依赖，核心能力聚焦于「结构化解析」和「标准生成」，可满足绝大多数业务场景的 PPT 处理需求。如需扩展更多复杂功能（如动画、公式、批注等），可基于源码轻松二次开发。