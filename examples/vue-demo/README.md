# PPTX Parser Vue Demo

这是一个使用 Vue 3 + Vite 构建的 PPTX 解析器演示应用。

## 功能特性

- 📤 上传 PPTX 文件进行解析（支持点击与拖拽）
- 🎨 可视化展示幻灯片渲染结果
- 📊 解析进度显示（带进度条）
- 🖼️ 全屏查看模式
- 🧩 **JSON 数据查看**：以语法高亮展示解析出的结构化 JSON，支持按页切换、精简大字段（缩略图 / 媒体 base64）、一键复制
- 🔄 **格式转换**：将当前 PPTX 导出为 PDF / Word / Markdown / HTML；中文 PDF 自动加载所需字形的网络字体，也可手动上传 TTF/OTF 字体
- 🔥 **热加载支持**：修改上层库源码后自动重新加载

## 安装依赖

```bash
pnpm install
```

## 启动开发服务器

```bash
pnpm dev
```

应用将在 http://localhost:5173 启动（默认 Vite 端口）。

## 热加载说明

本项目通过本地包引用（`file:../../`）使用 pptx-parser 库。

如果你修改了 pptx-parser 的源码，需要：
1. 在根目录重新构建库：`pnpm build`
2. 或者在 vue-demo 目录运行：`pnpm install --force`

格式转换功能使用根目录构建产物 `dist/convert.browser.js`，首次运行前也需先执行根目录 `pnpm build`。

## API 说明

本示例使用最新版 `pptxToHtml` API：

```typescript
const result = await pptxToHtml(fileData, {
  mediaProcess: true,      // 处理媒体文件
  themeProcess: true,      // 处理主题样式
  callbacks: {
    onProgress: (percent: number) => {
      // 解析进度回调
    }
  }
})

// result 包含：
// - slides: 幻灯片数组 { html, slideNum, fileName }
// - slideSize: 幻灯片尺寸 { width, height }
// - styles: 全局样式 { global: string }
// - metadata: 文件元数据
// - charts: 图表数据
```

格式转换使用 `@fefeding/ppt-parser/convert` 暴露的 `convert`：

```typescript
import { convert } from '../../../dist/convert.browser.js'
import { collectCjkChars, loadCjkFont } from '../../cjk-font.js'

// 浏览器无法探测系统字体，示例用 Google Fonts 按实际汉字下载子集字体
const font = await loadCjkFont(collectCjkChars(JSON.stringify(jsonResult)))
const bytes = await convert(fileData, 'pdf', {
  title: '演示文稿',
  fonts: font ? { regular: font } : undefined
}) // 目标还可为 docx / md / html
```

## 使用说明

1. 点击上传区域选择 PPTX 文件
2. 等待解析完成（显示进度百分比）
3. 所有幻灯片会以垂直排列方式展示
4. 点击"全屏"按钮可进入全屏模式
5. 在结果区的"格式转换"栏选择目标格式下载；中文 PDF 会自动加载网络字体，离线时可手动上传 TTF/OTF 字体

## 构建生产版本

```bash
pnpm build
```

## 预览生产构建

```bash
pnpm preview
```

## 技术栈

- Vue 3 (Composition API + TypeScript)
- Vite
- pptx-parser
