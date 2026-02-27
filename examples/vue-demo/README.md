# PPTX Parser Vue Demo

这是一个使用 Vue 3 + Vite 构建的 PPTX 解析器演示应用。

## 功能特性

- 📤 上传 PPTX 文件进行解析
- 🎨 可视化展示幻灯片内容
- 📊 解析进度显示
- 🖼️ 全屏查看模式
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

## 使用说明

1. 点击上传区域选择 PPTX 文件
2. 等待解析完成（显示进度百分比）
3. 所有幻灯片会以垂直排列方式展示
4. 点击"全屏"按钮可进入全屏模式

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
