# PPTX Parser Vue Demo

这是一个使用 Vue 3 + Vite 构建的 PPTX 解析器演示应用。

## 功能特性

- 📤 上传 PPTX 文件进行解析
- 🎨 可视化展示幻灯片内容
- ⌨️ 支持键盘方向键切换幻灯片
- 📱 响应式设计
- 🔍 查看原始 JSON 数据
- 🔥 **热加载支持**：修改上层库源码后自动重新加载

## 安装依赖

```bash
pnpm install
```

## 启动开发服务器

```bash
pnpm dev
```

应用将在 http://localhost:3000 启动。

## 热加载说明

本项目配置了 Vite 直接引用上层库的源码（`../../src`），当你修改 `pptx-parser` 的源码时，Vue Demo 会自动热更新，无需重新构建库文件。

如果热加载没有生效，可以：
1. 刷新浏览器页面
2. 重启 Vite 开发服务器（`Ctrl+C` 然后再次运行 `pnpm dev`）

## 构建生产版本

```bash
pnpm build
```

## 预览生产构建

```bash
pnpm preview
```

## 使用说明

1. 点击上传区域选择 PPTX 文件
2. 等待解析完成
3. 使用上一页/下一页按钮或点击缩略图切换幻灯片
4. 查看可视化展示效果
5. 展开"查看原始数据"可查看解析后的 JSON 数据

## 支持的元素类型

- 文本（Text）
- 图片（Image）
- 形状（Shape）
- 表格（Table）
- 图表（Chart，显示为占位符）
- 其他元素类型会显示为"暂不支持"

## 技术栈

- Vue 3
- TypeScript
- Vite
- pptx-parser
