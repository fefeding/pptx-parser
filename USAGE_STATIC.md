# 运行 PPTXjs 静态页面

## 快速开始

1. **安装依赖**（如果还没有安装）：
```bash
pnpm install
```

2. **启动开发服务器**：
```bash
pnpm run static
```

这将：
- 在 `http://localhost:3001` 启动 Vite 开发服务器
- 自动打开浏览器
- 提供热更新支持

3. **使用页面**：
   - 点击 "选择文件" 按钮上传 PPTX 文件
   - 或者直接拖拽 PPTX 文件到页面上
   - 查看解析后的幻灯片
   - 点击 "Fullscreen" 按钮进入全屏模式

## 构建静态文件

如果需要构建生产版本：

```bash
pnpm run static:build
```

静态文件将输出到 `dist-static/` 目录。

## 文件说明

- `vite.static.config.ts` - Vite 配置文件
- `src/index.html` - 主页面
- `src/css/` - 样式文件
- `src/js/` - JavaScript 库文件
- `src/README_STATIC.md` - 详细的配置说明

## 注意事项

1. **PPTX 文件上传**：
   - 页面默认需要手动上传 PPTX 文件
   - 支持 `.pptx` 格式
   - 文件大小建议不超过 100MB

2. **浏览器兼容性**：
   - 推荐使用 Chrome、Firefox、Edge 等现代浏览器
   - 需要支持 ES5 和 HTML5

3. **依赖说明**：
   - 使用 jQuery 1.11.3（PPTXjs 原始版本要求）
   - JSZip 用于解析 PPTX 文件
   - D3.js 和 NVD3 用于图表渲染
