# PPTXjs 静态页面说明

这是一个基于 jQuery 的 PPTXjs 演示页面，使用 Vite 作为开发服务器。

## 运行方式

### 开发模式
```bash
npm run static
```
这将在 http://localhost:3001 启动开发服务器，并自动打开浏览器。

### 生产构建
```bash
npm run static:build
```
这将生成静态文件到 `dist-static/` 目录。

## 功能说明

- 上传 PPTX 文件并查看解析结果
- 全屏模式查看演示
- 支持幻灯片模式（配置 `slideMode: true`）
- 支持键盘快捷键（配置 `keyBoardShortCut: true`）

## 文件结构

```
src/
├── index.html          # 主页面
├── css/
│   ├── pptxjs.css     # PPTXjs 样式
│   └── nv.d3.min.css # D3 图表库样式
└── js/
    ├── jquery-1.11.3.min.js    # jQuery 1.11.3
    ├── jszip.min.js            # JSZip 库
    ├── filereader.js          # 文件读取器
    ├── d3.min.js             # D3.js 数据可视化库
    ├── nv.d3.min.js          # NVD3 图表库
    ├── pptxjs.js             # PPTXjs 核心库
    ├── divs2slides.js        # Div 转 Slides 工具
    └── jquery.fullscreen-min.js # 全屏插件
```

## 配置选项

```javascript
$("#result").pptxToHtml({
  pptxFileUrl: "Sample_12.pptx",  // 默认 PPTX 文件路径
  fileInputId: "uploadFileInput",  // 文件输入框 ID
  slideMode: false,                // 是否使用幻灯片模式
  keyBoardShortCut: false,          // 是否启用键盘快捷键
  slideModeConfig: {
    first: 1,                     // 起始幻灯片
    nav: false,                   // 显示导航按钮
    navTxtColor: "white",         // 导航文字颜色
    navNextTxt: ">",              // 下一个按钮文字
    navPrevTxt: "<",              // 上一个按钮文字
    showPlayPauseBtn: false,       // 显示播放/暂停按钮
    keyBoardShortCut: false,       // 键盘快捷键
    showSlideNum: false,          // 显示幻灯片编号
    showTotalSlideNum: false,      // 显示总幻灯片数
    autoSlide: false,             // 自动播放（秒数）
    randomAutoSlide: false,        // 随机自动播放
    loop: false,                  // 循环播放
    background: "black",          // 背景颜色
    transition: "default",         // 过渡效果
    transitionTime: 1             // 过渡时间（秒）
  }
});
```

## 注意事项

- 该页面使用 jQuery 1.11.3，这是 PPTXjs 原始版本的要求
- 所有依赖都是预编译的 min.js 文件，无需额外打包
- 如果需要修改 PPTXjs 源码，请参考 `examples/PPTXjs/js/pptxjs.js`
