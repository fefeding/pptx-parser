# PPT-Parser

一个轻量级的 PPTX 解析库，让处理 PowerPoint 文件变得简单。

## 特性

- 📦 **简单易用** - 几行代码即可完成 PPTX 文件的解析和生成
- 🔧 **纯 TypeScript** - 完整的类型定义，优秀的开发体验
- 🎯 **零框架依赖** - 可在任何 JavaScript/TypeScript 项目中使用
- 📱 **双向支持** - 支持 PPTX 文件 → HTML/JSON、HTML/JSON → PPTX 双向转换
- 🎨 **支持多种元素** - 文本、形状、表格、图片等常见元素
- 🔄 **智能转换** - 自动处理 EMU ↔ PX 单位转换
- 📦 **双格式输出** - 同时支持 ESM 和 CommonJS 模块
- 🌐 **浏览器/Node.js 双支持** - 可在浏览器环境和 Node.js 中使用

## 安装

```bash
npm install @fefeding/ppt-parser
```

或者直接下载 [`dist`](./dist) 目录下的文件使用。

## 快速开始

### 解析 PPTX 文件为 HTML（推荐）

```javascript
import pptxParser from '@fefeding/ppt-parser';

// 上传并解析 PPTX 文件为 HTML
const fileInput = document.querySelector('#ppt-upload');

fileInput.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  const result = await pptxParser.parseToHtml(file, {
    parseImages: true,    // 解析图片为Base64
    verbose: true         // 详细日志
  });

  console.log('HTML:', result.html);
  console.log('样式:', result.styles);
  
  // 直接获取转换后的HTML内容
  document.getElementById('preview').innerHTML = result.html;
});
```

### 解析 PPTX 文件为 JSON

```javascript
import { pptxToJson } from '@fefeding/ppt-parser';

// 解析 PPTX 文件为 JSON 数据
const fileInput = document.querySelector('#ppt-upload');

fileInput.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  const result = await pptxToJson(file);
  console.log('JSON:', result);
});
```

### 解析 PPTX 文件获取所有文件索引和内容

```javascript
import { pptxToFiles } from '@fefeding/ppt-parser';

// 解析 PPTX 文件获取所有文件的索引和内容
const fileInput = document.querySelector('#ppt-upload');

fileInput.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  const result = await pptxToFiles(file);

  // 查看文件索引
  console.log('文件列表:', result.files);
  // [
  //   { name: 'ppt/slides/slide1.xml', dir: false, size: 12345 },
  //   { name: 'ppt/media/image1.png', dir: false, size: 6789 },
  //   ...
  // ]

  // 获取特定文件内容
  const slide1Content = result.content['ppt/slides/slide1.xml'];
  console.log('Slide1 内容:', slide1Content.content);

  // 获取图片
  const image1 = result.content['ppt/media/image1.png'];
  console.log('图片 Data URL:', image1.dataUrl);
});
```

`pptxToFiles` 返回值结构：
```javascript
{
  files: [
    {
      name: "ppt/slides/slide1.xml",    // 文件路径
      dir: false,                         // 是否为目录
      size: 12345                         // 解压后大小
    }
  ],
  content: {
    "ppt/slides/slide1.xml": {
      type: "text",
      content: "<?xml version=\"1.0\"..."  // XML 文件内容
    },
    "ppt/media/image1.png": {
      type: "image",
      format: "png",
      base64: "iVBORw0KGgoAAAANSUhEUg...",   // Base64 编码
      dataUrl: "data:image/png;base64,iVBORw0KGgo..."  // Data URL
    }
  }
}
```

### 导出 PPTX 文件

> 注意：当前版本导出功能正在完善中，主要支持解析功能

### 使用工具函数

```javascript
import { utils } from '@fefeding/ppt-parser';

// 像素转 EMU
const emu = utils.px2emu(100);

// EMU 转像素
const px = utils.emu2px(914400);

// 生成唯一 ID
const id = utils.generateId('slide');
```

## 输出格式

`parseToHtml` 方法返回以下结构：

```javascript
{
  html: '<div class="pptx-preview">...</div>',  // 转换后的HTML内容
  styles: {                                     // 全局样式表
    global: '._css_1 { ... }',
    table: '._tbl_cell_css_1 { ... }'
  },
  slides: [                                     // 幻灯片数据
    {
      id: 'slide-1',
      elements: [...]
    }
  ]
}
```

## 功能特性

本库提供完整的PPTX解析能力，支持标准PPTX文件的所有元素类型。

### 支持的元素类型

- 📝 **文本** - 富文本、超链接、项目符号、编号列表
- 🖼️ **图片** - JPG、PNG、SVG 等格式
- 🔷 **形状** - 矩形、圆形、三角形、自定义形状等
- 📊 **表格** - 自定义表格样式
- 📈 **图表** - 柱状图、折线图、饼图等
- 🎬 **媒体** - 视频和音频支持（计划中）
- 🎨 **效果** - 阴影、渐变、3D 效果（计划中）

### 解析选项

```javascript
const result = await pptxParser.parseToHtml(file, {
  parseImages: true,    // 解析图片为Base64
  verbose: true,       // 详细日志
  slideHeight: 540,    // 幻灯片高度
  slideWidth: 960      // 幻灯片宽度
});
```

### 浏览器中使用

```html
<script src="./dist/ppt-parser.browser.js"></script>
<script>
  const fileInput = document.querySelector('#ppt-upload');
  
  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    const result = await pptxParser.parseToHtml(file);
    document.getElementById('preview').innerHTML = result.html;
  });
</script>
```

### Node.js 中使用

```javascript
const fs = require('fs');
const { pptxToHtml } = require('@fefeding/ppt-parser');

const buffer = fs.readFileSync('presentation.pptx');
const result = await pptxToHtml(buffer);
console.log(result.html);
```

## 使用场景

- 📊 在线 PPT 编辑器
- 📑 PPT 文件内容提取
- 🔄 PPT 格式转换
- 📤 PPT 报表导出
- 🎨 PPT 模板生成
- 📱 移动端 PPT 查看

## 路线图

查看 [docs/FEATURES.md](./docs/FEATURES.md) 了解功能规划和实现进度。

## 浏览器兼容性

- Chrome ≥ 80
- Firefox ≥ 75
- Edge ≥ 80
- Safari ≥ 14

## Node.js 支持

```javascript
const { pptxToHtml } = require('@fefeding/ppt-parser');
const fs = require('fs');

async function parsePptx() {
  const buffer = fs.readFileSync('presentation.pptx');
  const result = await pptxToHtml(buffer);
  console.log(result.html);
}

parsePptx();
```

## 开发

```bash
# 克隆项目
git clone https://github.com/fefeding/pptx-parser.git

# 安装依赖
npm install

# 开发模式
npm run dev

# 构建
npm run build

# 运行测试
npm test
```

## 文档

- [API 文档](./docs/API.md) - 完整的 API 参考
- [功能规划](./docs/FEATURES.md) - 功能开发和路线图

## 贡献

欢迎提交 Issue 和 Pull Request！

## 许可证

[MIT License](LICENSE)

## 致谢

本库在开发过程中参考和借鉴了 [pptxjs](https://github.com/meshesha/pptxjs) 项目的部分实现思路，特此表示感谢。pptxjs 是一个优秀的客户端PPTX解析库，为本项目的架构设计提供了重要参考。

---

**Made with ❤️ for developers**