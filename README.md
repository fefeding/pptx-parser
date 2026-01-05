# PPT-Parser

一个轻量级的 PPTX 解析与序列化库，让处理 PowerPoint 文件变得简单。

## 特性

- 📦 **简单易用** - 几行代码即可完成 PPTX 文件的解析和生成
- 🔧 **纯 TypeScript** - 完整的类型定义，优秀的开发体验
- 🎯 **零框架依赖** - 可在任何 JavaScript/TypeScript 项目中使用
- 📱 **双向支持** - 支持 PPTX 文件 → JSON、JSON → PPTX 双向转换
- 🎨 **支持多种元素** - 文本、形状、表格、图片等常见元素
- 🔄 **智能转换** - 自动处理 EMU ↔ PX 单位转换
- 📦 **双格式输出** - 同时支持 ESM 和 CommonJS 模块

## 安装

```bash
npm install @fefeding/ppt-parser
```

或者直接下载 [`dist`](./dist) 目录下的文件使用。

## 快速开始

### 解析 PPTX 文件

```typescript
import PptParserCore from '@fefeding/ppt-parser';

// 上传并解析 PPTX 文件
const fileInput = document.querySelector('#ppt-upload') as HTMLInputElement;

fileInput.addEventListener('change', async (e) => {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (!file) return;

  const pptJson = await PptParserCore.parse(file);
  console.log(pptJson);
});
```

### 增强版解析（推荐使用）

> 💡 **增强版**提供完整的PPTX解析能力，支持标准PPTX文件的所有元素类型，包括OLE对象、分组元素、图片Base64解析等

```typescript
import { parsePptx } from 'pptx-parser';

// 上传并解析 PPTX 文件（增强版）
const fileInput = document.querySelector('#ppt-upload') as HTMLInputElement;

fileInput.addEventListener('change', async (e) => {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (!file) return;

  const result = await parsePptx(file, {
    parseImages: true,    // 解析图片为Base64
    verbose: true          // 详细日志
  });

  console.log('PPT标题:', result.title);
  console.log('作者:', result.author);
  console.log('幻灯片数量:', result.slides.length);

  // 遍历所有元素
  result.slides.forEach((slide, index) => {
    console.log(`幻灯片 ${index + 1}: ${slide.title}`);
    slide.elements.forEach(element => {
      console.log(`  ${element.type}: ${element.text || ''}`);
    });
  });
});
```

📖 **查看增强版文档**：[ENHANCED_README.md](./docs/ENHANCED_README.md) | [ENHANCED_GUIDE.md](./docs/ENHANCED_GUIDE.md)

### 导出 PPTX 文件

```typescript
import PptParserCore from '@fefeding/ppt-parser';

async function exportPptx(pptJson) {
  const pptBlob = await PptParserCore.serialize(pptJson);
  
  const url = URL.createObjectURL(pptBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${pptJson.title || 'presentation'}.pptx`;
  a.click();
  URL.revokeObjectURL(url);
}
```

### 使用工具函数

```typescript
import PptParserCore from '@fefeding/ppt-parser';

const { utils } = PptParserCore;

// 像素转 EMU
const emu = utils.px2emu(100);

// EMU 转像素
const px = utils.emu2px(914400);

// 生成唯一 ID
const id = utils.generateId('slide');
```

## 数据结构

解析后的数据结构如下：

```typescript
// 完整文档
{
  id: string;
  title: string;
  slides: Array<{
    id: string;
    title: string;
    bgColor: string;
    elements: Array<{
      id: string;
      type: 'text' | 'image' | 'shape' | 'table' | 'chart' | 'container' | 'media';
      rect: { x, y, width, height };
      style: { fontSize, color, textAlign, ... };
      content: any;
      props: object;
    }>;
  }>;
  props: { width, height, ratio };
}
```

详细类型定义请查看 [docs/API.md](./docs/API.md)。

## 增强版功能

增强版 `parsePptx` 提供以下额外功能：

- ✅ **完整元素解析** - 支持形状、图片、OLE对象、图表、分组等所有标准元素
- ✅ **命名空间处理** - 遵循 ECMA-376 OpenXML 标准
- ✅ **图片Base64** - 自动解析图片为Base64格式
- ✅ **文本样式** - 解析字体大小、颜色、加粗、斜体等样式
- ✅ **元数据提取** - 提取标题、作者、创建时间等信息
- ✅ **关联关系** - 解析rels文件，正确引用资源
- ✅ **完善容错** - 节点不存在时返回默认值，不抛异常

详细文档请查看：
- [增强版文档](./docs/ENHANCED_README.md) - 完整API和功能说明
- [使用指南](./docs/ENHANCED_GUIDE.md) - 实战示例和最佳实践

## 功能特性

### 支持的元素类型

- 📝 **文本** - 富文本、超链接、项目符号、编号列表
- 🖼️ **图片** - JPG、PNG、SVG 等格式
- 🔷 **形状** - 矩形、圆形、三角形、自定义形状等
- 📊 **表格** - 自定义表格样式
- 📈 **图表** - 柱状图、折线图、饼图等
- 🎬 **媒体** - 视频和音频支持（计划中）
- 🎨 **效果** - 阴影、渐变、3D 效果（计划中）

### 解析选项

```typescript
const pptDoc = await PptParserCore.parse(file, {
  extractImages: true,   // 提取图片二进制数据
  parseMedia: true,      // 解析媒体文件
  parseTheme: false,     // 解析主题
  onProgress: (p, msg) => console.log(`${p}%: ${msg}`)
});
```

### 序列化选项

```typescript
const blob = await PptParserCore.serialize(pptDoc, {
  includeNotes: true,    // 包含演讲者备注
  compress: true,         // 压缩输出
  compressionLevel: 6     // 压缩级别 (0-9)
});
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
const PptParserCore = require('@fefeding/ppt-parser');

// 解析本地文件
const fs = require('fs');
const pptJson = await PptParserCore.parse(fs.readFileSync('presentation.pptx'));
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

---

**Made with ❤️ for developers**