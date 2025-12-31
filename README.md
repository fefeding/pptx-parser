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

## 使用场景

- 📊 在线 PPT 编辑器
- 📑 PPT 文件内容提取
- 🔄 PPT 格式转换
- 📤 PPT 报表导出
- 🎨 PPT 模板生成

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

## 贡献

欢迎提交 Issue 和 Pull Request！

## 许可证

[MIT License](LICENSE)

---

**Made with ❤️ for developers**