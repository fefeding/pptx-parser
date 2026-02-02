# PPTX.js 重构拆分计划

## 目标
将 `src/js/pptxjs.js`（14104行）拆分成模块化结构，提高代码可维护性和可读性。

## 已完成的部分

### 1. 基础结构
- ✅ 创建目录结构：`modules/{utils,core,shapes}`
- ✅ 创建 `constants.js` - 常量定义

### 2. Utils 模块（已全部完成）
- ✅ `file-utils.js` - 文件处理工具（readXmlFile, getContentTypes, getSlideSizeAndSetDefaultTextStyle）
- ✅ `progress-utils.js` - 进度条工具
- ✅ `xml-utils.js` - XML处理工具（getTextByPathList, eachElement, angleToDegrees等）
- ✅ `color-utils.js` - 颜色处理工具（applyShade, applyTint, getSvgGradient等）
- ✅ `text-utils.js` - 文本处理工具（alphaNumeric, romanize, setNumericBullets等）
- ✅ `image-utils.js` - 图片和媒体工具（getMimeType, base64ArrayBuffer等）
- ✅ `chart-utils.js` - 图表处理工具（extractChartData, processSingleMsg等）

### 3. Core 模块
- ✅ `pptx-processor.js` - PPTX主处理逻辑框架
- ✅ `slide-processor.js` - 幻灯片处理框架
- ✅ `node-processors.js` - 节点处理器框架

### 4. Shape 模块（基础框架）
- ✅ `shape-generator.js` - 形状生成器框架（genShape等核心函数签名）

## 待完成的部分

### 1. 剩余 Utils 模块
需要创建以下工具模块，并从原始文件中迁移相应函数：

#### `modules/utils/xml-utils.js`
- `getTextByPathList()` - 通过路径列表获取XML文本
- `indexNodes()` - 索引节点
- `angleToDegrees()` - 角度转换
- `degreesToRadians()` - 弧度转换

#### `modules/utils/color-utils.js`
- `getSolidFill()` - 获取实心填充
- `getGradientFill()` - 获取渐变填充
- `getPatternFill()` - 获取图案填充
- `getBackgroundColor()` - 获取背景颜色

#### `modules/utils/style-utils.js`
- `getPosition()` - 获取位置样式
- `getSize()` - 获取尺寸样式
- `getFontStyle()` - 获取字体样式
- `getTextStyle()` - 获取文本样式

### 2. Shape 模块
#### `modules/shapes/shape-generator.js`
- `genShape()` - 生成形状（主函数）
- `processSpNode()` - 处理形状节点
- `processCxnSpNode()` - 处理连接形状
- `processPicNode()` - 处理图片
- `processGraphicFrameNode()` - 处理图形框架
- `processGroupSpNode()` - 处理组合形状

### 3. 完善 Core 模块
#### `pptx-processor.js`
- 实现 `base64ArrayBuffer()` 函数
- 实现 `genGlobalCSS()` 函数

#### `slide-processor.js`
- 实现 `getBackground()` 函数
- 实现 `getSlideBackgroundFill()` 函数
- 实现 `getTextByPathList()` 函数
- 补全 diagram 处理逻辑

### 4. 主入口文件重构
重构 `src/js/pptxjs.js`：
- 保留 jQuery 插件入口
- 导入并使用各个模块
- 移除所有内部函数定义
- 保持配置和初始化逻辑

### 5. 辅助模块
根据代码分析，可能还需要：
- `modules/utils/text-utils.js` - 文本处理工具
- `modules/utils/chart-utils.js` - 图表处理
- `modules/utils/table-utils.js` - 表格处理
- `modules/utils/media-utils.js` - 媒体处理

## 迁移步骤

### 步骤1：迁移工具函数
1. 从原始文件底部开始，找到所有工具函数
2. 按功能分类，移动到对应的 utils 文件
3. 确保导出和导入正确

### 步骤2：迁移形状处理函数
1. 找到 `genShape()` 函数及其相关函数
2. 迁移到 `shape-generator.js`
3. 处理所有依赖关系

### 步骤3：完善处理器模块
1. 补全 `pptx-processor.js` 中的 TODO
2. 补全 `slide-processor.js` 中的 TODO
3. 补全 `node-processors.js` 中的 TODO

### 步骤4：重构主文件
1. 在 `pptxjs.js` 中导入所有模块
2. 移除所有函数定义
3. 使用模块函数替代
4. 测试确保功能正常

## 注意事项

1. **依赖关系**：注意函数之间的依赖，确保正确的导入/导出顺序
2. **全局变量**：原始文件使用了很多全局变量（如 `slideFactor`, `settings` 等），需要作为参数传递
3. **jQuery 依赖**：保持 jQuery 插件的兼容性
4. **测试**：每次迁移后都要测试，确保功能正常
5. **注释**：保留原始注释，便于理解代码逻辑

## 迁移示例

### 原始代码
```javascript
function updateProgressBar(percent) {
    var progressBarElemtnt = $(".slides-loading-progress-bar");
    progressBarElemtnt.width(percent + "%");
    progressBarElemtnt.html("<span style='text-align: center;'>Loading...(" + percent + "%)</span>");
}
```

### 迁移后
**progress-utils.js:**
```javascript
export function updateProgressBar(percent) {
    var progressBarElemtnt = $(".slides-loading-progress-bar");
    progressBarElemtnt.width(percent + "%");
    progressBarElemtnt.html("<span style='text-align: center;'>Loading...(" + percent + "%)</span>");
}
```

**pptxjs.js:**
```javascript
import { updateProgressBar } from './modules/utils/progress-utils.js';

// 直接使用
updateProgressBar(percent);
```

## 预计工作量

- 工具函数：约 30 个函数，需要 2-3 小时
- 形状处理：约 20 个函数，需要 3-4 小时
- 完善核心：约 10 个函数，需要 1-2 小时
- 重构主文件：需要 1-2 小时
- 测试调试：需要 2-3 小时

**总计：约 10-15 小时完成完整重构**

## 优势

1. **可维护性**：模块化结构，易于理解和维护
2. **可测试性**：可以单独测试各个模块
3. **可重用性**：工具模块可以在其他项目中使用
4. **可读性**：文件更小，功能更集中
5. **协作性**：多人可以同时开发不同模块
