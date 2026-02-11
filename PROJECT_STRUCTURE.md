# PPTX Parser 项目结构说明

## 目录结构

```
src/js/
├── core/               # 核心工具
│   ├── constants.js    # 常量定义（SLIDE_FACTOR, FONT_SIZE_FACTOR 等）
│   ├── tXml.js         # XML 解析库
│   └── tinycolor.js    # 颜色处理库
│
├── shape/              # 形状渲染模块
│   ├── shape.js        # 主形状渲染模块（4875 行，208 个形状类型）
│   ├── arrow-shapes.js # 箭头形状（基础和双向箭头）
│   ├── star-shapes.js  # 星形和多边形
│   ├── bracket-shapes.js # 括号形状（大括号、方括号等）
│   ├── pie-shapes.js   # 饼图和弧形
│   ├── math-symbols.js # 数学符号
│   ├── misc-shapes.js  # 杂项形状
│   ├── action-buttons.js # 动作按钮
│   ├── custom-shape.js # 自定义形状
│   ├── path-generators.js # 路径生成器（纯数学函数）
│   └── shape-categories.js # 形状分类常量
│
├── utils/              # 工具函数
│   ├── xml.js          # XML 节点遍历和查询
│   ├── style.js        # 样式处理（填充、边框、阴影等）
│   ├── text.js         # 文本处理（样式、段落、RTL 支持）
│   └── node.js         # 节点处理（幻灯片、图表、SmartArt）
│
└── index.js            # 主入口文件
```

## 模块说明

### core/

#### constants.js
定义项目使用的所有常量，包括：
- `SLIDE_FACTOR`: 幻灯片缩放因子
- `FONT_SIZE_FACTOR`: 字体大小缩放因子
- `RTL_LANGS_ARRAY`: RTL 语言列表
- `DINGBAT_UNICODE`: 装饰字符 Unicode 码点

#### tXml.js
轻量级 XML 解析库，用于解析 PPTX 文件中的 XML 内容。

#### tinycolor.js
颜色处理库，用于颜色的转换和操作。

---

### shape/

#### shape.js（主模块）
形状渲染的核心模块，通过 IIFE 导出 `PPTXShapeUtils` 对象。

**主要功能:**
- `genShape()`: 主入口函数，处理单个形状的完整渲染流程
- 坐标变换和尺寸计算
- 形状类型识别和路由
- 基础几何形状的 SVG 生成（矩形、圆形、三角形等）

**注意:**
- 代码量较大（4875 行），包含约 208 个形状类型
- 使用 ES5 语法以保持兼容性
- 复杂形状已拆分到独立子模块

#### arrow-shapes.js
箭头形状渲染模块，处理各种箭头的 SVG 生成。

**箭头分类:**
- 基础箭头: `rightArrow`, `leftArrow`, `upArrow`, `downArrow`
- 双向箭头: `leftRightArrow`, `upDownArrow`
- 复杂箭头: `quadArrow`, `bentArrow`, `curvedArrow`, `circularArrow` 等
- 标注箭头: `xxxArrowCallout`

**导出函数:**
- `isArrow()`: 判断形状是否为箭头
- `renderArrow()`: 渲染箭头形状
- `renderBasicArrow()`: 渲染基础方向箭头
- `renderDoubleArrow()`: 渲染双向箭头

#### star-shapes.js
星形和多边形形状渲染模块。

#### bracket-shapes.js
括号形状渲染模块，处理大括号、方括号等。

#### pie-shapes.js
饼图和弧形形状渲染模块。

#### math-symbols.js
数学符号渲染模块。

#### misc-shapes.js
杂项形状渲染模块。

#### action-buttons.js
动作按钮形状渲染模块。

#### custom-shape.js
自定义形状渲染模块。

#### path-generators.js
路径生成器模块，纯数学计算函数，无外部依赖，无副作用。

**导出函数:**
- `polarToCartesian()`: 极坐标转笛卡尔坐标
- `shapeArc()`: 生成圆弧路径
- `shapeArcAlt()`: 生成圆弧路径（替代版本）
- `shapeSnipRoundRect()`: 生成切角圆角矩形路径
- `shapeSnipRoundRectAlt()`: 生成切角圆角矩形路径（替代版本）
- `shapePie()`: 生成饼图路径
- `shapeGear()`: 生成齿轮路径

#### shape-categories.js
形状分类常量模块。

**导出的常量:**
- `RECT_SHAPES`: 基础矩形类形状
- `ROUND_RECT_SHAPES`: 圆角矩形类
- `SNIP_RECT_SHAPES`: 切角矩形类
- `FLOWCHART_SHAPES`: 流程图形状
- `ACTION_BUTTONS`: 按钮类
- `BASIC_SHAPES`: 基础几何形状
- `STAR_SHAPES`: 星形
- `ARROW_SHAPES`: 箭头类
- `CALLOUT_SHAPES`: 标注/气泡类
- `BRACKET_SHAPES`: 括号类
- `SPECIAL_SHAPES`: 特殊形状

**导出函数:**
- `getShapeCategory()`: 获取形状所属分类
- `isComplexShape()`: 检查形状是否需要特殊处理

---

### utils/

#### xml.js
XML 工具函数模块，提供 XML 节点遍历和查询功能。

**导出:**
- `PPTXXmlUtils`: IIFE 对象
  - `getTextByPathList()`: 通过路径数组访问嵌套的 XML 节点
  - `getTextByPathStr()`: 通过路径字符串访问嵌套的 XML 节点
  - `readXmlFile()`: 从 ZIP 文件中读取 XML 文件

#### style.js
样式处理模块，处理 PPTX 文件中的各种样式属性。

**处理功能:**
- 填充类型（纯色、渐变、图片、图案等）
- 边框样式
- 阴影效果
- 3D 效果
- 反射效果

**导出:**
- `PPTXStyleUtils`: IIFE 对象

#### text.js
文本处理模块，处理 PPTX 中的文本内容。

**处理功能:**
- 文本样式解析（字体、大小、颜色、对齐等）
- 段落和文本运行处理
- 项目符号和编号
- 超链接处理
- 文本宽度计算
- RTL（从右到左）语言支持

**导出:**
- `PPTXTextUtils`: IIFE 对象

#### node.js
节点工具函数模块，处理 PPTX 节点的各种操作。

**处理功能:**
- 幻灯片节点处理
- 图表生成
- SmartArt 图表处理
- 节点索引和查询

**导出:**
- `PPTXNodeUtils`: IIFE 对象

---

## 命名规范

### 变量命名
项目使用匈牙利命名法：

| 前缀 | 含义 | 示例 |
|------|------|------|
| `shp` | Shape（形状） | `shpId`, `shapType` |
| `img` | Image（图片） | `imgFillFlg` |
| `grnd` | Gradient（渐变） | `grndFillFlg` |
| `adj` | Adjustment（调整参数） | `adj1`, `adj2`, `adj3` |
| `cnst` | Constant（常量） | `cnstVal1`, `cnstVal2` |
| `d` | Dimension（尺寸） | `dVal`, `d_val` |
| `w` | Width（宽度） | `w` |
| `h` | Height（高度） | `h` |
| `vc` | Vertical Center（垂直中心） | `vc` |
| `hc` | Horizontal Center（水平中心） | `hc` |

### 函数命名
- 模块导出函数使用完整名称：`renderArrow`, `isArrow`, `getShapeCategory`
- 内部函数使用驼峰命名：`renderBasicArrow`, `readAdjustmentParams`

### 文件命名
- 模块文件使用 kebab-case：`arrow-shapes.js`, `path-generators.js`
- 工具模块使用单数名词：`xml.js`, `style.js`, `text.js`

---

## 代码风格

### 模块格式
所有模块都使用 ES6 模块语法：
```javascript
/**
 * 模块描述
 * 
 * 详细说明...
 * @module module/name
 */

import { ... } from './path.js';

/**
 * 函数描述
 * @param {type} param - 参数说明
 * @returns {type} 返回值说明
 */
export function functionName() { ... }
```

### IIFE 格式
工具模块使用 IIFE 导出对象：
```javascript
export const PPTXXmlUtils = (function() {
    // 私有函数
    function privateFunc() { ... }
    
    // 公开接口
    return {
        publicFunc: privateFunc
    };
})();
```

### 注释规范
- 模块级注释：描述模块职责、主要功能和注意事项
- 函数级注释：使用 JSDoc 格式，包含参数和返回值说明
- 行内注释：简要说明复杂逻辑

---

## 开发建议

1. **保持模块化**: 将新功能拆分到独立模块，避免 `shape.js` 继续膨胀
2. **遵循命名规范**: 使用项目既定的匈牙利命名法和函数命名规则
3. **添加文档注释**: 所有导出函数都应包含 JSDoc 注释
4. **避免副作用**: 保持 `path-generators.js` 等工具模块的纯函数特性
5. **向后兼容**: 工具模块使用 ES5 语法以保持兼容性

---

## 已知的重构点

1. **shape.js**: 代码量过大（4875 行），建议继续拆分复杂形状到独立模块
2. **复杂箭头**: `arrow-shapes.js` 中 23 个复杂箭头仍在 `shape.js` 中，可以迁移
3. **代码重复**: 部分 shape 模块存在重复的调整参数读取逻辑，可以提取共享函数
