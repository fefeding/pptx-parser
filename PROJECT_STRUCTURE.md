# PPTX Parser 项目结构说明

## 目录结构

```
src/
├── core/               # 核心工具
│   ├── constants.ts    # 常量定义（SLIDE_FACTOR, FONT_SIZE_FACTOR 等）
│   ├── tXml.ts         # XML 解析库
│   └── tinycolor.ts    # 颜色处理库
│
├── shape/              # 形状渲染模块
│   ├── shape.ts        # 主形状渲染模块（4875 行，208 个形状类型）
│   ├── arrow-shapes.ts # 箭头形状（基础和双向箭头）
│   ├── star-shapes.ts  # 星形和多边形
│   ├── bracket-shapes.ts # 括号形状（大括号、方括号等）
│   ├── pie-shapes.ts   # 饼图和弧形
│   ├── math-symbols.ts # 数学符号
│   ├── misc-shapes.ts  # 杂项形状
│   ├── action-buttons.ts # 动作按钮
│   ├── custom-shape.ts # 自定义形状
│   ├── path-generators.ts # 路径生成器（纯数学函数）
│   └── shape-categories.ts # 形状分类常量
│
├── serializer/         # JSON→PPTX 序列化模块
│   ├── xml-builder.ts  # XML 构建工具（转义、单位换算、节点生成）
│   ├── templates.ts    # OOXML 静态模板（主题/母版/版式/Content-Types 等）
│   ├── element-builders.ts # 元素构建器（文本/形状/图片 → OOXML 节点）
│   ├── composer.ts     # 流式构建 API（PPTXComposer）
│   └── json-to-pptx.ts # 序列化主入口（jsonToPptx / editPptx 编辑器）
│
├── utils/              # 工具函数
│   ├── xml.ts          # XML 节点遍历和查询
│   ├── style.ts        # 样式处理（填充、边框、阴影等）
│   ├── text.ts         # 文本处理（样式、段落、RTL 支持）
│   └── node.ts         # 节点处理（幻灯片、图表、SmartArt）
│
└── index.ts            # 主入口文件
```

## 模块说明

### core/

#### constants.ts
定义项目使用的所有常量，包括：
- `SLIDE_FACTOR`: 幻灯片缩放因子
- `FONT_SIZE_FACTOR`: 字体大小缩放因子
- `RTL_LANGS_ARRAY`: RTL 语言列表
- `DINGBAT_UNICODE`: 装饰字符 Unicode 码点

#### tXml.ts
轻量级 XML 解析库，用于解析 PPTX 文件中的 XML 内容。

#### tinycolor.ts
颜色处理库，用于颜色的转换和操作。

---

### shape/

#### shape.ts（主模块）
形状渲染的核心模块，以具名导出 `PPTXShapeUtils` 对象。

**主要功能:**
- `genShape()`: 主入口函数，处理单个形状的完整渲染流程
- 坐标变换和尺寸计算
- 形状类型识别和路由
- 基础几何形状的 SVG 生成（矩形、圆形、三角形等）

**注意:**
- 代码量较大（4875 行），包含约 208 个形状类型
- 源码为 TypeScript（编译目标 ES2020），经 Rollup 打包以保持浏览器/Node 兼容
- 复杂形状已拆分到独立子模块

#### arrow-shapes.ts
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

#### star-shapes.ts
星形和多边形形状渲染模块。

#### bracket-shapes.ts
括号形状渲染模块，处理大括号、方括号等。

#### pie-shapes.ts
饼图和弧形形状渲染模块。

#### math-symbols.ts
数学符号渲染模块。

#### misc-shapes.ts
杂项形状渲染模块。

#### action-buttons.ts
动作按钮形状渲染模块。

#### custom-shape.ts
自定义形状渲染模块。

#### path-generators.ts
路径生成器模块，纯数学计算函数，无外部依赖，无副作用。

**导出函数:**
- `polarToCartesian()`: 极坐标转笛卡尔坐标
- `shapeArc()`: 生成圆弧路径
- `shapeArcAlt()`: 生成圆弧路径（替代版本）
- `shapeSnipRoundRect()`: 生成切角圆角矩形路径
- `shapeSnipRoundRectAlt()`: 生成切角圆角矩形路径（替代版本）
- `shapePie()`: 生成饼图路径
- `shapeGear()`: 生成齿轮路径

#### shape-categories.ts
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

#### xml.ts
XML 工具函数模块，提供 XML 节点遍历和查询功能。

**导出:**
- `PPTXXmlUtils`: 具名导出对象
  - `getTextByPathList()`: 通过路径数组访问嵌套的 XML 节点
  - `getTextByPathStr()`: 通过路径字符串访问嵌套的 XML 节点
  - `readXmlFile()`: 从 ZIP 文件中读取 XML 文件

#### style.ts
样式处理模块，处理 PPTX 文件中的各种样式属性。

**处理功能:**
- 填充类型（纯色、渐变、图片、图案等）
- 边框样式
- 阴影效果
- 3D 效果
- 反射效果

**导出:**
- `PPTXStyleUtils`: 具名导出对象

#### text.ts
文本处理模块，处理 PPTX 中的文本内容。

**处理功能:**
- 文本样式解析（字体、大小、颜色、对齐等）
- 段落和文本运行处理
- 项目符号和编号
- 超链接处理
- 文本宽度计算
- RTL（从右到左）语言支持

**导出:**
- `PPTXTextUtils`: 具名导出对象

#### node.ts
节点工具函数模块，处理 PPTX 节点的各种操作。

**处理功能:**
- 幻灯片节点处理
- 图表生成
- SmartArt 图表处理
- 节点索引和查询

**导出:**
- `PPTXNodeUtils`: 具名导出对象

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
- 模块文件使用 kebab-case：`arrow-shapes.ts`, `path-generators.ts`
- 工具模块使用单数名词：`xml.ts`, `style.ts`, `text.ts`

---

## 代码风格

### 模块格式
所有模块都使用 ES 模块语法（TypeScript）：
```typescript
/**
 * 模块描述
 * 
 * 详细说明...
 * @module module/name
 */

import { ... } from './path.ts';

/**
 * 函数描述
 * @param {type} param - 参数说明
 * @returns {type} 返回值说明
 */
export function functionName() { ... }
```

### 模块导出格式
工具模块使用具名导出（named export）：
```typescript
// 私有函数
function privateFunc() { ... }

// 具名导出
export const PPTXXmlUtils = {
    publicFunc: privateFunc
};
```

### 注释规范
- 模块级注释：描述模块职责、主要功能和注意事项
- 函数级注释：使用 JSDoc 格式，包含参数和返回值说明
- 行内注释：简要说明复杂逻辑

---

## 开发建议

1. **保持模块化**: 将新功能拆分到独立模块，避免 `shape.ts` 继续膨胀
2. **遵循命名规范**: 使用项目既定的匈牙利命名法和函数命名规则
3. **添加文档注释**: 所有导出函数都应包含 JSDoc 注释
4. **避免副作用**: 保持 `path-generators.ts` 等工具模块的纯函数特性
5. **向后兼容**: 源码为 TypeScript（目标 ES2020），经 Rollup 打包输出适配浏览器与 Node 环境

---

## 已知的重构点

1. **shape.ts**: 代码量过大（4875 行），建议继续拆分复杂形状到独立模块
2. **复杂箭头**: `arrow-shapes.ts` 中 23 个复杂箭头仍在 `shape.ts` 中，可以迁移
3. **代码重复**: 部分 shape 模块存在重复的调整参数读取逻辑，可以提取共享函数
