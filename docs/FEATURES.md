# 功能扩展路线图

## 基于 PPTXjs 的功能对标

本文档详细说明了参考 [PPTXjs](https://github.com/meshesha/PPTXjs) 项目的功能设计和标准支持。

---

## ✅ 已支持的功能

### 基础功能
- ✅ PPTX 文件解析（XML 结构解析）
- ✅ PPTX 文件序列化（JSON → PPTX）
- ✅ 文本元素解析
- ✅ 形状元素解析
- ✅ 图片元素解析（ID 占位符）
- ✅ 表格元素解析
- ✅ 基础样式解析（字体、颜色、对齐等）
- ✅ EMU ↔ PX 单位转换
- ✅ 幻灯片背景解析
- ✅ 幻灯片尺寸解析

---

## 🎯 阶段一：核心元素扩展（高优先级）

### 1.1 文本功能增强

#### 项目符号和编号列表
```typescript
interface PptTextParagraph {
  bullet?: {
    type?: 'none' | 'bullet' | 'numbered';
    char?: string;
    level?: number;
  };
}
```

**实现要点：**
- 解析 `<a:buChar>` 标签获取项目符号字符
- 解析 `<a:buAutoNum>` 获取编号列表
- 支持多级列表（1-9 级）
- 支持自定义项目符号字符

#### 超链接支持
```typescript
interface PptTextParagraph {
  hyperlink?: {
    url: string;
    tooltip?: string;
  };
}
```

**实现要点：**
- 解析 `<a:hlinkClick>` 标签
- 通过 `r:id` 映射获取实际 URL
- 支持内部链接（跳转到其他幻灯片）

#### 富文本样式
```typescript
interface PptTextStyle {
  textDecoration?: 'none' | 'underline' | 'line-through';
  textVerticalAlign?: 'top' | 'middle' | 'bottom';
  lineHeight?: number;
  letterSpacing?: number;
  textShadow?: string;
}
```

**实现要点：**
- 解析 `<a:u>` (underline)、`<a:strike>` (line-through)
- 解析 `<a:baseline>` (vertical align)
- 解析 `<a:spc>` (letter spacing)
- 解析 `<a:effectLst>` (text shadow)

### 1.2 形状功能增强

#### 更多形状类型
```typescript
type PptShapeType =
  | 'rectangle'
  | 'roundRectangle'
  | 'ellipse'
  | 'circle'
  | 'triangle'
  | 'diamond'
  | 'star'
  | 'arrow'
  | 'line'
  | 'curve'
  | 'polygon'
  | 'custom';
```

**实现要点：**
- 解析 `<a:prstGeom>` 获取预设形状
- 支持约 180+ 种 Office 预设形状
- 支持自定义 SVG 路径形状

#### 变换效果
```typescript
interface PptTransform {
  rotate?: number;
  flipH?: boolean;
  flipV?: boolean;
}
```

**实现要点：**
- 解析 `<a:xfrm>` 标签
- 计算旋转变换矩阵
- 处理翻转效果

### 1.3 填充效果扩展

#### 渐变填充
```typescript
interface PptFill {
  type?: 'solid' | 'gradient' | 'pattern' | 'picture' | 'none';
  gradientStops?: Array<{ position: number; color: string }>;
  gradientDirection?: number;
}
```

**实现要点：**
- 解析 `<a:gradFill>` 标签
- 支持线性和径向渐变
- 支持多色渐变停止点

#### 图片填充
```typescript
interface PptFill {
  image?: string;
  opacity?: number;
}
```

**实现要点：**
- 解析 `<a:blipFill>` 标签
- 通过 `r:embed` 获取图片资源
- 支持填充模式和透明度

### 1.4 边框效果扩展

```typescript
interface PptBorder {
  color?: string;
  width?: number;
  style?: 'solid' | 'dashed' | 'dotted' | 'double';
  dashStyle?: string;
}
```

**实现要点：**
- 解析 `<a:ln>` 标签
- 支持多种边框样式
- 支持自定义虚线模式

---

## 🚀 阶段二：高级功能（中优先级）

### 2.1 媒体支持

#### 视频支持
```typescript
type PptNodeType = 'video';

interface PptVideoContent {
  src: string;
  poster?: string;
  autoplay?: boolean;
  loop?: boolean;
  muted?: boolean;
  controls?: boolean;
}
```

**实现要点：**
- 解析 `<p:videoFile>` 标签
- 支持嵌入视频和外部视频链接
- 提取视频缩略图
- 生成 HTML5 `<video>` 元素

#### 音频支持
```typescript
type PptNodeType = 'audio';

interface PptAudioContent {
  src: string;
  autoplay?: boolean;
  loop?: boolean;
  volume?: number;
}
```

**实现要点：**
- 解析 `<p:audioFile>` 标签
- 支持自动播放和循环
- 生成 HTML5 `<audio>` 元素

### 2.2 图表增强

#### 更多图表类型
```typescript
type PptChartType =
  | 'bar'
  | 'column'
  | 'line'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'area'
  | 'radar'
  | 'bubble';
```

**实现要点：**
- 解析 `<c:chart>` 相关标签
- 支持 Office 图表 XML 格式
- 集成图表库（如 ECharts、Chart.js）渲染

#### 图表配置
```typescript
interface PptChartContent {
  chartType: PptChartType;
  title?: string;
  categories: string[];
  series: PptChartSeries[];
  showLegend?: boolean;
  showDataLabels?: boolean;
  showGrid?: boolean;
}
```

### 2.3 SmartArt 图表

```typescript
type PptNodeType = 'smartart';

type PptSmartArtType =
  | 'process'
  | 'cycle'
  | 'hierarchy'
  | 'relationship'
  | 'matrix'
  | 'pyramid'
  | 'timeline';
```

**实现要点：**
- 解析 `<p:smartArt>` 标签
- 解析 `dml` (DrawingML) 图形数据
- 渲染层次化结构

### 2.4 公式和方程式

```typescript
type PptNodeType = 'equation';

interface PptEquationContent {
  latex?: string;      // LaTeX 格式
  mathML?: string;     // MathML 格式
  image?: string;      // 公式图片
}
```

**实现要点：**
- 解析 Office MathML 格式
- 转换为 LaTeX（使用 MathJax 或 KaTeX）
- 或直接渲染为图片

---

## 🎨 阶段三：视觉效果（中优先级）

### 3.1 阴影效果

```typescript
interface PptShadow {
  color?: string;
  blur?: number;
  offsetX?: number;
  offsetY?: number;
  opacity?: number;
}
```

**实现要点：**
- 解析 `<a:effectLst><a:outerShdw>` 标签
- 支持 CSS `box-shadow` 转换

### 3.2 反射效果

```typescript
interface PptReflection {
  blur?: number;
  opacity?: number;
  offset?: number;
}
```

**实现要点：**
- 解析 `<a:reflection>` 标签
- 使用 CSS `box-reflect` 或 SVG 滤镜

### 3.3 发光效果

```typescript
interface PptGlow {
  color?: string;
  radius?: number;
  opacity?: number;
}
```

**实现要点：**
- 解析 `<a:glow>` 标签
- 使用 CSS `filter: drop-shadow()` 或 SVG 滤镜

### 3.4 3D 效果

```typescript
interface PptEffect3D {
  material?: 'matte' | 'plastic' | 'metal' | 'wireframe';
  lightRig?: 'harsh' | 'flat' | 'normal' | 'soft';
  bevel?: { type?: string; width?: number; height?: number };
}
```

**实现要点：**
- 解析 `<a:sp3d>` 标签
- 使用 CSS 3D transforms 或 WebGL

---

## 📊 阶段四：幻灯片功能（低优先级）

### 4.1 幻灯片过渡效果

```typescript
interface PptSlideTransition {
  type?: 'none' | 'fade' | 'slide' | 'push' | 'wipe' | 'zoom';
  duration?: number;
  direction?: 'left' | 'right' | 'up' | 'down';
}
```

**实现要点：**
- 解析 `<p:transition>` 标签
- 使用 CSS transitions 或动画

### 4.2 幻灯片布局

```typescript
type PptSlideLayout =
  | 'blank'
  | 'title'
  | 'titleOnly'
  | 'titleAndContent'
  | 'sectionHeader'
  | 'twoContent'
  | 'comparison'
  | 'verticalText'
  | 'contentWithCaption';
```

**实现要点：**
- 解析 `<p:sldLayout>` 标签
- 支持母版幻灯片继承

### 4.3 演讲者备注

```typescript
interface PptSlide {
  props: {
    notes?: string;
  };
}
```

**实现要点：**
- 解析 `ppt/notesSlides/notesSlideX.xml` 文件

---

## 🎯 阶段五：主题和母版（低优先级）

### 5.1 主题定义

```typescript
interface PptTheme {
  name?: string;
  colors?: {
    background?: string;
    text?: string;
    accent1?: string;
    accent2?: string;
    accent3?: string;
    accent4?: string;
    accent5?: string;
    accent6?: string;
  };
  fonts?: {
    heading?: string;
    body?: string;
  };
}
```

**实现要点：**
- 解析 `ppt/theme/themeX.xml` 文件
- 解析 `ppt/slideMasters/slideMasterX.xml` 文件
- 支持主题颜色继承

### 5.2 母版幻灯片

**实现要点：**
- 解析母版元素
- 将母版样式应用到幻灯片
- 处理占位符替换

---

## 📝 实现优先级

### P0（必须实现）
- 文本：项目符号、超链接、富文本样式
- 形状：更多形状类型、变换效果
- 填充：渐变填充、图片填充
- 边框：多种边框样式

### P1（重要功能）
- 媒体：视频、音频
- 图表：更多图表类型
- 连线：支持形状之间的连线

### P2（增强功能）
- SmartArt 图表
- 公式和方程式
- 阴影、反射、发光效果

### P3（可选功能）
- 3D 效果
- 幻灯片过渡效果
- 主题和母版

---

## 🔧 技术实现参考

### PPTXjs 的关键技术

1. **XML 解析**: 使用 `tXml` 库（轻量级 XML 解析）
2. **PPTX 结构**:
   - `[Content_Types].xml` - 文件类型映射
   - `ppt/presentation.xml` - 主文档结构
   - `ppt/slides/slideX.xml` - 幻灯片内容
   - `ppt/slideLayouts/slideLayoutX.xml` - 布局定义
   - `ppt/slideMasters/slideMasterX.xml` - 母版
   - `ppt/theme/themeX.xml` - 主题
   - `ppt/_rels/*` - 关系映射（图片、媒体等）

3. **单位转换**: EMU (914400) ↔ PX (96)
   ```javascript
   const px = emu * 96 / 914400;
   const emu = px * 914400 / 96;
   ```

4. **样式解析**:
   - 字体: `<a:rPr>` 标签
   - 段落: `<a:pPr>` 标签
   - 形状: `<a:spPr>` 标签
   - 填充: `<a:solidFill>`, `<a:gradFill>`, `<a:blipFill>`
   - 边框: `<a:ln>` 标签

---

## 📚 参考资源

- [Office Open XML 规范](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [PPTXjs 源码](https://github.com/meshesha/PPTXjs)
- [DrawingML 参考文档](https://docs.microsoft.com/en-us/openspecs/office_standards/)
