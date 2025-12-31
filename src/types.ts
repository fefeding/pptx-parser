/**
 * PPT解析/序列化核心类型定义 - 纯TS，无任何耦合，全覆盖所有结构
 */

// ============ 元素类型定义 ============
export type PptNodeType =
  | 'text'
  | 'image'
  | 'shape'
  | 'table'
  | 'chart'
  | 'container'
  | 'media'
  | 'video'
  | 'audio'
  | 'line'
  | 'connector'
  | 'group'
  | 'smartart'
  | 'equation';

// ============ 坐标/尺寸基础结构 ============
/** 坐标/尺寸基础结构 - PPT核心基础模型 */
export interface PptRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** 变换属性（旋转、翻转等） */
export interface PptTransform {
  rotate?: number; // 旋转角度（度）
  flipH?: boolean; // 水平翻转
  flipV?: boolean; // 垂直翻转
}

// ============ 样式基础结构 ============
/** 文本样式 */
export interface PptTextStyle {
  fontSize?: number;
  fontFamily?: string;
  fontStyle?: 'normal' | 'italic';
  fontWeight?: 'normal' | 'bold';
  textDecoration?: 'none' | 'underline' | 'line-through';
  color?: string;
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  textVerticalAlign?: 'top' | 'middle' | 'bottom';
  lineHeight?: number;
  letterSpacing?: number;
  textShadow?: string;
}

/** 填充样式 */
export interface PptFill {
  type?: 'solid' | 'gradient' | 'pattern' | 'picture' | 'none';
  color?: string;
  gradientStops?: Array<{ position: number; color: string }>;
  gradientDirection?: number; // 渐变角度
  image?: string;
  opacity?: number;
}

/** 边框样式 */
export interface PptBorder {
  color?: string;
  width?: number;
  style?: 'solid' | 'dashed' | 'dotted' | 'double';
  dashStyle?: string; // 自定义虚线样式
}

/** 阴影效果 */
export interface PptShadow {
  color?: string;
  blur?: number;
  offsetX?: number;
  offsetY?: number;
  opacity?: number;
}

/** 反射效果 */
export interface PptReflection {
  blur?: number;
  opacity?: number;
  offset?: number;
}

/** 发光效果 */
export interface PptGlow {
  color?: string;
  radius?: number;
  opacity?: number;
}

/** 3D 效果 */
export interface PptEffect3D {
  material?: 'matte' | 'plastic' | 'metal' | 'wireframe';
  lightRig?: 'harsh' | 'flat' | 'normal' | 'soft';
  bevel?: {
    type?: 'relaxed' | 'slope' | 'angle' | 'circle' | 'convex';
    width?: number;
    height?: number;
  };
  contour?: {
    color?: string;
    width?: number;
  };
}

/** 完整样式基础结构 */
export interface PptStyle extends PptTextStyle {
  backgroundColor?: string | PptFill;
  borderColor?: string;
  borderWidth?: number;
  borderStyle?: 'solid' | 'dashed' | 'dotted' | 'double';
  border?: PptBorder;
  fill?: PptFill;
  shadow?: PptShadow;
  reflection?: PptReflection;
  glow?: PptGlow;
  effect3d?: PptEffect3D;
  opacity?: number;
  zIndex?: number;
}

// ============ 文本元素 ============
/** 文本段落 */
export interface PptTextParagraph {
  text: string;
  style?: Partial<PptTextStyle>;
  bullet?: {
    type?: 'none' | 'bullet' | 'numbered';
    char?: string; // 项目符号字符
    level?: number; // 层级
  };
  hyperlink?: {
    url: string;
    tooltip?: string;
  };
}

/** 文本元素内容 */
export type PptTextContent = PptTextParagraph[];

// ============ 图片元素 ============
/** 图片元素 */
export interface PptImageContent {
  src: string; // 图片URL或Base64
  alt?: string;
  mimeType?: string;
}

// ============ 形状元素 ============
/** 形状类型 */
export type PptShapeType =
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

/** 形状元素内容 */
export interface PptShapeContent {
  shapeType: PptShapeType;
  text?: string | PptTextContent; // 形状内的文本
  path?: string; // SVG路径（自定义形状）
  roundedCorners?: number; // 圆角半径
}

// ============ 表格元素 ============
/** 表格单元格样式 */
export interface PptTableCellStyle {
  backgroundColor?: string | PptFill;
  borderColor?: string;
  borderWidth?: number;
  verticalAlign?: 'top' | 'middle' | 'bottom';
  padding?: { top?: number; bottom?: number; left?: number; right?: number };
}

/** 表格单元格 */
export interface PptTableCell {
  text: string;
  style?: PptTableCellStyle;
  colspan?: number;
  rowspan?: number;
}

/** 表格元素内容 */
export type PptTableContent = PptTableCell[][];

// ============ 图表元素 ============
/** 图表类型 */
export type PptChartType =
  | 'bar'
  | 'column'
  | 'line'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'area'
  | 'radar'
  | 'bubble';

/** 图表数据系列 */
export interface PptChartSeries {
  name: string;
  data: number[];
  color?: string;
}

/** 图表元素内容 */
export interface PptChartContent {
  chartType: PptChartType;
  title?: string;
  categories: string[];
  series: PptChartSeries[];
  showLegend?: boolean;
  showDataLabels?: boolean;
  showGrid?: boolean;
}

// ============ 媒体元素 ============
/** 视频内容 */
export interface PptVideoContent {
  src: string; // 视频URL或Base64
  poster?: string; // 封面图
  mimeType?: string;
  autoplay?: boolean;
  loop?: boolean;
  muted?: boolean;
  controls?: boolean;
}

/** 音频内容 */
export interface PptAudioContent {
  src: string; // 音频URL或Base64
  mimeType?: string;
  autoplay?: boolean;
  loop?: boolean;
  volume?: number; // 0-100
}

/** 媒体内容类型 */
export type PptMediaContent = PptVideoContent | PptAudioContent;

// ============ 连线元素 ============
/** 连线样式 */
export interface PptConnectorStyle {
  startArrow?: 'none' | 'arrow' | 'stealth' | 'diamond' | 'oval';
  endArrow?: 'none' | 'arrow' | 'stealth' | 'diamond' | 'oval';
  lineType?: 'straight' | 'elbow' | 'curved';
}

/** 连线元素内容 */
export interface PptConnectorContent {
  startElementId?: string; // 起始元素ID
  endElementId?: string; // 结束元素ID
  startPoint?: { x: number; y: number };
  endPoint?: { x: number; y: number };
  style?: PptConnectorStyle;
}

// ============ SmartArt元素 ============
/** SmartArt类型 */
export type PptSmartArtType =
  | 'process'
  | 'cycle'
  | 'hierarchy'
  | 'relationship'
  | 'matrix'
  | 'pyramid'
  | 'timeline';

/** SmartArt元素内容 */
export interface PptSmartArtContent {
  smartArtType: PptSmartArtType;
  nodes: Array<{
    text: string;
    children?: PptSmartArtContent['nodes'];
    level?: number;
  }>;
}

// ============ 公式元素 ============
/** 公式元素内容 */
export interface PptEquationContent {
  latex?: string; // LaTeX格式
  mathML?: string; // MathML格式
  image?: string; // 公式图片
}

// ============ 元素内容联合类型 ============
/** 元素内容联合类型 */
export type PptElementContent =
  | string
  | PptTextContent
  | PptImageContent
  | PptShapeContent
  | PptTableContent
  | PptChartContent
  | PptMediaContent
  | PptConnectorContent
  | PptSmartArtContent
  | PptEquationContent;

// ============ 单个PPT元素节点 ============
/** 单个PPT元素节点（前端编辑器的标准数据结构） */
export interface PptElement {
  id: string;
  type: PptNodeType;
  rect: PptRect;
  transform?: PptTransform;
  style: PptStyle;
  content: PptElementContent;
  props: Record<string, unknown>;
  children?: PptElement[];
  parentId?: string;
}

// ============ 幻灯片结构 ============
/** 幻灯片过渡效果 */
export interface PptSlideTransition {
  type?: 'none' | 'fade' | 'slide' | 'push' | 'wipe' | 'zoom';
  duration?: number; // 毫秒
  direction?: 'left' | 'right' | 'up' | 'down';
}

/** 幻灯片布局类型 */
export type PptSlideLayout =
  | 'blank'
  | 'title'
  | 'titleOnly'
  | 'titleAndContent'
  | 'sectionHeader'
  | 'twoContent'
  | 'comparison'
  | 'verticalText'
  | 'contentWithCaption';

/** 单张幻灯片结构 */
export interface PptSlide {
  id: string;
  title: string;
  bgColor: string | PptFill;
  backgroundImage?: string; // 背景图片
  elements: PptElement[];
  props: {
    width: number;
    height: number;
    slideLayout: PptSlideLayout;
    transition?: PptSlideTransition;
    notes?: string; // 演讲者备注
    slideNumber?: number;
  };
}

// ============ PPT文档结构 ============
/** 主题定义 */
export interface PptTheme {
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

/** 完整PPT文档结构（前端编辑器的顶层数据结构） */
export interface PptDocument {
  id: string;
  title: string;
  author?: string;
  subject?: string;
  keywords?: string;
  description?: string;
  created?: string;
  modified?: string;
  slides: PptSlide[];
  theme?: PptTheme;
  props: {
    width: number;
    height: number;
    ratio: number;
    pageSize?: '4:3' | '16:9' | '16:10' | 'custom';
  };
}

// ============ XML中间结构 ============
/** PPTX文件解析的中间XML结构（标准Office XML映射） */
export interface XmlSlide {
  xml: string;
  slideId: string;
  rId: string;
  layout: string;
}

export interface XmlElement {
  tag: string;
  attrs: Record<string, string>;
  children: XmlElement[];
  text: string;
}

// ============ 关系映射 ============
/** 关系映射（用于解析图片、媒体等资源） */
export interface Relationship {
  id: string;
  type: string;
  target: string;
}

export interface RelationshipMap {
  [key: string]: Relationship;
}