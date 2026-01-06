/**
 * PPTX 核心类型定义
 */

import type { BaseElement } from '../elements/BaseElement';

/** 关联关系映射表 */
export interface RelsMap {
  [relId: string]: {
    id: string;
    type: string;
    target: string;
  };
}

/** 解析选项 */
export interface ParseOptions {
  /** 是否解析图片资源 */
  parseImages?: boolean;
  /** 是否保留原始XML */
  keepRawXml?: boolean;
  /** 详细日志 */
  verbose?: boolean;
}

/** 元数据对象 */
export interface Metadata {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  description?: string;
  created?: string;
  modified?: string;
}

/** 页面尺寸 */
export interface SlideSize {
  width: number;
  height: number;
}

/** 幻灯片页面属性 */
export interface SlideProps {
  width: number;
  height: number;
  ratio: number;
  pageSize: '4:3' | '16:9' | '16:10' | 'custom';
}

/** PPTX 解析结果 */
export interface PptxParseResult {
  id: string;
  title: string;
  author?: string;
  subject?: string;
  keywords?: string;
  description?: string;
  created?: string;
  modified?: string;
  slides: SlideParseResult[];
  props: SlideProps;
  globalRelsMap: RelsMap;
  theme?: ThemeResult;
  masterSlides?: MasterSlideResult[];
  slideLayouts?: Record<string, SlideLayoutResult>;
  notesMasters?: NotesMasterResult[];
  notesSlides?: NotesSlideResult[];
  charts?: ChartResult[];
  diagrams?: DiagramResult[];
  tags?: TagsResult[];
}

/** 导入SlideLayoutResult类型以便在PptxParseResult中使用 */
/**
 * 幻灯片版式解析结果
 * 对应 PPTXjs 的 layout 解析能力
 */
export interface SlideLayoutResult {
  id: string;
  name?: string;
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string; schemeRef?: string };
  elements: any[];
  /** 占位符定义（布局规则） */
  placeholders?: Placeholder[];
  relsMap: RelsMap;
  colorMap?: Record<string, string>;
  /** 文本样式（从 p:txStyles 解析） */
  textStyles?: TextStyles;
  /** 对 master 的引用（从 layout 的 _rels 解析） */
  masterRef?: string;
  /** master 对象（由 parser 填充） */
  master?: MasterSlideResult;
}

/**
 * 占位符定义
 * 对应 PPTXjs 的 placeholder 布局规则
 */
export interface Placeholder {
  id: string;
  type: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';
  name?: string;
  /** 位置尺寸（EMU单位） */
  rect: { x: number; y: number; width: number; height: number };
  /** 水平对齐 */
  hAlign?: 'left' | 'center' | 'right';
  /** 垂直对齐 */
  vAlign?: 'top' | 'middle' | 'bottom';
  /** 占位符索引 */
  idx?: number;
  /** 原始XML节点 */
  rawNode?: Element;
}

/** 主题解析结果 */
export interface ThemeResult {
  colors: ThemeColors;
}

/** 主题颜色方案 */
export interface ThemeColors {
  bg1?: string;
  tx1?: string;
  bg2?: string;
  tx2?: string;
  accent1?: string;
  accent2?: string;
  accent3?: string;
  accent4?: string;
  accent5?: string;
  accent6?: string;
  hlink?: string;
  folHlink?: string;
}

/** 幻灯片母版解析结果 */
export interface MasterSlideResult {
  id: string;
  /** 母版文件名（如 slideMaster1） */
  masterId?: string;
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string; schemeRef?: string };
  elements: any[];
  /** 母版元素（footer, slide number等）的位置和样式 */
  placeholders?: any[];
  colorMap: Record<string, string>;
  /** 文本样式（从 p:txStyles 解析） */
  textStyles?: TextStyles;
  /** 对 theme 的引用（从 master 的 _rels 解析） */
  themeRef?: string;
  /** 关联关系映射表 */
  relsMap?: any;
}

/** 文本样式（从 master 或 layout 的 txStyles 解析） */
export interface TextStyles {
  /** 标题样式 */
  titleParaPr?: any;
  /** 正文样式 */
  bodyPr?: any;
  /** 其他样式 */
  otherPr?: any;
}

/** 幻灯片背景 */
export interface Background {
  type: 'color' | 'image' | 'none';
  value?: string;
  relId?: string;
  schemeRef?: string; // 引用主题颜色的标识
}

/** 单个幻灯片解析结果 */
export interface SlideParseResult {
  id: string;
  title: string;
  background?: string | Background;
  elements: BaseElement[];
  relsMap: RelsMap;
  rawXml?: string;
  index?: number;
  layoutId?: string;
  /** 引用的布局对象（由 parser 填充） */
  layout?: SlideLayoutResult;
  /** 引用的母版对象（由 parser 填充） */
  master?: MasterSlideResult;
}

/** 关联关系 */
export interface Relationship {
  id: string;
  type: string;
  target: string;
}

// ============ 图表/绘图相关类型 ============

/** 图表数据系列 */
export interface ChartSeriesData {
  name?: string;
  idx?: number;
  order?: number;
  points?: ChartDataPoint[];
  color?: string;
}

/** 图表数据点 */
export interface ChartDataPoint {
  idx?: number;
  value?: number;
  category?: string;
}

/** 图表解析结果 */
export interface ChartResult {
  id: string;
  chartType: 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart' | 'unknown';
  title?: string;
  series?: ChartSeriesData[];
  categories?: string[];
  xTitle?: string;
  yTitle?: string;
  showLegend?: boolean;
  showDataLabels?: boolean;
  relsMap: RelsMap;
}

/** SmartArt/Diagram 形状 */
export interface DiagramShapeData {
  id: string;
  type: string;
  position?: { x: number; y: number };
  size?: { width: number; height: number };
  text?: string;
}

/** SmartArt/Diagram 解析结果 */
export interface DiagramResult {
  id: string;
  diagramType?: string;
  layout?: string;
  colors?: Record<string, string>;
  data?: Record<string, any>;
  shapes?: DiagramShapeData[];
  relsMap: RelsMap;
}

// ============ 讲演者备注相关类型 ============

/** 备注占位符定义 */
export interface NotesPlaceholder {
  id: string;
  type: 'header' | 'body' | 'dateTime' | 'slideImage' | 'footer' | 'other';
  name?: string;
  rect: { x: number; y: number; width: number; height: number };
}

/** 备注母版解析结果 */
export interface NotesMasterResult {
  id: string;
  elements: any[];
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string };
  placeholders?: NotesPlaceholder[];
  relsMap: RelsMap;
}

/** 备注页解析结果 */
export interface NotesSlideResult {
  id: string;
  slideId?: string; // 关联的幻灯片ID
  text?: string; // 备注文本
  elements: any[];
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string };
  relsMap: RelsMap;
  masterRef?: string; // 引用的母版
  master?: NotesMasterResult; // 母版对象
}

// ============ 标签相关类型 ============

/** 幻灯片标签 */
export interface SlideTag {
  name: string;
  value: string;
}

/** 扩展数据 */
export interface ExtensionData {
  uri?: string;
  data?: any;
}

/** 自定义属性 */
export interface CustomProperty {
  name: string;
  value: any;
  type?: 'string' | 'number' | 'boolean' | 'date';
}

/** 标签解析结果 */
export interface TagsResult {
  id: string;
  slideId?: string;
  tags: SlideTag[];
  extensions: ExtensionData[];
  customProperties: CustomProperty[];
  relsMap: RelsMap;
}
