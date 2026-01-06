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
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string; schemeRef?: string };
  elements: any[];
  colorMap: Record<string, string>;
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
}

/** 关联关系 */
export interface Relationship {
  id: string;
  type: string;
  target: string;
}
