/**
 * PPTX解析增强类型定义
 * 基于原库 types.ts 增量扩展，保持完全兼容
 */

import type { PptElement, PptRect, PptStyle, PptTextStyle } from './types';

// ============ 解析结果扩展类型 ============
/** 幻灯片元素解析结果 */
export interface SlideParseResult {
  /** 幻灯片ID */
  id: string;
  /** 幻灯片标题 */
  title: string;
  /** 背景颜色 */
  background: string;
  /** 元素列表 */
  elements: ParsedSlideElement[];
  /** 关联关系映射 */
  relsMap: Record<string, Relation>;
  /** 幻灯片原始XML（可选，用于调试） */
  rawXml?: string;
}

/** 解析后的单个元素（增强版） */
export interface ParsedSlideElement extends PptElement {
  /** 元素名称 */
  name?: string;
  /** 是否隐藏 */
  hidden?: boolean;
  /** 完整文本内容（纯文本，用于搜索） */
  text?: string;
  /** 关联ID（用于引用图片、OLE对象等） */
  relId?: string;
  /** 原始属性集合 */
  attrs?: Record<string, string>;
  /** 原始XML节点（可选） */
  rawNode?: Element;
}

/** 形状元素（增强版） */
export interface ParsedShapeElement extends ParsedSlideElement {
  type: 'shape' | 'text';
  /** 形状类型（矩形、圆形等） */
  shapeType?: string;
  /** 文本内容 */
  text?: string;
  /** 文本样式（运行级别） */
  textStyle?: Array<{ text: string; style: Partial<PptTextStyle> }>;
  /** 是否占位符 */
  isPlaceholder?: boolean;
  /** 占位符类型 */
  placeholderType?: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';
}

/** 图片元素（增强版） */
export interface ParsedImageElement extends ParsedSlideElement {
  type: 'image';
  /** 图片URL或Base64 */
  src: string;
  /** 图片关联ID */
  relId: string;
  /** MIME类型 */
  mimeType?: string;
  /** 替代文本 */
  altText?: string;
}

/** OLE嵌入对象元素 */
export interface ParsedOleElement extends ParsedSlideElement {
  type: 'ole';
  /** OLE对象类型标识符 */
  progId?: string;
  /** 关联ID */
  relId: string;
  /** 对象名称 */
  name?: string;
  /** 是否包含降级图片 */
  hasFallback?: boolean;
}

/** 图表元素 */
export interface ParsedChartElement extends ParsedSlideElement {
  type: 'chart';
  /** 图表类型 */
  chartType?: string;
  /** 关联ID */
  relId: string;
}

/** 分组元素 */
export interface ParsedGroupElement extends ParsedSlideElement {
  type: 'group';
  /** 分组内的子元素 */
  children: ParsedSlideElement[];
}

/** 元素类型联合 */
export type SlideElementType =
  | ParsedShapeElement
  | ParsedImageElement
  | ParsedOleElement
  | ParsedChartElement
  | ParsedGroupElement;

// ============ 关联关系类型 ============
/** 关联关系 */
export interface Relation {
  id: string;
  type: string;
  target: string;
}

/** 关联关系映射表 */
export type RelsMap = Record<string, Relation>;

// ============ PPTX解析完整结果 ============
/** PPTX解析完整结果 */
export interface PptxParseResult {
  /** PPT文档ID */
  id: string;
  /** PPT标题 */
  title: string;
  /** 作者 */
  author?: string;
  /** 主题 */
  subject?: string;
  /** 关键词 */
  keywords?: string;
  /** 描述 */
  description?: string;
  /** 创建时间 */
  created?: string;
  /** 修改时间 */
  modified?: string;
  /** 幻灯片列表 */
  slides: SlideParseResult[];
  /** 幻灯片尺寸（像素） */
  props: {
    width: number;
    height: number;
    ratio: number;
    pageSize?: '4:3' | '16:9' | '16:10' | 'custom';
  };
  /** 全局关联关系映射 */
  globalRelsMap?: RelsMap;
}

// ============ 位置尺寸类型 ============
/** 位置和尺寸（像素单位） */
export interface Position {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** PPTX原始位置尺寸（EMU单位） */
export interface EmuPosition {
  x: string;
  y: string;
  cx: string;
  cy: string;
}

// ============ 文本样式扩展 ============
/** 文本运行样式 */
export interface TextRunStyle {
  /** 字体大小（pt） */
  fontSize?: number;
  /** 字体名称 */
  fontFamily?: string;
  /** 加粗 */
  bold?: boolean;
  /** 斜体 */
  italic?: boolean;
  /** 下划线 */
  underline?: boolean;
  /** 删除线 */
  strike?: boolean;
  /** 颜色（十六进制） */
  color?: string;
}

/** 文本运行 */
export interface TextRun {
  text: string;
  style?: TextRunStyle;
}

// ============ XML解析辅助类型 ============
/** 幻灯片XML信息 */
export interface SlideXmlInfo {
  xml: string;
  relsXml?: string;
  slideId: string;
  rId: string;
}

/** 解析选项 */
export interface ParseOptions {
  /** 是否解析图片Base64 */
  parseImages?: boolean;
  /** 是否保留原始XML */
  keepRawXml?: boolean;
  /** 是否详细日志 */
  verbose?: boolean;
  /** 自定义命名空间映射 */
  customNS?: Record<string, string>;
}

/** 图片解析结果 */
export interface ImageParseResult {
  relId: string;
  mimeType?: string;
  base64?: string;
  blob?: Blob;
}
