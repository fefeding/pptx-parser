/**
 * PPT解析/序列化核心类型定义 - 纯TS，无任何耦合，全覆盖所有结构
 */
export type PptNodeType =
  | 'text'
  | 'image'
  | 'shape'
  | 'table'
  | 'chart'
  | 'container'
  | 'media';

/** 坐标/尺寸基础结构 - PPT核心基础模型 */
export interface PptRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** 样式基础结构 */
export interface PptStyle {
  fontSize?: number;
  color?: string;
  fontWeight?: 'normal' | 'bold';
  textAlign?: 'left' | 'center' | 'right';
  backgroundColor?: string;
  borderColor?: string;
  borderWidth?: number;
}

/** 单个PPT元素节点（前端编辑器的标准数据结构） */
export interface PptElement {
  id: string;
  type: PptNodeType;
  rect: PptRect;
  style: PptStyle;
  content: string | string[][] | Record<string, any>;
  props: Record<string, unknown>;
  children?: PptElement[];
}

/** 单张幻灯片结构 */
export interface PptSlide {
  id: string;
  title: string;
  bgColor: string;
  elements: PptElement[];
  props: {
    width: number;
    height: number;
    slideLayout: string;
  };
}

/** 完整PPT文档结构（前端编辑器的顶层数据结构） */
export interface PptDocument {
  id: string;
  title: string;
  slides: PptSlide[];
  props: {
    width: number;
    height: number;
    ratio: number;
  };
}

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