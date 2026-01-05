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
}

/** 幻灯片背景 */
export interface Background {
  type: 'color' | 'image' | 'none';
  value?: string;
  relId?: string;
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
}

/** 关联关系 */
export interface Relationship {
  id: string;
  type: string;
  target: string;
}
