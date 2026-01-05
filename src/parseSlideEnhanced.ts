/**
 * PPTX幻灯片增强解析入口
 * 返回包含toHTML能力的元素实例
 */

import { parseSlide, parseSlideElements } from './parseSlide';
import { SlideElement, PptxDocument } from './elements';
import type { SlideParseResult, RelsMap } from './types-enhanced';

/**
 * 解析单个幻灯片，返回带toHTML能力的元素实例
 * @param slideXml 幻灯片XML字符串
 * @param relsMap 关联关系映射表
 * @param slideIndex 幻灯片索引
 * @returns 带toHTML能力的幻灯片元素
 */
export function parseSlideWithElements(
  slideXml: string,
  relsMap: RelsMap = {},
  slideIndex: number = 0
): SlideElement {
  // 先解析原始数据
  const slideResult = parseSlide(slideXml, relsMap, slideIndex);

  // 再解析元素实例
  const parser = new DOMParser();
  const doc = parser.parseFromString(slideXml, 'application/xml');
  const root = doc.documentElement;

  const elementInstances = parseSlideElements(root, relsMap);

  return new SlideElement(slideResult, elementInstances);
}

/**
 * 创建PPTX文档对象（带toHTML能力）
 */
export function createPptxDocument(
  slidesResult: SlideParseResult[],
  slidesElements: SlideElement[],
  title: string = 'PPTX Presentation'
): PptxDocument {
  const { width, height } = slidesResult[0]?.props || { width: 960, height: 540 };

  return new PptxDocument(
    `pptx_${Date.now()}`,
    title,
    slidesElements,
    width,
    height
  );
}

// 导出类型
export type { SlideElement, PptxDocument };
