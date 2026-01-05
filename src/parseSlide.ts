/**
 * PPTX幻灯片解析核心函数
 * 遵循 ECMA-376 OpenXML 标准，完整解析 slide.xml 中的所有元素
 *
 * 核心功能：
 * 1. 完整解析 <p:spTree> 下的4类核心节点：p:sp, p:pic, p:graphicFrame, p:grpSp
 * 2. 支持命名空间查询，避免错误匹配
 * 3. 完善的容错处理，节点不存在时返回默认值
 * 4. 解析文本内容、位置尺寸、关联关系等核心属性
 * 5. 返回可调用toHTML的元素实例
 */

import { NS } from './constants';
import {
  getChildrenByTagNS,
  getFirstChildByTagNS,
  generateId,
  log
} from './utils';
import type {
  SlideParseResult,
  RelsMap
} from './types-enhanced';
import {
  ShapeElement,
  ImageElement,
  OleElement,
  ChartElement,
  GroupElement,
  type BaseElement
} from './elements';

/**
 * 解析单个幻灯片（核心函数）
 * @param slideXml 幻灯片XML字符串
 * @param relsMap 关联关系映射表
 * @param slideIndex 幻灯片索引（用于生成标题）
 * @returns 幻灯片解析结果
 */
export function parseSlide(
  slideXml: string,
  relsMap: RelsMap = {},
  slideIndex: number = 0
): SlideParseResult {
  try {
    // 解析XML文档
    const parser = new DOMParser();
    const doc = parser.parseFromString(slideXml, 'application/xml');

    // 验证根节点
    const root = doc.documentElement;
    if (root.tagName !== 'p:sld' && !root.tagName.includes('sld')) {
      log('warn', 'Invalid slide XML root element', root.tagName);
      return createEmptySlide(slideIndex);
    }

    // 解析背景色
    const background = parseSlideBackground(root);

    // 解析幻灯片标题
    const title = parseSlideTitle(root, slideIndex);

    // 解析元素列表（返回可调用toHTML的元素实例）
    const elements = parseSlideElements(root, relsMap);

    log('info', `Parsed slide ${slideIndex + 1}: ${elements.length} elements`);

    return {
      id: generateId('slide'),
      title,
      background,
      elements: elements.map(el => el.toParsedElement ? el.toParsedElement() : el),
      relsMap
    };
  } catch (error) {
    log('error', `Failed to parse slide ${slideIndex}`, error);
    return createEmptySlide(slideIndex);
  }
}

/**
 * 解析幻灯片所有元素，返回可调用toHTML的元素实例
 * @param root 幻灯片根元素
 * @param relsMap 关联关系映射表
 * @returns 元素数组
 */
export function parseSlideElements(root: Element, relsMap: RelsMap): BaseElement[] {
  const elements: BaseElement[] = [];

  // 查找 <p:spTree> 容器
  const cSld = getFirstChildByTagNS(root, 'cSld', NS.p);
  if (!cSld) {
    log('warn', 'cSld element not found');
    return elements;
  }

  const spTree = getFirstChildByTagNS(cSld, 'spTree', NS.p);
  if (!spTree) {
    log('warn', 'spTree element not found');
    return elements;
  }

  log('info', 'Processing spTree children...');

  // 遍历 spTree 的所有子元素
  Array.from(spTree.children).forEach((child, index) => {
    if (child.nodeType !== 1) return; // 跳过非元素节点

    const tagName = child.tagName;
    log('info', `Processing element ${index}: ${tagName}`);

    // 根据标签类型分发解析，返回元素实例
    let element: BaseElement | null = null;

    if (tagName === 'p:sp' || tagName === 'sp') {
      element = ShapeElement.fromNode(child, relsMap);
    } else if (tagName === 'p:pic' || tagName === 'pic') {
      element = ImageElement.fromNode(child, relsMap);
    } else if (tagName === 'p:graphicFrame' || tagName === 'graphicFrame') {
      element = parseGraphicFrameElement(child, relsMap);
    } else if (tagName === 'p:grpSp' || tagName === 'grpSp') {
      element = GroupElement.fromNode(child, relsMap);
    } else {
      log('info', `Skipping unknown element type: ${tagName}`);
    }

    if (element) {
      elements.push(element);
    }
  });

  log('info', `Total elements parsed: ${elements.length}`);

  return elements;
}

/**
 * 解析图形框元素（<p:graphicFrame>）
 * 判断是OLE对象还是图表，并返回对应的元素实例
 */
function parseGraphicFrameElement(
  graphicFrameNode: Element,
  relsMap: RelsMap
): BaseElement | null {
  try {
    // 查找 graphicData 判断类型
    const graphic = getFirstChildByTagNS(graphicFrameNode, 'graphic', NS.a);
    const graphicData = graphic ? getFirstChildByTagNS(graphic, 'graphicData', NS.a) : null;

    if (!graphicData) {
      log('warn', 'graphicData not found in graphicFrame');
      return null;
    }

    const uri = graphicData.getAttribute('uri') || '';

    // 判断类型
    if (uri.includes('oleObject')) {
      return OleElement.fromNode(graphicFrameNode, relsMap);
    } else if (uri.includes('chart')) {
      return ChartElement.fromNode(graphicFrameNode, relsMap);
    } else {
      log('info', `Unknown graphicFrame type: ${uri}`);
      return null;
    }
  } catch (error) {
    log('error', 'Failed to parse graphicFrame element', error);
    return null;
  }
}

/**
 * 解析幻灯片背景色
 * @param root 幻灯片根元素
 * @returns 背景颜色（十六进制）
 */
function parseSlideBackground(root: Element): string {
  // 查找 <p:bgPr> 节点
  const bgPr = getFirstChildByTagNS(root, 'bgPr', NS.p);
  if (!bgPr) {
    return '#ffffff'; // 默认白色
  }

  // 查找填充类型
  const solidFill = getFirstChildByTagNS(bgPr, 'solidFill', NS.a);
  if (!solidFill) {
    return '#ffffff';
  }

  // 提取颜色值
  const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
  if (srgbClr?.getAttribute('val')) {
    return `#${srgbClr.getAttribute('val')}`;
  }

  // 检查方案引用
  const schemeClr = getFirstChildByTagNS(solidFill, 'schemeClr', NS.a);
  if (schemeClr?.getAttribute('val')) {
    return schemeClr.getAttribute('val') || '#ffffff';
  }

  return '#ffffff';
}

/**
 * 解析幻灯片标题
 * @param root 幻灯片根元素
 * @param defaultIndex 默认索引
 * @returns 标题文本
 */
function parseSlideTitle(root: Element, defaultIndex: number): string {
  // 查找 <p:cSld> -> <p:spTree> 下的元素
  const cSld = getFirstChildByTagNS(root, 'cSld', NS.p);
  if (!cSld) {
    return `幻灯片 ${defaultIndex + 1}`;
  }

  const spTree = getFirstChildByTagNS(cSld, 'spTree', NS.p);
  if (!spTree) {
    return `幻灯片 ${defaultIndex + 1}`;
  }

  // 遍历所有形状，查找标题占位符
  const shapes = getChildrenByTagNS(spTree, 'sp', NS.p);
  for (const sp of shapes) {
    const nvSpPr = getFirstChildByTagNS(sp, 'nvSpPr', NS.p);
    if (!nvSpPr) continue;

    const nvPr = getFirstChildByTagNS(nvSpPr, 'nvPr', NS.p);
    if (!nvPr) continue;

    const ph = getFirstChildByTagNS(nvPr, 'ph', NS.p);
    if (!ph) continue;

    // 检查是否是标题占位符
    const phType = ph.getAttribute('type');
    if (phType === 'title' || phType === 'ctrTitle') {
      // 提取标题文本
      const txBody = getFirstChildByTagNS(sp, 'txBody', NS.p);
      if (!txBody) continue;

      const titleText = extractTextFromTxBody(txBody);
      if (titleText && titleText.trim()) {
        return titleText.trim();
      }
    }
  }

  return `幻灯片 ${defaultIndex + 1}`;
}

/**
 * 从txBody中提取文本
 */
function extractTextFromTxBody(txBody: Element): string {
  const paragraphs = Array.from(txBody.children).filter(
    child => child.tagName === 'a:p' || child.tagName.includes(':p')
  );

  const texts: string[] = [];

  for (const p of paragraphs) {
    const runs = Array.from(p.children).filter(
      child => child.tagName === 'a:r' || child.tagName.includes(':r')
    );

    for (const r of runs) {
      const t = getFirstChildByTagNS(r, 't', 'http://schemas.openxmlformats.org/drawingml/2006/main');
      if (t?.textContent) {
        texts.push(t.textContent);
      }
    }
  }

  return texts.join('');
}

/**
 * 创建空幻灯片（容错处理）
 * @param slideIndex 幻灯片索引
 * @returns 空幻灯片对象
 */
function createEmptySlide(slideIndex: number): SlideParseResult {
  return {
    id: generateId('slide'),
    title: `幻灯片 ${slideIndex + 1}`,
    background: '#ffffff',
    elements: [],
    relsMap: {}
  };
}

// 导出类型
export type {
  SlideParseResult
};
