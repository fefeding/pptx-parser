/**
 * 幻灯片解析器
 * 处理单个和多个幻灯片的解析
 */

import JSZip from 'jszip';
import { PATHS, NS } from '../constants';
import {
  getChildrenByTagNS,
  getFirstChildByTagNS,
  generateId,
  log
} from '../utils/index';
import {
  ShapeElement,
  ImageElement,
  OleElement,
  ChartElement,
  GroupElement,
  TableElement,
  DiagramElement,
  type BaseElement
} from '../elements';
import { parseSlideRels } from './relationships-parser';
import type { ParseOptions, RelsMap, SlideParseResult } from './types';



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

    // 解析背景（支持颜色和图片）
    const background = parseSlideBackground(root, relsMap);

    // 解析幻灯片标题
    const title = parseSlideTitle(root, slideIndex);

    // 解析元素列表（返回可调用toHTML的元素实例）
    const elements = parseSlideElements(root, relsMap);

    log('info', `Parsed slide ${slideIndex + 1}: ${elements.length} elements`);

    return {
      id: generateId('slide'),
      title,
      background,
      elements: elements.map(el => el),  // 直接返回元素实例，保留toHTML方法
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
    const localName = child.localName || tagName.split(':').pop() || tagName;
    log('info', `Processing element ${index}: tagName=${tagName}, localName=${localName}`);

    // 根据标签类型分发解析，返回元素实例
    let element: BaseElement | null = null;

    // 使用本地名称进行判断
    if (localName === 'sp') {
      element = ShapeElement.fromNode(child, relsMap);
    } else if (localName === 'pic') {
      element = ImageElement.fromNode(child, relsMap);
    } else if (localName === 'graphicFrame') {
      element = parseGraphicFrameElement(child, relsMap);
    } else if (localName === 'grpSp') {
      element = GroupElement.fromNode(child, relsMap);
    } else {
      log('info', `Skipping unknown element type: tagName=${tagName}, localName=${localName}`);
    }

    if (element) {
      elements.push(element);
    }
  });

  log('info', `Total elements parsed: ${elements.length}`);

  return elements;
}

/**
 * 解析所有幻灯片
 * @param zip JSZip对象
 * @param options 解析选项
 * @returns 幻灯片数组
 */
export async function parseAllSlides(
  zip: JSZip,
  options: ParseOptions
): Promise<SlideParseResult[]> {
  try {
    // 获取所有幻灯片文件
    const slideFiles = Object.keys(zip.files)
      .filter(path => path.startsWith(PATHS.SLIDES))
      .filter(path => path.endsWith('.xml'))
      .filter(path => !path.includes('_rels'))
      .sort((a, b) => {
        // 按文件名数字排序
        const numA = parseInt(a.match(/slide(\d+)\.xml/)?.[1] || '0', 10);
        const numB = parseInt(b.match(/slide(\d+)\.xml/)?.[1] || '0', 10);
        return numA - numB;
      });

    log('info', `Found ${slideFiles.length} slide files`);

    const slides: SlideParseResult[] = [];

    for (let i = 0; i < slideFiles.length; i++) {
      const slidePath = slideFiles[i];
      log('info', `Parsing slide ${i + 1}: ${slidePath}`);

      // 读取幻灯片XML
      const slideXml = await zip.file(slidePath)?.async('string');
      if (!slideXml) {
        log('warn', `Failed to read slide: ${slidePath}`);
        continue;
      }

      // 读取幻灯片的关联关系文件
      const slideNumber = slidePath.match(/slide(\d+)\.xml/)?.[1];
      let relsMap: RelsMap = {};

      if (slideNumber) {
        relsMap = await parseSlideRels(zip, slideNumber);
        log('info', `Loaded ${Object.keys(relsMap).length} relationships for slide ${slideNumber}`);
      }

      // 解析幻灯片
      const slide = parseSlide(slideXml, relsMap, i);

      // 保存原始XML（如果需要）
      if (options.keepRawXml) {
        slide.rawXml = slideXml;
      }

      slides.push(slide);
    }

    return slides;
  } catch (error) {
    log('error', 'Failed to parse slides', error);
    return [];
  }
}

/**
 * 解析图形框元素（<p:graphicFrame>）
 * 判断是OLE对象、图表、表格还是图解，并返回对应的元素实例
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
    } else if (uri.includes('diagram')) {
      return DiagramElement.fromNode(graphicFrameNode, relsMap);
    } else if (uri.includes('table')) {
      return TableElement.fromNode(graphicFrameNode, relsMap);
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
 * 解析背景引用 <p:bgRef>
 * @param bgRef bgRef元素
 * @param relsMap 关联关系映射表
 * @returns 背景对象
 */
function parseBgRef(bgRef: Element, relsMap: RelsMap): { type: 'color' | 'image' | 'none'; value?: string; relId?: string; schemeRef?: string } {
  // bgRef引用的是主题中的背景，返回默认白色
  const idx = bgRef.getAttribute('idx');
  const schemeClr = getFirstChildByTagNS(bgRef, 'schemeClr', NS.a);
  
  if (schemeClr) {
    const val = schemeClr.getAttribute('val');
    if (val) {
      // 返回方案颜色值，后续可以解析主题文件获取实际颜色
      return { type: 'color', value: val, schemeRef: val };
    }
  }
  
  return { type: 'color', value: '#ffffff' };
}

/**
 * 解析幻灯片背景（支持颜色和图片）
 * @param root 幻灯片根元素
 * @param relsMap 关联关系映射表
 * @returns 背景对象 { type: 'color'|'image', value: string }
 */
export function parseSlideBackground(root: Element, relsMap: RelsMap = {}): { type: 'color' | 'image' | 'none'; value?: string; relId?: string; schemeRef?: string } {
  // 查找 <p:bg> 节点
  const bg = getFirstChildByTagNS(root, 'bg', NS.p);
  if (!bg) {
    return { type: 'color', value: '#ffffff' }; // 默认白色
  }

  // 1. 检查背景引用 <p:bgRef>
  const bgRef = getFirstChildByTagNS(bg, 'bgRef', NS.p);
  if (bgRef) {
    return parseBgRef(bgRef, relsMap);
  }

  // 2. 检查背景属性 <p:bgPr>
  const bgPr = getFirstChildByTagNS(bg, 'bgPr', NS.p);
  if (bgPr) {
    // 检查图片填充 <a:blipFill>
    const blipFill = getFirstChildByTagNS(bgPr, 'blipFill', NS.a);
    if (blipFill) {
      const blip = getFirstChildByTagNS(blipFill, 'blip', NS.a);
      if (blip) {
        const relId = blip.getAttribute('r:embed') || blip.getAttributeNS(NS.r, 'embed');
        if (relId && relsMap[relId]) {
          return {
            type: 'image',
            value: relsMap[relId].target,
            relId
          };
        }
      }
    }

    // 检查纯色填充 <a:solidFill>
    const solidFill = getFirstChildByTagNS(bgPr, 'solidFill', NS.a);
    if (solidFill) {
      // 提取颜色值
      const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
      if (srgbClr?.getAttribute('val')) {
        return { type: 'color', value: `#${srgbClr.getAttribute('val')}` };
      }

      // 检查方案引用
      const schemeClr = getFirstChildByTagNS(solidFill, 'schemeClr', NS.a);
      if (schemeClr) {
        const val = schemeClr.getAttribute('val');
        if (val) {
          return { type: 'color', value: val || '#ffffff', schemeRef: val || undefined };
        }
      }
    }

    // 检查渐变填充 <a:gradFill>
    const gradFill = getFirstChildByTagNS(bgPr, 'gradFill', NS.a);
    if (gradFill) {
      // 简化处理：渐变背景返回白色
      return { type: 'color', value: '#ffffff' };
    }
  }

  return { type: 'color', value: '#ffffff' };
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
