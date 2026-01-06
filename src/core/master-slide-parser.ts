/**
 * 幻灯片母版解析器
 * 解析slideMaster文件，提取母版元素和背景
 */

import JSZip from 'jszip';
import { log, getFirstChildByTagNS, getChildrenByTagNS, parseRels, parseRelsWithBase } from '../utils/index';
import { NS } from '../constants';
import type { RelsMap, TextStyles } from './types';
import { ShapeElement } from '../elements/ShapeElement';
import { ImageElement } from '../elements/ImageElement';
import { GroupElement } from '../elements/GroupElement';
import { TableElement } from '../elements/TableElement';
import { ChartElement } from '../elements/ChartElement';
import { DiagramElement } from '../elements/DiagramElement';
import type { BaseElement } from '../elements/BaseElement';

export interface MasterSlideResult {
  id: string;
  /** 母版文件名（如 slideMaster1） */
  masterId?: string;
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string; schemeRef?: string };
  elements: any[];
  /** 母版元素（footer, slide number等）的位置和样式 */
  placeholders?: any[];
  colorMap: Record<string, string>;
  /** 对 theme 的引用（从 master 的 _rels 解析） */
  themeRef?: string;
  /** 关联关系映射表 */
  relsMap?: any;
}

/**
 * 解析所有幻灯片母版
 * @param zip JSZip对象
 * @returns 母版数组
 */
export async function parseAllMasterSlides(zip: JSZip): Promise<MasterSlideResult[]> {
  try {
    const masterFiles = Object.keys(zip.files)
      .filter(path => path.startsWith('ppt/slideMasters/') && path.endsWith('.xml'))
      .sort();

    log('info', `Found ${masterFiles.length} master slide files`);

    const masters: MasterSlideResult[] = [];

    for (const masterFile of masterFiles) {
      const masterNumber = masterFile.match(/slideMaster(\d+)\.xml/)?.[1];
      if (!masterNumber) continue;

      // 读取母版XML
      const masterXml = await zip.file(masterFile)?.async('string');
      if (!masterXml) continue;

      // 读取母版关联关系（关键：PPTXjs 使用 master 的 _rels 获取对 theme 的引用）
      const relsPath = masterFile.replace('slideMasters/', 'slideMasters/_rels/')
        .replace('.xml', '.xml.rels');

      let relsMap: RelsMap = {};
      try {
        const relsXml = await zip.file(relsPath)?.async('string');
        if (relsXml) {
          relsMap = parseRelsWithBase(relsXml, relsPath);
          log('info', `Loaded ${Object.keys(relsMap).length} relationships for master ${masterNumber}`);
          // 打印关系链
          Object.entries(relsMap).forEach(([id, rel]) => {
            log('info', `  - ${id}: type=${rel.type}, target=${rel.target}`);
          });
        }
      } catch (e) {
        log('warn', `Failed to read master rels: ${relsPath}`, e);
      }

      // 解析母版
      const master = parseMasterSlide(masterXml, relsMap, masterNumber);
      masters.push(master);
    }

    return masters;
  } catch (error) {
    log('error', 'Failed to parse master slides', error);
    return [];
  }
}

/**
 * 解析单个幻灯片母版
 * @param masterXml 母版XML字符串
 * @param relsMap 关联关系映射表
 * @param masterNumber 母版编号
 * @returns 母版解析结果
 */
function parseMasterSlide(
  masterXml: string,
  relsMap: RelsMap,
  masterNumber: string
): MasterSlideResult {
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(masterXml, 'application/xml');
    const root = doc.documentElement;

    // 解析背景
    const background = parseMasterBackground(root, relsMap);

    // 解析颜色映射
    const colorMap = parseColorMap(root);

    // 解析文本样式（关键：PPTXjs 的样式继承）
    const textStyles = parseTextStyles(root);

    // 解析元素 (footer, slide numbers, shapes, images等)
    const elements = parseMasterElements(root, relsMap);
    const placeholders = elements.filter((el: any) => el.placeholder);  // 提取占位符元素

    // 获取 theme 引用（从 relsMap 中）
    let themeRef: string | undefined;
    for (const rel of Object.values(relsMap)) {
      if (rel.type.includes('theme')) {
        // 提取 theme 文件名，例如 "../theme/theme1.xml" -> "theme1"
        const match = rel.target.match(/theme(\d+)\.xml/);
        if (match) {
          themeRef = `theme${match[1]}`;
          log('info', `Master references theme: ${themeRef}`);
        }
      }
    }

    return {
      id: `master-${masterNumber}`,
      masterId: `slideMaster${masterNumber}`,
      background,
      elements,
      placeholders,
      colorMap,
      textStyles,
      themeRef,
      relsMap
    };
  } catch (error) {
    log('error', `Failed to parse master slide ${masterNumber}`, error);
    return {
      id: `master-${masterNumber}`,
      masterId: `slideMaster${masterNumber}`,
      elements: [],
      colorMap: {}
    };
  }
}

/**
 * 解析母版背景
 * @param root 母版根元素
 * @param relsMap 关联关系映射表
 * @returns 背景对象
 */
function parseMasterBackground(
  root: Element,
  relsMap: RelsMap
): { type: 'color' | 'image' | 'none'; value?: string; relId?: string; schemeRef?: string } {
  // 查找 <p:bg>
  const bg = getFirstChildByTagNS(root, 'bg', 
    'http://schemas.openxmlformats.org/presentationml/2006/main');
  
  if (!bg) {
    return { type: 'color', value: '#ffffff' };
  }

  // 1. 检查背景引用 <p:bgRef>
  const bgRef = getFirstChildByTagNS(bg, 'bgRef', 
    'http://schemas.openxmlformats.org/presentationml/2006/main');
  
  if (bgRef) {
    const idx = bgRef.getAttribute('idx');
    const schemeClr = getFirstChildByTagNS(bgRef, 'schemeClr', 
      'http://schemas.openxmlformats.org/drawingml/2006/main');
    
    if (schemeClr) {
      const val = schemeClr.getAttribute('val');
      if (val) {
        // 返回方案颜色引用，需要后续解析主题文件获取实际颜色
        return { type: 'color', value: val, schemeRef: val };
      }
    }
  }

  // 2. 检查背景属性 <p:bgPr>
  const bgPr = getFirstChildByTagNS(bg, 'bgPr', 
    'http://schemas.openxmlformats.org/presentationml/2006/main');
  
  if (bgPr) {
    // 图片填充
    const blipFill = getFirstChildByTagNS(bgPr, 'blipFill', 
      'http://schemas.openxmlformats.org/drawingml/2006/main');
    
    if (blipFill) {
      const blip = getFirstChildByTagNS(blipFill, 'blip', 
        'http://schemas.openxmlformats.org/drawingml/2006/main');
      
      if (blip) {
        const relId = blip.getAttribute('r:embed') || blip.getAttributeNS(
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
        
        if (relId && relsMap[relId]) {
          return {
            type: 'image',
            value: relsMap[relId].target,
            relId
          };
        }
      }
    }

    // 纯色填充
    const solidFill = getFirstChildByTagNS(bgPr, 'solidFill', 
      'http://schemas.openxmlformats.org/drawingml/2006/main');
    
    if (solidFill) {
      const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', 
        'http://schemas.openxmlformats.org/drawingml/2006/main');
      
      if (srgbClr?.getAttribute('val')) {
        return { type: 'color', value: `#${srgbClr.getAttribute('val')}` };
      }

      const schemeClr = getFirstChildByTagNS(solidFill, 'schemeClr', 
        'http://schemas.openxmlformats.org/drawingml/2006/main');
      
      if (schemeClr) {
        const val = schemeClr.getAttribute('val');
        if (val) {
          return { 
            type: 'color', 
            value: val || '#ffffff',
            schemeRef: val || undefined
          };
        }
      }
    }
  }

  return { type: 'color', value: '#ffffff' };
}

/**
 * 解析母版元素（如页脚、幻灯片编号）
 * @param root 母版根元素
 * @returns 元素数组
 */
function parseMasterElements(root: Element, relsMap: RelsMap): any[] {
  const elements: any[] = [];

  const cSld = getFirstChildByTagNS(root, 'cSld', 
    'http://schemas.openxmlformats.org/presentationml/2006/main');
  
  if (!cSld) return elements;

  const spTree = getFirstChildByTagNS(cSld, 'spTree', 
    'http://schemas.openxmlformats.org/presentationml/2006/main');
  
  if (!spTree) return elements;

  // 遍历所有子元素
  Array.from(spTree.children).forEach((child: any) => {
    if (child.nodeType !== 1) return;

    const localName = child.localName || child.tagName.split(':').pop();

    if (localName === 'sp') {
      // 查找占位符
      const nvSpPr = getFirstChildByTagNS(child, 'nvSpPr', NS.p);
      const nvPr = nvSpPr ? getFirstChildByTagNS(nvSpPr, 'nvPr', NS.p) : null;
      const ph = nvPr ? getFirstChildByTagNS(nvPr, 'ph', NS.p) : null;

      if (ph) {
        // 占位符元素（footer, slide number等）
        const phType = ph.getAttribute('type');
        const supportedTypes = ['sldNum', 'ftr', 'dt'];
        if (supportedTypes.includes(phType || '')) {
          const element = parsePlaceholderShape(child);
          if (element) {
            elements.push(element);
          }
        }
      } else {
        // 非占位符的形状
        const element = ShapeElement.fromNode(child, relsMap);
        if (element) {
          elements.push(element);
        }
      }
    } else if (localName === 'pic') {
      const element = ImageElement.fromNode(child, relsMap);
      if (element) {
        elements.push(element);
      }
    } else if (localName === 'grpSp') {
      const element = GroupElement.fromNode(child, relsMap);
      if (element) {
        elements.push(element);
      }
    } else if (localName === 'graphicFrame') {
      const element = parseMasterGraphicFrame(child, relsMap);
      if (element) {
        elements.push(element);
      }
    }
  });

  return elements;
}


/**
 * 解析母版中的图片
 * @param picEl 图片元素
 * @returns 图片对象
 */
function parsePictureForMaster(picEl: Element): any | null {
  const nvPicPr = getFirstChildByTagNS(picEl, 'nvPicPr', 
    'http://schemas.openxmlformats.org/presentationml/2006/main');
  
  if (!nvPicPr) return null;

  const cNvPr = getFirstChildByTagNS(nvPicPr, 'cNvPr', 
    'http://schemas.openxmlformats.org/presentationml/2006/main');
  
  const id = cNvPr?.getAttribute('id');
  const name = cNvPr?.getAttribute('name');

  const blipFill = getFirstChildByTagNS(picEl, 'blipFill', 
    'http://schemas.openxmlformats.org/drawingml/2006/main');
  
  if (!blipFill) return null;

  const blip = getFirstChildByTagNS(blipFill, 'blip', 
    'http://schemas.openxmlformats.org/drawingml/2006/main');
  
  if (!blip) return null;

  const relId = blip.getAttribute('r:embed') || blip.getAttributeNS(
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');

  return {
    id,
    name,
    type: 'image',
    relId,
    // 可以添加更多属性（位置、样式等）
  };
}

/**
 * 解析颜色映射
 * @param root 根元素
 * @returns 颜色映射对象
 */
function parseColorMap(root: Element): Record<string, string> {
  const clrMapOvr = getFirstChildByTagNS(root, 'clrMapOvr',
    'http://schemas.openxmlformats.org/presentationml/2006/main');

  const clrMap: Record<string, string> = {};

  if (!clrMapOvr) {
    return clrMap;
  }

  // 解析 masterClrMapping
  const masterClrMapping = getFirstChildByTagNS(clrMapOvr, 'masterClrMapping',
    'http://schemas.openxmlformats.org/drawingml/2006/main');

  if (masterClrMapping) {
    // 读取所有属性作为映射
    Array.from(masterClrMapping.attributes).forEach(attr => {
      clrMap[attr.name] = attr.value;
    });
  }

  return clrMap;
}

/**
 * 解析文本样式（从 master 或 layout 的 p:txStyles）
 * 对应 PPTXjs 的文本样式继承机制
 * @param root 根元素
 * @returns 文本样式对象
 */
function parseTextStyles(root: Element): TextStyles {
  const txStyles = getFirstChildByTagNS(root, 'txStyles',
    'http://schemas.openxmlformats.org/presentationml/2006/main');

  if (!txStyles) {
    return {};
  }

  return {
    titleParaPr: getFirstChildByTagNS(txStyles, 'titleP',
      'http://schemas.openxmlformats.org/presentationml/2006/main'),
    bodyPr: getFirstChildByTagNS(txStyles, 'body',
      'http://schemas.openxmlformats.org/presentationml/2006/main'),
    otherPr: getFirstChildByTagNS(txStyles, 'other',
      'http://schemas.openxmlformats.org/presentationml/2006/main')
  };
}

/**
 * 解析母版中的占位符形状
 */
function parsePlaceholderShape(shapeEl: Element): any | null {
  const nvSpPr = getFirstChildByTagNS(shapeEl, 'nvSpPr', NS.p);
  if (!nvSpPr) return null;

  const cNvPr = getFirstChildByTagNS(nvSpPr, 'cNvPr', NS.p);
  const id = cNvPr?.getAttribute('id');
  const name = cNvPr?.getAttribute('name');

  const nvPr = getFirstChildByTagNS(nvSpPr, 'nvPr', NS.p);
  if (!nvPr) return null;

  const ph = getFirstChildByTagNS(nvPr, 'ph', NS.p);
  if (!ph) return null;

  const phType = ph.getAttribute('type');

  // 解析位置和尺寸
  const spPr = getFirstChildByTagNS(shapeEl, 'spPr', NS.a);
  const xfrm = getFirstChildByTagNS(spPr, 'xfrm', NS.a);

  let rect = { x: 0, y: 0, width: 0, height: 0 };
  if (xfrm) {
    const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
    const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

    if (off) {
      rect.x = parseInt(off.getAttribute('x') || '0');
      rect.y = parseInt(off.getAttribute('y') || '0');
    }
    if (ext) {
      rect.width = parseInt(ext.getAttribute('cx') || '0');
      rect.height = parseInt(ext.getAttribute('cy') || '0');
    }
  }

  return {
    id,
    name,
    type: 'shape',
    placeholder: phType,
    rect,
  };
}

/**
 * 解析母版中的 graphicFrame 元素
 */
function parseMasterGraphicFrame(graphicFrameNode: Element, relsMap: RelsMap): BaseElement | null {
  try {
    const graphic = getFirstChildByTagNS(graphicFrameNode, 'graphic', NS.a);
    const graphicData = graphic ? getFirstChildByTagNS(graphic, 'graphicData', NS.a) : null;

    if (!graphicData) {
      return null;
    }

    const uri = graphicData.getAttribute('uri') || '';

    if (uri.includes('table')) {
      return TableElement.fromNode(graphicFrameNode, relsMap);
    } else if (uri.includes('chart')) {
      return ChartElement.fromNode(graphicFrameNode, relsMap);
    } else if (uri.includes('diagram')) {
      return DiagramElement.fromNode(graphicFrameNode, relsMap);
    }

    return null;
  } catch (error) {
    log('warn', 'Failed to parse master graphicFrame element', error);
    return null;
  }
}



