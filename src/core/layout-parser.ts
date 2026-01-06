/**
 * 幻灯片布局解析器
 * 解析 slideLayouts 文件，处理布局继承关系
 */

import JSZip from 'jszip';
import { NS } from '../constants';
import { PATHS as CONSTANTS_PATHS } from '../constants';
import {
  getFirstChildByTagNS,
  generateId,
  log,
  parseRels
} from '../utils';
import { parseSlideBackground } from './slide-parser';
import type { RelsMap, SlideLayoutResult } from './types';

// 本地PATHS常量，避免循环引用
const PATHS = {
  slideLayouts: CONSTANTS_PATHS.SLIDE_LAYOUTS,
  slideLayoutsRels: 'ppt/slideLayouts/_rels/'
};

/**
 * 解析单个幻灯片布局
 * @param layoutXml 布局XML字符串
 * @param relsMap 关联关系映射表（来自 layout 的 _rels 目录）
 * @returns 布局解析结果
 */
export function parseSlideLayout(
  layoutXml: string,
  relsMap: RelsMap = {}
): SlideLayoutResult | null {
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(layoutXml, 'application/xml');
    const root = doc.documentElement;

    if (root.tagName !== 'p:sldLayout' && !root.tagName.includes('sldLayout')) {
      log('warn', 'Invalid layout XML root element', root.tagName);
      return null;
    }

    // 解析布局名称
    const cSld = getFirstChildByTagNS(root, 'cSld', NS.p);
    const name = cSld?.getAttribute('name') || undefined;

    // 解析背景
    const background = parseSlideBackground(root, relsMap);

    // 解析颜色映射（可选）
    const clrMapOvr = getFirstChildByTagNS(root, 'clrMapOvr', NS.p);
    const masterClrMapping = getFirstChildByTagNS(clrMapOvr, 'masterClrMapping', NS.a);
    let colorMap: Record<string, string> = {};

    if (masterClrMapping) {
      // 解析颜色映射关系
      const attrs = Array.from(masterClrMapping.attributes);
      attrs.forEach(attr => {
        if (attr.name !== 'xmlns:a') {
          colorMap[attr.name] = attr.value;
        }
      });
    }

    // 解析布局中的元素（暂时简化，主要关注背景）
    const elements: any[] = [];

    // 解析布局占位符（关键：PPTXjs 核心能力）
    const placeholders = parseLayoutPlaceholders(root);

  // 获取 master 引用（从 relsMap 中）
  let masterRef: string | undefined;
  for (const rel of Object.values(relsMap)) {
    if (rel.type.includes('slideMaster')) {
      // 提取 master 文件名，例如 "../slideMasters/slideMaster1.xml" -> "slideMaster1"
      const match = rel.target.match(/slideMaster(\d+)\.xml/);
      if (match) {
        masterRef = `slideMaster${match[1]}`;
        log('info', `Layout references master: ${masterRef}`);
      }
    }
  }

    return {
      id: generateId('layout'),
      name,
      background,
      elements,
      placeholders,
      relsMap,
      colorMap,
      masterRef
    };
  } catch (error) {
    log('error', 'Failed to parse slide layout', error);
    return null;
  }
}

/**
 * 解析所有幻灯片布局
 * @param zip JSZip对象
 * @returns 布局映射表 (layoutId -> SlideLayoutResult)
 */
export async function parseAllSlideLayouts(zip: JSZip): Promise<Record<string, SlideLayoutResult>> {
  try {
    const layouts: Record<string, SlideLayoutResult> = {};

    // 获取所有 slideLayout 文件
    const layoutFiles = Object.keys(zip.files)
      .filter(path => path.startsWith(PATHS.slideLayouts) && path.endsWith('.xml'))
      .filter(path => !path.includes('_rels'))
      .sort((a, b) => {
        const numA = parseInt(a.match(/slideLayout(\d+)\.xml/)?.[1] || '0', 10);
        const numB = parseInt(b.match(/slideLayout(\d+)\.xml/)?.[1] || '0', 10);
        return numA - numB;
      });

    log('info', `Found ${layoutFiles.length} layout files`);

    for (let i = 0; i < layoutFiles.length; i++) {
      const layoutPath = layoutFiles[i];
      log('info', `Parsing layout ${i + 1}: ${layoutPath}`);

      const layoutXml = await zip.file(layoutPath)?.async('string');
      if (!layoutXml) {
        log('warn', `Failed to read layout: ${layoutPath}`);
        continue;
      }

      // 读取布局的关联关系（关键：PPTXjs 使用 layout 的 _rels 获取对 master 的引用）
      const layoutNumber = layoutPath.match(/slideLayout(\d+)\.xml/)?.[1];
      let relsMap: RelsMap = {};

      if (layoutNumber) {
        const relsPath = `ppt/slideLayouts/_rels/slideLayout${layoutNumber}.xml.rels`;
        try {
          const relsXml = await zip.file(relsPath)?.async('string');
          if (relsXml) {
            relsMap = parseRels(relsXml);
            log('info', `Loaded ${Object.keys(relsMap).length} relationships for layout ${layoutNumber}`);
            // 打印关系链
            Object.entries(relsMap).forEach(([id, rel]) => {
              log('info', `  - ${id}: type=${rel.type}, target=${rel.target}`);
            });
          }
        } catch (e) {
          log('warn', `Failed to read layout rels: ${relsPath}`, e);
        }
      }

      const layout = parseSlideLayout(layoutXml, relsMap);
      if (layout) {
        const layoutId = `slideLayout${layoutNumber}`;
        layouts[layoutId] = layout;
      }
    }

    return layouts;
  } catch (error) {
    log('error', 'Failed to parse slide layouts', error);
    return {};
  }
}

/**
 * 解析布局占位符
 * 对应 PPTXjs 的 parseLayout 功能
 * @param root 布局根元素
 * @returns 占位符数组
 */
function parseLayoutPlaceholders(root: Element): any[] {
  const placeholders: any[] = [];

  const cSld = getFirstChildByTagNS(root, 'cSld', NS.p);
  if (!cSld) return placeholders;

  const spTree = getFirstChildByTagNS(cSld, 'spTree', NS.p);
  if (!spTree) return placeholders;

  // 遍历所有子元素
  Array.from(spTree.children).forEach(child => {
    if (child.nodeType !== 1) return;

    const localName = child.localName || child.tagName.split(':').pop();

    if (localName === 'sp') {
      const placeholder = parsePlaceholderFromShape(child);
      if (placeholder) {
        placeholders.push(placeholder);
      }
    }
  });

  return placeholders;
}

/**
 * 从形状元素解析占位符
 * 对应 PPTXjs 的 placeholder 解析
 * @param shapeEl 形状元素
 * @returns 占位符对象或 null
 */
function parsePlaceholderFromShape(shapeEl: Element): any | null {
  // 查找非视觉属性
  const nvSpPr = getFirstChildByTagNS(shapeEl, 'nvSpPr', NS.p);
  if (!nvSpPr) return null;

  const cNvPr = getFirstChildByTagNS(nvSpPr, 'cNvPr', NS.p);
  const id = cNvPr?.getAttribute('id') || '';
  const name = cNvPr?.getAttribute('name') || '';

  const nvPr = getFirstChildByTagNS(nvSpPr, 'nvPr', NS.p);
  if (!nvPr) return null;

  const ph = getFirstChildByTagNS(nvPr, 'ph', NS.p);
  if (!ph) return null;

  // 解析占位符类型
  const phType = ph.getAttribute('type');
  const typeMap: Record<string, 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other'> = {
    'title': 'title',
    'ctrTitle': 'title',
    'body': 'body',
    'dt': 'dateTime',
    'sldNum': 'slideNumber',
    'ftr': 'footer'
  };

  const type = typeMap[phType || ''] || 'other';
  const idx = ph.getAttribute('idx');

  // 解析位置和尺寸（spPr/xfrm）
  const spPr = getFirstChildByTagNS(shapeEl, 'spPr', NS.p);
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

  // 解析对齐方式（txBody/bodyPr）
  const txBody = getFirstChildByTagNS(shapeEl, 'txBody', NS.p);
  let hAlign: 'left' | 'center' | 'right' | undefined;
  let vAlign: 'top' | 'middle' | 'bottom' | undefined;

  if (txBody) {
    const bodyPr = getFirstChildByTagNS(txBody, 'bodyPr', NS.a);
    if (bodyPr) {
      const algn = bodyPr.getAttribute('algn');
      const anchor = bodyPr.getAttribute('anchor');

      // 水平对齐
      if (algn === 'l') hAlign = 'left';
      else if (algn === 'ctr') hAlign = 'center';
      else if (algn === 'r') hAlign = 'right';

      // 垂直对齐
      if (anchor === 't') vAlign = 'top';
      else if (anchor === 'ctr') vAlign = 'middle';
      else if (anchor === 'b') vAlign = 'bottom';
    }
  }

  return {
    id,
    type,
    name,
    rect,
    hAlign,
    vAlign,
    idx: idx ? parseInt(idx) : undefined,
    rawNode: shapeEl
  };
}

/**
 * 合并背景信息：优先级 slide > layout > master
 * @param slideBackground 幻灯片背景
 * @param layoutBackground 布局背景
 * @param masterBackground 母版背景
 * @returns 合并后的背景
 */
export function mergeBackgrounds(
  slideBackground?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string; schemeRef?: string },
  layoutBackground?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string; schemeRef?: string },
  masterBackground?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string; schemeRef?: string }
): { type: 'color' | 'image' | 'none'; value?: string; relId?: string; schemeRef?: string } {
  // 如果幻灯片有明确的背景，优先使用
  if (slideBackground && slideBackground.type !== 'none') {
    return slideBackground;
  }

  // 否则使用布局背景
  if (layoutBackground && layoutBackground.type !== 'none') {
    return layoutBackground;
  }

  // 最后使用母版背景
  if (masterBackground) {
    return masterBackground;
  }

  // 默认白色背景
  return { type: 'color', value: '#ffffff' };
}
