/**
 * 讲演者备注解析器
 * 解析 notesMasters（备注母版）和 notesSlides（备注页）
 * 对齐 PPTXjs 的备注解析能力
 */

import JSZip from 'jszip';
import { getFirstChildByTagNS, getAttrSafe, emu2px, log } from '../utils';
import { NS } from '../constants';
import type { RelsMap } from './types';

/**
 * 备注母版解析结果
 */
export interface NotesMasterResult {
  id: string;
  elements: any[];
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string };
  placeholders?: Placeholder[];
  relsMap: RelsMap;
}

/**
 * 备注页解析结果
 */
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

/**
 * 备注占位符定义
 */
export interface Placeholder {
  id: string;
  type: 'header' | 'body' | 'dateTime' | 'slideImage' | 'footer' | 'other';
  name?: string;
  rect: { x: number; y: number; width: number; height: number };
}

/**
 * 解析所有备注母版
 * @param zip JSZip对象
 * @returns 备注母版数组
 */
export async function parseAllNotesMasters(zip: JSZip): Promise<NotesMasterResult[]> {
  try {
    const masterFiles = Object.keys(zip.files)
      .filter(path => path.startsWith('ppt/notesMasters/'))
      .filter(path => path.endsWith('.xml'))
      .filter(path => !path.includes('_rels'))
      .sort((a, b) => {
        const numA = parseInt(a.match(/notesMaster(\d+)\.xml/)?.[1] || '0', 10);
        const numB = parseInt(b.match(/notesMaster(\d+)\.xml/)?.[1] || '0', 10);
        return numA - numB;
      });

    log('info', `Found ${masterFiles.length} notes master files`);

    const masters: NotesMasterResult[] = [];

    for (const masterPath of masterFiles) {
      const masterXml = await zip.file(masterPath)?.async('string');
      if (!masterXml) {
        log('warn', `Failed to read notes master: ${masterPath}`);
        continue;
      }

      const master = await parseNotesMaster(zip, masterPath, masterXml);
      masters.push(master);
    }

    return masters;
  } catch (error) {
    log('error', 'Failed to parse notes masters', error);
    return [];
  }
}

/**
 * 解析单个备注母版
 * @param zip JSZip对象
 * @param masterPath 母版路径
 * @param masterXml 母版XML
 * @returns 备注母版解析结果
 */
async function parseNotesMaster(
  zip: JSZip,
  masterPath: string,
  masterXml: string
): Promise<NotesMasterResult> {
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(masterXml, 'application/xml');
    const root = doc.documentElement;

    // 解析背景
    const background = parseNotesBackground(root);

    // 解析元素
    const elements = parseNotesElements(root);

    // 解析占位符
    const placeholders = parseNotesPlaceholders(root);

    // 解析关联关系
    const masterId = masterPath.match(/notesMaster(\d+)\.xml/)?.[1] || '1';
    const relsMap = await parseNotesMasterRels(zip, `notesMaster${masterId}`);

    log('info', `Parsed notes master: ${masterId}`);

    return {
      id: `notesMaster_${masterId}`,
      elements,
      background,
      placeholders,
      relsMap
    };
  } catch (error) {
    log('error', `Failed to parse notes master: ${masterPath}`, error);
    throw error;
  }
}

/**
 * 解析所有备注页
 * @param zip JSZip对象
 * @returns 备注页数组
 */
export async function parseAllNotesSlides(zip: JSZip): Promise<NotesSlideResult[]> {
  try {
    const slideFiles = Object.keys(zip.files)
      .filter(path => path.startsWith('ppt/notesSlides/'))
      .filter(path => path.endsWith('.xml'))
      .filter(path => !path.includes('_rels'))
      .sort((a, b) => {
        const numA = parseInt(a.match(/notesSlide(\d+)\.xml/)?.[1] || '0', 10);
        const numB = parseInt(b.match(/notesSlide(\d+)\.xml/)?.[1] || '0', 10);
        return numA - numB;
      });

    log('info', `Found ${slideFiles.length} notes slide files`);

    const slides: NotesSlideResult[] = [];

    for (const slidePath of slideFiles) {
      const slideXml = await zip.file(slidePath)?.async('string');
      if (!slideXml) {
        log('warn', `Failed to read notes slide: ${slidePath}`);
        continue;
      }

      const slide = await parseNotesSlide(zip, slidePath, slideXml);
      slides.push(slide);
    }

    return slides;
  } catch (error) {
    log('error', 'Failed to parse notes slides', error);
    return [];
  }
}

/**
 * 解析单个备注页
 * @param zip JSZip对象
 * @param slidePath 备注页路径
 * @param slideXml 备注页XML
 * @returns 备注页解析结果
 */
async function parseNotesSlide(
  zip: JSZip,
  slidePath: string,
  slideXml: string
): Promise<NotesSlideResult> {
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(slideXml, 'application/xml');
    const root = doc.documentElement;

    // 解析背景
    const background = parseNotesBackground(root);

    // 解析文本内容
    const text = parseNotesText(root);

    // 解析元素
    const elements = parseNotesElements(root);

    // 解析关联关系
    const slideId = slidePath.match(/notesSlide(\d+)\.xml/)?.[1] || '1';
    const relsMap = await parseNotesSlideRels(zip, `notesSlide${slideId}`);

    // 获取母版引用
    const masterRef = relsMap['rId1']?.target || '';

    log('info', `Parsed notes slide: ${slideId}`);

    return {
      id: `notesSlide_${slideId}`,
      slideId: `slide${slideId}`, // 假设备注页与幻灯片一一对应
      text,
      elements,
      background,
      relsMap,
      masterRef
    };
  } catch (error) {
    log('error', `Failed to parse notes slide: ${slidePath}`, error);
    throw error;
  }
}

/**
 * 解析备注背景
 */
function parseNotesBackground(root: Element): { type: 'color' | 'image' | 'none'; value?: string; relId?: string } {
  const bg = getFirstChildByTagNS(root, 'bg', NS.p);
  if (!bg) {
    return { type: 'color', value: '#ffffff' };
  }

  const bgPr = getFirstChildByTagNS(bg, 'bgPr', NS.p);
  if (bgPr) {
    const solidFill = getFirstChildByTagNS(bgPr, 'solidFill', NS.a);
    if (solidFill) {
      const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
      if (srgbClr?.getAttribute('val')) {
        return { type: 'color', value: `#${srgbClr.getAttribute('val')}` };
      }
    }
  }

  return { type: 'color', value: '#ffffff' };
}

/**
 * 解析备注文本内容
 */
function parseNotesText(root: Element): string {
  const cSld = getFirstChildByTagNS(root, 'cSld', NS.p);
  if (!cSld) return '';

  const spTree = getFirstChildByTagNS(cSld, 'spTree', NS.p);
  if (!spTree) return '';

  // 查找主体文本
  const shapes = Array.from(spTree.children).filter(
    child => child.tagName === 'p:sp' || child.tagName.includes(':sp')
  );

  for (const sp of shapes) {
    const nvSpPr = getFirstChildByTagNS(sp, 'nvSpPr', NS.p);
    if (!nvSpPr) continue;

    const nvPr = getFirstChildByTagNS(nvSpPr, 'nvPr', NS.p);
    if (!nvPr) continue;

    const ph = getFirstChildByTagNS(nvPr, 'ph', NS.p);
    if (ph?.getAttribute('type') === 'body') {
      // 提取文本
      const txBody = getFirstChildByTagNS(sp, 'txBody', NS.p);
      if (!txBody) continue;

      const text = extractTextFromTxBody(txBody);
      if (text && text.trim()) {
        return text.trim();
      }
    }
  }

  return '';
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

  return texts.join('\n');
}

/**
 * 解析备注元素
 */
function parseNotesElements(root: Element): any[] {
  const elements: any[] = [];

  const cSld = getFirstChildByTagNS(root, 'cSld', NS.p);
  if (!cSld) return elements;

  const spTree = getFirstChildByTagNS(cSld, 'spTree', NS.p);
  if (!spTree) return elements;

  Array.from(spTree.children).forEach((child, index) => {
    if (child.nodeType !== 1) return;

    const tagName = child.tagName;

    if (tagName.includes('sp')) {
      elements.push(parseShapeElement(child));
    } else if (tagName.includes('pic')) {
      elements.push(parsePictureElement(child));
    }
  });

  return elements;
}

/**
 * 解析形状元素
 */
function parseShapeElement(node: Element): any {
  const cNvPr = getFirstChildByTagNS(node, 'cNvPr', NS.p);
  const id = getAttrSafe(cNvPr, 'id', `shape_${Date.now()}`);
  const name = getAttrSafe(cNvPr, 'name', '');

  const spPr = getFirstChildByTagNS(node, 'spPr', NS.p);
  const xfrm = spPr ? getFirstChildByTagNS(spPr, 'xfrm', NS.a) : null;

  let rect = { x: 0, y: 0, width: 0, height: 0 };
  if (xfrm) {
    const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
    const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

    if (off) {
      rect.x = emu2px(off.getAttribute('x') || '0');
      rect.y = emu2px(off.getAttribute('y') || '0');
    }
    if (ext) {
      rect.width = emu2px(ext.getAttribute('cx') || '0');
      rect.height = emu2px(ext.getAttribute('cy') || '0');
    }
  }

  return {
    id,
    name,
    type: 'shape',
    rect,
    content: {}
  };
}

/**
 * 解析图片元素
 */
function parsePictureElement(node: Element): any {
  const cNvPr = getFirstChildByTagNS(node, 'cNvPr', NS.p);
  const id = getAttrSafe(cNvPr, 'id', `image_${Date.now()}`);
  const name = getAttrSafe(cNvPr, 'name', '');

  const spPr = getFirstChildByTagNS(node, 'spPr', NS.p);
  const xfrm = spPr ? getFirstChildByTagNS(spPr, 'xfrm', NS.a) : null;

  let rect = { x: 0, y: 0, width: 0, height: 0 };
  if (xfrm) {
    const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
    const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

    if (off) {
      rect.x = emu2px(off.getAttribute('x') || '0');
      rect.y = emu2px(off.getAttribute('y') || '0');
    }
    if (ext) {
      rect.width = emu2px(ext.getAttribute('cx') || '0');
      rect.height = emu2px(ext.getAttribute('cy') || '0');
    }
  }

  return {
    id,
    name,
    type: 'image',
    rect,
    content: {}
  };
}

/**
 * 解析备注占位符
 */
function parseNotesPlaceholders(root: Element): Placeholder[] {
  const placeholders: Placeholder[] = [];

  const cSld = getFirstChildByTagNS(root, 'cSld', NS.p);
  if (!cSld) return placeholders;

  const spTree = getFirstChildByTagNS(cSld, 'spTree', NS.p);
  if (!spTree) return placeholders;

  const shapes = Array.from(spTree.children).filter(
    child => child.tagName === 'p:sp' || child.tagName.includes(':sp')
  );

  for (const sp of shapes) {
    const nvSpPr = getFirstChildByTagNS(sp, 'nvSpPr', NS.p);
    if (!nvSpPr) continue;

    const nvPr = getFirstChildByTagNS(nvSpPr, 'nvPr', NS.p);
    if (!nvPr) continue;

    const ph = getFirstChildByTagNS(nvPr, 'ph', NS.p);
    if (!ph) continue;

    const cNvPr = getFirstChildByTagNS(nvSpPr, 'cNvPr', NS.p);
    const id = getAttrSafe(cNvPr, 'id', `placeholder_${Date.now()}`);
    const name = getAttrSafe(cNvPr, 'name', '');

    const phType = ph.getAttribute('type') || 'other';

    const spPr = getFirstChildByTagNS(sp, 'spPr', NS.p);
    const xfrm = spPr ? getFirstChildByTagNS(spPr, 'xfrm', NS.a) : null;

    let rect = { x: 0, y: 0, width: 0, height: 0 };
    if (xfrm) {
      const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
      const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

      if (off) {
        rect.x = emu2px(off.getAttribute('x') || '0');
        rect.y = emu2px(off.getAttribute('y') || '0');
      }
      if (ext) {
        rect.width = emu2px(ext.getAttribute('cx') || '0');
        rect.height = emu2px(ext.getAttribute('cy') || '0');
      }
    }

    placeholders.push({
      id,
      type: phType as any,
      name,
      rect
    });
  }

  return placeholders;
}

/**
 * 解析备注母版关联关系
 */
async function parseNotesMasterRels(zip: JSZip, masterId: string): Promise<RelsMap> {
  try {
    const relsPath = `ppt/notesMasters/_rels/${masterId}.xml.rels`;
    const relsXml = await zip.file(relsPath)?.async('string');

    if (!relsXml) {
      log('warn', `Notes master rels not found: ${relsPath}`);
      return {};
    }

    return parseRelationshipsXml(relsXml);
  } catch (error) {
    log('warn', `Failed to parse notes master rels for ${masterId}`, error);
    return {};
  }
}

/**
 * 解析备注页关联关系
 */
async function parseNotesSlideRels(zip: JSZip, slideId: string): Promise<RelsMap> {
  try {
    const relsPath = `ppt/notesSlides/_rels/${slideId}.xml.rels`;
    const relsXml = await zip.file(relsPath)?.async('string');

    if (!relsXml) {
      log('warn', `Notes slide rels not found: ${relsPath}`);
      return {};
    }

    return parseRelationshipsXml(relsXml);
  } catch (error) {
    log('warn', `Failed to parse notes slide rels for ${slideId}`, error);
    return {};
  }
}

/**
 * 解析关联关系XML
 */
function parseRelationshipsXml(relsXml: string): RelsMap {
  const relsMap: RelsMap = {};

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(relsXml, 'application/xml');
    const root = doc.documentElement;

    const relationships = Array.from(root.children).filter(
      child => child.tagName === 'Relationship' || child.tagName.includes(':Relationship')
    );

    for (const rel of relationships) {
      const id = rel.getAttribute('Id') || '';
      const type = rel.getAttribute('Type') || '';
      const target = rel.getAttribute('Target') || '';

      if (id) {
        relsMap[id.replace('rId', '')] = {
          id,
          type,
          target
        };
      }
    }
  } catch (error) {
    log('warn', 'Failed to parse relationships XML', error);
  }

  return relsMap;
}

/**
 * 关联备注页和母版
 */
export function linkNotesToMasters(notesSlides: NotesSlideResult[], masters: NotesMasterResult[]): void {
  const masterMap = new Map<string, NotesMasterResult>();
  masters.forEach(master => {
    const masterId = master.id.replace('notesMaster_', '');
    masterMap.set(masterId, master);
  });

  notesSlides.forEach(notesSlide => {
    if (notesSlide.masterRef) {
      const masterId = notesSlide.masterRef.split('/').pop()?.replace('.xml', '') || '';
      if (masterMap.has(masterId)) {
        notesSlide.master = masterMap.get(masterId);
      }
    }
  });
}
