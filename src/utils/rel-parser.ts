/**
 * 关系文件解析工具
 * 解析 PPTX 关系文件（.rels）
 */

import { getFirstChildByTagNS } from './xml-parser';
import type { RelsMap } from '../core/types';

/**
 * 解析关系文件（rels文件）
 * @param relsXml 关系XML字符串
 * @returns 关系映射表 { relId: target }
 */
export function parseRels(relsXml: string): RelsMap {
  const parser = new DOMParser();
  const doc = parser.parseFromString(relsXml, 'application/xml');

  // 根元素是 <Relationships>，子元素是 <Relationship>
  // 使用不带命名空间前缀的方式查询，因为子元素继承父元素的默认命名空间
  const relationships = Array.from(doc.documentElement.children);

  const relsMap: RelsMap = {};

  relationships.forEach(rel => {
    const id = rel.getAttribute('Id');
    const type = rel.getAttribute('Type');
    const target = rel.getAttribute('Target');
    if (id) {
      relsMap[id] = { id, type: type || '', target: target || '' };
    }
  });

  return relsMap;
}

/**
 * 规范化目标路径，将其转换为相对于ZIP根目录的绝对路径
 * @param relsFilePath 关系文件路径（如 'ppt/slides/_rels/slide1.xml.rels'）
 * @param target 目标路径（如 '../media/image1.png'）
 * @returns 规范化后的绝对路径（如 'ppt/media/image1.png'）
 */
export function normalizeTargetPath(relsFilePath: string, target: string): string {
  // 如果目标是空字符串，直接返回
  if (!target) return target;
  
  // 如果目标以 '/' 开头，视为绝对路径，去掉开头的 '/'
  if (target.startsWith('/')) {
    target = target.substring(1);
  }
  
  // 找到 _rels/ 目录的位置
  const relsIndex = relsFilePath.indexOf('_rels/');
  let parentDir = relsFilePath;
  if (relsIndex !== -1) {
    // 截取到 _rels/ 之前的部分作为父目录
    parentDir = relsFilePath.substring(0, relsIndex);
  }
  
  // 将目标路径相对于父目录进行解析
  const fullPath = parentDir + target;
  
  // 规范化路径：移除 './'，处理 '../'
  const parts = fullPath.split('/');
  const result: string[] = [];
  for (const part of parts) {
    if (part === '' || part === '.') continue;
    if (part === '..') {
      if (result.length > 0) {
        result.pop();
      }
    } else {
      result.push(part);
    }
  }
  return result.join('/');
}

/**
 * 解析关系文件并规范化目标路径
 * @param relsXml 关系XML字符串
 * @param relsFilePath 关系文件路径（用于路径解析）
 * @returns 关系映射表 { relId: target }，其中target是绝对路径
 */
export function parseRelsWithBase(relsXml: string, relsFilePath: string): RelsMap {
  const relsMap = parseRels(relsXml);
  // 规范化每个目标路径
  for (const relId in relsMap) {
    const rel = relsMap[relId];
    rel.target = normalizeTargetPath(relsFilePath, rel.target);
  }
  return relsMap;
}

/**
 * 解析元数据（docProps/core.xml）
 * @param coreXml 核心属性XML字符串
 * @returns 元数据对象
 */
export function parseMetadata(coreXml: string): {
  title?: string;
  author?: string;
  created?: string;
  modified?: string;
  subject?: string;
  keywords?: string;
} {
  const parser = new DOMParser();
  const doc = parser.parseFromString(coreXml, 'application/xml');

  const dcNS = 'http://purl.org/dc/elements/1.1/';
  const dctermsNS = 'http://purl.org/dc/terms/';
  const cpNS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';

  const getTagText = (tagName: string, namespace: string): string | undefined => {
    const el = doc.getElementsByTagNameNS(namespace, tagName)[0];
    return el?.textContent || undefined;
  };

  return {
    title: getTagText('title', dcNS),
    author: getTagText('creator', dcNS),
    subject: getTagText('subject', dcNS),
    keywords: getTagText('keywords', cpNS),
    created: getTagText('created', dctermsNS),
    modified: getTagText('modified', dctermsNS)
  };
}

/**
 * 解析幻灯片背景色
 * @param slideXml 幻灯片XML字符串
 * @returns 背景颜色（十六进制）
 */
export function parseBackgroundColor(slideXml: string): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(slideXml, 'application/xml');

  const bgPr = getFirstChildByTagNS(doc.documentElement, 'bgPr',
    'http://schemas.openxmlformats.org/presentationml/2006/main');
  if (!bgPr) return '#ffffff';

  const solidFill = getFirstChildByTagNS(bgPr, 'solidFill',
    'http://schemas.openxmlformats.org/drawingml/2006/main');
  if (!solidFill) return '#ffffff';

  const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr',
    'http://schemas.openxmlformats.org/drawingml/2006/main');
  if (srgbClr?.getAttribute('val')) {
    return `#${srgbClr.getAttribute('val')}`;
  }

  return '#ffffff';
}

/**
 * 解析幻灯片尺寸
 * @param slideLayoutXml 幻灯片布局XML
 * @returns 尺寸对象（像素单位）
 */
export function parseSlideSize(slideLayoutXml: string): { width: number; height: number } {
  const parser = new DOMParser();
  const doc = parser.parseFromString(slideLayoutXml, 'application/xml');
  const NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';

  // 尝试从 slideLayout 中查找
  const sldSz = getFirstChildByTagNS(doc.documentElement, 'sldSz', NS) ||
                doc.getElementsByTagNameNS(NS, 'sldSz')[0];

  if (sldSz) {
    const cx = sldSz.getAttribute('cx');
    const cy = sldSz.getAttribute('cy');
    return {
      width: cx ? Math.round(parseInt(cx) * 96 / 914400) : 1280,
      height: cy ? Math.round(parseInt(cy) * 96 / 914400) : 720
    };
  }

  return { width: 1280, height: 720 };
}
