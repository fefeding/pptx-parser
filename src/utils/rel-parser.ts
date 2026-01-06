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
