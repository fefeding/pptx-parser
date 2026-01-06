/**
 * PPTX解析工具函数集合
 * 提供单位转换、XML解析、属性提取等核心工具
 */

import { NS, EMU_TO_PIXEL_RATIO, PIXEL_TO_EMU_RATIO } from './constants';

/**
 * EMU单位转像素（核心转换函数）
 * PPTX内部使用EMU (English Metric Unit)，1 EMU = 1/914400 英寸
 * 转换为像素时保留两位小数，避免精度溢出
 * @param emu EMU值（字符串或数字）
 * @returns 像素值
 */
export function emu2px(emu: string | number): number {
  const numEmu = typeof emu === 'string' ? parseInt(emu || '0', 10) : emu;
  return Math.round(numEmu * EMU_TO_PIXEL_RATIO * 100) / 100;
}

/**
 * 像素转EMU单位（用于序列化）
 * @param px 像素值
 * @returns EMU值
 */
export function px2emu(px: number): number {
  return Math.round(px * PIXEL_TO_EMU_RATIO);
}

/**
 * 提取XML节点的所有属性为标准对象
 * 支持命名空间属性（如 r:id）
 * @param node DOM元素节点
 * @returns 属性对象
 */
export function getAttrs(node: Element): Record<string, string> {
  const attrs: Record<string, string> = {};
  if (!node?.attributes) return attrs;

  Array.from(node.attributes).forEach(attr => {
    attrs[attr.name] = attr.value;
  });

  return attrs;
}

/**
 * 使用命名空间获取子元素（核心工具）
 * PPTX XML必须使用命名空间查询，避免错误匹配
 * @param parent 父元素
 * @param tagName 标签名（不含命名空间前缀）
 * @param namespaceURI 命名空间URI
 * @returns 子元素数组
 */
export function getChildrenByTagNS(
  parent: Element | null,
  tagName: string,
  namespaceURI: string
): Element[] {
  if (!parent) return [];
  return Array.from(parent.getElementsByTagNameNS(namespaceURI, tagName));
}

/**
 * 获取第一个匹配的子元素
 * @param parent 父元素
 * @param tagName 标签名
 * @param namespaceURI 命名空间URI
 * @returns 子元素或null
 */
export function getFirstChildByTagNS(
  parent: Element | null,
  tagName: string,
  namespaceURI: string
): Element | null {
  if (!parent) return null;
  return parent.getElementsByTagNameNS(namespaceURI, tagName)[0] || null;
}

/**
 * 解析文本内容（标准路径：a:p -> a:r -> a:t）
 * 支持多段落、多文本运行拼接，中英文混合文本自动合并
 * @param txBody 文本框节点（p:txBody）
 * @returns 完整文本字符串
 */
export function parseTextContent(txBody: Element): string {
  const paragraphs = getChildrenByTagNS(txBody, 'p', NS.a);
  const textParts: string[] = [];

  paragraphs.forEach(p => {
    const runs = getChildrenByTagNS(p, 'r', NS.a);
    runs.forEach(r => {
      const t = getFirstChildByTagNS(r, 't', NS.a);
      if (t?.textContent) {
        textParts.push(t.textContent);
      }
    });
  });

  return textParts.join('');
}

/**
 * 解析文本内容带样式（扩展版本）
 * @param txBody 文本框节点
 * @returns 文本运行数组，包含文本和样式
 */
export function parseTextWithStyle(txBody: Element): Array<{ text: string; style?: Record<string, unknown> }> {
  const paragraphs = getChildrenByTagNS(txBody, 'p', NS.a);
  const result: Array<{ text: string; style?: Record<string, unknown> }> = [];

  paragraphs.forEach(p => {
    const runs = getChildrenByTagNS(p, 'r', NS.a);
    runs.forEach(r => {
      const t = getFirstChildByTagNS(r, 't', NS.a);
      const rPr = getFirstChildByTagNS(r, 'rPr', NS.a);

      if (t?.textContent) {
        const style = rPr ? parseTextStyle(rPr) : undefined;
        result.push({ text: t.textContent, style });
      }
    });
  });

  return result;
}

/**
 * 解析文本样式属性（a:rPr节点）
 * @param rPr 文本运行属性节点
 * @returns 样式对象
 */
export function parseTextStyle(rPr: Element): Record<string, unknown> {
  const style: Record<string, unknown> = {};

  // 字体大小 (sz单位：1/100 pt)
  const sz = rPr.getAttribute('sz');
  if (sz) {
    style.fontSize = parseInt(sz, 10) / 100;
  }

  // 字体名称
  const latin = getFirstChildByTagNS(rPr, 'latin', NS.a);
  if (latin?.getAttribute('typeface')) {
    style.fontFamily = latin.getAttribute('typeface');
  }

  // 加粗
  if (rPr.getAttribute('b') === '1' || rPr.getAttribute('b') === 'true') {
    style.fontWeight = 'bold';
  }

  // 斜体
  if (rPr.getAttribute('i') === '1' || rPr.getAttribute('i') === 'true') {
    style.fontStyle = 'italic';
  }

  // 下划线
  if (rPr.getAttribute('u') === 'sng' || rPr.getAttribute('u') === '1') {
    style.textDecoration = 'underline';
  }

  // 删除线
  if (rPr.getAttribute('strike') === 'sngStrike' || rPr.getAttribute('strike') === '1') {
    style.textDecoration = style.textDecoration ? `${style.textDecoration} line-through` : 'line-through';
  }

  // 字体颜色
  const solidFill = getFirstChildByTagNS(rPr, 'solidFill', NS.a);
  if (solidFill) {
    const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
    if (srgbClr?.getAttribute('val')) {
      style.color = `#${srgbClr.getAttribute('val')}`;
    }
  }

  return style;
}

/**
 * 解析位置和尺寸（标准路径：a:xfrm -> a:off + a:ext）
 * @param spPr 形状属性节点（p:spPr）
 * @returns 位置尺寸对象（像素单位）
 */
export function parsePosition(spPr: Element): { x: number; y: number; width: number; height: number } {
  const xfrm = getFirstChildByTagNS(spPr, 'xfrm', NS.a);
  if (!xfrm) {
    return { x: 0, y: 0, width: 0, height: 0 };
  }

  const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
  const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

  const x = off ? emu2px(off.getAttribute('x') || '0') : 0;
  const y = off ? emu2px(off.getAttribute('y') || '0') : 0;
  const width = ext ? emu2px(ext.getAttribute('cx') || '0') : 0;
  const height = ext ? emu2px(ext.getAttribute('cy') || '0') : 0;

  return { x, y, width, height };
}

/**
 * 解析关系文件（rels文件）
 * @param relsXml 关系XML字符串
 * @returns 关系映射表 { relId: target }
 */
export function parseRels(relsXml: string): Record<string, { id: string; type: string; target: string }> {
  const parser = new DOMParser();
  const doc = parser.parseFromString(relsXml, 'application/xml');

  // 根元素是 <Relationships>，子元素是 <Relationship>
  // 使用不带命名空间前缀的方式查询，因为子元素继承父元素的默认命名空间
  const relationships = Array.from(doc.documentElement.children);

  const relsMap: Record<string, { id: string; type: string; target: string }> = {};

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

  const bgPr = getFirstChildByTagNS(doc.documentElement, 'bgPr', NS.p);
  if (!bgPr) return '#ffffff';

  const solidFill = getFirstChildByTagNS(bgPr, 'solidFill', NS.a);
  if (!solidFill) return '#ffffff';

  const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
  if (srgbClr?.getAttribute('val')) {
    return `#${srgbClr.getAttribute('val')}`;
  }

  return '#ffffff';
}

/**
 * 安全获取属性值（容错处理）
 * @param element DOM元素
 * @param attrName 属性名
 * @param defaultValue 默认值
 * @returns 属性值或默认值
 */
export function getAttrSafe(
  element: Element | null,
  attrName: string,
  defaultValue: string = ''
): string {
  if (!element) return defaultValue;
  return element.getAttribute(attrName) || defaultValue;
}

/**
 * 解析布尔属性（兼容多种格式）
 * @param element DOM元素
 * @param attrName 属性名
 * @returns 布尔值
 */
export function getBoolAttr(element: Element | null, attrName: string): boolean {
  if (!element) return false;
  const value = element.getAttribute(attrName);
  return value === '1' || value === 'true';
}

/**
 * 解析幻灯片尺寸
 * @param slideLayoutXml 幻灯片布局XML
 * @returns 尺寸对象（像素单位）
 */
export function parseSlideSize(slideLayoutXml: string): { width: number; height: number } {
  const parser = new DOMParser();
  const doc = parser.parseFromString(slideLayoutXml, 'application/xml');

  // 尝试从 slideLayout 中查找
  const sldSz = getFirstChildByTagNS(doc.documentElement, 'sldSz', NS.p) ||
                doc.getElementsByTagNameNS(NS.p, 'sldSz')[0];

  if (sldSz) {
    const cx = sldSz.getAttribute('cx');
    const cy = sldSz.getAttribute('cy');
    return {
      width: cx ? emu2px(cx) : 1280,
      height: cy ? emu2px(cy) : 720
    };
  }

  return { width: 1280, height: 720 };
}

/**
 * 生成唯一ID
 * @param prefix 前缀
 * @returns 唯一ID
 */
export function generateId(prefix: string = 'ppt-node'): string {
  return `${prefix}-${Date.now()}-${Math.floor(Math.random() * 10000)}`;
}

/**
 * 日志工具（避免生产环境输出）
 * @param level 日志级别
 * @param message 消息
 * @param data 附加数据
 */
export function log(level: 'info' | 'warn' | 'error', message: string, data?: unknown): void {
    // @ts-ignore
    const showLog = (typeof process !== 'undefined' && process?.env?.NODE_ENV === 'development') || (import.meta?.env?.NODE_ENV === 'development');
  if (showLog) {
    const prefix = `[pptx-parser ${level.toUpperCase()}]`;
    if (data !== undefined) {
      console[level](prefix, message, data);
    } else {
      console[level](prefix, message);
    }
  }
}
