/**
 * XML 解析工具
 * 提供 XML 节点查询和属性提取功能
 */

import { NS } from '../constants';

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
