/**
 * 分组元素类
 */

import { BaseElement } from './BaseElement';
import { getFirstChildByTagNS } from '../utils';
import { NS } from '../constants';
import type { ParsedGroupElement, RelsMap } from '../types-enhanced';
import { ShapeElement } from './ShapeElement';
import { ImageElement } from './ImageElement';
import { OleElement } from './OleElement';
import { ChartElement } from './ChartElement';
import type { BaseElement as BaseElementType } from './BaseElement';

/**
 * 分组元素类
 */
export class GroupElement extends BaseElement {
  type = 'group' as const;

  /** 子元素列表 */
  children: BaseElementType[] = [];

  /**
   * 从XML节点创建分组元素
   */
  static fromNode(node: Element, relsMap: RelsMap): GroupElement | null {
    try {
      const element = new GroupElement(`group_${Date.now()}`, { x: 0, y: 0, width: 0, height: 0 }, [], {}, relsMap);

      // 解析分组位置尺寸
      const grpSpPr = getFirstChildByTagNS(node, 'grpSpPr', NS.p);
      if (grpSpPr) {
        const xfrm = getFirstChildByTagNS(grpSpPr, 'xfrm', NS.a);
        if (xfrm) {
          const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
          const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

          if (off && ext) {
            element.rect.x = parseInt(off.getAttribute('x') || '0') / 914400;
            element.rect.y = parseInt(off.getAttribute('y') || '0') / 914400;
            element.rect.width = parseInt(ext.getAttribute('cx') || '0') / 914400;
            element.rect.height = parseInt(ext.getAttribute('cy') || '0') / 914400;
          }
        }
      }

      // 递归解析分组内的子元素
      Array.from(node.children).forEach((child) => {
        if (child.nodeType !== 1) return;

        const tagName = child.tagName;
        let childElement: BaseElementType | null = null;

        if (tagName === 'p:sp' || tagName === 'sp') {
          childElement = ShapeElement.fromNode(child, relsMap);
        } else if (tagName === 'p:pic' || tagName === 'pic') {
          childElement = ImageElement.fromNode(child, relsMap);
        } else if (tagName === 'p:graphicFrame' || tagName === 'graphicFrame') {
          // 判断是OLE还是图表
          const graphic = getFirstChildByTagNS(child, 'graphic', NS.a);
          const graphicData = graphic ? getFirstChildByTagNS(graphic, 'graphicData', NS.a) : null;

          if (graphicData) {
            const uri = getAttrSafe(graphicData, 'uri', '');

            if (uri.includes('oleObject')) {
              childElement = OleElement.fromNode(child, relsMap);
            } else if (uri.includes('chart')) {
              childElement = ChartElement.fromNode(child, relsMap);
            }
          }
        } else if (tagName === 'p:grpSp' || tagName === 'grpSp') {
          // 递归解析嵌套分组
          childElement = GroupElement.fromNode(child, relsMap);
        }

        if (childElement) {
          element.children.push(childElement);
        }
      });

      element.content = {};
      element.props = {};
      element.name = '分组';
      element.rawNode = node;

      return element;
    } catch (error) {
      console.error('Failed to parse group element:', error);
      return null;
    }
  }

  /**
   * 辅助函数：获取属性
   */
  private static getAttrSafe(node: Element | null, attr: string, defaultValue: string = ''): string {
    return node?.getAttribute(attr) || defaultValue;
  }

  /**
   * 构造函数
   */
  constructor(
    id: string,
    rect: { x: number; y: number; width: number; height: number },
    children: BaseElementType[],
    props: any = {},
    relsMap: Record<string, any> = {}
  ) {
    super(id, 'group', rect, props, relsMap);
    this.children = children;
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const style = this.getContainerStyle();

    const childrenHTML = this.children
      .map(child => child.toHTML())
      .join('\n');

    return `<div style="${style}">
${childrenHTML}
    </div>`;
  }

  /**
   * 转换为ParsedGroupElement格式
   */
  toParsedElement(): ParsedGroupElement {
    return {
      id: this.id,
      type: 'group',
      rect: this.rect,
      style: this.style,
      content: this.content,
      props: this.props,
      name: this.name,
      hidden: this.hidden,
      children: this.children.map(child => {
        if (child instanceof ShapeElement) return child.toParsedElement();
        if (child instanceof ImageElement) return child.toParsedElement();
        if (child instanceof OleElement) return child.toParsedElement();
        if (child instanceof ChartElement) return child.toParsedElement();
        if (child instanceof GroupElement) return child.toParsedElement();
        return child as any;
      }),
      attrs: this.getAttributes(this.rawNode!),
      rawNode: this.rawNode
    };
  }
}
