/**
 * 分组元素类
 * 支持旋转、翻转、缩放等变换
 * 对齐 PPTXjs 的 Group 解析能力
 */

import { BaseElement } from './BaseElement';
import type { BaseElement as BaseElementType } from './BaseElement';
import { ShapeElement } from './ShapeElement';
import { ImageElement } from './ImageElement';
import { OleElement } from './OleElement';
import { ChartElement } from './ChartElement';
import { DiagramElement } from './DiagramElement';
import { TableElement } from './TableElement';
import { getFirstChildByTagNS, getAttrSafe, getBoolAttr, emu2px } from '../utils';
import { NS } from '../constants';
import type { ParsedGroupElement, RelsMap } from '../types';

/**
 * 分组元素类
 */
export class GroupElement extends BaseElement {
  type = 'group' as const;

  /** 子元素列表 */
  children: BaseElementType[] = [];

  /** 旋转角度（度） */
  rotation?: number;

  /** 是否水平翻转 */
  flipH?: boolean;

  /** 是否垂直翻转 */
  flipV?: boolean;

  /** 子偏移量 */
  childOffset?: {
    x: number;
    y: number;
  };

  /**
   * 从XML节点创建分组元素
   */
  static fromNode(node: Element, relsMap: RelsMap): GroupElement | null {
    try {
      const element = new GroupElement(`group_${Date.now()}`, 'group', { x: 0, y: 0, width: 0, height: 0 }, [], {}, relsMap);

      // 解析分组位置尺寸和变换
      const grpSpPr = getFirstChildByTagNS(node, 'grpSpPr', NS.p);
      element.parseGroupProperties(grpSpPr, node);

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
          childElement = element.parseGraphicFrame(child, relsMap);
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
   * 解析图形框
   */
  private parseGraphicFrame(node: Element, relsMap: RelsMap): BaseElementType | null {
    const graphic = getFirstChildByTagNS(node, 'graphic', NS.a);
    const graphicData = graphic ? getFirstChildByTagNS(graphic, 'graphicData', NS.a) : null;

    if (!graphicData) {
      return null;
    }

    const uri = graphicData.getAttribute('uri') || '';

    // 判断类型，直接使用已导入的类
    if (uri.includes('oleObject')) {
      return OleElement.fromNode(node, relsMap);
    } else if (uri.includes('chart')) {
      return ChartElement.fromNode(node, relsMap);
    } else if (uri.includes('diagram')) {
      return DiagramElement.fromNode(node, relsMap);
    } else if (uri.includes('table')) {
      return TableElement.fromNode(node, relsMap);
    }

    return null;
  }

  /**
   * 解析分组属性
   */
  private parseGroupProperties(grpSpPr: Element | null, node: Element): void {
    if (!grpSpPr) return;

    const xfrm = getFirstChildByTagNS(grpSpPr, 'xfrm', NS.a);
    if (xfrm) {
      const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
      const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);
      const chOff = getFirstChildByTagNS(xfrm, 'chOff', NS.a);
      const chExt = getFirstChildByTagNS(xfrm, 'chExt', NS.a);

      // 子元素偏移
      if (chOff) {
        const chx = parseInt(chOff.getAttribute('x') || '0');
        const chy = parseInt(chOff.getAttribute('y') || '0');
        this.childOffset = {
          x: emu2px(chx),
          y: emu2px(chy)
        };
      }

      // 子元素尺寸
      if (ext && chExt) {
        const cx = parseInt(ext.getAttribute('cx') || '0');
        const cy = parseInt(ext.getAttribute('cy') || '0');
        const chcx = parseInt(chExt.getAttribute('cx') || '0');
        const chcy = parseInt(chExt.getAttribute('cy') || '0');

        // 计算实际偏移
        if (!this.childOffset) {
          this.childOffset = { x: 0, y: 0 };
        }
        // 获取 chx 和 chy（如果 chOff 存在）
        let chx = 0;
        let chy = 0;
        if (chOff) {
          chx = parseInt(chOff.getAttribute('x') || '0');
          chy = parseInt(chOff.getAttribute('y') || '0');
        }
        // 检查 off 是否存在
        if (off) {
          this.childOffset.x += emu2px(parseInt(off.getAttribute('x') || '0') - chx);
          this.childOffset.y += emu2px(parseInt(off.getAttribute('y') || '0') - chy);
        }

        this.rect.x = emu2px(chcx);
        this.rect.y = emu2px(chcy);
        this.rect.width = emu2px(cx - chcx);
        this.rect.height = emu2px(cy - chcy);
      }

      // 旋转
      const rot = xfrm.getAttribute('rot');
      if (rot !== null) {
        this.rotation = parseInt(rot) / 60000; // 60000 EMU = 90度
      }

      // 翻转
      this.flipH = xfrm.getAttribute('flipH') === '1';
      this.flipV = xfrm.getAttribute('flipV') === '1';
    }
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const style = this.getGroupStyle();
    const dataAttrs = this.formatDataAttributes();

    const childrenHTML = this.children
      .map(child => child.toHTML())
      .join('\n');

    return `<div class="ppt-group" ${dataAttrs} style="${style}">
${childrenHTML}
    </div>`;
  }

  /**
   * 获取分组样式
   */
  private getGroupStyle(): string {
    const styles = [
      `position: absolute`,
      `left: ${this.rect.x}px`,
      `top: ${this.rect.y}px`,
      `width: ${this.rect.width}px`,
      `height: ${this.rect.height}px`
    ];

    // 应用旋转变换
    if (this.rotation !== undefined && this.rotation !== 0) {
      styles.push(`transform: rotate(${this.rotation}deg)`);
      styles.push(`transform-origin: center center`);
    }

    // 应用翻转
    if (this.flipH && this.flipV) {
      styles.push(`transform: rotate(${this.rotation || 0}deg) scale(-1, -1)`);
      styles.push(`transform-origin: center center`);
    } else if (this.flipH) {
      styles.push(`transform: rotate(${this.rotation || 0}deg) scale(-1, 1)`);
      styles.push(`transform-origin: center center`);
    } else if (this.flipV) {
      styles.push(`transform: rotate(${this.rotation || 0}deg) scale(1, -1)`);
      styles.push(`transform-origin: center center`);
    }

    // 应用子元素偏移
    if (this.childOffset) {
      styles.push(`margin-left: ${this.childOffset.x}px`);
      styles.push(`margin-top: ${this.childOffset.y}px`);
    }

    if (this.hidden) {
      styles.push('display: none');
    }

    return styles.join('; ');
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
      props: {
        rotation: this.rotation,
        flipH: this.flipH,
        flipV: this.flipV,
        childOffset: this.childOffset
      },
      name: this.name,
      hidden: this.hidden,
      children: this.children.map(child => {
        if (child instanceof ShapeElement) return child.toParsedElement();
        if (child instanceof ImageElement) return child.toParsedElement();
        if (child instanceof OleElement) return child.toParsedElement();
        if (child instanceof ChartElement) return child.toParsedElement();
        if (child instanceof DiagramElement) return child.toParsedElement();
        if (child instanceof GroupElement) return child.toParsedElement();
        return child as any;
      }),
      attrs: this.getAttributes(this.rawNode!),
      rawNode: this.rawNode
    };
  }
}
