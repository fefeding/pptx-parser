/**
 * OLE对象元素类
 */

import { BaseElement } from './BaseElement';
import { getFirstChildByTagNS, getAttrSafe, getBoolAttr } from '../utils';
import { NS } from '../constants';
import type { ParsedOleElement, RelsMap } from '../types-enhanced';
import { ImageElement } from './ImageElement';

/**
 * OLE对象元素类
 */
export class OleElement extends BaseElement {
  type = 'ole' as const;

  /** OLE对象类型标识符 */
  progId?: string;

  /** 关联ID */
  relId: string;

  /** 对象名称 */
  oleName?: string;

  /** 是否包含降级图片 */
  hasFallback?: boolean;

  /**
   * 从XML节点创建OLE元素
   */
  static fromNode(node: Element, relsMap: RelsMap): OleElement | null {
    try {
      // 查找 oleObj 节点
      const oleObj = getFirstChildByTagNS(node, 'oleObj', NS.p);

      if (!oleObj) {
        // 检查是否有降级图片（Fallback）
        const fallback = getFirstChildByTagNS(node, 'Fallback', NS.mc);
        if (fallback) {
          const pic = getFirstChildByTagNS(fallback, 'pic', NS.p);
          if (pic) {
            // 转换图片为OLE对象
            const imageElement = ImageElement.fromNode(pic, relsMap);
            if (imageElement) {
              return new OleElement(
                imageElement.id,
                imageElement.rect,
                '',
                imageElement.relId,
                {
                  ...imageElement.props,
                  isOle: true,
                  hasFallback: true
                },
                relsMap
              );
            }
          }
        }
        return null;
      }

      const progId = oleObj.getAttribute('progId') || '';
      const relId = oleObj.getAttributeNS(NS.r, 'id') || oleObj.getAttribute('r:id') || '';

      const element = new OleElement('', { x: 0, y: 0, width: 0, height: 0 }, progId, relId, {}, relsMap);

      // 解析ID和名称
      const nvGraphicFramePr = getFirstChildByTagNS(node, 'nvGraphicFramePr', NS.p);
      const cNvPr = nvGraphicFramePr ? getFirstChildByTagNS(nvGraphicFramePr, 'cNvPr', NS.p) : null;

      element.id = getAttrSafe(cNvPr, 'id', element.generateId());
      element.name = getAttrSafe(cNvPr, 'name', '');
      element.hidden = getBoolAttr(cNvPr, 'hidden');
      element.oleName = element.name;

      // 解析位置尺寸（使用xfrm而不是spPr）
      const xfrm = getFirstChildByTagNS(node, 'xfrm', NS.p);
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

      element.content = {};
      element.props = {
        progId,
        relId,
        hasFallback: false
      };

      element.rawNode = node;

      return element;
    } catch (error) {
      console.error('Failed to parse OLE element:', error);
      return null;
    }
  }

  /**
   * 构造函数
   */
  constructor(
    id: string,
    rect: { x: number; y: number; width: number; height: number },
    progId: string,
    relId: string,
    props: any = {},
    relsMap: Record<string, any> = {}
  ) {
    super(id, 'ole', rect, props, relsMap);
    this.progId = progId;
    this.relId = relId;
    this.oleName = name;
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const style = this.getContainerStyle();
    const innerStyle = [
      `width: 100%`,
      `height: 100%`,
      `display: flex`,
      `align-items: center`,
      `justify-content: center`,
      `background-color: #f5f5f5`,
      `border: 1px dashed #ccc`,
      `color: #666`,
      `font-size: 12px`,
      `text-align: center`,
      `padding: 10px`
    ].join('; ');

    const label = this.progId || this.oleName || 'OLE Object';

    return `<div style="${style}">
      <div style="${innerStyle}">
        ${label}
      </div>
    </div>`;
  }

  /**
   * 转换为ParsedOleElement格式
   */
  toParsedElement(): ParsedOleElement {
    return {
      id: this.id,
      type: 'ole',
      rect: this.rect,
      style: this.style,
      content: this.content,
      props: this.props,
      name: this.name,
      hidden: this.hidden,
      relId: this.relId,
      progId: this.progId,
      hasFallback: this.hasFallback,
      attrs: this.getAttributes(this.rawNode!),
      rawNode: this.rawNode
    };
  }
}
