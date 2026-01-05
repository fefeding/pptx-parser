/**
 * 图片元素类
 */

import { BaseElement } from './BaseElement';
import { getFirstChildByTagNS, getAttrSafe, getBoolAttr } from '../utils';
import { NS } from '../constants';
import type { ParsedImageElement, RelsMap } from '../types-enhanced';

/**
 * 图片元素类
 */
export class ImageElement extends BaseElement {
  type = 'image' as const;

  /** 图片URL或Base64 */
  src: string;

  /** 图片关联ID */
  relId: string;

  /** MIME类型 */
  mimeType?: string;

  /** 替代文本 */
  altText?: string;

  /**
   * 从XML节点创建图片元素
   */
  static fromNode(node: Element, relsMap: RelsMap): ImageElement | null {
    try {
      const element = new ImageElement('', { x: 0, y: 0, width: 0, height: 0 }, '', '', {}, relsMap);

      // 解析ID和名称
      const nvPicPr = getFirstChildByTagNS(node, 'nvPicPr', NS.p);
      const cNvPr = nvPicPr ? getFirstChildByTagNS(nvPicPr, 'cNvPr', NS.p) : null;

      element.id = getAttrSafe(cNvPr, 'id', element.generateId());
      element.name = getAttrSafe(cNvPr, 'name', '');
      element.hidden = getBoolAttr(cNvPr, 'hidden');
      element.altText = getAttrSafe(cNvPr, 'descr', '');

      // 解析位置尺寸
      element.rect = element.parsePosition(node);

      // 解析图片关联ID
      const blipFill = getFirstChildByTagNS(node, 'blipFill', NS.p);
      if (blipFill) {
        const blip = getFirstChildByTagNS(blipFill, 'blip', NS.a);
        element.relId = blip?.getAttributeNS(NS.r, 'embed') || blip?.getAttribute('r:embed') || '';
      }

      // 从关系映射中获取图片路径
      if (element.relId && relsMap[element.relId]) {
        const target = relsMap[element.relId].target;
        element.src = target;
        // 从路径推断MIME类型
        if (target.endsWith('.png')) element.mimeType = 'image/png';
        else if (target.endsWith('.jpg') || target.endsWith('.jpeg')) element.mimeType = 'image/jpeg';
        else if (target.endsWith('.gif')) element.mimeType = 'image/gif';
      }

      element.content = {
        src: element.src,
        alt: element.altText
      };

      element.props = {
        imgId: element.relId,
        mimeType: element.mimeType
      };

      element.rawNode = node;

      return element;
    } catch (error) {
      console.error('Failed to parse image element:', error);
      return null;
    }
  }

  /**
   * 构造函数
   */
  constructor(
    id: string,
    rect: { x: number; y: number; width: number; height: number },
    src: string,
    relId: string,
    props: any = {},
    relsMap: Record<string, any> = {}
  ) {
    super(id, 'image', rect, props, relsMap);
    this.src = src;
    this.relId = relId;
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const style = this.getContainerStyle();
    const imgStyle = [
      `width: 100%`,
      `height: 100%`,
      `object-fit: contain`
    ].join('; ');

    return `<div style="${style}">
      <img src="${this.src}" alt="${this.altText || ''}" style="${imgStyle}" />
    </div>`;
  }

  /**
   * 转换为ParsedImageElement格式
   */
  toParsedElement(): ParsedImageElement {
    return {
      id: this.id,
      type: 'image',
      rect: this.rect,
      style: this.style,
      content: this.content,
      props: this.props,
      name: this.name,
      hidden: this.hidden,
      relId: this.relId,
      src: this.src,
      altText: this.altText,
      mimeType: this.mimeType,
      attrs: this.getAttributes(this.rawNode!),
      rawNode: this.rawNode
    };
  }
}
