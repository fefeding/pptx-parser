/**
 * 形状元素类
 * 支持文本框、自定义形状、占位符等
 */

import { BaseElement } from './BaseElement';
import { getFirstChildByTagNS, parseTextContent } from '../utils';
import { NS } from '../constants';
import type { ParsedShapeElement, RelsMap } from '../types-enhanced';

/**
 * 形状元素类
 */
export class ShapeElement extends BaseElement {
  type: 'shape' | 'text' = 'shape';

  /** 形状类型（矩形、圆形等） */
  shapeType?: string;

  /** 文本内容 */
  text?: string;

  /** 文本样式（运行级别） */
  textStyle?: Array<{ text: string; style: any }>;

  /** 是否占位符 */
  isPlaceholder?: boolean;

  /** 占位符类型 */
  placeholderType?: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';

  /**
   * 从XML节点创建形状元素
   */
  static fromNode(node: Element, relsMap: RelsMap): ShapeElement | null {
    try {
      const element = new ShapeElement('', 'shape', { x: 0, y: 0, width: 0, height: 0 }, {}, relsMap);

      // 解析ID和名称
      const { id, name, hidden } = element.parseIdAndName(node, 'nvSpPr');
      element.id = id;
      element.name = name;
      element.hidden = hidden;

      // 检查是否是占位符
      const nvSpPr = getFirstChildByTagNS(node, 'nvSpPr', NS.p);
      const nvPr = nvSpPr ? getFirstChildByTagNS(nvSpPr, 'nvPr', NS.p) : null;
      const ph = nvPr ? getFirstChildByTagNS(nvPr, 'ph', NS.p) : null;
      element.isPlaceholder = !!ph;
      element.placeholderType = ph?.getAttribute('type') || undefined;

      // 解析位置尺寸
      element.rect = element.parsePosition(node);

      // 解析文本内容
      const txBody = getFirstChildByTagNS(node, 'txBody', NS.p);
      if (txBody) {
        element.text = parseTextContent(txBody);
        if (element.text) {
          element.type = 'text';
          element.content = element.text;
          element.style.color = '#000000';
          element.style.fontSize = 18;
        }
      }

      element.shapeType = 'rectangle';
      element.rawNode = node;

      return element;
    } catch (error) {
      console.error('Failed to parse shape element:', error);
      return null;
    }
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const style = this.getContainerStyle();
    const textStyle = this.getTextStyle();

    if (this.type === 'text' && this.text) {
      // 文本框
      return `<div style="${style}${textStyle}">${this.escapeHtml(this.text)}</div>`;
    } else {
      // 形状（矩形、圆形等）
      const shapeStyle = this.getShapeStyle();
      return `<div style="${style}${shapeStyle}"></div>`;
    }
  }

  /**
   * 获取文本样式
   */
  private getTextStyle(): string {
    const styles = [
      `display: flex`,
      `align-items: center`,
      `justify-content: ${this.style.textAlign === 'center' ? 'center' : this.style.textAlign === 'right' ? 'flex-end' : 'flex-start'}`,
      `padding: 10px`,
      `color: ${this.style.color}`,
      `font-size: ${this.style.fontSize}px`
    ];

    if (this.style.fontWeight === 'bold') {
      styles.push(`font-weight: bold`);
    }

    if (this.style.backgroundColor && this.style.backgroundColor !== 'transparent') {
      styles.push(`background-color: ${this.style.backgroundColor}`);
    }

    if (this.style.borderWidth && this.style.borderWidth > 0) {
      styles.push(`border: ${this.style.borderWidth}px solid ${this.style.borderColor}`);
    }

    return styles.join('; ');
  }

  /**
   * 获取形状样式
   */
  private getShapeStyle(): string {
    const styles = [
      `background-color: ${this.style.backgroundColor || '#ffffff'}`,
      `border: ${this.style.borderWidth}px solid ${this.style.borderColor}`
    ];

    return styles.join('; ');
  }

  /**
   * HTML转义
   */
  private escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  /**
   * 转换为ParsedShapeElement格式
   */
  toParsedElement(): ParsedShapeElement {
    return {
      id: this.id,
      type: this.type,
      rect: this.rect,
      style: this.style,
      content: this.content,
      props: this.props,
      name: this.name,
      hidden: this.hidden,
      text: this.text,
      textStyle: this.textStyle,
      attrs: this.getAttributes(this.rawNode!),
      rawNode: this.rawNode
    };
  }
}
