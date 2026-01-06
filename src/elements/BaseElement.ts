/**
 * 基础元素类
 * 所有PPTX元素的基类，提供公共逻辑和默认toHTML实现
 */

import { emu2px, getAttrs, getFirstChildByTagNS, parsePosition, parseTextWithStyle, getAttrSafe, getBoolAttr } from '../utils';
import { NS } from '../constants';
import type { PptRect, PptStyle, Position } from '../types';

/**
 * 基础元素类
 */
export abstract class BaseElement {
  /** 元素ID */
  id: string;

  /** 元素类型 */
  abstract type: string;

  /** 位置和尺寸 */
  rect: PptRect;

  /** 样式 */
  style: PptStyle;

  /** 内容（兼容原库） */
  content: any;

  /** 属性（兼容原库） */
  props: any;

  /** 元素名称 */
  name?: string;

  /** 是否隐藏 */
  hidden?: boolean;

  /** 原始XML节点 */
  rawNode?: Element;

  /** 是否为占位符 */
  isPlaceholder?: boolean;

  /** 关联关系映射表 */
  protected relsMap: Record<string, any>;

  constructor(
    id: string,
    type: string,
    rect: PptRect,
    content: any = {},
    props: any = {},
    relsMap: Record<string, any> = {}
  ) {
    this.id = id;
    this.rect = rect;
    this.style = {
      fontSize: 14,
      color: '#333333',
      fontWeight: 'normal',
      textAlign: 'left',
      backgroundColor: 'transparent',
      borderColor: '#000000',
      borderWidth: 1
    };
    this.content = content;
    this.props = props;
    this.relsMap = relsMap;
    this.name = '';
    this.hidden = false;
  }

  /**
   * 转换为HTML字符串
   */
  abstract toHTML(): string;

  /**
   * 获取容器样式字符串
   */
  protected getContainerStyle(): string {
    const { x, y, width, height } = this.rect;
    const style = [
      `position: absolute`,
      `left: ${x}px`,
      `top: ${y}px`,
      `width: ${width}px`,
      `height: ${height}px`
    ];

    if (this.hidden) {
      style.push('display: none');
    }

    return style.join('; ');
  }

  /**
   * 获取用于调试的 data-* 属性
   * 子类可以重写此方法以添加更多属性
   */
  protected getDataAttributes(): Record<string, string> {
    const attrs: Record<string, string> = {};
    if (this.id) {
      attrs['data-id'] = this.id;
    }
    // type 是抽象属性，子类必须实现
    attrs['data-type'] = this.type;
    return attrs;
  }

  /**
   * 格式化 data-* 属性为字符串，用于 HTML 输出
   */
  protected formatDataAttributes(): string {
    const attrs = this.getDataAttributes();
    const attrString = Object.entries(attrs)
      .map(([key, value]) => `${key}="${value}"`)
      .join(' ');
    return attrString;
  }

  /**
   * 解析位置尺寸
   */
  protected parsePosition(node: Element, tag = 'spPr', namespace = NS.p): Position {
    const spPr = getFirstChildByTagNS(node, tag, namespace);
    return spPr ? parsePosition(spPr) : { x: 0, y: 0, width: 0, height: 0 };
  }

  /**
   * 解析ID和名称
   */
  protected parseIdAndName(node: Element, nonVisualTag: string, namespace = NS.p): { id: string; name: string; hidden: boolean } {
    const nvPr = getFirstChildByTagNS(node, nonVisualTag, namespace);
    const cNvPr = nvPr ? getFirstChildByTagNS(nvPr, 'cNvPr', namespace) : null;

    const id = getAttrSafe(cNvPr, 'id', this.generateId());
    const name = getAttrSafe(cNvPr, 'name', '');
    const hidden = getBoolAttr(cNvPr, 'hidden');

    return { id, name, hidden };
  }

  /**
   * 生成ID
   */
  protected generateId(): string {
    return `el_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * 获取属性对象
   */
  protected getAttributes(node: Element): Record<string, string> {
    return getAttrs(node);
  }
}
