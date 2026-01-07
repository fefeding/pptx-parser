/**
 * 母版元素类
 * 表示幻灯片母版，包含所有幻灯片共享的元素和样式
 */

import { BaseElement } from './BaseElement';
import { PlaceholderElement } from './LayoutElement';
import { createElementFromData } from './index';
import type { MasterSlideResult } from '../core/types';

/**
 * 母版元素类
 */
export class MasterElement extends BaseElement {
  type: 'master' = 'master';

  /** 母版文件名 */
  masterId?: string;

  /** 母版元素（页脚、页码等） */
  elements: BaseElement[];

  /** 占位符定义 */
  placeholders: PlaceholderElement[];

  /** 文本样式 */
  textStyles?: any;

  /** 背景样式 */
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string };

  /** 颜色映射 */
  colorMap: Record<string, string>;

  /** 媒体资源映射表（relId -> base64 URL） */
  mediaMap?: Map<string, string>;

  constructor(
    id: string,
    elements: BaseElement[] = [],
    placeholders: PlaceholderElement[] = [],
    props: any = {}
  ) {
    super(id, 'master', { x: 0, y: 0, width: 960, height: 540 }, {}, props, {});
    this.masterId = props.masterId;
    this.elements = elements;
    this.placeholders = placeholders;
    this.textStyles = props.textStyles;
    this.background = props.background;
    this.colorMap = props.colorMap || {};
    this.mediaMap = props.mediaMap;
  }

  /**
   * 从 MasterSlideResult 创建 MasterElement
   */
  static fromResult(result: MasterSlideResult, mediaMap?: Map<string, string>): MasterElement {
    // 解析占位符
    const placeholders = (result.placeholders || []).map((ph: any) => {
      const phEl = new PlaceholderElement(
        ph.id,
        ph.type || 'other',
        ph.rect || { x: 0, y: 0, width: 100, height: 50 },
        { idx: ph.idx, name: ph.name }
      );
      return phEl;
    });

    // 将 result.elements 转换为 BaseElement 实例
    const elements: BaseElement[] = (result.elements || []).map(elementData =>
      createElementFromData(elementData, result.relsMap, mediaMap)
    ).filter((el): el is BaseElement => el !== null);

    return new MasterElement(
      result.id,
      elements,
      placeholders,
      {
        masterId: result.masterId,
        textStyles: result.textStyles,
        background: result.background,
        colorMap: result.colorMap,
        relsMap: result.relsMap,
        mediaMap
      }
    );
  }

  toHTML(): string {
    const background = this.getBackgroundStyle();
    const style = [
      `width: 100%`,
      `height: 100%`,
      `position: absolute`,
      `top: 0`,
      `left: 0`,
      background,
      `pointer-events: none`, // 母版元素不响应鼠标事件
      `z-index: 0`
    ].join('; ');

    const elementsHTML = this.elements
      .map(el => el.toHTML())
      .join('\n');

    const placeholdersHTML = this.placeholders
      .map(ph => ph.toHTML())
      .join('\n');

    return `<div class="ppt-master" style="${style}" data-master-id="${this.id}">
${placeholdersHTML}
${elementsHTML}
</div>`;
  }

  /**
   * 获取背景样式
   */
  private getBackgroundStyle(): string {
    if (!this.background) {
      return 'background-color: transparent;';
    }

    if (this.background.type === 'color' && this.background.value) {
      return `background-color: ${this.background.value};`;
    }

    if (this.background.type === 'image' && this.background.relId) {
      // 通过 mediaMap 解析 relId 到实际的 base64 URL
      const imageUrl = this.mediaMap ? this.mediaMap.get(this.background.relId) : this.background.relId;
      if (imageUrl) {
        return `background-image: url('${imageUrl}'); background-size: cover;`;
      }
    }

    return 'background-color: transparent;';
  }

  /**
   * 获取占位符样式
   * @param placeholderType 占位符类型
   */
  getPlaceholderStyle(placeholderType: 'title' | 'body' | 'other'): any {
    if (!this.textStyles) {
      return null;
    }

    switch (placeholderType) {
      case 'title':
        return this.parseParagraphProperties(this.textStyles.titleParaPr);

      case 'body':
        return this.parseParagraphProperties(this.textStyles.bodyPr);

      default:
        return this.parseParagraphProperties(this.textStyles.otherPr);
    }
  }

  /**
   * 解析段落属性
   */
  private parseParagraphProperties(paraPr: any): any {
    if (!paraPr) return null;

    // TODO: 实现段落属性解析
    return {};
  }
}
