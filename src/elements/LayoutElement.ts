/**
 * 布局元素类
 * 表示幻灯片版式，包含占位符定义和样式
 */

import { BaseElement } from './BaseElement';
import type { SlideLayoutResult } from '../core/types';
import type { TextRun } from './ShapeElement';
import { createElementFromData } from './element-factory';

/**
 * 占位符元素
 */
export class PlaceholderElement extends BaseElement {
  type: 'placeholder' = 'placeholder';

  /** 占位符类型 */
  placeholderType: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';

  /** 占位符索引 */
  idx?: number;

  /** 占位符名称 */
  name?: string;

  /** 文本内容 */
  text?: string;

  /** 文本样式 */
  textStyle?: TextRun[];

  constructor(
    id: string,
    placeholderType: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other',
    rect: { x: number; y: number; width: number; height: number },
    props: any = {}
  ) {
    super(id, 'placeholder', rect, {}, props, {});
    this.placeholderType = placeholderType;
    this.isPlaceholder = true;
    this.idx = props.idx;
    this.name = props.name;
    this.text = props.text;
    this.textStyle = props.textStyle;
  }

  toHTML(): string {
    const containerStyle = [
      this.getContainerStyle(),
      `pointer-events: none`, // 占位符不响应鼠标事件
      `user-select: none`
    ].join('; ');

    const content = this.getPlaceholderContent();

    return `<div class="ppt-placeholder" data-placeholder-type="${this.placeholderType}" style="${containerStyle}">
${content}
</div>`;
  }

  /**
   * 获取占位符内容
   */
  private getPlaceholderContent(): string {
    switch (this.placeholderType) {
      case 'title':
        return this.isContentSet ? this.renderTextContent() : '<span class="placeholder-label">点击添加标题</span>';

      case 'body':
        return this.isContentSet ? this.renderTextContent() : '<span class="placeholder-label">点击添加文本</span>';

      case 'slideNumber':
        return '<span class="slide-number"></span>';

      case 'footer':
        return this.text || '<span class="footer-text"></span>';

      case 'dateTime':
        return this.text || '<span class="date-time"></span>';

      default:
        return this.text || '';
    }
  }

  /**
   * 渲染文本内容
   */
  private renderTextContent(): string {
    if (!this.text) return '';

    // 如果有文本样式，应用样式
    if (this.textStyle && this.textStyle.length > 0) {
      return this.textStyle.map(run => {
        const style = this.getTextRunStyle(run);
        return `<span style="${style}">${run.text}</span>`;
      }).join('');
    }

    return `<span style="color: inherit;">${this.text}</span>`;
  }

  /**
   * 检查内容是否已设置
   */
  private get isContentSet(): boolean {
    return !!(this.text && this.text.trim());
  }

  /**
   * 获取文本运行样式
   */
  private getTextRunStyle(run: TextRun): string {
    const styles: string[] = [];

    if (run.fontSize) {
      styles.push(`font-size: ${run.fontSize}px`);
    }
    if (run.fontFamily) {
      styles.push(`font-family: ${run.fontFamily}`);
    }
    if (run.color) {
      styles.push(`color: ${run.color}`);
    }
    if (run.bold) {
      styles.push('font-weight: bold');
    }
    if (run.italic) {
      styles.push('font-style: italic');
    }
    if (run.underline) {
      styles.push('text-decoration: underline');
    }

    return styles.join('; ');
  }
}

/**
 * 布局元素类
 */
export class LayoutElement extends BaseElement {
  type: 'layout' = 'layout';

  /** 布局名称 */
  name?: string;

  /** 占位符列表 */
  placeholders: PlaceholderElement[];

  /** 实际元素（图片、形状等，非占位符） */
  elements: BaseElement[];

  /** 文本样式 */
  textStyles?: any;

  /** 背景样式 */
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string };

  /** 关联关系映射表 */
  relsMap: Record<string, any>;

  /** 媒体资源映射表（relId -> base64 URL） */
  mediaMap?: Map<string, string>;

  constructor(
    id: string,
    name?: string,
    placeholders: PlaceholderElement[] = [],
    elements: BaseElement[] = [],
    props: any = {}
  ) {
    super(id, 'layout', { x: 0, y: 0, width: 960, height: 540 }, {}, props, {});
    this.name = name;
    this.placeholders = placeholders;
    this.elements = elements;
    this.textStyles = props.textStyles;
    this.background = props.background;
    this.relsMap = props.relsMap || {};
    this.mediaMap = props.mediaMap;
  }

  /**
   * 从 SlideParseResult 创建 LayoutElement
   */
  static fromResult(result: SlideLayoutResult, mediaMap?: Map<string, string>): LayoutElement {
    const placeholders = (result.placeholders || []).map(ph => {
      return new PlaceholderElement(
        ph.id,
        ph.type,
        ph.rect,
        { idx: ph.idx, name: ph.name, rawNode: ph.rawNode }
      );
    });

    // 将解析的元素数据转换为 BaseElement 实例
    const elements = (result.elements || []).map((el: any) => {
      if (el instanceof BaseElement) {
        return el;
      }
      return createElementFromData(el, result.relsMap || {}, mediaMap);
    }).filter((el: any) => el !== null) as BaseElement[];

    return new LayoutElement(
      result.id,
      result.name,
      placeholders,
      elements,
      {
        textStyles: result.textStyles,
        background: result.background,
        relsMap: result.relsMap,
        colorMap: result.colorMap,
        mediaMap
      }
    );
  }

  toHTML(): string {
    const background = this.getBackgroundStyle();
    // 布局背景应该是全屏的，不受 rect 尺寸限制
    const style = [
      `position: absolute`,
      `left: 0`,
      `top: 0`,
      `width: 100%`,
      `height: 100%`,
      `pointer-events: none`, // 布局元素不响应鼠标事件
      `z-index: 1`, // 在 master 之上，slide 之下
      background
    ].join('; ');

    const placeholdersHTML = this.placeholders
      .map(ph => ph.toHTML())
      .join('\n');

    // 渲染实际元素（图片、形状等）
    const elementsHTML = this.elements
      .map(el => el.toHTML())
      .join('\n');

    return `<div class="ppt-layout" style="${style}" data-layout-id="${this.id}" data-layout-name="${this.name || ''}">
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
}
