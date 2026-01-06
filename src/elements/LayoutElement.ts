/**
 * 布局元素类
 * 表示幻灯片版式，包含占位符定义和样式
 */

import { BaseElement } from './BaseElement';
import type { SlideLayoutResult } from '../core/types';
import type { TextRun } from './ShapeElement';

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

  /** 文本样式 */
  textStyles?: any;

  /** 背景样式 */
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string };

  constructor(
    id: string,
    name?: string,
    placeholders: PlaceholderElement[] = [],
    props: any = {}
  ) {
    super(id, 'layout', { x: 0, y: 0, width: 960, height: 540 }, {}, props, {});
    this.name = name;
    this.placeholders = placeholders;
    this.textStyles = props.textStyles;
    this.background = props.background;
  }

  /**
   * 从 SlideParseResult 创建 LayoutElement
   */
  static fromResult(result: SlideLayoutResult): LayoutElement {
    const placeholders = (result.placeholders || []).map(ph => {
      return new PlaceholderElement(
        ph.id,
        ph.type,
        ph.rect,
        { idx: ph.idx, name: ph.name, rawNode: ph.rawNode }
      );
    });

    return new LayoutElement(
      result.id,
      result.name,
      placeholders,
      {
        textStyles: result.textStyles,
        background: result.background,
        relsMap: result.relsMap,
        colorMap: result.colorMap
      }
    );
  }

  toHTML(): string {
    const background = this.getBackgroundStyle();
    const style = [
      this.getContainerStyle(),
      background
    ].join('; ');

    const placeholdersHTML = this.placeholders
      .map(ph => ph.toHTML())
      .join('\n');

    return `<div class="ppt-layout" style="${style}" data-layout-id="${this.id}" data-layout-name="${this.name || ''}">
${placeholdersHTML}
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
      return `background-image: url('${this.background.relId}'); background-size: cover;`;
    }

    return 'background-color: transparent;';
  }
}
