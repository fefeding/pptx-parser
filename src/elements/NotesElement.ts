/**
 * 备注元素类
 * 表示备注页和备注母版
 */

import { BaseElement } from './BaseElement';
import { createElementFromData } from './element-factory';
import type { NotesSlideResult, NotesMasterResult } from '../core/types';

/**
 * 备注母版元素类
 */
export class NotesMasterElement extends BaseElement {
  type: 'notesMaster' = 'notesMaster';

  /** 备注母版元素 */
  elements: BaseElement[];

  /** 占位符定义 */
  placeholders: any[];

  /** 背景样式 */
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string };

  constructor(
    id: string,
    elements: BaseElement[] = [],
    placeholders: any[] = [],
    props: any = {}
  ) {
    super(id, 'notesMaster', { x: 0, y: 0, width: 800, height: 600 }, {}, props, {});
    this.elements = elements;
    this.placeholders = placeholders;
    this.background = props.background;
  }

  /**
   * 从 NotesMasterResult 创建 NotesMasterElement
   */
  static fromResult(result: NotesMasterResult): NotesMasterElement {
    // 将 elements 转换为 BaseElement 实例
    const elements: BaseElement[] = (result.elements || []).map(elementData => 
      createElementFromData(elementData, result.relsMap)
    ).filter((el): el is BaseElement => el !== null);

    return new NotesMasterElement(
      result.id,
      elements,
      result.placeholders || [],
      { background: result.background, relsMap: result.relsMap }
    );
  }

  toHTML(): string {
    const background = this.getBackgroundStyle();
    const style = [
      this.getContainerStyle(),
      background
    ].join('; ');

    const elementsHTML = this.elements
      .map(el => el.toHTML())
      .join('\n');

    return `<div class="ppt-notes-master" style="${style}" data-notes-master-id="${this.id}">
${elementsHTML}
</div>`;
  }

  /**
   * 获取背景样式
   */
  private getBackgroundStyle(): string {
    if (!this.background) {
      return 'background-color: #ffffff;';
    }

    if (this.background.type === 'color' && this.background.value) {
      return `background-color: ${this.background.value};`;
    }

    if (this.background.type === 'image' && this.background.relId) {
      return `background-image: url('${this.background.relId}'); background-size: cover;`;
    }

    return 'background-color: #ffffff;';
  }
}

/**
 * 备注页元素类
 */
export class NotesSlideElement extends BaseElement {
  type: 'notesSlide' = 'notesSlide';

  /** 备注文本 */
  text?: string;

  /** 备注母版引用 */
  masterRef?: string;

  /** 备注母版对象 */
  master?: NotesMasterElement;

  /** 关联的幻灯片ID */
  slideId?: string;

  constructor(
    id: string,
    rect: { x: number; y: number; width: number; height: number },
    props: any = {}
  ) {
    super(id, 'notesSlide', rect, {}, props, {});
    this.text = props.text;
    this.masterRef = props.masterRef;
    this.slideId = props.slideId;
  }

  /**
   * 从 NotesSlideResult 创建 NotesSlideElement
   */
  static fromResult(result: NotesSlideResult): NotesSlideElement {
    return new NotesSlideElement(
      result.id,
      { x: 0, y: 0, width: 800, height: 600 },
      {
        text: result.text,
        masterRef: result.masterRef,
        slideId: result.slideId,
        elements: result.elements,
        background: result.background,
        relsMap: result.relsMap
      }
    );
  }

  toHTML(): string {
    const style = [
      this.getContainerStyle(),
      `background: #fff`,
      `padding: 20px`,
      `border: 1px solid #ddd`,
      `border-radius: 4px`
    ].join('; ');

    let content = '';

    if (this.slideId) {
      content += `<div style="font-weight: bold; margin-bottom: 10px;">幻灯片 ${this.slideId} 备注</div>`;
    }

    if (this.text) {
      content += `<div style="font-size: 14px; line-height: 1.6;">${this.text}</div>`;
    } else {
      content += `<div style="color: #999; font-style: italic;">无备注内容</div>`;
    }

    return `<div class="ppt-notes-slide" style="${style}" data-notes-id="${this.id}" data-slide-id="${this.slideId || ''}">
${content}
</div>`;
  }

  /**
   * 设置关联的备注母版
   */
  setMaster(master: NotesMasterElement): void {
    this.master = master;
  }
}
