/**
 * 标签元素类
 * 表示幻灯片标签、扩展属性和自定义属性
 */

import { BaseElement } from './BaseElement';
import type { TagsResult, SlideTag, CustomProperty, ExtensionData } from '../core/types';

/**
 * 标签元素类
 */
export class TagsElement extends BaseElement {
  type: 'tags' = 'tags';

  /** 标签列表 */
  tags: SlideTag[];

  /** 扩展数据 */
  extensions: ExtensionData[];

  /** 自定义属性 */
  customProperties: CustomProperty[];

  /** 关联的幻灯片ID */
  slideId?: string;

  constructor(
    id: string,
    tags: SlideTag[] = [],
    customProperties: CustomProperty[] = [],
    extensions: ExtensionData[] = [],
    props: any = {}
  ) {
    super(id, 'tags', { x: 0, y: 0, width: 0, height: 0 }, {}, props, {});
    this.tags = tags;
    this.customProperties = customProperties;
    this.extensions = extensions;
    this.slideId = props.slideId;
  }

  /**
   * 从 TagsResult 创建 TagsElement
   */
  static fromResult(result: TagsResult): TagsElement {
    return new TagsElement(
      result.id,
      result.tags,
      result.customProperties,
      result.extensions,
      { slideId: result.slideId }
    );
  }

  toHTML(): string {
    const hasContent = this.tags.length > 0 || this.customProperties.length > 0;

    if (!hasContent) {
      // 没有内容时返回空字符串或占位符
      return '';
    }

    const style = [
      `display: none`, // 标签通常不显示，用于数据存储
      `data-tags-id="${this.id}"`,
      `data-slide-id="${this.slideId || ''}"`
    ].join('; ');

    // 序列化为 data 属性
    const tagsData = this.tags.length > 0
      ? `data-tags='${JSON.stringify(this.tags)}'`
      : '';

    const propsData = this.customProperties.length > 0
      ? `data-custom-props='${JSON.stringify(this.customProperties)}'`
      : '';

    const extensionsData = this.extensions.length > 0
      ? `data-extensions='${JSON.stringify(this.extensions)}'`
      : '';

    return `<div class="ppt-tags" style="${style}" ${tagsData} ${propsData} ${extensionsData}>
</div>`;
  }

  /**
   * 渲染为可见的调试信息（仅用于开发）
   */
  toHTMLDebug(): string {
    const style = [
      `position: absolute`,
      `top: 10px`,
      `right: 10px`,
      `background: rgba(0, 0, 0, 0.8)`,
      `color: white`,
      `padding: 10px`,
      `border-radius: 4px`,
      `font-size: 12px`,
      `max-width: 300px`,
      `z-index: 1000`,
      `pointer-events: none`
    ].join('; ');

    let content = `<strong>Tags (${this.tags.length}):</strong>`;
    if (this.tags.length > 0) {
      content += '<ul>';
      this.tags.forEach(tag => {
        content += `<li>${tag.name}: ${tag.value}</li>`;
      });
      content += '</ul>';
    }

    if (this.customProperties.length > 0) {
      content += `<strong>Properties (${this.customProperties.length}):</strong><ul>`;
      this.customProperties.forEach(prop => {
        content += `<li>${prop.name}: ${prop.value} (${prop.type})</li>`;
      });
      content += '</ul>';
    }

    return `<div class="ppt-tags-debug" style="${style}">
${content}
</div>`;
  }

  /**
   * 获取标签值
   * @param name 标签名
   */
  getTag(name: string): string | undefined {
    return this.tags.find(tag => tag.name === name)?.value;
  }

  /**
   * 获取自定义属性
   * @param name 属性名
   */
  getProperty(name: string): any | undefined {
    return this.customProperties.find(prop => prop.name === name)?.value;
  }

  /**
   * 设置标签
   */
  setTag(name: string, value: string): void {
    const existing = this.tags.find(tag => tag.name === name);
    if (existing) {
      existing.value = value;
    } else {
      this.tags.push({ name, value });
    }
  }

  /**
   * 设置自定义属性
   */
  setProperty(name: string, value: any, type?: 'string' | 'number' | 'boolean' | 'date'): void {
    const existing = this.customProperties.find(prop => prop.name === name);
    if (existing) {
      existing.value = value;
      if (type) existing.type = type;
    } else {
      this.customProperties.push({ name, value, type });
    }
  }
}

