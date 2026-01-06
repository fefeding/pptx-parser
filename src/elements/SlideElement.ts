/**
 * 幻灯片元素类
 * 封装幻灯片的完整信息，包括toHTML能力
 */

import type { SlideParseResult } from '../core/types';
import { BaseElement } from './BaseElement';
import { applyStyleInheritance } from '../core/style-inheritance';
import type { SlideLayoutResult, MasterSlideResult } from '../core/types';
import { createElementFromData } from './element-factory';

/**
 * 幻灯片元素类
 */
export class SlideElement {
  /** 幻灯片ID */
  id: string;

  /** 幻灯片标题 */
  title: string;

  /** 背景颜色 */
  background: string;

  /** 元素列表 */
  elements: BaseElement[];

  /** 原始解析结果 */
  rawResult: SlideParseResult;

  /** 布局对象 */
  layout?: SlideLayoutResult;

  /** 母版对象 */
  master?: MasterSlideResult;

  /** 媒体资源映射表 */
  mediaMap?: Map<string, string>;

  constructor(
    result: SlideParseResult,
    elements: BaseElement[],
    layout?: SlideLayoutResult,
    master?: MasterSlideResult,
    mediaMap?: Map<string, string>
  ) {
    this.id = result.id;
    this.title = result.title;
    this.background = typeof result.background === 'string' ? result.background : '';
    this.elements = elements;
    this.rawResult = result;
    this.layout = layout;
    this.master = master;
    this.mediaMap = mediaMap;
  }

  /**
   * 转换为HTML
   * 支持 layout 和 master 样式继承
   */
  toHTML(): string {
    // 确保样式已应用（从 layout 和 master 继承）
    if (!this.rawResult.styleApplied && this.layout) {
      applyStyleInheritance(this.rawResult, this.layout, this.master);
      this.rawResult.styleApplied = true;
    }

    // 获取背景样式
    const background = this.getSlideBackground();

    // 获取幻灯片尺寸
    const size = { width: 960, height: 540 };

    const slideStyle = [
      `width: ${size.width}px`,
      `height: ${size.height}px`,
      `position: relative`,
      background,
      `overflow: hidden`
    ].join('; ');

    // 渲染布局和母版元素
    const layoutElementsHTML = this.renderLayoutElements();

    // 渲染幻灯片元素
    const elementsHTML = this.elements
      .map(element => element.toHTML())
      .join('\n');

    return `<div class="ppt-slide" style="${slideStyle}" data-slide-id="${this.id}" data-slide-title="${this.escapeHtml(this.title)}">
${layoutElementsHTML}
${elementsHTML}
    </div>`;
  }

  /**
   * 渲染布局和母版元素
   * PPTXjs 中，layout 和 master 中的元素会被渲染到 slide 上
   * 渲染顺序：master elements -> layout elements -> slide elements
   */
  private renderLayoutElements(): string {
    const elements: string[] = [];

    // 渲染母版元素（如页脚、页码、背景图片等）
    if (this.master?.elements && this.master.elements.length > 0) {
      this.master.elements.forEach(el => {
        // 将原始数据转换为 BaseElement 实例（如果还没有转换）
        if (!(el instanceof BaseElement)) {
          const relsMap = (this.master as any).relsMap || {};
          const element = createElementFromData(el, relsMap, this.mediaMap);
          // 只渲染非隐藏的元素
          if (element && !el.hidden) {
            const html = element.toHTML();
            elements.push(`<div class="ppt-master-element">${html}</div>`);
          }
        } else if (!el.hidden) {
          elements.push(`<div class="ppt-master-element">${el.toHTML()}</div>`);
        }
      });
    }

    // 渲染布局元素（如果有实际的形状元素，而不仅仅是占位符定义）
    if (this.layout?.elements && this.layout.elements.length > 0) {
      this.layout.elements.forEach(el => {
        // 将原始数据转换为 BaseElement 实例（如果还没有转换）
        if (!(el instanceof BaseElement)) {
          const relsMap = (this.layout as any).relsMap || {};
          const element = createElementFromData(el, relsMap, this.mediaMap);
          // 只渲染有 type 属性的元素（形状、图片等），跳过纯占位符定义
          if (element && el.type && !el.hidden && !el.isPlaceholder) {
            const html = element.toHTML();
            elements.push(`<div class="ppt-layout-element">${html}</div>`);
          }
        } else if (!el.isPlaceholder && !el.hidden) {
          elements.push(`<div class="ppt-layout-element">${el.toHTML()}</div>`);
        }
      });
    }

    return elements.join('\n');
  }

  /**
   * 获取幻灯片背景样式
   */
  private getSlideBackground(): string {
    const bg = this.rawResult.background;

    if (!bg) {
      return 'background-color: #ffffff;';
    }

    if (typeof bg === 'string') {
      return `background-color: ${bg};`;
    }

    // 背景对象
    if (bg.type === 'color' && bg.value) {
      return `background-color: ${bg.value};`;
    }

    if (bg.type === 'image') {
      // 图片背景，优先使用解析后的 base64 URL，否则使用 relId
      const imageUrl = bg.value || bg.relId;
      if (imageUrl) {
        return `background-image: url('${imageUrl}'); background-size: cover; background-position: center;`;
      }
    }

    return 'background-color: #ffffff;';
  }

  /**
   * 转换为HTML字符串（带包裹容器）
   */
  toHTMLString(): string {
    const containerStyle = [
      `width: 960px`,
      `height: 540px`,
      `margin: 20px auto`,
      `border: 1px solid #ccc`,
      `box-shadow: 0 2px 8px rgba(0,0,0,0.1)`,
      `background-color: ${this.background}`,
      `position: relative`,
      `overflow: hidden`
    ].join('; ');

    const slideHTML = this.toHTML();

    return `<div class="ppt-slide-container" style="${containerStyle}">
      <div class="ppt-slide-header" style="padding: 10px; background: #f5f5f5; border-bottom: 1px solid #ddd; font-size: 14px; font-weight: bold;">
        ${this.escapeHtml(this.title)}
      </div>
      ${slideHTML}
    </div>`;
  }

  /**
   * HTML转义
   */
  private escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }
}

/**
 * PPTX文档类
 */
export class PptxDocument {
  /** 文档ID */
  id: string;

  /** 文档标题 */
  title: string;

  /** 作者 */
  author?: string;

  /** 幻灯片列表 */
  slides: SlideElement[];

  /** 文档宽度（像素） */
  width: number;

  /** 文档高度（像素） */
  height: number;

  /** 宽高比 */
  ratio: number;

  constructor(
    id: string,
    title: string,
    slides: SlideElement[],
    width: number = 960,
    height: number = 540,
    author?: string
  ) {
    this.id = id;
    this.title = title;
    this.slides = slides;
    this.width = width;
    this.height = height;
    this.ratio = width / height;
    this.author = author;
  }

  /**
   * 从 PptxParseResult 创建 PptxDocument
   */
  static fromParseResult(result: any): PptxDocument {
    const slides = result.slides.map((slide: any) => {
      // 将 slide.elements 转换为 BaseElement 实例
      const { createElementFromData } = require('../elements');
      const elements = (slide.elements || []).map((el: any) => {
        if (el instanceof BaseElement) {
          return el;
        }
        return createElementFromData(el, slide.relsMap || {});
      }).filter((el: any) => el !== null);

      return new SlideElement(slide, elements, slide.layout, slide.master);
    });

    return new PptxDocument(
      result.id,
      result.title,
      slides,
      result.props.width || 960,
      result.props.height || 540,
      result.author
    );
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const containerStyle = [
      `width: ${this.width}px`,
      `min-height: ${this.height}px`,
      `margin: 20px auto`,
      `padding: 20px`,
      `background: #f9f9f9`,
      `border: 1px solid #ddd`
    ].join('; ');

    const slidesHTML = this.slides
      .map(slide => slide.toHTMLString())
      .join('\n\n');

    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${this.escapeHtml(this.title)}</title>
  <style>
    .ppt-slide-container {
      page-break-after: always;
    }
    .ppt-slide-container:last-child {
      page-break-after: auto;
    }
  </style>
</head>
<body>
  <div style="${containerStyle}">
    <h1 style="text-align: center; margin-bottom: 30px;">${this.escapeHtml(this.title)}</h1>
${slidesHTML}
  </div>
</body>
</html>`;
  }

  /**
   * 转换为单个页面的HTML（带导航）
   */
  toHTMLWithNavigation(): string {
    const slidesHTML = this.slides.map((slide, index) => {
      return `
        <div class="ppt-slide-page" data-index="${index}" style="display: ${index === 0 ? 'block' : 'none'};">
          ${slide.toHTML()}
        </div>
      `;
    }).join('\n');

    const navHTML = `
      <div style="position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); z-index: 1000;">
        <button onclick="prevSlide()" style="padding: 10px 20px; margin-right: 10px; cursor: pointer;">上一页</button>
        <span id="slideCounter" style="padding: 10px;">1 / ${this.slides.length}</span>
        <button onclick="nextSlide()" style="padding: 10px 20px; margin-left: 10px; cursor: pointer;">下一页</button>
      </div>
    `;

    const script = `
      <script>
        let currentSlide = 0;
        const slides = document.querySelectorAll('.ppt-slide-page');

        function showSlide(index) {
          slides.forEach((slide, i) => {
            slide.style.display = i === index ? 'block' : 'none';
          });
          document.getElementById('slideCounter').textContent = (index + 1) + ' / ${this.slides.length}';
          currentSlide = index;
        }

        function prevSlide() {
          showSlide(Math.max(0, currentSlide - 1));
        }

        function nextSlide() {
          showSlide(Math.min(${this.slides.length} - 1, currentSlide + 1));
        }

        document.addEventListener('keydown', (e) => {
          if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') prevSlide();
          if (e.key === 'ArrowRight' || e.key === 'ArrowDown') nextSlide();
        });
      <\/script>
    `;

    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${this.escapeHtml(this.title)}</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      display: flex;
      justify-content: center;
      align-items: center;
      min-height: 100vh;
      background: #333;
    }
    .ppt-wrapper {
      width: ${this.width}px;
      height: ${this.height}px;
      background: #fff;
      box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    }
  </style>
</head>
<body>
  <div class="ppt-wrapper">
${slidesHTML}
  </div>
${navHTML}
${script}
</body>
</html>`;
  }

  /**
   * HTML转义
   */
  private escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }
}
