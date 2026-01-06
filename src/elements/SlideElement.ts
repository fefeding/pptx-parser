/**
 * 幻灯片元素类
 * 封装幻灯片的完整信息，包括toHTML能力
 */

import type { SlideParseResult } from '../core/types';
import type { BaseElement } from './BaseElement';

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

  constructor(result: SlideParseResult, elements: BaseElement[]) {
    this.id = result.id;
    this.title = result.title;
    this.background = typeof result.background === 'string' ? result.background : '';
    this.elements = elements;
    this.rawResult = result;
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const slideStyle = [
      `width: 100%`,
      `height: 100%`,
      `position: relative`,
      `background-color: ${this.background}`,
      `overflow: hidden`
    ].join('; ');

    const elementsHTML = this.elements
      .map(element => element.toHTML())
      .join('\n');

    return `<div class="ppt-slide" style="${slideStyle}" data-slide-id="${this.id}" data-slide-title="${this.escapeHtml(this.title)}">
${elementsHTML}
    </div>`;
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
