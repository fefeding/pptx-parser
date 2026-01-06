/**
 * HTML 渲染器模块
 * 将 PPTX 解析结果转换为 HTML
 * 支持 layout、master 样式继承
 */

import type { PptxParseResult, SlideParseResult, SlideLayoutResult, MasterSlideResult } from '../core/types';
import { BaseElement } from '../elements/BaseElement';
import { createElementFromData } from '../elements';
import { applyStyleInheritance } from '../core/style-inheritance';

/**
 * HTML 渲染选项
 */
export interface HtmlRenderOptions {
  /** 是否包含样式 */
  includeStyles?: boolean;
  /** 是否包含脚本 */
  includeScripts?: boolean;
  /** 是否包含布局和母版元素 */
  includeLayoutElements?: boolean;
  /** 自定义 CSS */
  customCss?: string;
}

/**
 * 将单个幻灯片转换为 HTML
 * @param slide 幻灯片解析结果
 * @param options 渲染选项
 * @returns HTML 字符串
 */
export function slide2HTML(slide: SlideParseResult, options: HtmlRenderOptions = {}): string {
  // 确保样式已应用
  if (!slide.styleApplied && slide.layout) {
    applyStyleInheritance(slide, slide.layout, slide.master);
    slide.styleApplied = true;
  }

  // 创建元素列表
  const elements = slide.elements.map(el => {
    if (el instanceof BaseElement) {
      return el;
    }
    // 将普通对象转换为元素实例
    return createElementFromData(el, slide.relsMap || {});
  }).filter(el => el !== null) as BaseElement[];

  // 获取背景样式
  const background = getSlideBackground(slide);

  // 获取幻灯片尺寸
  const size = getSlideSize(slide);

  // 渲染幻灯片容器
  const slideStyle = [
    `position: relative`,
    `width: ${size.width}px`,
    `height: ${size.height}px`,
    `overflow: hidden`,
    background
  ].join('; ');

  // 渲染布局和母版元素（如果启用）
  const layoutElementsHTML = renderLayoutElements(slide, options);

  // 渲染幻灯片元素
  const elementsHTML = elements
    .map(el => el.toHTML())
    .join('\n');

  const slideHTML = `<div class="ppt-slide" style="${slideStyle}" data-slide-id="${slide.id}" data-slide-index="${slide.index ?? 0}">
${layoutElementsHTML}
${elementsHTML}
</div>`;

  return slideHTML;
}

/**
 * 将整个 PPT 文档转换为 HTML 数组
 * @param result PPTX 解析结果
 * @param options 渲染选项
 * @returns HTML 字符串数组（每个幻灯片一个）
 */
export function ppt2HTML(result: PptxParseResult, options: HtmlRenderOptions = {}): string[] {
  return result.slides.map(slide => slide2HTML(slide, options));
}

/**
 * 将整个 PPT 文档转换为单个 HTML 文档
 * @param result PPTX 解析结果
 * @param options 渲染选项
 * @returns 完整的 HTML 文档字符串
 */
export function ppt2HTMLDocument(result: PptxParseResult, options: HtmlRenderOptions = {}): string {
  const slidesHTML = ppt2HTML(result, options).join('\n\n');

  const size = {
    width: result.props.width || 960,
    height: result.props.height || 540
  };

  const containerStyle = [
    `max-width: ${size.width}px`,
    `margin: 0 auto`,
    `padding: 20px`,
    `background: #f5f5f5`
  ].join('; ');

  const styles = options.includeStyles !== false ? generateStyles(result, options) : '';

  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(result.title)}</title>
  ${styles ? '<style>\n' + styles + '\n</style>' : ''}
</head>
<body>
  <div class="ppt-container" style="${containerStyle}">
    <div class="ppt-header" style="text-align: center; padding: 20px 0; margin-bottom: 20px; border-bottom: 1px solid #ddd;">
      <h1>${escapeHtml(result.title)}</h1>
      ${result.author ? `<p style="color: #666; font-size: 14px;">作者: ${escapeHtml(result.author)}</p>` : ''}
    </div>
    <div class="ppt-slides">
${slidesHTML}
    </div>
  </div>
</body>
</html>`;
}

/**
 * 渲染布局和母版元素
 * @param slide 幻灯片对象
 * @param options 渲染选项
 * @returns HTML 字符串
 */
function renderLayoutElements(slide: SlideParseResult, options: HtmlRenderOptions): string {
  if (options.includeLayoutElements === false) {
    return '';
  }

  const elements: string[] = [];

  // 渲染母版元素（如页脚、页码等）
  if (slide.master?.elements && slide.master.elements.length > 0) {
    slide.master.elements.forEach(el => {
      if (el instanceof BaseElement) {
        // 只渲染标记为"显示在所有幻灯片"的元素
        if (!el.hidden) {
          elements.push(`<div class="ppt-master-element">${el.toHTML()}</div>`);
        }
      }
    });
  }

  // 渲染布局元素（占位符）
  if (slide.layout?.elements && slide.layout.elements.length > 0) {
    slide.layout.elements.forEach(el => {
      if (el instanceof BaseElement) {
        // 只渲染占位符（且未被 slide 元素覆盖）
        if (el.isPlaceholder && !el.hidden) {
          elements.push(`<div class="ppt-layout-element">${el.toHTML()}</div>`);
        }
      }
    });
  }

  return elements.join('\n');
}

/**
 * 获取幻灯片背景样式
 * @param slide 幻灯片对象
 * @returns CSS 背景样式
 */
function getSlideBackground(slide: SlideParseResult): string {
  const bg = slide.background;

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

  if (bg.type === 'image' && bg.relId) {
    // 图片背景需要从 relsMap 获取实际 URL
    return `background-image: url('${bg.relId}'); background-size: cover; background-position: center;`;
  }

  return 'background-color: #ffffff;';
}

/**
 * 获取幻灯片尺寸
 * @param slide 幻灯片对象
 * @returns 尺寸对象
 */
function getSlideSize(slide: SlideParseResult): { width: number; height: number } {
  // 使用默认尺寸（16:9）
  return {
    width: 960,
    height: 540
  };
}

/**
 * 生成 CSS 样式
 * @param result PPTX 解析结果
 * @param options 渲染选项
 * @returns CSS 字符串
 */
function generateStyles(result: PptxParseResult, options: HtmlRenderOptions): string {
  const css: string[] = [
    '/* PPTX 默认样式 */',
    '* { box-sizing: border-box; }',
    '',
    '.ppt-slide {',
    '  position: relative;',
    '  margin: 20px auto;',
    '  border: 1px solid #ddd;',
    '  background: #fff;',
    '}',
    '',
    '.ppt-master-element,',
    '.ppt-layout-element {',
    '  pointer-events: none;',
    '}',
    '',
    '/* 元素通用样式 */',
    '.ppt-element {',
    '  position: absolute;',
    '}',
    '',
    '/* 文本元素 */',
    '.ppt-text {',
    '  overflow: hidden;',
    '}',
    ''
  ];

  // 自定义 CSS
  if (options.customCss) {
    css.push('/* 自定义样式 */');
    css.push(options.customCss);
  }

  return css.join('\n');
}

/**
 * HTML 转义
 * @param text 文本
 * @returns 转义后的文本
 */
function escapeHtml(text: string): string {
  const map: Record<string, string> = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, m => map[m]);
}
