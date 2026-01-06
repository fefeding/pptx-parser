/**
 * 文档元素类
 * 整个PPT文档的主对象，包含所有文档信息、基础样式和公共样式
 */

import type { PptxParseResult, SlideLayoutResult, MasterSlideResult } from '../core/types';
import { BaseElement } from './BaseElement';
import { SlideElement } from './SlideElement';
import { LayoutElement } from './LayoutElement';
import { MasterElement } from './MasterElement';
import { TagsElement } from './TagsElement';
import { NotesMasterElement, NotesSlideElement } from './NotesElement';
import { createElementFromData } from './index';
import type { BaseElement as BaseElementType } from './BaseElement';

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
  /** 是否带导航 */
  withNavigation?: boolean;
  /** 自定义 CSS */
  customCss?: string;
}

/**
 * 文档元素类
 */
export class DocumentElement extends BaseElement {
  type: 'document' = 'document';

  /** 文档标题 */
  title: string;

  /** 作者 */
  author?: string;

  /** 主题 */
  subject?: string;

  /** 关键词 */
  keywords?: string;

  /** 描述 */
  description?: string;

  /** 创建时间 */
  created?: string;

  /** 修改时间 */
  modified?: string;

  /** 幻灯片列表 */
  slides: SlideElement[];

  /** 布局列表 */
  layouts: Record<string, LayoutElement>;

  /** 母版列表 */
  masters: MasterElement[];

  /** 标签列表 */
  tags: TagsElement[];

  /** 备注母版 */
  notesMasters: NotesMasterElement[];

  /** 备注页 */
  notesSlides: NotesSlideElement[];

  /** 文档宽度（像素） */
  width: number;

  /** 文档高度（像素） */
  height: number;

  /** 宽高比 */
  ratio: number;

  /** 页面尺寸类型 */
  pageSize: '4:3' | '16:9' | '16:10' | 'custom';

  /** 全局关联关系映射表 */
  globalRelsMap: Record<string, any>;

  /** 媒体资源映射表（relId -> base64 URL） */
  mediaMap?: Map<string, string>;

  constructor(
    id: string,
    title: string,
    width: number = 960,
    height: number = 540,
    props: any = {}
  ) {
    super(id, 'document', { x: 0, y: 0, width, height }, {}, props, {});
    this.title = title;
    this.width = width;
    this.height = height;
    this.ratio = width / height;
    this.pageSize = props.pageSize || '16:9';
    this.author = props.author;
    this.subject = props.subject;
    this.keywords = props.keywords;
    this.description = props.description;
    this.created = props.created;
    this.modified = props.modified;
    this.slides = [];
    this.layouts = {};
    this.masters = [];
    this.tags = [];
    this.notesMasters = [];
    this.notesSlides = [];
    this.globalRelsMap = props.globalRelsMap || {};
    this.mediaMap = props.mediaMap;
  }

  /**
   * 从 PptxParseResult 创建 DocumentElement
   */
  static fromParseResult(result: PptxParseResult): DocumentElement {
    const doc = new DocumentElement(
      result.id,
      result.title,
      result.props.width || 960,
      result.props.height || 540,
      {
        pageSize: result.props.pageSize,
        author: result.author,
        subject: result.subject,
        keywords: result.keywords,
        description: result.description,
        created: result.created,
        modified: result.modified,
        globalRelsMap: result.globalRelsMap,
        mediaMap: result.mediaMap
      }
    );

    // 解析母版
    if (result.masterSlides && result.masterSlides.length > 0) {
      doc.masters = result.masterSlides.map(master => MasterElement.fromResult(master));
    }

    // 解析布局
    if (result.slideLayouts) {
      Object.entries(result.slideLayouts).forEach(([layoutId, layout]) => {
        doc.layouts[layoutId] = LayoutElement.fromResult(layout);
      });
    }

    // 解析备注母版
    if (result.notesMasters && result.notesMasters.length > 0) {
      doc.notesMasters = result.notesMasters.map(nm => NotesMasterElement.fromResult(nm));
    }

    // 解析备注页
    if (result.notesSlides && result.notesSlides.length > 0) {
      doc.notesSlides = result.notesSlides.map(ns => {
        const notesElement = NotesSlideElement.fromResult(ns);
        // 设置关联的母版
        if (ns.masterRef) {
          const master = doc.notesMasters.find(m => m.id === ns.masterRef);
          if (master) {
            notesElement.setMaster(master);
          }
        }
        return notesElement;
      });
    }

    // 解析标签
    if (result.tags && result.tags.length > 0) {
      doc.tags = result.tags.map(tag => TagsElement.fromResult(tag));
    }

    // 解析幻灯片
    if (result.slides && result.slides.length > 0) {
      doc.slides = result.slides.map((slide: any) => {
        // 将 slide.elements 转换为 BaseElement 实例
        const elements = (slide.elements || []).map((el: any) => {
          if (el instanceof BaseElement) {
            return el;
          }
          return createElementFromData(el, slide.relsMap || {}, doc.mediaMap);
        }).filter((el: any) => el !== null) as BaseElement[];

        // 使用 parser 已关联的 layout 和 master 对象（SlideLayoutResult 和 MasterSlideResult 类型）
        const layout = slide.layout as SlideLayoutResult | undefined;
        const master = slide.master as MasterSlideResult | undefined;

        return new SlideElement(slide, elements, layout, master, doc.mediaMap);
      });
    }

    return doc;
  }

  /**
   * 转换为HTML文档
   * @param options 渲染选项
   */
  toHTML(options: HtmlRenderOptions = {}): string {
    // 默认带导航
    const withNav = options.withNavigation !== false;

    if (withNav) {
      return this.toHTMLWithNavigation(options);
    }

    return this.toHTMLDocument(options);
  }

  /**
   * 转换为完整的HTML文档（静态展示）
   */
  toHTMLDocument(options: HtmlRenderOptions = {}): string {
    const containerStyle = [
      `max-width: ${this.width}px`,
      `margin: 0 auto`,
      `padding: 20px`,
      `background: #f5f5f5`
    ].join('; ');

    const slidesHTML = this.slides
      .map(slide => slide.toHTML())
      .join('\n\n');

    const styles = options.includeStyles !== false ? this.generateStyles(options) : '';

    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${this.escapeHtml(this.title)}</title>
  ${styles ? '<style>\n' + styles + '\n</style>' : ''}
</head>
<body>
  <div class="ppt-container" style="${containerStyle}">
    <div class="ppt-header" style="text-align: center; padding: 20px 0; margin-bottom: 20px; border-bottom: 1px solid #ddd;">
      <h1>${this.escapeHtml(this.title)}</h1>
      ${this.author ? `<p style="color: #666; font-size: 14px;">作者: ${this.escapeHtml(this.author)}</p>` : ''}
    </div>
    <div class="ppt-slides">
${slidesHTML}
    </div>
  </div>
</body>
</html>`;
  }

  /**
   * 转换为带导航的HTML文档（交互式展示）
   */
  toHTMLWithNavigation(options: HtmlRenderOptions = {}): string {
    const slidesHTML = this.slides.map((slide, index) => {
      return `
        <div class="ppt-slide-page" data-index="${index}" style="display: ${index === 0 ? 'block' : 'none'};">
          ${slide.toHTML()}
        </div>
      `;
    }).join('\n');

    const navHTML = `
      <div class="ppt-navigation" style="position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); z-index: 1000; display: flex; align-items: center; gap: 15px; background: rgba(0,0,0,0.7); padding: 10px 20px; border-radius: 25px;">
        <button onclick="prevSlide()" style="padding: 8px 16px; cursor: pointer; background: #fff; border: none; border-radius: 4px; font-size: 14px;">上一页</button>
        <span id="slideCounter" style="color: #fff; font-size: 14px; min-width: 60px; text-align: center;">1 / ${this.slides.length}</span>
        <button onclick="nextSlide()" style="padding: 8px 16px; cursor: pointer; background: #fff; border: none; border-radius: 4px; font-size: 14px;">下一页</button>
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

        // 触摸滑动支持
        let touchStartX = 0;
        let touchEndX = 0;
        document.addEventListener('touchstart', e => {
          touchStartX = e.changedTouches[0].screenX;
        });
        document.addEventListener('touchend', e => {
          touchEndX = e.changedTouches[0].screenX;
          if (touchStartX - touchEndX > 50) nextSlide();
          if (touchEndX - touchStartX > 50) prevSlide();
        });
      <\/script>
    `;

    const styles = options.includeStyles !== false ? this.generateStyles(options) : this.generateNavigationStyles();

    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${this.escapeHtml(this.title)}</title>
  ${styles ? '<style>\n' + styles + '\n</style>' : ''}
</head>
<body>
  <div class="ppt-wrapper">
${slidesHTML}
  </div>
${navHTML}
${options.includeScripts !== false ? script : ''}
</body>
</html>`;
  }

  /**
   * 生成CSS样式
   */
  private generateStyles(options: HtmlRenderOptions): string {
    const css: string[] = [
      '/* PPTX 文档基础样式 */',
      '* { box-sizing: border-box; margin: 0; padding: 0; }',
      '',
      'body {',
      '  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;',
      '  background: #333;',
      '  min-height: 100vh;',
      '}',
      '',
      '.ppt-wrapper {',
      '  display: flex;',
      '  justify-content: center;',
      '  align-items: center;',
      '  min-height: 100vh;',
      '  padding: 20px;',
      '}',
      '',
      '.ppt-container {',
      '  background: #fff;',
      '  box-shadow: 0 2px 8px rgba(0,0,0,0.1);',
      '}',
      '',
      '/* 幻灯片样式 */',
      '.ppt-slide {',
      '  position: relative;',
      '  margin: 20px auto;',
      '  border: 1px solid #ddd;',
      '  background: #fff;',
      '}',
      '',
      '.ppt-slide-page {',
      `  width: ${this.width}px;`,
      `  height: ${this.height}px;`,
      '  background: #fff;',
      '  box-shadow: 0 4px 20px rgba(0,0,0,0.3);',
      '}',
      '',
      '/* 布局和母版元素 */',
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
      '',
      '/* 占位符样式 */',
      '.ppt-placeholder {',
      '  pointer-events: none;',
      '  user-select: none;',
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
   * 生成导航模式样式
   */
  private generateNavigationStyles(): string {
    return `* { margin: 0; padding: 0; box-sizing: border-box; }
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
}`;
  }

  /**
   * 获取指定幻灯片
   */
  getSlide(index: number): SlideElement | undefined {
    return this.slides[index];
  }

  /**
   * 获取指定布局
   */
  getLayout(layoutId: string): LayoutElement | undefined {
    return this.layouts[layoutId];
  }

  /**
   * 获取指定母版
   */
  getMaster(masterId: string): MasterElement | undefined {
    return this.masters.find(m => m.id === masterId);
  }

  /**
   * HTML转义
   */
  private escapeHtml(text: string): string {
    const map: Record<string, string> = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
  }
}

/**
 * 从 PptxParseResult 创建文档元素的便捷函数
 */
export function createDocument(result: PptxParseResult): DocumentElement {
  return DocumentElement.fromParseResult(result);
}
