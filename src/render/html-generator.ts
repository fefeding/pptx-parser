/**
 * HTML生成器 - 对齐PPTXjs渲染逻辑
 * 将PPTX解析结果转换为HTML
 * 
 * 核心特性：
 * 1. 绝对定位元素（对齐PPTXjs的position:absolute）
 * 2. 富文本样式继承（对齐PPTXjs的样式层次）
 * 3. 单位转换（使用EMU↔PX转换工具）
 */

import {
  emu2px,
  fontUnits2px,
  pt2px
} from '../utils/unit-converter';
import type { PptDocument, PptSlide, PptTextElement, PptStyle } from '../types';

export interface HtmlGenerationOptions {
  slideType?: 'div' | 'section'; // 支持revealjs格式
  includeGlobalCSS?: boolean; // 是否包含全局CSS
  containerClass?: string; // 容器CSS类名
}

export class HtmlGenerator {
  private styleTable: Map<string, { name: string; css: string }> = new Map();
  private styleCounter = 0;

  constructor(private options: HtmlGenerationOptions = {}) {}

  /**
   * 生成完整的HTML
   * 对齐PPTXjs的输出结构
   */
  generate(document: PptDocument): string {
    const containerClass = this.options.containerClass || 'pptxjs-container';
    let html = `<div class="${containerClass}">`;

    // 生成幻灯片
    for (const slide of document.slides) {
      html += this.generateSlide(slide, document.props.width, document.props.height);
    }

    html += '</div>';

    // 添加全局CSS
    if (this.options.includeGlobalCSS !== false) {
      html += `<style>\n${this.generateGlobalCSS()}\n</style>`;
    }

    return html;
  }

  /**
   * 生成单张幻灯片HTML
   * 对齐PPTXjs的绝对定位结构
   */
  private generateSlide(slide: PptSlide, docWidth: number, docHeight: number): string {
    const isRevealJs = this.options.slideType === 'section';
    const tag = isRevealJs ? 'section' : 'div';
    const slideWidth = slide.width || docWidth;
    const slideHeight = slide.height || docHeight;

    let html = `<${tag} class="slide" data-slide-id="${slide.id}" style="width:${slideWidth}px;height:${slideHeight}px;position:relative;overflow:hidden;">`;

    // 背景处理 - 支持纯色、渐变、图片
    if (slide.bgFill) {
      html += this.generateBackground(slide.bgFill);
    } else if (slide.bgColor && slide.bgColor !== 'transparent') {
      // 向后兼容：纯色背景
      html += `<div class="slide-background" style="width:100%;height:100%;position:absolute;top:0;left:0;background-color:${slide.bgColor};"></div>`;
    }

    // 生成所有元素
    if (slide.elements) {
      for (const element of slide.elements) {
        html += this.generateElement(element);
      }
    }

    html += `</${tag}>`;
    return html;
  }

  /**
   * 生成元素HTML
   * 对齐PPTXjs的元素渲染逻辑
   */
  private generateElement(element: any): string {
    switch (element.type) {
      case 'text':
        return this.generateTextElement(element);
      case 'image':
        return this.generateImageElement(element);
      case 'shape':
        return this.generateShapeElement(element);
      case 'table':
        return this.generateTableElement(element);
      default:
        return '';
    }
  }

  /**
   * 生成文本元素HTML
   * 对齐PPTXjs的genSpanElement逻辑
   * 
   * 关键特性：
   * 1. 绝对定位（position: absolute）
   * 2. 富文本样式（fontSize, fontFamily, color等）
   * 3. 单位转换（EMU→PX）
   * 4. 文本对齐（text-align）
   */
  private generateTextElement(element: PptTextElement): string {
    const style = element.style || {};
    const position = element.position || { x: 0, y: 0, width: 100, height: 50 };

    // 生成样式
    let styleStr = this.generateTextStyle(style, position);
    const cssClass = this.getStyleClass(styleStr);

    // 计算位置（EMU转PX）
    const x = emu2px(position.x || 0);
    const y = emu2px(position.y || 0);
    const width = emu2px(position.width || 100);
    const height = emu2px(position.height || 50);

    let html = `<div class="text-element ${cssClass}" style="position:absolute;left:${x}px;top:${y}px;width:${width}px;height:${height}px;${styleStr}">`;

    // 生成文本内容（富文本）
    if (element.text) {
      html += this.escapeHtml(element.text);
    }

    html += '</div>';
    return html;
  }

  /**
   * 生成文本样式
   * 对齐PPTXjs的样式属性
   */
  private generateTextStyle(style: PptStyle, position: any): string {
    const styles: string[] = [];

    // 字体大小（PPTX字体单位转PX）
    if (style.fontSize) {
      const fontSizePx = fontUnits2px(style.fontSize);
      styles.push(`font-size:${fontSizePx}px`);
    }

    // 字体颜色
    if (style.color) {
      styles.push(`color:${style.color}`);
    }

    // 字体族
    if (style.fontFamily) {
      styles.push(`font-family:${style.fontFamily}`);
    }

    // 字体粗细
    if (style.fontWeight === 'bold') {
      styles.push('font-weight:bold');
    }

    // 字体样式
    if (style.fontStyle === 'italic') {
      styles.push('font-style:italic');
    }

    // 文本装饰（下划线、删除线）
    if (style.textDecoration) {
      styles.push(`text-decoration:${style.textDecoration}`);
    }

    // 文本对齐
    if (style.textAlign) {
      styles.push(`text-align:${style.textAlign}`);
    }

    // 垂直对齐
    if (style.textVerticalAlign) {
      styles.push(`vertical-align:${style.textVerticalAlign}`);
    }

    // 行高
    if (style.lineHeight) {
      styles.push(`line-height:${style.lineHeight}`);
    }

    // 背景色
    if (style.backgroundColor && style.backgroundColor !== 'transparent') {
      styles.push(`background-color:${style.backgroundColor}`);
    }

    // 内边距
    if (style.padding) {
      styles.push(`padding:${style.padding}`);
    }

    return styles.join(';');
  }

  /**
   * 生成图片元素HTML
   * 对齐PPTXjs的图片渲染逻辑
   * 支持base64嵌入和外部URL
   */
  private generateImageElement(element: any): string {
    const position = element.position || { x: 0, y: 0 };
    const size = element.size || { width: 100, height: 100 };

    const x = emu2px(position.x);
    const y = emu2px(position.y);
    const width = emu2px(size.width);
    const height = emu2px(size.height);

    // 处理图片源：支持base64和URL
    let imageSrc = element.src;
    if (element.data) {
      // 如果是base64数据，直接使用
      if (element.data.startsWith('data:')) {
        imageSrc = element.data;
      } else {
        // 否则添加base64前缀
        const mimeType = element.mimeType || 'image/png';
        imageSrc = `data:${mimeType};base64,${element.data}`;
      }
    }

    return `<img class="slide-image" style="position:absolute;left:${x}px;top:${y}px;width:${width}px;height:${height}px;" src="${imageSrc}" alt="slide image" />`;
  }

  /**
   * 生成背景HTML
   * 对齐PPTXjs的背景渲染逻辑
   * 支持纯色、渐变、图片三种类型
   */
  private generateBackground(bgFill: any): string {
    const style: string[] = ['width:100%', 'height:100%', 'position:absolute', 'top:0', 'left:0'];

    switch (bgFill.type) {
      case 'solid':
        // 纯色背景
        if (bgFill.color) {
          style.push(`background-color:${bgFill.color}`);
        }
        break;

      case 'gradient':
        // 渐变背景
        if (bgFill.colors && bgFill.colors.length >= 2) {
          const gradient = `linear-gradient(${bgFill.direction || 'to bottom'}, ${bgFill.colors.join(', ')})`;
          style.push(`background:${gradient}`);
        }
        break;

      case 'image':
        // 图片背景
        if (bgFill.src) {
          style.push('background-size:cover');
          style.push('background-position:center');
          style.push('background-repeat:no-repeat');
          
          // 处理base64和URL
          if (bgFill.src.startsWith('data:')) {
            style.push(`background-image:url('${bgFill.src}')`);
          } else {
            style.push(`background-image:url('${bgFill.src}')`);
          }
        }
        break;

      default:
        // 默认纯色
        if (bgFill.color) {
          style.push(`background-color:${bgFill.color}`);
        }
    }

    return `<div class="slide-background" style="${style.join(';')}"></div>`;
  }

  /**
   * 生成形状元素HTML
   * 对齐PPTXjs的形状渲染逻辑
   */
  private generateShapeElement(element: any): string {
    const style = element.style || {};
    const position = element.position || { x: 0, y: 0 };
    const size = element.size || { width: 100, height: 100 };

    const x = emu2px(position.x);
    const y = emu2px(position.y);
    const width = emu2px(size.width);
    const height = emu2px(size.height);

    let styleStr = '';
    if (style.backgroundColor) {
      styleStr += `background-color:${style.backgroundColor};`;
    }
    if (style.borderColor) {
      styleStr += `border:${style.borderWidth || 1}px solid ${style.borderColor};`;
    }

    return `<div class="slide-shape" style="position:absolute;left:${x}px;top:${y}px;width:${width}px;height:${height}px;${styleStr}"></div>`;
  }

  /**
   * 生成表格元素HTML
   * 对齐PPTXjs的表格渲染逻辑
   */
  private generateTableElement(element: any): string {
    const position = element.position || { x: 0, y: 0 };
    const x = emu2px(position.x);
    const y = emu2px(position.y);

    let html = `<table class="slide-table" style="position:absolute;left:${x}px;top:${y}px;">`;

    if (element.rows) {
      element.rows.forEach((row: any) => {
        html += '<tr>';
        if (row.cells) {
          row.cells.forEach((cell: any) => {
            html += '<td>';
            if (cell.text) {
              html += this.escapeHtml(cell.text);
            }
            html += '</td>';
          });
        }
        html += '</tr>';
      });
    }

    html += '</table>';
    return html;
  }

  /**
   * 生成全局CSS
   * 对齐PPTXjs的CSS结构
   */
  private generateGlobalCSS(): string {
    let css = '';

    // 基础容器样式
    css += `
.pptxjs-container {
  position: relative;
  width: 100%;
  height: 100%;
  overflow: auto;
  background-color: #f0f0f0;
  padding: 20px;
  box-sizing: border-box;
}
`;

    // 幻灯片样式
    css += `
.slide {
  position: relative;
  margin-bottom: 20px;
  background-color: #ffffff;
  box-shadow: 0 2px 8px rgba(0,0,0,0.1);
  overflow: hidden;
}
`;

    // 文本元素样式
    css += `
.text-element {
  display: flex;
  align-items: center;
  word-wrap: break-word;
  overflow-wrap: break-word;
  box-sizing: border-box;
}
`;

    // 图片样式
    css += `
.slide-image {
  display: block;
  object-fit: contain;
}
`;

    // 形状样式
    css += `
.slide-shape {
  box-sizing: border-box;
}
`;

    // 表格样式
    css += `
.slide-table {
  border-collapse: collapse;
}
.slide-table td {
  border: 1px solid #cccccc;
  padding: 8px;
  text-align: left;
}
`;

    // 添加所有生成的样式类
    for (const [, style] of this.styleTable) {
      css += `.${style.name} { ${style.css} }\n`;
    }

    return css;
  }

  /**
   * 获取或创建CSS类名
   * 对齐PPTXjs的样式类生成逻辑
   */
  private getStyleClass(styleText: string): string {
    // 查找是否已存在相同样式
    for (const [key, value] of this.styleTable.entries()) {
      if (key === styleText) {
        return value.name;
      }
    }

    // 创建新样式类
    const className = `_pptx_style_${this.styleCounter++}`;
    this.styleTable.set(styleText, { name: className, css: styleText });
    return className;
  }

  /**
   * HTML转义
   */
  private escapeHtml(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }
}

/**
 * 便捷函数：生成HTML
 */
export function generateHtml(
  document: PptDocument,
  options?: HtmlGenerationOptions
): string {
  const generator = new HtmlGenerator(options);
  return generator.generate(document);
}