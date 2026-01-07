/**
 * 样式继承工具
 * 处理 master、layout 和 slide 之间的样式继承和覆盖
 * 对齐 PPTXjs 的样式继承机制
 */

import type { TextStyles, SlideLayoutResult, MasterSlideResult } from './types';
import { getFirstChildByTagNS } from '../utils';
import { NS } from '../constants';

/**
 * 样式继承上下文
 * 包含 slide、layout 和 master 的所有样式信息
 */
export interface StyleContext {
  slide?: any;
  layout?: SlideLayoutResult;
  master?: MasterSlideResult;
  theme?: any;
}

  /**
   * 获取占位符样式
   * 按优先级查找：slide > layout > master > default
   * @param element 元素对象
   * @param context 样式继承上下文
   * @returns 样式对象
   */
  export function getPlaceholderStyle(element: any, context: StyleContext): any {
    // 如果元素本身定义了文本样式，优先使用
    if (element.textStyle || element.style) {
      return element.textStyle || element.style;
    }

    // 检查是否是占位符
    if (!element.isPlaceholder || !element.placeholderType) {
      return getDefaultTextStyle();
    }

    // 获取占位符大小属性（sz），用于确定默认字体大小
    const placeholderSize = element.phSize || element.getAttribute?.('sz');

    // 根据占位符类型和大小获取样式
    const style = getPlaceholderStyleByType(element.placeholderType, context);
    
    // 如果没有继承到样式，使用基于占位符大小和类型的默认样式
    if (!style) {
      return getDefaultTextStyle(element.placeholderType, placeholderSize);
    }
    
    return style;
  }

/**
 * 根据占位符类型获取样式
 * @param placeholderType 占位符类型
 * @param context 样式继承上下文
 * @returns 样式对象
 */
function getPlaceholderStyleByType(placeholderType: string, context: StyleContext): any {
  // 优先级：layout > master
  if (context.layout?.textStyles) {
    const layoutStyle = getStyleFromTextStyles(context.layout.textStyles, placeholderType);
    if (layoutStyle) return layoutStyle;
  }

  if (context.master?.textStyles) {
    const masterStyle = getStyleFromTextStyles(context.master.textStyles, placeholderType);
    if (masterStyle) return masterStyle;
  }

  // 默认样式
  return getDefaultTextStyle();
}

/**
 * 从文本样式对象中获取指定类型的样式
 * @param textStyles 文本样式对象
 * @param placeholderType 占位符类型
 * @returns 样式对象
 */
function getStyleFromTextStyles(textStyles: TextStyles, placeholderType: string): any | null {
  switch (placeholderType) {
    case 'title':
    case 'ctrTitle':
    case 'subTitle':
      return parseParagraphProperties(textStyles.titleParaPr);

    case 'body':
    case 'obj':
    case 'chart':
    case 'table':
    case 'dgm':
      return parseParagraphProperties(textStyles.bodyPr);

    default:
      return parseParagraphProperties(textStyles.otherPr);
  }
}

/**
 * 解析段落属性为样式对象
 * @param paraPr 段落属性节点
 * @returns 样式对象
 */
function parseParagraphProperties(paraPr: any): any {
  if (!paraPr) return null;

  const style: any = {};

  // 解析默认运行属性（defaultRunProperties）
  const defRPr = getFirstChildByTagNS(paraPr, 'defRPr', NS.a);
  if (defRPr) {
    // 字体大小
    const sz = getFirstChildByTagNS(defRPr, 'sz', NS.a);
    if (sz) {
      // PPTX中sz单位是百分之一磅（1/100 pt）
      // 需要将磅转换为像素：1 pt = 4/3 px（96 DPI下）
      const ptSize = parseInt(sz.getAttribute('val') || '0', 10) / 100;
      style.fontSize = ptSize * (4 / 3);
    }

    // 字体颜色
    const solidFill = getFirstChildByTagNS(defRPr, 'solidFill', NS.a);
    if (solidFill) {
      const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
      if (srgbClr?.getAttribute('val')) {
        style.color = `#${srgbClr.getAttribute('val')}`;
      }

      // 主题颜色
      const schemeClr = getFirstChildByTagNS(solidFill, 'schemeClr', NS.a);
      if (schemeClr?.getAttribute('val')) {
        style.color = schemeClr.getAttribute('val');
      }
    }

    // 字体名称
    const latin = getFirstChildByTagNS(defRPr, 'latin', NS.a);
    if (latin?.getAttribute('typeface')) {
      style.fontFamily = latin.getAttribute('typeface');
    }

    // 粗体
    const b = getFirstChildByTagNS(defRPr, 'b', NS.a);
    if (b) {
      style.fontWeight = b.getAttribute('val') === '1' || b.getAttribute('val') === 'true' ? 'bold' : 'normal';
    }

    // 斜体
    const i = getFirstChildByTagNS(defRPr, 'i', NS.a);
    if (i) {
      style.italic = i.getAttribute('val') === '1' || i.getAttribute('val') === 'true';
    }

    // 下划线
    const u = getFirstChildByTagNS(defRPr, 'u', NS.a);
    if (u) {
      style.underline = u.getAttribute('val') || 'sng';
    }

    // 删除线
    const strike = getFirstChildByTagNS(defRPr, 'strike', NS.a);
    if (strike) {
      style.strike = strike.getAttribute('val') === '1' || strike.getAttribute('val') === 'true';
    }
  }

  // 解析对齐方式
  const align = paraPr.getAttribute('algn');
  if (align) {
    style.align = align;
  }

  return Object.keys(style).length > 0 ? style : null;
}

  /**
   * 获取默认文本样式
   * @param placeholderType 占位符类型
   * @param placeholderSize 占位符大小（quarter, half, full等）
   * @returns 默认样式对象
   */
  function getDefaultTextStyle(placeholderType?: string, placeholderSize?: string): any {
    // 根据占位符大小设置默认字体大小
    let fontSize = 14; // 默认14px
    
    if (placeholderSize) {
      switch (placeholderSize.toLowerCase()) {
        case 'quarter':
          fontSize = 40; // 30pt * 4/3 = 40px
          break;
        case 'half':
          fontSize = 32; // 24pt * 4/3 = 32px
          break;
        case 'full':
          fontSize = 24; // 18pt * 4/3 = 24px
          break;
        default:
          fontSize = 14;
      }
    } else if (placeholderType === 'title') {
      fontSize = 44; // 标题默认44px (33pt)
    } else if (placeholderType === 'body') {
      fontSize = 18; // 正文默认18px (13.5pt)
    }
    
    return {
      fontSize,
      color: '#333333',
      fontWeight: 'normal',
      fontFamily: 'Arial',
      align: 'left'
    };
  }

/**
 * 合并样式
 * 优先级：newStyle > baseStyle
 * @param baseStyle 基础样式
 * @param newStyle 新样式
 * @returns 合并后的样式
 */
export function mergeStyles(baseStyle: any, newStyle: any): any {
  if (!baseStyle) return newStyle;
  if (!newStyle) return baseStyle;

  return {
    ...baseStyle,
    ...newStyle
  };
}

/**
 * 解析颜色（支持主题颜色）
 * @param colorNode 颜色节点
 * @param themeColors 主题颜色映射
 * @returns 颜色值
 */
export function resolveColor(colorNode: any, themeColors?: Record<string, string>): string | undefined {
  if (!colorNode) return undefined;

  // 1. 检查纯色（srgbClr）
  const srgbClr = getFirstChildByTagNS(colorNode, 'srgbClr', NS.a);
  if (srgbClr?.getAttribute('val')) {
    return `#${srgbClr.getAttribute('val')}`;
  }

  // 2. 检查主题颜色（schemeClr）
  const schemeClr = getFirstChildByTagNS(colorNode, 'schemeClr', NS.a);
  if (schemeClr?.getAttribute('val') && themeColors) {
    const schemeRef = schemeClr.getAttribute('val');
    return themeColors[schemeRef];
  }

  return undefined;
}

/**
 * 创建样式继承上下文
 * @param slide 幻灯片对象
 * @param layout 布局对象
 * @param master 母版对象
 * @param theme 主题对象
 * @returns 样式继承上下文
 */
export function createStyleContext(
  slide?: any,
  layout?: SlideLayoutResult,
  master?: MasterSlideResult,
  theme?: any
): StyleContext {
  return {
    slide,
    layout,
    master,
    theme
  };
}

/**
 * 应用样式继承到 slide 的所有元素
 * @param slide 幻灯片对象
 * @param layout 布局对象
 * @param master 母版对象
 * @param theme 主题对象
 */
export function applyStyleInheritance(
  slide: any,
  layout?: SlideLayoutResult,
  master?: MasterSlideResult,
  theme?: any
): void {
  if (!slide.elements || !Array.isArray(slide.elements)) {
    return;
  }

  const context = createStyleContext(slide, layout, master, theme);

  // 遍历所有元素并应用样式继承
  slide.elements.forEach((element: any) => {
    // 只处理形状元素（可能包含占位符）
    if (element.type === 'shape' || element.type === 'text') {
      applyElementStyle(element, context);
    }
  });
}

  /**
   * 应用样式到单个元素
   * @param element 元素对象
   * @param context 样式继承上下文
   */
  function applyElementStyle(element: any, context: StyleContext): void {
    // 存储占位符大小属性，用于后续样式计算
    if (element.isPlaceholder && element.rawNode) {
      const phNode = element.rawNode.querySelector('p\\:ph, ph');
      if (phNode) {
        element.phSize = phNode.getAttribute('sz');
      }
    }
    
    // 如果元素有textStyle，需要为缺失的属性应用继承样式
    if (element.textStyle && element.textStyle.length > 0) {
      // 获取继承的样式（无论是否是占位符）
      const inheritedStyle = getPlaceholderStyle(element, context);

      if (inheritedStyle) {
        // 为每个文本运行应用缺失的样式
        element.textStyle.forEach((run: any) => {
          if (!run.fontSize && inheritedStyle.fontSize) {
            run.fontSize = inheritedStyle.fontSize;
          }
          if (!run.color && inheritedStyle.color) {
            run.color = inheritedStyle.color;
          }
          if (!run.fontFamily && inheritedStyle.fontFamily) {
            run.fontFamily = inheritedStyle.fontFamily;
          }
          if (!run.bold && inheritedStyle.fontWeight === 'bold') {
            run.bold = true;
          }
          if (!run.italic && inheritedStyle.italic) {
            run.italic = inheritedStyle.italic;
          }
        });
      }
      return;
    }

    // 如果元素是占位符且没有textStyle，获取继承的样式
    if (element.isPlaceholder && element.placeholderType) {
      const inheritedStyle = getPlaceholderStyle(element, context);
      if (inheritedStyle) {
        // 合并继承的样式到元素的 style 属性
        element.style = {
          ...element.style,
          ...inheritedStyle
        };
      }
    }
  }

