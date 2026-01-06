/**
 * 样式继承工具
 * 实现完整的样式继承链条：Slide > Layout > Master
 * 对应 PPTXjs 的样式合并逻辑
 */

import type { PptRect, PptStyle } from '../types';
import type { Placeholder } from '../core/types';
import type { BaseElement } from '../elements/BaseElement';
import { emu2px } from '../utils';

/**
 * 合并占位符样式
 * 优先级：Slide元素 > Layout占位符 > Master占位符
 *
 * @param element 幻灯片元素
 * @param layoutPlaceholder 布局占位符
 * @param masterPlaceholder 母版占位符（可选）
 * @returns 合并后的样式和位置
 */
export function mergePlaceholderStyles(
  element: BaseElement | null,
  layoutPlaceholder?: Placeholder,
  masterPlaceholder?: Placeholder
): { rect: PptRect; style: PptStyle; alignmentClass?: string } {
  // 默认位置和样式
  const result = {
    rect: element?.rect || { x: 0, y: 0, width: 0, height: 0 },
    style: element?.style || {
      fontSize: 14,
      color: '#333333',
      fontWeight: 'normal',
      textAlign: 'left',
      backgroundColor: 'transparent',
      borderColor: '#000000',
      borderWidth: 1
    },
    alignmentClass: ''
  };

  // 应用布局占位符位置（优先级2）
  if (layoutPlaceholder) {
    result.rect = {
      x: emu2px(layoutPlaceholder.rect.x),
      y: emu2px(layoutPlaceholder.rect.y),
      width: emu2px(layoutPlaceholder.rect.width),
      height: emu2px(layoutPlaceholder.rect.height)
    };

    // 生成对齐类
    if (layoutPlaceholder.hAlign || layoutPlaceholder.vAlign) {
      result.alignmentClass = getAlignmentClass(
        layoutPlaceholder.hAlign,
        layoutPlaceholder.vAlign
      );
    }
  }

  // 如果元素有自定义位置，优先使用（优先级1）
  if (element?.rect) {
    result.rect = element.rect;
  }

  return result;
}

/**
 * 获取对齐类名
 * @param hAlign 水平对齐
 * @param vAlign 垂直对齐
 * @returns CSS类名
 */
function getAlignmentClass(
  hAlign?: 'left' | 'center' | 'right',
  vAlign?: 'top' | 'middle' | 'bottom'
): string {
  const hMap: Record<string, string> = {
    'left': 'h-left',
    'center': 'h-mid',
    'right': 'h-right'
  };
  const vMap: Record<string, string> = {
    'top': 'v-up',
    'middle': 'v-mid',
    'bottom': 'v-down'
  };

  const hClass = hAlign ? hMap[hAlign] : '';
  const vClass = vAlign ? vMap[vAlign] : '';

  if (hClass && vClass) {
    // 组合类名（如 'up-center'）
    return `${vClass.replace('v-', '')}_${hClass.replace('h-', '')}`;
  }

  return `${vClass} ${hClass}`.trim();
}

/**
 * 合并背景样式
 * 优先级：Slide背景 > Layout背景 > Master背景
 *
 * @param slideBackground 幻灯片背景
 * @param layoutBackground 布局背景
 * @param masterBackground 母版背景
 * @returns 合并后的背景
 */
export function mergeBackgroundStyles(
  slideBackground?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string },
  layoutBackground?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string },
  masterBackground?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string }
): { type: 'color' | 'image' | 'none'; value?: string; relId?: string } {
  // 幻灯片背景优先
  if (slideBackground && slideBackground.type !== 'none') {
    return slideBackground;
  }

  // 其次布局背景
  if (layoutBackground && layoutBackground.type !== 'none') {
    return layoutBackground;
  }

  // 最后母版背景
  if (masterBackground && masterBackground.type !== 'none') {
    return masterBackground;
  }

  // 默认白色
  return { type: 'color', value: '#ffffff' };
}

/**
 * 合并文本样式
 * 优先级：Slide元素 > Layout样式 > Master样式
 *
 * @param elementStyle 元素样式
 * @param layoutStyle 布局样式（可选）
 * @param masterStyle 母版样式（可选）
 * @returns 合并后的样式
 */
export function mergeTextStyles(
  elementStyle?: Partial<PptStyle>,
  layoutStyle?: Partial<PptStyle>,
  masterStyle?: Partial<PptStyle>
): PptStyle {
  const defaults: PptStyle = {
    fontSize: 14,
    color: '#333333',
    fontWeight: 'normal',
    textAlign: 'left',
    backgroundColor: 'transparent',
    borderColor: '#000000',
    borderWidth: 1
  };

  // 从 Master 开始合并
  let merged: PptStyle = { ...defaults, ...masterStyle };

  // 应用 Layout 样式（覆盖 Master）
  merged = { ...merged, ...layoutStyle };

  // 应用 Element 样式（最高优先级）
  merged = { ...merged, ...elementStyle };

  return merged;
}

/**
 * 查找占位符
 * @param placeholders 占位符数组
 * @param type 占位符类型
 * @param idx 占位符索引（可选）
 * @returns 匹配的占位符或 undefined
 */
export function findPlaceholder(
  placeholders: Placeholder[] | undefined,
  type: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other',
  idx?: number
): Placeholder | undefined {
  if (!placeholders || placeholders.length === 0) {
    return undefined;
  }

  // 如果指定了索引，优先匹配索引
  if (idx !== undefined) {
    const byIndex = placeholders.find(p => p.idx === idx);
    if (byIndex) {
      return byIndex;
    }
  }

  // 按类型匹配
  const byType = placeholders.find(p => p.type === type);
  return byType;
}
