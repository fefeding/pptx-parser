/**
 * 布局工具类
 * 对应 PPTXjs 的布局类映射功能
 */

import { emu2px } from '.';

/**
 * 对齐类映射表（PPTXjs 核心 CSS 规则）
 */
export const ALIGNMENT_CLASSES = {
  // 垂直对齐
  v_up: 'v-up',
  v_mid: 'v-mid',
  v_down: 'v-down',
  // 水平对齐
  h_left: 'h-left',
  h_mid: 'h-mid',
  h_right: 'h-right',
  // 复合对齐（高频组合）
  up_center: 'up-center',
  up_left: 'up-left',
  up_right: 'up-right',
  center_center: 'center-center',
  center_left: 'center-left',
  center_right: 'center-right',
  down_center: 'down-center',
  down_left: 'down-left',
  down_right: 'down-right'
} as const;

/**
 * 获取垂直对齐类名
 * @param vAlign 垂直对齐方式
 * @returns CSS类名
 */
export function getVerticalAlignClass(vAlign?: 'top' | 'middle' | 'bottom'): string {
  if (!vAlign) return '';
  const vMap: Record<string, string> = {
    'top': ALIGNMENT_CLASSES.v_up,
    'middle': ALIGNMENT_CLASSES.v_mid,
    'bottom': ALIGNMENT_CLASSES.v_down
  };
  return vMap[vAlign] || '';
}

/**
 * 获取水平对齐类名
 * @param hAlign 水平对齐方式
 * @returns CSS类名
 */
export function getHorizontalAlignClass(hAlign?: 'left' | 'center' | 'right'): string {
  if (!hAlign) return '';
  const hMap: Record<string, string> = {
    'left': ALIGNMENT_CLASSES.h_left,
    'center': ALIGNMENT_CLASSES.h_mid,
    'right': ALIGNMENT_CLASSES.h_right
  };
  return hMap[hAlign] || '';
}

/**
 * 获取组合对齐类名
 * @param hAlign 水平对齐
 * @param vAlign 垂直对齐
 * @returns 组合CSS类名
 */
export function getAlignmentClass(
  hAlign?: 'left' | 'center' | 'right',
  vAlign?: 'top' | 'middle' | 'bottom'
): string {
  const hClass = getHorizontalAlignClass(hAlign);
  const vClass = getVerticalAlignClass(vAlign);

  if (hClass && vClass) {
    // 返回组合类名（如 'up-center', 'center-center'）
    return `${vClass.replace('v-', '')}_${hClass.replace('h-', '')}`;
  }

  return `${vClass} ${hClass}`.trim();
}

/**
 * 生成占位符布局样式
 * 对应 PPTXjs 的 .block 样式生成
 * @param placeholder 占位符对象
 * @returns CSS 样式对象
 */
export function getPlaceholderLayoutStyle(
  placeholder: {
    rect: { x: number; y: number; width: number; height: number };
    hAlign?: 'left' | 'center' | 'right';
    vAlign?: 'top' | 'middle' | 'bottom';
  }
): { style: string; className: string } {
  const { rect, hAlign, vAlign } = placeholder;

  // 转换 EMU 到 px
  const x = emu2px(rect.x);
  const y = emu2px(rect.y);
  const width = emu2px(rect.width);
  const height = emu2px(rect.height);

  // 生成定位样式
  const style = [
    `position: absolute`,
    `top: ${y}px`,
    `left: ${x}px`,
    `width: ${width}px`,
    `height: ${height}px`
  ];

  // 生成对齐类名
  const className = getAlignmentClass(hAlign, vAlign);

  return {
    style: style.join('; '),
    className: className ? `block ${className}` : 'block'
  };
}

/**
 * 获取幻灯片容器样式
 * 对应 PPTXjs 的 .slide 样式
 * @param width 宽度（EMU）
 * @param height 高度（EMU）
 * @param background 背景对象
 * @param scale 缩放比例（可选）
 * @returns CSS 样式字符串
 */
export function getSlideContainerStyle(
  width: number,
  height: number,
  background?: { type: 'color' | 'image' | 'none'; value?: string; relId?: string },
  scale?: number
): string {
  const styles = [
    `position: relative`,
    `overflow: hidden`,
    `margin: 0 auto`,
    `width: ${emu2px(width)}px`,
    `height: ${emu2px(height)}px`
  ];

  // 背景样式
  if (background) {
    if (background.type === 'color' && background.value) {
      styles.push(`background: ${background.value}`);
    } else if (background.type === 'image' && background.value) {
      styles.push(`background: url('${background.value}') no-repeat center/cover`);
    }
  } else {
    styles.push(`background: #ffffff`);
  }

  // 缩放样式
  if (scale && scale !== 1) {
    styles.push(`transform: scale(${scale})`);
    styles.push(`transform-origin: center center`);
  }

  return styles.join('; ');
}

/**
 * 在占位符列表中查找匹配的占位符
 * @param placeholders 占位符数组
 * @param type 占位符类型
 * @param idx 占位符索引，可选
 * @returns 匹配的占位符或 undefined
 */
export function findPlaceholder(
  placeholders: Array<{ type?: string | 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other'; idx?: number } & Record<string, any>>,
  type: string,
  idx?: number
): (Record<string, any> & { type?: string | 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other'; idx?: number }) | undefined {
  return placeholders.find(p => {
    if (p.type !== type) return false;
    if (idx !== undefined && p.idx !== idx) return false;
    return true;
  });
}

/**
 * 合并占位符的基础样式与继承样式
 * @param baseStyle 占位符自身的样式对象
 * @param inheritedStyle 从布局或母版继承的样式对象
 * @returns 合并后的样式对象
 */
export function mergePlaceholderStyles(baseStyle: Record<string, any>, inheritedStyle: Record<string, any>): Record<string, any> {
  return {
    ...baseStyle,
    ...inheritedStyle
  };
}
