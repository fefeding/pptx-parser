/**
 * 主题元素类
 * 解析和管理PPTX主题，包括颜色方案、字体方案和效果方案
 */

import { BaseElement } from './BaseElement';
import type { ThemeResult, ThemeColors } from '../core/types';

/**
 * 字体方案
 */
export interface FontScheme {
  /** 主要字体（标题等） */
  majorFont?: {
    latin?: string;
    ea?: string;
    cs?: string;
    [script: string]: string | undefined;
  };
  /** 次要字体（正文等） */
  minorFont?: {
    latin?: string;
    ea?: string;
    cs?: string;
    [script: string]: string | undefined;
  };
}

/**
 * 效果方案
 */
export interface EffectScheme {
  /** 填充效果 */
  fillStyles?: any[];
  /** 线条效果 */
  lineStyles?: any[];
  /** 特效样式 */
  effectStyles?: any[];
  /** 背景填充样式 */
  bgFillStyles?: any[];
}

/**
 * 主题元素类
 */
export class ThemeElement extends BaseElement {
  type: 'theme' = 'theme';

  /** 主题名称 */
  name: string;

  /** 颜色方案 */
  colors: ThemeColors;

  /** 字体方案 */
  fonts?: FontScheme;

  /** 效果方案 */
  effects?: EffectScheme;

  /** 主题ID (用于生成CSS类前缀) */
  themeId: string;

  constructor(
    id: string,
    name: string,
    colors: ThemeColors,
    themeId?: string,
    fonts?: FontScheme,
    effects?: EffectScheme
  ) {
    super(id, 'theme', { x: 0, y: 0, width: 0, height: 0 }, {}, {}, {});
    this.name = name;
    this.colors = colors;
    this.fonts = fonts;
    this.effects = effects;
    // 使用主题名称或ID生成CSS友好的前缀
    this.themeId = themeId || this.sanitizeThemeName(name);
  }

  /**
   * 从 ThemeResult 创建 ThemeElement
   */
  static fromResult(result: ThemeResult, themeName: string = 'theme1'): ThemeElement {
    return new ThemeElement(
      `theme_${themeName}`,
      themeName,
      result.colors,
      themeName
    );
  }

  /**
   * 生成主题样式的CSS类前缀
   */
  getThemeClassPrefix(): string {
    return `theme-${this.themeId}`;
  }

  /**
   * 生成完整的CSS类名
   */
  getThemeClass(suffix: string): string {
    return `${this.getThemeClassPrefix()}-${suffix}`;
  }

  /**
   * 将主题名称转换为CSS友好的格式
   */
  private sanitizeThemeName(name: string): string {
    return name
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '-')
      .replace(/-+/g, '-')
      .replace(/^-|-$/g, '');
  }

  /**
   * 生成主题CSS样式
   */
  generateThemeCSS(): string {
    const prefix = this.getThemeClassPrefix();
    const css: string[] = [];

    // 主题颜色变量
    css.push(`/* ===== ${this.name} 主题样式 ===== */`);
    css.push(`.${prefix} {`);

    // CSS变量定义 (使用 --theme-* 前缀)
    if (this.colors.bg1) css.push(`  --${prefix}-bg1: ${this.colors.bg1};`);
    if (this.colors.tx1) css.push(`  --${prefix}-tx1: ${this.colors.tx1};`);
    if (this.colors.bg2) css.push(`  --${prefix}-bg2: ${this.colors.bg2};`);
    if (this.colors.tx2) css.push(`  --${prefix}-tx2: ${this.colors.tx2};`);
    if (this.colors.accent1) css.push(`  --${prefix}-accent1: ${this.colors.accent1};`);
    if (this.colors.accent2) css.push(`  --${prefix}-accent2: ${this.colors.accent2};`);
    if (this.colors.accent3) css.push(`  --${prefix}-accent3: ${this.colors.accent3};`);
    if (this.colors.accent4) css.push(`  --${prefix}-accent4: ${this.colors.accent4};`);
    if (this.colors.accent5) css.push(`  --${prefix}-accent5: ${this.colors.accent5};`);
    if (this.colors.accent6) css.push(`  --${prefix}-accent6: ${this.colors.accent6};`);
    if (this.colors.hlink) css.push(`  --${prefix}-hlink: ${this.colors.hlink};`);
    if (this.colors.folHlink) css.push(`  --${prefix}-fol-hlink: ${this.colors.folHlink};`);

    css.push('}');
    css.push('');

    // 颜色类 (类似于 Bootstrap 的颜色工具类)
    const colorClasses = [
      'bg1', 'tx1', 'bg2', 'tx2',
      'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
      'hlink', 'fol-hlink'
    ];

    colorClasses.forEach(color => {
      const colorVar = color === 'fol-hlink' ? 'folHlink' : color;
      const colorValue = this.colors[colorVar as keyof ThemeColors];
      if (colorValue) {
        // 背景色类
        css.push(`.${prefix}-bg-${color} {`);
        css.push(`  background-color: var(--${prefix}-${color});`);
        css.push('}');

        // 文本色类
        css.push(`.${prefix}-text-${color} {`);
        css.push(`  color: var(--${prefix}-${color});`);
        css.push('}');

        // 边框色类
        css.push(`.${prefix}-border-${color} {`);
        css.push(`  border-color: var(--${prefix}-${color});`);
        css.push('}');
      }
    });

    css.push('');

    // 字体方案
    if (this.fonts) {
      css.push(`/* ${this.name} 字体方案 */`);

      if (this.fonts.majorFont) {
        css.push(`.${prefix}-font-major {`);
        if (this.fonts.majorFont.latin) {
          css.push(`  font-family: "${this.fonts.majorFont.latin}", sans-serif;`);
        }
        if (this.fonts.majorFont.ea) {
          css.push(`  font-family: "${this.fonts.majorFont.ea}", sans-serif;`);
        }
        css.push('}');
      }

      if (this.fonts.minorFont) {
        css.push(`.${prefix}-font-minor {`);
        if (this.fonts.minorFont.latin) {
          css.push(`  font-family: "${this.fonts.minorFont.latin}", sans-serif;`);
        }
        if (this.fonts.minorFont.ea) {
          css.push(`  font-family: "${this.fonts.minorFont.ea}", sans-serif;`);
        }
        css.push('}');
      }

      css.push('');
    }

    // 预设样式类
    css.push(`/* ${this.name} 预设样式类 */`);

    // 标题样式 (使用主题文字色和字体)
    css.push(`.${prefix}-title {`);
    css.push(`  color: var(--${prefix}-tx1);`);
    if (this.fonts?.majorFont?.latin) {
      css.push(`  font-family: "${this.fonts.majorFont.latin}", sans-serif;`);
    }
    css.push('}');

    // 正文样式 (使用次要文字色和字体)
    css.push(`.${prefix}-body {`);
    css.push(`  color: var(--${prefix}-tx2);`);
    if (this.fonts?.minorFont?.latin) {
      css.push(`  font-family: "${this.fonts.minorFont.latin}", sans-serif;`);
    }
    css.push('}');

    // 链接样式
    if (this.colors.hlink) {
      css.push(`.${prefix}-link {`);
      css.push(`  color: var(--${prefix}-hlink);`);
      css.push(`  text-decoration: none;`);
      css.push('}');
    }

    if (this.colors.folHlink) {
      css.push(`.${prefix}-link:visited {`);
      css.push(`  color: var(--${prefix}-fol-hlink);`);
      css.push('}');
    }

    // 强调色样式
    ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].forEach((accent, index) => {
      const colorValue = this.colors[accent as keyof ThemeColors];
      if (colorValue) {
        css.push(`.${prefix}-accent-${index + 1} {`);
        css.push(`  color: var(--${prefix}-${accent});`);
        css.push('}');
      }
    });

    return css.join('\n');
  }

  /**
   * 获取指定颜色值
   */
  getColor(colorKey: keyof ThemeColors): string | undefined {
    return this.colors[colorKey];
  }

  /**
   * 获取主字体
   */
  getMajorFont(): string | undefined {
    return this.fonts?.majorFont?.latin || this.fonts?.majorFont?.ea;
  }

  /**
   * 获取次要字体
   */
  getMinorFont(): string | undefined {
    return this.fonts?.minorFont?.latin || this.fonts?.minorFont?.ea;
  }

  /**
   * 转换为HTML字符串 (主题元素通常不需要输出HTML)
   */
  toHTML(): string {
    // 主题元素本身不输出HTML，而是通过CSS样式应用
    return '';
  }
}
