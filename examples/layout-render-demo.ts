/**
 * 示例：如何使用布局和样式继承功能
 * 对应 PPTXjs 的渲染逻辑
 */

import type { SlideParseResult } from '../core/types';
import { mergePlaceholderStyles, mergeBackgroundStyles, findPlaceholder } from '../utils/style-inheritance';
import { getPlaceholderLayoutStyle, getSlideContainerStyle } from '../utils/layout-utils';
import type { BaseElement } from '../elements/BaseElement';

/**
 * 渲染幻灯片为 HTML
 * 完整复刻 PPTXjs 的渲染逻辑
 */
export function renderSlideHTML(slide: SlideParseResult): string {
  // 1. 获取布局和母版信息（已在 parser.ts 中附加）
  const layout = (slide as any).layout;
  const master = (slide as any).master;

  // 2. 生成幻灯片容器样式（.slide）
  // 对应 PPTXjs：.slide 容器，承载母版样式
  const slideStyle = getSlideContainerStyle(
    9144000, // 默认宽度 10英寸 = 960px
    6858000, // 默认高度 7.5英寸 = 720px (16:9)
    slide.background as any
  );

  // 3. 渲染元素，应用布局占位符样式
  const elementsHTML = slide.elements
    .map((element: BaseElement) => {
      // 查找对应的布局占位符
      const layoutPlaceholder = findPlaceholder(
        layout?.placeholders,
        element.type === 'shape' ? 'body' : 'other'
      );

      // 查找对应的母版占位符
      const masterPlaceholder = findPlaceholder(
        master?.placeholders,
        element.type === 'shape' ? 'body' : 'other'
      );

      // 合并样式（优先级：Slide > Layout > Master）
      const { rect, style, alignmentClass } = mergePlaceholderStyles(
        element,
        layoutPlaceholder,
        masterPlaceholder
      );

      // 生成布局类样式
      const placeholderStyle = layoutPlaceholder
        ? getPlaceholderLayoutStyle(layoutPlaceholder)
        : { style: '', className: 'block' };

      // 渲染元素 HTML
      return `
        <div class="${placeholderStyle.className}" style="${placeholderStyle.style}">
          <div class="content slide-prgrph" style="text-align: ${style.textAlign || 'left'};">
            ${element.toHTML()}
          </div>
        </div>
      `;
    })
    .join('\n');

  // 4. 生成完整 HTML
  // 对应 PPTXjs：.slide > .block > .content 结构
  return `
    <div class="slide" style="${slideStyle}">
      ${elementsHTML}
    </div>
  `;
}

/**
 * 生成 CSS 样式表
 * 对应 PPTXjs 的布局类 CSS 规则
 */
export function generateLayoutCSS(): string {
  return `
    /* 幻灯片容器基础样式 */
    .slide {
      position: relative;
      overflow: hidden;
      margin: 0 auto;
      box-sizing: border-box;
    }

    /* 占位符基础样式（对应版式占位符） */
    .slide div.block {
      position: absolute;
      top: 0;
      left: 0;
      width: 100%;
      line-height: 1;
      box-sizing: border-box;
      display: flex;
      flex-direction: column;
    }

    /* 垂直对齐类（映射版式垂直规则） */
    .slide div.v-up { justify-content: flex-start; }
    .slide div.v-mid { justify-content: center; }
    .slide div.v-down { justify-content: flex-end; }

    /* 水平对齐类（映射版式水平规则） */
    .slide div.h-left { align-items: flex-start; text-align: left; }
    .slide div.h-mid { align-items: center; text-align: center; }
    .slide div.h-right { align-items: flex-end; text-align: right; }

    /* 复合对齐类（高频组合） */
    .slide div.up-left { justify-content: flex-start; align-items: flex-start; }
    .slide div.up-center { justify-content: flex-start; align-items: center; text-align: center; }
    .slide div.up-right { justify-content: flex-start; align-items: flex-end; text-align: right; }
    .slide div.center-left { justify-content: center; align-items: flex-start; }
    .slide div.center-center { justify-content: center; align-items: center; text-align: center; }
    .slide div.center-right { justify-content: center; align-items: flex-end; text-align: right; }
    .slide div.down-left { justify-content: flex-end; align-items: flex-start; }
    .slide div.down-center { justify-content: flex-end; align-items: center; text-align: center; }
    .slide div.down-right { justify-content: flex-end; align-items: flex-end; text-align: right; }

    /* 内容容器 */
    .content {
      width: 100%;
      height: 100%;
      box-sizing: border-box;
    }

    /* 文本样式（全局字体） */
    .slide-prgrph {
      font-family: Arial, sans-serif;
      font-size: 14px;
      color: #333333;
    }
  `;
}
