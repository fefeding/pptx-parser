/**
 * 文本解析工具
 * 解析 PPTX 文本内容和样式
 */

import { NS } from '../constants';
import { getFirstChildByTagNS, getChildrenByTagNS } from './xml-parser';

/**
 * 解析文本内容（标准路径：a:p -> a:r -> a:t）
 * 支持多段落、多文本运行拼接，中英文混合文本自动合并
 * @param txBody 文本框节点（p:txBody）
 * @returns 完整文本字符串
 */
export function parseTextContent(txBody: Element): string {
  const paragraphs = getChildrenByTagNS(txBody, 'p', NS.a);
  const textParts: string[] = [];

  paragraphs.forEach(p => {
    const runs = getChildrenByTagNS(p, 'r', NS.a);
    runs.forEach(r => {
      const t = getFirstChildByTagNS(r, 't', NS.a);
      if (t?.textContent) {
        textParts.push(t.textContent);
      }
    });
  });

  return textParts.join('');
}

/**
 * 解析文本内容带样式（扩展版本）
 * @param txBody 文本框节点
 * @returns 文本运行数组，包含文本和样式
 */
export function parseTextWithStyle(txBody: Element): Array<{ text: string; style?: Record<string, unknown> }> {
  const paragraphs = getChildrenByTagNS(txBody, 'p', NS.a);
  const result: Array<{ text: string; style?: Record<string, unknown> }> = [];

  paragraphs.forEach(p => {
    const runs = getChildrenByTagNS(p, 'r', NS.a);
    runs.forEach(r => {
      const t = getFirstChildByTagNS(r, 't', NS.a);
      const rPr = getFirstChildByTagNS(r, 'rPr', NS.a);

      if (t?.textContent) {
        const style = rPr ? parseTextStyle(rPr) : undefined;
        result.push({ text: t.textContent, style });
      }
    });
  });

  return result;
}

/**
 * 解析文本样式属性（a:rPr节点）
 * @param rPr 文本运行属性节点
 * @returns 样式对象
 */
export function parseTextStyle(rPr: Element): Record<string, unknown> {
  const style: Record<string, unknown> = {};

  // 字体大小 (sz单位：1/100 pt)
  const sz = rPr.getAttribute('sz');
  if (sz) {
    style.fontSize = parseInt(sz, 10) / 100;
  }

  // 字体名称
  const latin = getFirstChildByTagNS(rPr, 'latin', NS.a);
  if (latin?.getAttribute('typeface')) {
    style.fontFamily = latin.getAttribute('typeface');
  }

  // 加粗
  if (rPr.getAttribute('b') === '1' || rPr.getAttribute('b') === 'true') {
    style.fontWeight = 'bold';
  }

  // 斜体
  if (rPr.getAttribute('i') === '1' || rPr.getAttribute('i') === 'true') {
    style.fontStyle = 'italic';
  }

  // 下划线
  if (rPr.getAttribute('u') === 'sng' || rPr.getAttribute('u') === '1') {
    style.textDecoration = 'underline';
  }

  // 删除线
  if (rPr.getAttribute('strike') === 'sngStrike' || rPr.getAttribute('strike') === '1') {
    style.textDecoration = (style.textDecoration ? `${style.textDecoration} line-through` : 'line-through') as string;
  }

  // 字体颜色
  const solidFill = getFirstChildByTagNS(rPr, 'solidFill', NS.a);
  if (solidFill) {
    const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
    if (srgbClr?.getAttribute('val')) {
      style.color = `#${srgbClr.getAttribute('val')}`;
    }
  }

  return style;
}
