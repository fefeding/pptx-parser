/**
 * 位置和尺寸解析工具
 * 解析 PPTX 元素的位置和尺寸信息
 */

import { NS } from '../constants';
import { getFirstChildByTagNS } from './xml-parser';
import { emu2px } from './convert';

/**
 * 解析位置和尺寸（标准路径：a:xfrm -> a:off + a:ext）
 * @param spPr 形状属性节点（p:spPr）
 * @returns 位置尺寸对象（像素单位）
 */
export function parsePosition(spPr: Element): { x: number; y: number; width: number; height: number } {
  const xfrm = getFirstChildByTagNS(spPr, 'xfrm', NS.a);
  if (!xfrm) {
    return { x: 0, y: 0, width: 0, height: 0 };
  }

  const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
  const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

  const x = off ? emu2px(off.getAttribute('x') || '0') : 0;
  const y = off ? emu2px(off.getAttribute('y') || '0') : 0;
  const width = ext ? emu2px(ext.getAttribute('cx') || '0') : 0;
  const height = ext ? emu2px(ext.getAttribute('cy') || '0') : 0;

  return { x, y, width, height };
}
