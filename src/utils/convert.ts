/**
 * 单位转换工具
 * 提供 EMU 和像素之间的转换
 */

import { EMU_TO_PIXEL_RATIO, PIXEL_TO_EMU_RATIO } from '../constants';

/**
 * EMU单位转像素（核心转换函数）
 * PPTX内部使用EMU (English Metric Unit)，1 EMU = 1/914400 英寸
 * 转换为像素时保留两位小数，避免精度溢出
 * @param emu EMU值（字符串或数字）
 * @returns 像素值
 */
export function emu2px(emu: string | number): number {
  const numEmu = typeof emu === 'string' ? parseInt(emu || '0', 10) : emu;
  return Math.round(numEmu * EMU_TO_PIXEL_RATIO * 100) / 100;
}

/**
 * 像素转EMU单位（用于序列化）
 * @param px 像素值
 * @returns EMU值
 */
export function px2emu(px: number): number {
  return Math.round(px * PIXEL_TO_EMU_RATIO);
}
