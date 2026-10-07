/**
 * 单位换算：pt/px/in。与 core/constants.ts 的 FONT_SIZE_FACTOR / DPI 保持一致，
 * 不复造魔法数（96 DPI 下 1pt = 96/72 px）。
 */
import { FONT_SIZE_FACTOR, DPI } from '../core/constants';

/** pt → px（96dpi） */
export const ptToPx = (pt: number): number => (Number(pt) || 0) * FONT_SIZE_FACTOR;
/** px → pt */
export const pxToPt = (px: number): number => (Number(px) || 0) / FONT_SIZE_FACTOR;
/** in → px */
export const inToPx = (i: number): number => (Number(i) || 0) * DPI;
