/**
 * 单位转换工具模块
 * 对齐PPTXjs的单位转换逻辑
 * 
 * EMU (English Metric Unit) 是PPTX中使用的内部单位
 * 1 inch = 914400 EMU
 * 1 cm = 360000 EMU
 * 1 inch = 96 px (在96DPI下)
 * 
 * 因此：1 EMU = 96 / 914400 px = 0.000105 px
 */

/**
 * EMU转PX单位转换
 * 对齐PPTXjs的slideFactor = 96 / 914400
 * 
 * @param emu - EMU单位值
 * @returns 像素值
 */
export function emu2px(emu: number): number {
  if (typeof emu !== 'number' || isNaN(emu)) {
    return 0;
  }
  // PPTXjs使用：parseInt(emu * slideFactor)
  // slideFactor = 96 / 914400
  return Math.round(emu * (96 / 914400));
}

/**
 * PX转EMU单位转换
 * 对齐PPTXjs的逆转换逻辑
 * 
 * @param px - 像素值
 * @returns EMU单位值
 */
export function px2emu(px: number): number {
  if (typeof px !== 'number' || isNaN(px)) {
    return 0;
  }
  // 逆转换：px / slideFactor
  return Math.round(px / (96 / 914400));
}

/**
 * EMU转PT单位转换（用于字体大小）
 * PPTX字体大小单位：百分之一磅
 * 1 pt = 100 font units
 * 转换因子：fontSizeFactor = 4 / 3.2 = 1.25
 * 
 * @param fontUnits - PPTX字体单位（百分之一磅）
 * @returns 像素值
 */
export function fontUnits2px(fontUnits: number): number {
  if (typeof fontUnits !== 'number' || isNaN(fontUnits)) {
    return 0;
  }
  // PPTXjs使用：parseInt(sz) / 100 * fontSizeFactor
  const fontSizeFactor = 4 / 3.2;
  return (fontUnits / 100) * fontSizeFactor;
}

/**
 * PT转EMU单位转换
 * 1 pt = 12700 EMU
 * 
 * @param pt - 磅值
 * @returns EMU单位值
 */
export function pt2emu(pt: number): number {
  if (typeof pt !== 'number' || isNaN(pt)) {
    return 0;
  }
  return Math.round(pt * 12700);
}

/**
 * EMU转PT单位转换
 * 
 * @param emu - EMU单位值
 * @returns 磅值
 */
export function emu2pt(emu: number): number {
  if (typeof emu !== 'number' || isNaN(emu)) {
    return 0;
  }
  return emu / 12700;
}

/**
 * PX转PT单位转换
 * 在96DPI下：1 px = 0.75 pt
 * 
 * @param px - 像素值
 * @returns 磅值
 */
export function px2pt(px: number): number {
  if (typeof px !== 'number' || isNaN(px)) {
    return 0;
  }
  return px * 0.75;
}

/**
 * PT转PX单位转换
 * 在96DPI下：1 pt = 1.333 px
 * 
 * @param pt - 磅值
 * @returns 像素值
 */
export function pt2px(pt: number): number {
  if (typeof pt !== 'number' || isNaN(pt)) {
    return 0;
  }
  return pt * (4 / 3);
}

/**
 * 百分比转像素值
 * 用于相对尺寸计算
 * 
 * @param percent - 百分比值 (0-100)
 * @param total - 总值
 * @returns 计算后的像素值
 */
export function percentToPx(percent: number, total: number): number {
  if (typeof percent !== 'number' || isNaN(percent) || typeof total !== 'number' || isNaN(total)) {
    return 0;
  }
  return (percent / 100) * total;
}

/**
 * 计算两点之间的距离（EMU单位）
 * 
 * @param x1 - 第一个点的X坐标
 * @param y1 - 第一个点的Y坐标
 * @param x2 - 第二个点的X坐标
 * @param y2 - 第二个点的Y坐标
 * @returns 距离（EMU单位）
 */
export function distanceEmu(x1: number, y1: number, x2: number, y2: number): number {
  const dx = x2 - x1;
  const dy = y2 - y1;
  return Math.sqrt(dx * dx + dy * dy);
}

/**
 * 计算矩形的对角线长度（EMU单位）
 * 
 * @param width - 宽度（EMU单位）
 * @param height - 高度（EMU单位）
 * @returns 对角线长度（EMU单位）
 */
export function diagonalEmu(width: number, height: number): number {
  return distanceEmu(0, 0, width, height);
}

/**
 * 检查EMU值是否有效
 * PPTX中EMU值通常在合理范围内
 * 
 * @param emu - EMU单位值
 * @returns 是否有效
 */
export function isValidEmu(emu: number): boolean {
  return typeof emu === 'number' && 
         !isNaN(emu) && 
         emu >= 0 && 
         emu <= 5278760; // 典型PPTX最大值
}