/**
 * 颜色工具：与编辑器此前 util.js 的纯手写实现同源，下沉到库层以便复用。
 * 与 utils/style.ts 的 tinycolor2 互补——本模块是无依赖的轻量解析/转换，
 * 适合仅需「任意写法 → #RRGGBB」或简单混色/亮度计算的场景。
 */

/** 常见命名色（含 PowerPoint 调色板子集）；transparent 映射到 null（透明） */
const NAMED: Record<string, string | null> = {
  white: '#ffffff', black: '#000000', red: '#d93025', green: '#1e8e3e', blue: '#1a73e8',
  yellow: '#f9ab00', gray: '#80868b', grey: '#80868b', orange: '#f29900', purple: '#8430ce',
  pink: '#ff6d9e', cyan: '#12b5cb', transparent: null
};

/** 任意颜色写法 → '#RRGGBB'，无法识别返回 null */
export function normalizeColor(input: string | null | undefined): string | null {
  if (input == null) return null;
  let c = String(input).trim().toLowerCase();
  if (!c) return null;
  if (NAMED[c] !== undefined) return NAMED[c];
  if (c === 'none' || c === 'transparent') return null;
  if (c[0] === '#') c = c.slice(1);
  // 裸 hex（OOXML srgbClr 输出不带 #）与带 # 的写法都接受
  if (c.length === 3) c = c.split('').map((x) => x + x).join('');
  if (c.length === 8) c = c.slice(0, 6);
  if (/^[0-9a-f]{6}$/.test(c)) return '#' + c.toUpperCase();
  if (c.startsWith('rgb')) {
    const nums = c.replace(/[^0-9.,]/g, '').split(',').map(Number);
    if (nums.length >= 3) return rgbToHex(nums[0], nums[1], nums[2]);
  }
  return null;
}

export function rgbToHex(r: number, g: number, b: number): string {
  const f = (v: number) => clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0');
  return `#${f(r)}${f(g)}${f(b)}`.toUpperCase();
}

export function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const c = normalizeColor(hex) || '#000000';
  return { r: parseInt(c.slice(1, 3), 16), g: parseInt(c.slice(3, 5), 16), b: parseInt(c.slice(5, 7), 16) };
}

/** 亮度 0(黑)~1(白) */
export function luminance(hex: string): number {
  const { r, g, b } = hexToRgb(hex);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
}

/** 在其上叠加白/黑得到新的明度（用于生成渐变第二色） */
export function shade(hex: string, amount = 0.2): string {
  const { r, g, b } = hexToRgb(hex);
  const t = amount < 0 ? 0 : 255;
  const p = Math.abs(amount);
  return rgbToHex(r + (t - r) * p, g + (t - g) * p, b + (t - b) * p);
}

/** 带透明度的 css 颜色 */
export function withAlpha(hex: string, alphaPct?: number): string {
  const c = normalizeColor(hex);
  if (!c) return 'transparent';
  if (!alphaPct) return c;
  const { r, g, b } = hexToRgb(c);
  return `rgba(${r},${g},${b},${clamp(1 - alphaPct / 100, 0, 1)})`;
}

export function mixHex(a: string, b: string, t = 0.5): string {
  const A = hexToRgb(a), B = hexToRgb(b);
  return rgbToHex(A.r + (B.r - A.r) * t, A.g + (B.g - A.g) * t, A.b + (B.b - A.b) * t);
}

function clamp(v: number, lo: number, hi: number): number {
  return Math.min(Math.max(v, lo), hi);
}
