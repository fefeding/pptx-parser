/**
 * 几何工具：包围盒与旋转。下沉自编辑器 util.js，与 DOM 无关，可在库层复用。
 */

export interface ElementBox {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  rotation?: number;
}

export interface BBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** 元素（含可选旋转）的轴对齐包围盒 */
export function bboxOf(el: ElementBox): BBox {
  const w = el.width || 0, h = el.height || 0;
  if (!el.rotation) return { x: el.x || 0, y: el.y || 0, width: w, height: h };
  const cx = (el.x || 0) + w / 2, cy = (el.y || 0) + h / 2;
  const rad = (el.rotation * Math.PI) / 180;
  const cos = Math.abs(Math.cos(rad)), sin = Math.abs(Math.sin(rad));
  const nw = w * cos + h * sin, nh = w * sin + h * cos;
  return { x: cx - nw / 2, y: cy - nh / 2, width: nw, height: nh };
}

/** 多元素并集包围盒；空数组返回 null */
export function unionBBox(els: ElementBox[]): BBox | null {
  if (!els.length) return null;
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const el of els) {
    const b = bboxOf(el);
    x0 = Math.min(x0, b.x); y0 = Math.min(y0, b.y);
    x1 = Math.max(x1, b.x + b.width); y1 = Math.max(y1, b.y + b.height);
  }
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
}

/** 把点绕 (cx,cy) 旋转 deg 度 */
export function rotatePoint(px: number, py: number, cx: number, cy: number, deg: number): { x: number; y: number } {
  const r = (deg * Math.PI) / 180, cos = Math.cos(r), sin = Math.sin(r);
  const dx = px - cx, dy = py - cy;
  return { x: cx + dx * cos - dy * sin, y: cy + dx * sin + dy * cos };
}
