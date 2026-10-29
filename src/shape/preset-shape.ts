/**
 * 预设几何 → SVG 路径。
 *
 * 从编辑器 examples/editor/src/preset-paths.js 下沉而来，统一到库层，
 * 供编辑器渲染与任何需要「OOXML preset geometry → SVG path」的消费者复用。
 * 底层几何（shapeArc/shapeArcAlt/shapePie）复用 path-generators 的单一实现，
 * 仅以 moveTo 语义薄包装 shapeArcAlt 以匹配编辑器调用约定。
 */

import { shapeArc, shapeArcAlt as _shapeArcAlt, shapePie, shapeGear } from './path-generators';

const fmt = (n: number) => parseFloat(Number(n).toFixed(2));

/**
 * 复用 path-generators 的几何实现（单一权威）。库版 shapeArcAlt 第 7 参为 isClose
 * （控制末尾是否 Z 闭合），此处以 moveTo 语义薄包装：moveTo=true 首点 M（独立子路径），
 * moveTo=false 首点 L（续接当前子路径）——与编辑器原调用约定一致，等价于库版后 .replace("M","L")。
 */
function shapeArcAlt(cX: number, cY: number, rX: number, rY: number, stAng: number, endAng: number, moveTo = true) {
  const d = _shapeArcAlt(cX, cY, rX, rY, stAng, endAng, false);
  return moveTo ? d : d.replace('M', 'L');
}

type Adj = Record<string, number | string> | undefined;
const num = (adj: Adj, key: string, def: number) => {
  const v = adj && adj[key] != null ? Number((adj as Record<string, number>)[key]) : def;
  return Number.isFinite(v) ? v : def;
};
/**
 * 依次尝试多个调整值名称取第一个有效值。
 * 预览端部分预设（如 sun/cloud）只读「第一个 a:gd」而非按名字取，
 * 而编辑模型按名字归一化（a:gd@name=adj / adj1），
 * 故这里同时兼容两种来源，避免同一形状在两端取到不同的调整值。
 */
const numAny = (adj: Adj, keys: string[], def: number) => {
  for (const k of keys) {
    const v = adj && (adj as Record<string, number>)[k];
    if (v != null) {
      const n = Number(v);
      if (Number.isFinite(n)) return n;
    }
  }
  return def;
};
const clampV = (v: number, lo: number, hi: number) => Math.min(Math.max(v, lo), hi);

/** 矩形闭合路径 */
const rectD = (w: number, h: number) => `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z`;

export interface PresetShapeResult {
  d: string;
  transform?: string;
  fillRule?: 'evenodd' | 'nonzero';
  strokes?: { d: string; width?: number }[];
  strokeOnly?: boolean;
  noFill?: boolean;
}

/**
 * 主入口：返回 { d, transform, fillRule, strokes:[{d,width}], strokeOnly }
 * 不支持的 preset 返回 null（由调用方回退 CSS 近似）。
 */
export function presetShapePath(prst: string, w: number, h: number, adj: Adj = {}, opts: Record<string, unknown> = {}): PresetShapeResult | null {
  switch (prst) {
    // 矩形：与预览端 <rect> 等价，供 presetShapeSvg 渲染虚线边框（CSS border-style 无法表达固定 dash 图案）
    case 'rect': return { d: rectD(w, h) };
    // 星形成（star4/5/6/8/10/12/16/24/32）：外/内顶点交替，内半径 = 外半径 * adj/50000
    case 'star4': case 'star5': case 'star6': case 'star8': case 'star10':
    case 'star12': case 'star16': case 'star24': case 'star32': {
      const N = { star4: 4, star5: 5, star6: 6, star8: 8, star10: 10, star12: 12, star16: 16, star24: 24, star32: 32 }[prst];
      const defAdj = prst === 'star5' ? 19098 : 12500;
      const ratio = num(adj, 'adj', defAdj) / 50000;
      const rx = w / 2, ry = h / 2;
      const pts: string[] = [];
      for (let i = 0; i < N * 2; i++) {
        const ang = -90 + i * (180 / N);
        const r = (i % 2 === 0) ? 1 : ratio;
        const a = ang * Math.PI / 180;
        pts.push(`${(rx + rx * r * Math.cos(a)).toFixed(2)},${(ry + ry * r * Math.sin(a)).toFixed(2)}`);
      }
      return { d: 'M' + pts.join(' L') + ' z' };
    }
    case 'pie':
    case 'pieWedge':
    case 'arc': {
      const isClose = prst !== 'arc';
      let adj1 = prst === 'pieWedge' ? 180 : prst === 'arc' ? 270 : 0;
      let adj2 = prst === 'pie' ? 270 : prst === 'pieWedge' ? 270 : 0;
      let H = prst === 'pieWedge' ? 2 * h : h;
      if (adj && (adj as Record<string, number>).adj1 != null) adj1 = num(adj, 'adj1', adj1) / 60000;
      if (adj && (adj as Record<string, number>).adj2 != null) adj2 = num(adj, 'adj2', adj2) / 60000;
      const [d, rot] = shapePie(H, w, adj1, adj2, isClose);
      // 预览端 arc 恒为空心描边（fill=none）
      return { d, transform: rot, noFill: !isClose || undefined };
    }
    case 'chord': {
      const a1 = num(adj, 'adj1', 45 * 60000) / 60000;
      const a2 = num(adj, 'adj2', 270 * 60000) / 60000;
      return { d: shapeArc(w / 2, h / 2, w / 2, h / 2, a1, a2, true) };
    }
    case 'noSmoking': {
      const a = clampV(num(adj, 'adj', 18750), 0, 50000);
      const dr = Math.min(w, h) * a / 100000;
      const iwd2 = w / 2 - dr, ihd2 = h / 2 - dr;
      const ang = Math.atan(h / w);
      const ct = ihd2 * Math.cos(ang), st = iwd2 * Math.sin(ang);
      const m = Math.sqrt(ct * ct + st * st);
      const n = iwd2 * ihd2 / m;
      const dang = Math.atan((dr / 2) / n);
      const swAng = -Math.PI + dang * 2;
      const stAng1 = ang - dang;
      const stAng2 = stAng1 - Math.PI;
      const dx1 = n * Math.cos(stAng1), dy1 = n * Math.sin(stAng1);
      const x1 = w / 2 + dx1, y1 = h / 2 + dy1;
      const x2 = w / 2 - dx1, y2 = h / 2 - dy1;
      const a1deg = stAng1 * 180 / Math.PI;
      const a2deg = stAng2 * 180 / Math.PI;
      const swDeg = swAng * 180 / Math.PI;
      const d = `M${0},${h / 2}${shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 180, 270, false)}` +
        `${shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 270, 360, false)}` +
        `${shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 0, 90, false)}` +
        `${shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 90, 180, false)} z` +
        `M${x1},${y1}${shapeArcAlt(w / 2, h / 2, iwd2, ihd2, a1deg, a1deg + swDeg, false)} z` +
        `M${x2},${y2}${shapeArcAlt(w / 2, h / 2, iwd2, ihd2, a2deg, a2deg + swDeg, false)} z`;
      // 两条内部弧线构成斜杠“镂空”，需用 evenOdd 填充规则让背景透出
      return { d, fillRule: 'evenodd' };
    }
    case 'smileyFace': {
      // 与预览端 misc-shapes 逐式移植（眼睛为退化弧线，实际只显示脸 + 嘴）
      const a = clampV(num(adj, 'adj', 4653), -4653, 4653);
      const wd2 = w / 2, hd2 = h / 2;
      const x1 = w * 4969 / 21699, x2 = w * 6215 / 21600, x3 = w * 13135 / 21600, x4 = w * 16640 / 21600;
      const y1 = h * 7570 / 21600, y3 = h * 16515 / 21600;
      const y2 = y3 - h * a / 100000;
      const y4 = y3 + h * a / 100000;
      const y5 = y4 + h * a / 50000;
      const wR = w * 1125 / 21600, hR = h * 1125 / 21600;
      const cX1 = x2 - wR * Math.cos(Math.PI);
      const cY1 = y1 - hR * Math.sin(Math.PI);
      const cX2 = x3 - wR * Math.cos(Math.PI);
      const d = `${shapeArc(cX1, cY1, wR, hR, 180, 540)}` +
        `${shapeArc(cX2, cY1, wR, hR, 180, 540)}` +
        ` M${x1},${y2} Q${wd2},${y5} ${x4},${y2} Q${wd2},${y5} ${x1},${y2}` +
        ` M${0},${hd2}${shapeArc(wd2, hd2, wd2, hd2, 180, 540).replace('M', 'L')} z`;
      return { d };
    }
    case 'plus': {
      const a1 = num(adj, 'adj', 25000) / 100000;
      const a2 = 1 - a1;
      const p = [
        [a1 * w, 0], [a1 * w, a1 * h], [0, a1 * h], [0, a2 * h],
        [a1 * w, a2 * h], [a1 * w, h], [a2 * w, h], [a2 * w, a2 * h],
        [w, a2 * h], [w, a1 * h], [a2 * w, a1 * h], [a2 * w, 0]
      ];
      return { d: 'M' + p.map((pt) => `${fmt(pt[0])},${fmt(pt[1])}`).join(' L') + ' z' };
    }
    case 'plaque': {
      let adjVal = clampV(num(adj, 'adj', 25000), 0, 50000);
      const r = (adjVal / 100000) * Math.min(w, h);
      return { d: `M${r},0A${r} ${r} 0 0 1 0,${r}L0,${(h - r)}A${r} ${r} 0 0 1 ${r},${h}L${(w - r)},${h}A${r} ${r} 0 0 1 ${w},${(h - r)}L${w},${r}A${r} ${r} 0 0 1 ${(w - r)},0 z` };
    }
    case 'quadArrow': {
      const cn1 = 50000, cn2 = 100000, cn3 = 200000;
      const minWH = Math.min(w, h);
      const vc = h / 2, hc = w / 2;
      const a2 = clampV(num(adj, 'adj2', 22500), 0, cn1);
      const a1 = clampV(num(adj, 'adj1', 22500), 0, 2 * a2);
      const a3 = clampV(num(adj, 'adj3', 22500), 0, (cn2 - 2 * a2) / 2);
      const x1 = minWH * a3 / cn2, dx2 = minWH * a2 / cn2;
      const x2 = hc - dx2, x5 = hc + dx2;
      const dx3 = minWH * a1 / cn3;
      const x3 = hc - dx3, x4 = hc + dx3, x6 = w - x1;
      const y2 = vc - dx2, y5 = vc + dx2, y3 = vc - dx3, y4 = vc + dx3, y6 = h - x1;
      const d = `M${0},${vc} L${x1},${y2} L${x1},${y3} L${x3},${y3} L${x3},${x1} L${x2},${x1} L${hc},${0} L${x5},${x1} L${x4},${x1} L${x4},${y3} L${x6},${y3} L${x6},${y2} L${w},${vc} L${x6},${y5} L${x6},${y4} L${x4},${y4} L${x4},${y6} L${x5},${y6} L${hc},${h} L${x2},${y6} L${x3},${y6} L${x3},${y4} L${x1},${y4} L${x1},${y5} z`;
      return { d };
    }
    case 'uturnArrow': {
      const cn1 = 25000, cn2 = 100000;
      const minWH = Math.min(w, h);
      const a2 = clampV(num(adj, 'adj2', 25000), 0, cn1);
      const a1 = clampV(num(adj, 'adj1', 25000), 0, 2 * a2);
      const q2 = a1 * minWH / h, q3 = cn2 - q2;
      const a3 = clampV(num(adj, 'adj3', 25000), 0, q3 * h / minWH);
      const q1 = a3 + a1;
      const a5 = clampV(num(adj, 'adj5', 75000), q1 * minWH / h, cn2);
      const th = minWH * a1 / cn2;
      const aw2 = minWH * a2 / cn2;
      const th2 = th / 2, dh2 = aw2 - th2;
      const y5 = h * a5 / cn2;
      const ah = minWH * a3 / cn2;
      const y4 = y5 - ah;
      const x9 = w - dh2;
      const bw = x9 / 2;
      const bs = Math.min(bw, y4);
      const a4 = clampV(num(adj, 'adj4', 43750), 0, cn2 * bs / minWH);
      const bd = minWH * a4 / cn2;
      const bd2 = Math.max(bd - th, 0);
      const x3 = th + bd2, x8 = w - aw2, x6 = x8 - aw2, x7 = x6 + dh2;
      const x4 = x9 - bd, x5 = x7 - bd2;
      const d = `M${0},${h} L${0},${bd}${shapeArcAlt(bd, bd, bd, bd, 180, 270, false)}` +
        ` L${x4},${0}${shapeArcAlt(x4, bd, bd, bd, 270, 360, false)}` +
        ` L${x9},${y4} L${w},${y4} L${x8},${y5} L${x6},${y4} L${x7},${y4} L${x7},${x3}` +
        `${shapeArcAlt(x5, x3, bd2, bd2, 0, -90, false)} L${x3},${th}${shapeArcAlt(x3, x3, bd2, bd2, 270, 180, false)} L${th},${h} z`;
      return { d };
    }
    case 'wedgeRectCallout': {
      const cn1 = 100000;
      const vc = h / 2, hc = w / 2;
      const adj1 = num(adj, 'adj1', -20833), adj2 = num(adj, 'adj2', 62500);
      const dxPos = w * adj1 / cn1, dyPos = h * adj2 / cn1;
      const xPos = hc + dxPos, yPos = vc + dyPos;
      const dq = dxPos * h / w;
      const dz = Math.abs(dyPos) - Math.abs(dq);
      const xg1 = dxPos > 0 ? 7 : 2, xg2 = dxPos > 0 ? 10 : 5;
      const x1 = w * xg1 / 12, x2 = w * xg2 / 12;
      const yg1 = dyPos > 0 ? 7 : 2, yg2 = dyPos > 0 ? 10 : 5;
      const y1 = h * yg1 / 12, y2 = h * yg2 / 12;
      const xl = dz > 0 ? 0 : (dxPos > 0 ? 0 : xPos);
      const xt = dz > 0 ? (dyPos > 0 ? x1 : xPos) : x1;
      const xr = dz > 0 ? w : (dxPos > 0 ? xPos : w);
      const xb = dz > 0 ? (dyPos > 0 ? xPos : x1) : x1;
      const yl = dz > 0 ? y1 : (dxPos > 0 ? y1 : yPos);
      const yt = dz > 0 ? (dyPos > 0 ? 0 : yPos) : 0;
      const yr = dz > 0 ? y1 : (dxPos > 0 ? yPos : y1);
      const yb = dz > 0 ? (dyPos > 0 ? yPos : h) : h;
      const d = `M${0},${0} L${x1},${0} L${xt},${yt} L${x2},${0} L${w},${0} L${w},${y1} L${xr},${yr} L${w},${y2} L${w},${h} L${x2},${h} L${xb},${yb} L${x1},${h} L${0},${h} L${0},${y2} L${xl},${yl} L${0},${y1} z`;
      return { d };
    }
    case 'accentBorderCallout1':
    case 'accentCallout1':
    case 'borderCallout1':
    case 'callout1': {
      const cn1 = 100000;
      const y1 = h * num(adj, 'adj1', 18750) / cn1, x1 = w * num(adj, 'adj2', -8333) / cn1;
      const y2 = h * num(adj, 'adj3', 112500) / cn1, x2 = w * num(adj, 'adj4', -38333) / cn1;
      const isAccent = prst.startsWith('accent');
      const strokes = [{ d: `M${x1},${y1} L${x2},${y2}` }];
      if (isAccent) strokes.push({ d: `M${x1},0 L${x1},${h}` });
      return { d: rectD(w, h), strokes };
    }
    case 'accentBorderCallout2':
    case 'accentCallout2':
    case 'borderCallout2':
    case 'callout2': {
      const cn1 = 100000;
      const y1 = h * num(adj, 'adj1', 18750) / cn1, x1 = w * num(adj, 'adj2', -8333) / cn1;
      const y2 = h * num(adj, 'adj3', 18750) / cn1, x2 = w * num(adj, 'adj4', -16667) / cn1;
      const y3 = h * num(adj, 'adj5', 112500) / cn1, x3 = w * num(adj, 'adj6', -46667) / cn1;
      const isAccent = prst.startsWith('accent');
      const strokes = [{ d: `M${x1},${y1} L${x2},${y2} L${x3},${y3} L${x2},${y2}` }];
      if (isAccent) strokes.push({ d: `M${x1},0 L${x1},${h}` });
      return { d: rectD(w, h), strokes };
    }
    case 'accentBorderCallout3':
    case 'accentCallout3':
    case 'borderCallout3':
    case 'callout3': {
      const cn1 = 100000;
      const y1 = h * num(adj, 'adj1', 18750) / cn1, x1 = w * num(adj, 'adj2', -8333) / cn1;
      const y2 = h * num(adj, 'adj3', 18750) / cn1, x2 = w * num(adj, 'adj4', -16667) / cn1;
      const y3 = h * num(adj, 'adj5', 100000) / cn1, x3 = w * num(adj, 'adj6', -16667) / cn1;
      const y4 = h * num(adj, 'adj7', 112963) / cn1, x4 = w * num(adj, 'adj8', -8333) / cn1;
      const isAccent = prst.startsWith('accent');
      const strokes = [{ d: `M${x1},${y1} L${x2},${y2} L${x3},${y3} L${x4},${y4} L${x3},${y3} L${x2},${y2}` }];
      if (isAccent) strokes.push({ d: `M${x1},0 L${x1},${h}` });
      return { d: rectD(w, h), strokes };
    }
    case 'ellipseRibbon':
    case 'ellipseRibbon2': {
      const cn1 = 25000, cn3 = 75000, cn4 = 100000, cn5 = 200000;
      const hc = w / 2, t = 0, l = 0, b = h, r = w, wd8 = w / 8;
      const a1 = clampV(num(adj, 'adj1', 25000), 0, cn4);
      const a2 = clampV(num(adj, 'adj2', 50000), cn1, cn3);
      const q10 = cn4 - a1, q11 = q10 / 2, q12 = a1 - q11;
      const a3 = clampV(num(adj, 'adj3', 12500), Math.max(0, q12), a1);
      const dx2 = w * a2 / cn5;
      const x2 = hc - dx2, x3 = x2 + wd8, x4 = r - x3, x5 = r - x2, x6 = r - wd8;
      const dy1 = h * a3 / cn4;
      const f1 = 4 * dy1 / w;
      let q1 = x3 * x3 / w;
      const q2 = x3 - q1;
      const cx1 = x3 / 2, cx2 = r - cx1;
      q1 = h * a1 / cn4;
      const dy3 = q1 - dy1;
      const q3 = x2 * x2 / w;
      const q4 = x2 - q3;
      const q5 = f1 * q4;
      const rh = b - q1;
      const q8 = dy1 * 14 / 16;
      const cx4 = x2 / 2;
      const q9 = f1 * cx4;
      const cx5 = r - cx4;
      if (prst === 'ellipseRibbon') {
        const y1v = f1 * q2, cy1 = f1 * cx1;
        const y3v = q5 + dy3;
        const q6 = dy1 + dy3 - y3v;
        const q7 = q6 + dy1;
        const cy3 = q7 + dy3;
        const y2v = (q8 + rh) / 2;
        const y5v = q5 + rh, y6 = y3v + rh;
        const cy4 = q9 + rh, cy6 = cy3 + rh;
        const y7 = y1v + dy3, cy7 = q1 + q1 - y7;
        const y8 = b - dy1;
        const d = `M${l},${t} Q${cx1},${cy1} ${x3},${y1v} L${x2},${y3v} Q${hc},${cy3} ${x5},${y3v} L${x4},${y1v} Q${cx2},${cy1} ${r},${t} L${x6},${y2v} L${r},${rh} Q${cx5},${cy4} ${x5},${y5v} L${x5},${y6} Q${hc},${cy6} ${x2},${y6} L${x2},${y5v} Q${cx4},${cy4} ${l},${rh} L${wd8},${y2v} z` +
          `M${x2},${y5v} L${x2},${y3v}M${x5},${y3v} L${x5},${y5v}M${x3},${y1v} L${x3},${y7}M${x4},${y7} L${x4},${y1v}`;
        return { d };
      }
      // ellipseRibbon2
      const u1 = f1 * q2, cu1 = f1 * cx1;
      const u3 = q5 + dy3;
      const q6 = dy1 + dy3 - u3;
      const q7 = q6 + dy1;
      const cu3 = q7 + dy3;
      const u2 = (q8 + rh) / 2;
      const u5 = q5 + rh, u6 = u3 + rh;
      const cu4 = q9 + rh, cu6 = cu3 + rh;
      const u7 = u1 + dy3, cu7 = q1 + q1 - u7;
      const d2 = `M${l},${b} Q${cx1},${h - cu1} ${x3},${h - u1} L${x2},${h - u3} Q${hc},${h - cu3} ${x5},${h - u3} L${x4},${h - u1} Q${cx2},${h - cu1} ${r},${b} L${x6},${h - u2} L${r},${h - rh} Q${cx5},${h - cu4} ${x5},${h - u5} L${x5},${h - u6} Q${hc},${h - cu6} ${x2},${h - u6} L${x2},${h - u5} Q${cx4},${h - cu4} ${l},${h - rh} L${wd8},${h - u2} z` +
        `M${x2},${h - u5} L${x2},${h - u3}M${x5},${h - u3} L${x5},${h - u5}M${x3},${h - u1} L${x3},${h - u7}M${x4},${h - u7} L${x4},${h - u1}`;
      return { d: d2 };
    }
    case 'can':
    case 'cylinder':
    case 'flowChartMagneticDisk':
    case 'flowChartMagneticDrum': {
      const ss = Math.min(w, h);
      const maxAdj = 50000 * h / ss;
      let adjVal = (prst === 'can' || prst === 'cylinder') ? num(adj, 'adj', 25000) : 50000;
      if (adjVal < 0) adjVal = 0;
      if (adjVal > maxAdj) adjVal = maxAdj;
      const y1 = ss * adjVal / 200000;
      const y3 = h - y1;
      const wd2 = w / 2;
      const rot = prst === 'flowChartMagneticDrum' ? `rotate(90 ${w / 2},${h / 2})` : '';
      const d = `${shapeArcAlt(wd2, y1, wd2, y1, 0, 180)}` +
        `${shapeArcAlt(wd2, y1, wd2, y1, 180, 360, false)}` +
        ` L${w},${y3}${shapeArcAlt(wd2, y3, wd2, y1, 0, 180, false)} L${0},${y1}`;
      return { d, transform: rot || undefined };
    }
    case 'flowChartMultidocument': {
      const y1 = h * 18022 / 21600, y2 = h * 3675 / 21600, y3 = h * 23542 / 21600;
      const y4 = h * 1815 / 21600, y5 = h * 16252 / 21600, y6 = h * 16352 / 21600;
      const y7 = h * 14392 / 21600, y8 = h * 20782 / 21600, y9 = h * 14467 / 21600;
      const x1 = w * 1532 / 21600, x2 = w * 20000 / 21600, x3 = w * 9298 / 21600;
      const x4 = w * 19298 / 21600, x5 = w * 18595 / 21600, x6 = w * 2972 / 21600;
      const x7 = w * 20800 / 21600;
      const d = `M${0},${y2} L${x5},${y2} L${x5},${y1} C${x3},${y1} ${x3},${y3} ${0},${y8} z` +
        `M${x1},${y2} L${x1},${y4} L${x2},${y4} L${x2},${y5} C${x4},${y5} ${x5},${y6} ${x5},${y6}` +
        `M${x6},${y4} L${x6},${0} L${w},${0} L${w},${y7} C${x7},${y7} ${x2},${y9} ${x2},${y9}`;
      return { d };
    }
    case 'teardrop': {
      const cn1 = 100000, cn2 = 200000;
      const a1 = clampV(num(adj, 'adj', 100000), 0, cn2);
      const r2 = Math.sqrt(2);
      const tw = r2 * (w / 2), th2 = r2 * (h / 2);
      const sw = (tw * a1) / cn1, sh = (th2 * a1) / cn1;
      const rd45 = 45 * Math.PI / 180;
      const dx1 = sw * Math.cos(rd45), dy1 = sh * Math.cos(rd45);
      const x1 = (w / 2) + dx1, y1 = (h / 2) - dy1;
      const x2 = ((w / 2) + x1) / 2, y2 = ((h / 2) + y1) / 2;
      const d = `${shapeArc(w / 2, h / 2, w, h, 180, 270)}Q ${x2},0 ${x1},${y1}Q ${w},${y2} ${w},${h / 2}` +
        `${shapeArc(w / 2, h / 2, w, h, 0, 90)}${shapeArc(w / 2, h / 2, w, h, 90, 180)} z`;
      return { d };
    }
    // ===== 复杂轮廓预设：与预览端 src/shape/shape.ts 同源移植 =====
    // 此前这些形状在编辑器里只有 CSS 近似（border-radius / clip-path 多边形），
    // 月亮渲染成椭圆、太阳渲染成椭圆、云朵渲染成圆角矩形，与预览端和 PowerPoint 都不符。
    case 'moon': {
      // 月牙：外弧 = 形状椭圆左半（自上尖端绕左到下尖端），
      // 内弧 = 一个更大的椭圆在左侧切出的弧（尖端为尖角，与 PowerPoint 一致）。
      // adj（0~1，默认 0.5）控制内弧深度：0 → 半椭圆（D 形），越大月牙越细。
      // 注：预览端旧公式的弧半径与端点半径不一致（差 2 倍），会画出越界尖刺，此处按正确几何重写。
      const a = clampV(numAny(adj, ['adj', 'adj1'], 50000) / 100000, 0, 1);
      const hw = w / 2, hh = h / 2;
      // 内弧椭圆：缩放系数 k，使内弧恰好通过两个尖端（推导见 k = [-1-(a-1)^2] / [2(a-1)]）
      const k = clampV((-1 - (a - 1) * (a - 1)) / (2 * (a - 1)), 1, 1000);
      const rxI = k * hw, ryI = k * hh;
      // a→0 时 k→1，内弧退化为竖直线（无缺口）
      const inner = a <= 0.0001
        ? `L${hw},0`
        : `A${fmt(rxI)},${fmt(ryI)} 0 0 1 ${hw},0`;
      const d = `M${hw},0 A${fmt(hw)},${fmt(hh)} 0 0 0 ${hw},${h} ${inner} z`;
      return { d };
    }
    case 'sun': {
      const cn3 = 50000, cn4 = 100000;
      const a1 = clampV(numAny(adj, ['adj', 'adj1'], 25000), 12500, 46875);
      const g0 = cn3 - a1;
      const g1 = g0 * 30274 / 32768;
      const g2 = g0 * 12540 / 32768;
      const g5 = cn3 - g1;
      const g6 = cn3 - g2;
      const g7 = g0 * 23170 / 32768;
      const g8 = cn3 + g7;
      const g9 = cn3 - g7;
      const g10 = g5 * 3 / 4;
      const g11 = g6 * 3 / 4;
      const g12 = g10 + 3662;
      const g13 = g11 + 36620;
      const g14 = g11 + 12500;
      const g15 = cn4 - g10;
      const g16 = cn4 - g12;
      const g17 = cn4 - g13;
      const g18 = cn4 - g14;
      const ox1 = w * 18436 / 21600, oy1 = h * 3163 / 21600;
      const ox2 = w * 3163 / 21600, oy2 = h * 18436 / 21600;
      const X = (g: number) => (w * g / cn4);
      const Y = (g: number) => (h * g / cn4);
      const x10 = X(g10), x12 = X(g12), x13 = X(g13), x14 = X(g14), x15 = X(g15);
      const x16 = X(g16), x17 = X(g17), x18 = X(g18);
      const y10 = Y(g10), y12 = Y(g12), y13 = Y(g13), y14 = Y(g14), y15 = Y(g15);
      const y16 = Y(g16), y17 = Y(g17), y18 = Y(g18);
      const wR = w * g0 / cn4, hR = h * g0 / cn4;
      const d = `M${w},${h / 2} L${x15},${y18} L${x15},${y14}z` +
        ` M${ox1},${oy1} L${x16},${y17} L${x13},${y12}z` +
        ` M${w / 2},${0} L${x18},${y10} L${x14},${y10}z` +
        ` M${ox2},${oy1} L${x17},${y12} L${x12},${y17}z` +
        ` M${0},${h / 2} L${x10},${y14} L${x10},${y18}z` +
        ` M${ox2},${oy2} L${x12},${y13} L${x17},${y16}z` +
        ` M${w / 2},${h} L${x14},${y15} L${x18},${y15}z` +
        ` M${ox1},${oy2} L${x13},${y16} L${x16},${y13} z` +
        // 中心圆：旧公式用 shapeArc 生成 360° 弧，起止点重合导致整段退化（中心不填充），
        // 此处改为两段半椭圆弧（半径 = (50000-adj)/100000 × 尺寸）。
        ` M${w / 2 - wR},${h / 2} A${fmt(wR)},${fmt(hR)} 0 0 1 ${w / 2 + wR},${h / 2}` +
        ` A${fmt(wR)},${fmt(hR)} 0 0 1 ${w / 2 - wR},${h / 2} z`;
      return { d };
    }
    case 'cloud': {
      // 云形基于 43200×43200 的正方形设计：X 用 w 缩放，Y 用 h 缩放
      const RX = (v: number) => (w * v / 43200);
      const RY = (v: number) => (h * v / 43200);
      const x0 = fmt(RX(3900)), y0 = fmt(RY(14370));
      const rX1 = RX(6753), rY1 = RY(9190);
      const rX2 = RX(5333), rY2 = RY(7267);
      const rX3 = RX(4365), rY3 = RY(5945);
      const rX4 = RX(4857), rY4 = RY(6595);
      const rY5 = RY(7273);
      const rX6 = RX(6775), rY6 = RY(9220);
      const rX7 = RX(5785), rY7 = RY(7867);
      const rX8 = RX(6752), rY8 = RY(9215);
      const rX9 = RX(7720), rY9 = RY(10543);
      const rX10 = RX(4360), rY10 = RY(5918);
      const rX11 = RX(4345);
      // 起始角 / 扫描角（单位：度）
      const sA1 = -11429249 / 60000, wA1 = 7426832 / 60000;
      const sA2 = -8646143 / 60000, wA2 = 5396714 / 60000;
      const sA3 = -8748475 / 60000, wA3 = 5983381 / 60000;
      const sA4 = -7859164 / 60000, wA4 = 7034504 / 60000;
      const sA5 = -4722533 / 60000, wA5 = 6541615 / 60000;
      const sA6 = -2776035 / 60000, wA6 = 7816140 / 60000;
      const sA7 = 37501 / 60000, wA7 = 6842000 / 60000;
      const sA8 = 1347096 / 60000, wA8 = 6910353 / 60000;
      const sA9 = 3974558 / 60000, wA9 = 4542661 / 60000;
      const sA10 = -16496525 / 60000, wA10 = 8804134 / 60000;
      const sA11 = -14809710 / 60000, wA11 = 9151131 / 60000;
      // 椭圆弧（SVG A 命令）：返回 { end, path }
      const arc = (cx: number, cy: number, rx: number, ry: number, startDeg: number, sweepDeg: number) => {
        const endDeg = startDeg + sweepDeg;
        const endX = cx + rx * Math.cos(endDeg * Math.PI / 180);
        const endY = cy + ry * Math.sin(endDeg * Math.PI / 180);
        const largeArc = Math.abs(sweepDeg) > 180 ? 1 : 0;
        const sweep = sweepDeg > 0 ? 1 : 0;
        return {
          end: { x: endX, y: endY },
          path: `A ${fmt(rx)} ${fmt(ry)} 0 ${largeArc} ${sweep} ${fmt(endX)} ${fmt(endY)}`
        };
      };
      // 弧心 = 起点 - 半径 * cos/sin(起始角)，逐段推导下一段弧心
      let px = x0, py = y0;
      const seg = (rx: number, ry: number, sDeg: number, wDeg: number) => {
        const cx = px - rx * Math.cos(sDeg * Math.PI / 180);
        const cy = py - ry * Math.sin(sDeg * Math.PI / 180);
        const r = arc(cx, cy, rx, ry, sDeg, wDeg);
        px = r.end.x; py = r.end.y;
        return r.path;
      };
      const parts = [
        seg(rX1, rY1, sA1, wA1),
        seg(rX2, rY2, sA2, wA2),
        seg(rX3, rY3, sA3, wA3),
        seg(rX4, rY4, sA4, wA4),
        seg(rX2, rY5, sA5, wA5),
        seg(rX6, rY6, sA6, wA6),
        seg(rX7, rY7, sA7, wA7),
        seg(rX8, rY8, sA8, wA8),
        seg(rX9, rY9, sA9, wA9),
        seg(rX10, rY10, sA10, wA10),
        seg(rX11, rY3, sA11, wA11)
      ];
      return { d: `M${x0},${y0} ${parts.join(' ')} z` };
    }
    case 'donut': {
      // 圆环：外椭圆 + 内孔（evenodd 填充挖孔）。
      // 旧公式的弧半径与端点半径不一致，内孔被画成四角星形，此处按正确几何重写。
      const a = clampV(numAny(adj, ['adj', 'adj1'], 25000), 0, 50000);
      const dr = Math.min(w, h) * a / 100000;
      const wd2 = w / 2, hd2 = h / 2;
      const rx = wd2 - dr, ry = hd2 - dr;
      const d = `M0,${hd2}` +
        ` A${fmt(wd2)},${fmt(hd2)} 0 0 1 ${w},${hd2}` +
        ` A${fmt(wd2)},${fmt(hd2)} 0 0 1 0,${hd2} z` +
        ` M${dr},${hd2}` +
        ` A${fmt(rx)},${fmt(ry)} 0 0 1 ${w - dr},${hd2}` +
        ` A${fmt(rx)},${fmt(ry)} 0 0 1 ${dr},${hd2} z`;
      return { d, fillRule: 'evenodd' };
    }
    case 'frame': {
      const a1 = clampV(numAny(adj, ['adj', 'adj1'], 12500), 0, 50000);
      const x1 = Math.min(w, h) * a1 / 100000;
      const x4 = w - x1, y4 = h - x1;
      // 外框 + 内框（反向绕行 → nonzero 填充下中间镂空）
      return { d: `M0,0 L${w},0 L${w},${h} L0,${h} zM${x1},${x1} L${x1},${y4} L${x4},${y4} L${x4},${x1} z` };
    }
    case 'halfFrame': {
      // 半框：与预览端 shape.ts 的 halfFrame 同式（右下角切出一个直角缺口）
      // 预览端旧默认值为 3.5（实际等价 0）会让路径退化成零面积，这里给回可用的缺省值
      const a1 = clampV(numAny(adj, ['adj1', 'adj'], 50000), 0, 100000);
      const a2 = clampV(numAny(adj, ['adj2', 'adj'], 50000), 0, 100000);
      const minWH = Math.min(w, h);
      const cn = 100000;
      const maxAdj2 = cn * w / minWH;
      const v2 = Math.min(a2, maxAdj2);
      const x1 = minWH * v2 / cn;
      const g1 = h * x1 / w;
      const g2 = h - g1;
      const maxAdj1 = cn * g2 / minWH;
      const v1 = Math.min(a1, maxAdj1);
      const y1 = minWH * v1 / cn;
      const dx2 = y1 * w / h;
      const x2 = w - dx2;
      const dy2 = x1 * h / w;
      const y2 = h - dy2;
      return { d: `M0,0 L${w},0 L${x2},${y1} L${x1},${y1} L${x1},${y2} L0,${h} z` };
    }
    case 'gear6':
    case 'gear9': {
      const pts = prst === 'gear6' ? 6 : 9;
      const d = shapeGear(w, h / 3.5, pts);
      return { d, transform: `rotate(20,${(3 / 7) * h},${(3 / 7) * h})` };
    }
    case 'heart': {
      const dx1 = w * 49 / 48;
      const dx2 = w * 10 / 48;
      const x1 = w / 2 - dx1, x2 = w / 2 - dx2;
      const x3 = w / 2 + dx2, x4 = w / 2 + dx1;
      const y1 = -h / 3;
      const d = `M${w / 2},${h / 4}C${x3},${y1} ${x4},${h / 4} ${w / 2},${h}C${x1},${h / 4} ${x2},${y1} ${w / 2},${h / 4} z`;
      return { d };
    }
    case 'lightningBolt': {
      const px = (v: number) => (w * v / 21600);
      const py = (v: number) => (h * v / 21600);
      const x1 = px(5022), x2 = px(11050), x3 = px(8472), x5 = px(10012);
      const x6 = px(14767), x7 = px(12222), x8 = px(12860), x10 = px(7602), x11 = px(16577);
      const y1 = py(3890), y2 = py(6080), y3 = py(6797), y5 = py(12877);
      const y6 = py(9705), y7 = py(12007), y8 = py(13987), y9 = py(8382), y11 = py(14915);
      const d = `M${x3},0 L${x8},${y2} L${x2},${y3} L${x11},${y7} L${x6},${y5} L${w},${h}` +
        ` L${x5},${y11} L${x7},${y8} L${x1},${y6} L${x10},${y9} L0,${y1} z`;
      return { d };
    }
    case 'cube': {
      const a = clampV(numAny(adj, ['adj', 'adj1'], 25000), 0, 100000);
      const ss = Math.min(w, h);
      const y1 = ss * a / 100000;
      const y4 = h - y1, x4 = w - y1;
      const d = `M0,${y1} L${y1},0 L${w},0 L${w},${y4} L${x4},${h} L0,${h} z` +
        `M0,${y1} L${x4},${y1} M${x4},${y1} L${w},0 M${x4},${y1} L${x4},${h}`;
      return { d };
    }
    case 'bevel': {
      const a = clampV(numAny(adj, ['adj', 'adj1'], 12500), 0, 50000);
      const ss = Math.min(w, h);
      const x1 = ss * a / 100000;
      const x2 = w - x1, y2 = h - x1;
      const d = `M0,0 L${w},0 L${w},${h} L0,${h} z` +
        ` M${x1},${x1} L${x2},${x1} L${x2},${y2} L${x1},${y2} z` +
        ` M0,0 L${x1},${x1} M0,${h} L${x1},${y2} M${w},0 L${x2},${x1} M${w},${h} L${x2},${y2}`;
      return { d };
    }
    case 'foldedCorner': {
      const a = clampV(numAny(adj, ['adj', 'adj1'], 16667), 0, 50000);
      const ss = Math.min(w, h);
      const dy2 = ss * a / 100000;
      const dy1 = dy2 / 5;
      const x1 = w - dy2, x2 = x1 + dy1;
      const y2 = h - dy2, y1 = y2 + dy1;
      const d = `M${x1},${h} L${x2},${y1} L${w},${y2} L${x1},${h} L0,${h} L0,0 L${w},0 L${w},${y2}`;
      return { d };
    }
    case 'corner': {
      // 缺角矩形：右下角被切掉一块
      const a1 = clampV(numAny(adj, ['adj1', 'adj'], 50000), 0, 100000);
      const a2 = clampV(numAny(adj, ['adj2', 'adj'], 50000), 0, 100000);
      const minWH = Math.min(w, h);
      const maxAdj1 = 100000 * h / minWH;
      const maxAdj2 = 100000 * w / minWH;
      const v1 = Math.min(a1, maxAdj1), v2 = Math.min(a2, maxAdj2);
      const x1 = minWH * v2 / 100000;
      const dy1 = minWH * v1 / 100000;
      const y1 = h - dy1;
      return { d: `M0,0 L${x1},0 L${x1},${y1} L${w},${y1} L${w},${h} L0,${h} z` };
    }
    case 'diagStripe': {
      const a1 = clampV(numAny(adj, ['adj', 'adj1'], 50000), 0, 100000);
      const x2 = w * a1 / 100000;
      const y2 = h * a1 / 100000;
      return { d: `M0,${y2} L${x2},0 L${w},0 L0,${h} z` };
    }
    default:
      return null;
  }
}
