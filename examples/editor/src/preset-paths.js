/**
 * OOXML 预设几何 → SVG 路径（与预览端 pptxToHtml / src/shape 同源公式移植）。
 * 覆盖编辑器此前用 CSS 近似失真的自定义形状：pie/arc/chord、noSmoking、smileyFace、
 * plus、plaque、quadArrow、uturnArrow、wedgeRectCallout、callout 系列、
 * ellipseRibbon、flowChartMagneticDisk/Drum、flowChartMultidocument、teardrop。
 */

/** 角度(度) → 圆上点（w/h 为直径、含 -90° 相位，与预览端 polarToCartesian 一致） */
function polarPt(cx, cy, w, h, angleDeg) {
  const a = (angleDeg - 90) * Math.PI / 180;
  return { x: cx + (w / 2) * Math.cos(a), y: cy + (h / 2) * Math.sin(a) };
}
const fmt = (n) => parseFloat(Number(n).toFixed(2));

/** 与预览端 shapeArc 相同：起点取 endAngle、终点取 startAngle（cw=false 逆时针标记） */
function shapeArc(cx, cy, w, h, startAngle, endAngle, clockwise) {
  const start = polarPt(cx, cy, w, h, endAngle);
  const end = polarPt(cx, cy, w, h, startAngle);
  const largeArcFlag = endAngle - startAngle <= 180 ? '0' : '1';
  return ['M', fmt(start.x), fmt(start.y), 'A', fmt(w), fmt(h), 0, largeArcFlag, clockwise ? '0' : '1', fmt(end.x), fmt(end.y)].join(' ');
}

/**
 * 与预览端 shapeArcAlt 相同：逐度折线逼近的弧（rX/rY 为半径）。
 * moveTo=false 时首点用 L 续接当前子路径（对应 src/shape/shape.ts 里的 .replace("M","L")），
 * 否则会多开子路径导致闭合图形出现内接多边形伪影。
 */
function shapeArcAlt(cX, cY, rX, rY, stAng, endAng, moveTo = true) {
  let d = '';
  let angle = stAng;
  const head = moveTo ? 'M' : 'L';
  if (endAng >= stAng) {
    while (angle <= endAng) {
      const rad = angle * Math.PI / 180;
      const x = cX + Math.cos(rad) * rX, y = cY + Math.sin(rad) * rY;
      if (angle === stAng) d = ` ${head}${fmt(x)} ${fmt(y)}`;
      d += ` L${fmt(x)} ${fmt(y)}`;
      angle++;
    }
  } else {
    while (angle > endAng) {
      const rad = angle * Math.PI / 180;
      const x = cX + Math.cos(rad) * rX, y = cY + Math.sin(rad) * rY;
      if (angle === stAng) d = ` ${head}${fmt(x)} ${fmt(y)}`;
      d += ` L ${fmt(x)} ${fmt(y)}`;
      angle--;
    }
  }
  return d;
}

/** 与预览端 shapePie 相同：饼形/弧形（H 为高，radius = H/2；返回 [d, transform]） */
function shapePie(H, w, adj1, adj2, isClose) {
  const pieVal = parseInt(String(adj2));
  const piAngle = parseInt(String(adj1));
  const radius = parseInt(String(H)) / 2;
  let value = pieVal - piAngle;
  if (value < 0) value = 360 + value;
  value = Math.min(Math.max(value, 0), 360);
  const x = Math.cos((2 * Math.PI) / (360 / value));
  const y = Math.sin((2 * Math.PI) / (360 / value));
  const longArc = value <= 180 ? 0 : 1;
  if (isClose) {
    const d = `M${radius},${radius} L${radius},${0} A${radius},${radius} 0 ${longArc},1 ${(radius + y * radius)},${(radius - x * radius)} z`;
    return [d, `rotate(${piAngle - 270}, ${radius}, ${radius})`];
  }
  const radius1 = radius, radius2 = w / 2;
  const d = `M${radius1},${0} A${radius2},${radius1} 0 ${longArc},1 ${(radius2 + y * radius2)},${(radius1 - x * radius1)}`;
  return [d, `rotate(${piAngle + 90}, ${radius}, ${radius})`];
}

const num = (adj, key, def) => {
  const v = adj && adj[key] != null ? Number(adj[key]) : def;
  return Number.isFinite(v) ? v : def;
};
const clampV = (v, lo, hi) => Math.min(Math.max(v, lo), hi);

/** 矩形闭合路径 */
const rectD = (w, h) => `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z`;

/**
 * 主入口：返回 { d, transform, fillRule, strokes:[{d,width}], strokeOnly }
 * 不支持的 preset 返回 null（由调用方回退 CSS 近似）。
 */
export function presetShapePath(prst, w, h, adj = {}, opts = {}) {
  switch (prst) {
    // 星形成（star4/5/6/8/10/12/16/24/32）：外/内顶点交替，内半径 = 外半径 * adj/50000
    case 'star4': case 'star5': case 'star6': case 'star8': case 'star10':
    case 'star12': case 'star16': case 'star24': case 'star32': {
      const N = { star4: 4, star5: 5, star6: 6, star8: 8, star10: 10, star12: 12, star16: 16, star24: 24, star32: 32 }[prst];
      const defAdj = prst === 'star5' ? 19098 : 12500;
      const ratio = num(adj, 'adj', defAdj) / 50000;
      const rx = w / 2, ry = h / 2;
      const pts = [];
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
      if (adj.adj1 != null) adj1 = num(adj, 'adj1', adj1) / 60000;
      if (adj.adj2 != null) adj2 = num(adj, 'adj2', adj2) / 60000;
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
    case 'flowChartMagneticDisk':
    case 'flowChartMagneticDrum': {
      const ss = Math.min(w, h);
      const maxAdj = 50000 * h / ss;
      const a = clampV(50000, 0, maxAdj);
      const y1 = ss * a / 200000;
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
    default:
      return null;
  }
}
