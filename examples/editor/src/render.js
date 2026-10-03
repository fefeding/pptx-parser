/**
 * 幻灯片渲染：内部模型 → DOM（画布 / 缩略图 / 演示共用）
 */
import { h, ptToPx, normalizeColor, withAlpha, hexToRgb, clamp } from './util.js';
import { renderChartSVG } from './charts.js';
import { getTheme } from './model.js';

/** 元素的实际矩形（组合取子元素的并集） */
export function elementRect(el) {
  if (el.type === 'group') {
    const kids = el.children || [];
    if (!kids.length) return { x: el.x, y: el.y, width: el.width, height: el.height };
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const c of kids) {
      const r = rotatedRect(c);
      x0 = Math.min(x0, r.x0); y0 = Math.min(y0, r.y0);
      x1 = Math.max(x1, r.x1); y1 = Math.max(y1, r.y1);
    }
    return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
  }
  return { x: el.x, y: el.y, width: el.width, height: el.height };
}

export function rotatedRect(el) {
  const w = el.width || 0, hh = el.height || 0;
  if (!el.rotation) return { x0: el.x, y0: el.y, x1: el.x + w, y1: el.y + hh };
  const cx = el.x + w / 2, cy = el.y + hh / 2;
  const rad = (el.rotation * Math.PI) / 180;
  const cos = Math.abs(Math.cos(rad)), sin = Math.abs(Math.sin(rad));
  const nw = w * cos + hh * sin, nh = w * sin + hh * cos;
  return { x0: cx - nw / 2, y0: cy - nh / 2, x1: cx + nw / 2, y1: cy + nh / 2 };
}

/* ======================= 背景 ======================= */
export function backgroundStyle(bg, theme) {
  if (!bg || bg === 'none') return { background: '#FFFFFF' };
  if (typeof bg === 'string') return { background: normalizeColor(bg) || '#FFFFFF' };
  if (bg.type === 'gradient') {
    const stops = (bg.stops || []).slice().sort((a, b) => a.position - b.position);
    const deg = bg.direction === 'vertical' ? 180 : bg.direction === 'diagonal' ? 135 : 90;
    const list = stops.length >= 2
      ? stops.map((s) => `${normalizeColor(s.color) || '#FFFFFF'} ${clamp(s.position, 0, 1) * 100}%`).join(',')
      : '#FFFFFF,#FFFFFF';
    return { background: `linear-gradient(${deg}deg, ${list})` };
  }
  if (bg.type === 'image') {
    const url = bg.data || bg.src || '';
    return {
      background: `${normalizeColor(theme?.bg) || '#FFFFFF'} url("${url}") center / cover no-repeat`
    };
  }
  return { background: normalizeColor(bg.color) || '#FFFFFF' };
}

/* ======================= 元素渲染 ======================= */
const DASH_MAP = { solid: 'solid', dash: 'dashed', dashDot: 'dashed', dotted: 'dotted', lgDash: 'dashed', sysDot: 'dotted' };

function boxStyle(el) {
  const r = elementRect(el);
  return {
    left: `${r.x}px`, top: `${r.y}px`, width: `${r.width}px`, height: `${r.height}px`,
    transform: el.rotation ? `rotate(${el.rotation}deg)` : '',
    transformOrigin: 'center center',
    zIndex: ''
  };
}

function shapeVisual(el) {
  const style = {};
  let clip = null, radius = null;
  const fill = el.fill;
  if (fill && fill !== 'none') {
    if (typeof fill === 'string') {
      style.background = normalizeColor(fill) || 'transparent';
    } else if (fill.type === 'gradient') {
      const stops = (fill.stops || []).slice().sort((a, b) => a.position - b.position);
      const deg = fill.direction === 'vertical' ? 180 : fill.direction === 'diagonal' ? 135 : 90;
      style.background = `linear-gradient(${deg}deg, ${stops.map((s) => `${withAlpha(s.color, 0)} ${clamp(s.position, 0, 1) * 100}%`).join(',')})`;
    } else {
      style.background = withAlpha(fill.color, fill.transparency || 0);
    }
  } else {
    style.background = 'transparent';
  }
  if (el.line && el.line !== 'none' && el.line.color) {
    const w = Math.max(0.5, ptToPx(el.line.width || 1));
    style.border = `${w}px ${DASH_MAP[el.line.dashType] || 'solid'} ${withAlpha(el.line.color, el.line.transparency || 0)}`;
  }
  if (el.shadow) {
    const s = el.shadow;
    const rad = ((s.angle ?? 45) * Math.PI) / 180;
    const dx = (s.distance ?? 4) * Math.sin(rad);
    const dy = (s.distance ?? 4) * Math.cos(rad);
    const { r, g, b } = hexToRgb(s.color || '#000000');
    const alpha = clamp(1 - (s.transparency ?? 60) / 100, 0, 1) * 0.6;
    style.boxShadow = `${dx.toFixed(1)}px ${dy.toFixed(1)}px ${(s.blur ?? 8).toFixed(1)}px rgba(${r},${g},${b},${alpha.toFixed(2)})`;
  }
  return { style, clip, radius };
}

function renderTextBody(el, ctx) {
  const body = h('div', { class: 'tb-body' });
  const vmap = { top: 'flex-start', middle: 'center', bottom: 'flex-end' };
  body.style.justifyContent = vmap[el.valign] || 'flex-start';
  body.style.whiteSpace = 'pre-wrap';
  body.style.counterReset = 'pnum 0';
  (el.paragraphs || []).forEach((p) => {
    const para = h('div', {});
    para.style.textAlign = p.align || el.align || 'left';
    if (p.lineSpacing) para.style.lineHeight = String(p.lineSpacing);
    const bullet = p.bullet ?? el.bullet;
    if (bullet) para.classList.add(bullet === 'number' ? 'num-para' : 'bullet-para');
    if (el.indent) para.style.marginLeft = `${el.indent}px`;
    (p.runs || []).forEach((run) => {
      const span = document.createElement('span');
      span.textContent = run.text == null ? '' : String(run.text);
      const st = span.style;
      st.fontSize = `${ptToPx(run.fontSize ?? el.fontSize ?? 18)}px`;
      st.fontFamily = quoteFont(run.fontFace || el.fontFace || '微软雅黑');
      st.color = normalizeColor(run.color) || normalizeColor(el.color) || '#202124';
      if (run.bold ?? el.bold) st.fontWeight = '700';
      if (run.italic ?? el.italic) st.fontStyle = 'italic';
      if (run.underline ?? el.underline) st.textDecoration = 'underline';
      para.appendChild(span);
    });
    if (!para.childNodes.length) para.appendChild(document.createElement('br'));
    body.appendChild(para);
  });
  if (ctx.editing) {
    body.contentEditable = 'true';
    body.spellcheck = false;
  }
  return body;
}

function quoteFont(f) {
  return /^[A-Za-z0-9 _-]+$/.test(f) ? f : `"${f}"`;
}

function renderTableBody(el) {
  const grid = h('div', { class: 'el-table', style: { width: '100%', height: '100%' } });
  const rows = el.rows || [];
  const cols = Math.max(1, ...rows.map((r) => (r.cells || []).length));
  let colWidths = (el.colWidths && el.colWidths.length === cols) ? el.colWidths : new Array(cols).fill(1);
  const sumW = colWidths.reduce((a, b) => a + (Number(b) || 0), 0) || 1;
  const wScale = (el.width || 400) / sumW;
  grid.style.gridTemplateColumns = colWidths.map((w) => `${(Number(w) || 1) * wScale}px`).join(' ');
  const sumH = rows.reduce((a, r) => a + (Number(r.height) || 0), 0) || 1;
  const hScale = (el.height || 200) / sumH;
  grid.style.gridTemplateRows = rows.map((r) => `${(Number(r.height) || 40) * hScale}px`).join(' ');

  const bw = Math.max(0.5, ptToPx(el.border?.width ?? 1));
  const bc = normalizeColor(el.border?.color) || '#CBD5E1';
  rows.forEach((row) => {
    (row.cells || []).forEach((cell) => {
      const d = h('div', { class: 'cell cell-' + (cell.valign || 'middle') });
      d.style.border = `${bw}px solid ${bc}`;
      if (cell.fill) d.style.background = normalizeColor(cell.fill) || 'transparent';
      d.style.fontSize = `${ptToPx(cell.fontSize || 14)}px`;
      d.style.color = normalizeColor(cell.color) || '#202124';
      if (cell.bold) d.style.fontWeight = '700';
      d.style.textAlign = cell.align || 'left';
      d.style.justifyContent = cell.align === 'center' ? 'center' : cell.align === 'right' ? 'flex-end' : 'flex-start';
      if (cell.colSpan > 1) d.style.gridColumn = `span ${cell.colSpan}`;
      if (cell.rowSpan > 1) d.style.gridRow = `span ${cell.rowSpan}`;
      d.textContent = cell.text || '';
      grid.appendChild(d);
    });
  });
  return grid;
}

export function renderElement(el, ctx = {}) {
  const node = h('div', {
    class: `el el-${el.type}${el.locked ? ' locked' : ''}${el.hidden ? ' hidden-el' : ''}`,
    dataset: { id: el.id, type: el.type }
  });
  Object.assign(node.style, boxStyle(el));

  switch (el.type) {
    case 'text': {
      node.appendChild(renderTextBody(el, ctx));
      break;
    }
    case 'shape': {
      if (el.shapeType === 'line') {
        const ln = h('div', { style: { position: 'absolute', left: '0', top: '50%', width: '100%' } });
        const w = Math.max(1, ptToPx((el.line && el.line.width) || 2));
        ln.style.height = `${w}px`;
        ln.style.background = normalizeColor((el.line && el.line.color) || '#000000') || '#000';
        ln.style.transform = 'translateY(-50%)';
        node.appendChild(ln);
      } else {
        const inner = h('div', { style: { position: 'absolute', inset: '0' } });
        const { style } = shapeVisual(el);
        Object.assign(inner.style, style);
        const geo = shapeGeometry(el);
        Object.assign(inner.style, geo);
        node.appendChild(inner);
      }
      break;
    }
    case 'image': {
      const img = document.createElement('img');
      img.src = el.data || el.src || '';
      img.draggable = false;
      const adj = el.imageAdjust || {};
      const filters = [];
      if (adj.brightness) filters.push(`brightness(${1 + adj.brightness / 100})`);
      if (adj.contrast) filters.push(`contrast(${1 + adj.contrast / 100})`);
      if (filters.length) img.style.filter = filters.join(' ');
      if (adj.transparency) img.style.opacity = String(clamp(1 - adj.transparency / 100, 0, 1));
      node.appendChild(img);
      break;
    }
    case 'table': {
      node.appendChild(renderTableBody(el));
      break;
    }
    case 'chart': {
      const wrap = h('div', { class: 'el-chart', style: { width: '100%', height: '100%' } });
      const theme = getTheme(ctx.theme);
      wrap.innerHTML = renderChartSVG(el, {
        width: el.width, height: el.height,
        palette: theme.accents,
        textColor: theme.text, titleColor: theme.title
      });
      node.appendChild(wrap);
      break;
    }
    case 'group': {
      const rect = elementRect(el);
      for (const child of el.children || []) {
        const cn = renderElement(child, ctx);
        cn.style.left = `${child.x - rect.x}px`;
        cn.style.top = `${child.y - rect.y}px`;
        node.appendChild(cn);
      }
      break;
    }
    default: {
      node.innerHTML = `<div class="el-raw" style="width:100%;height:100%">${el.label || el.type || '未支持元素'}</div>`;
    }
  }
  if (ctx.editing) node.dataset.editing = '1';
  return node;
}

function shapeGeometry(el) {
  const r = Math.min(el.width, el.height) * 0.16;
  switch (el.shapeType) {
    case 'ellipse': case 'moon': case 'sun': case 'donut': case 'arc': case 'smileyFace':
      return { borderRadius: '50%' };
    case 'roundRect': return { borderRadius: `${Math.min(r, 40)}px` };
    case 'gear6': return { borderRadius: '18%' };
    case 'can': return { borderRadius: '10% / 18%' };
    case 'cloud': return { borderRadius: '44% 44% 38% 38% / 50% 50% 50% 50%' };
    case 'triangle': return { clipPath: 'polygon(50% 0%,100% 100%,0% 100%)' };
    case 'rtTriangle': return { clipPath: 'polygon(0% 0%,100% 100%,0% 100%)' };
    case 'diamond': return { clipPath: 'polygon(50% 0%,100% 50%,50% 100%,0% 50%)' };
    case 'parallelogram': return { clipPath: 'polygon(20% 0%,100% 0%,80% 100%,0% 100%)' };
    case 'trapezoid': return { clipPath: 'polygon(20% 0%,80% 0%,100% 100%,0% 100%)' };
    case 'pentagon': return { clipPath: 'polygon(50% 0%,100% 38%,82% 100%,18% 100%,0% 38%)' };
    case 'hexagon': return { clipPath: 'polygon(25% 0%,75% 0%,100% 50%,75% 100%,25% 100%,0% 50%)' };
    case 'octagon': return { clipPath: 'polygon(30% 0%,70% 0%,100% 30%,100% 70%,70% 100%,30% 100%,0% 70%,0% 30%)' };
    case 'chevron': return { clipPath: 'polygon(0% 0%,62% 0%,100% 50%,62% 100%,0% 100%,38% 50%)' };
    case 'rightArrow': return { clipPath: 'polygon(0% 20%,58% 20%,58% 0%,100% 50%,58% 100%,58% 80%,0% 80%)' };
    case 'leftArrow': return { clipPath: 'polygon(100% 20%,42% 20%,42% 0%,0% 50%,42% 100%,42% 80%,100% 80%)' };
    case 'upArrow': return { clipPath: 'polygon(50% 0%,100% 100%,70% 100%,70% 58%,30% 58%,30% 100%,0% 100%)' };
    case 'downArrow': return { clipPath: 'polygon(50% 100%,0% 0%,30% 0%,30% 42%,70% 42%,70% 0%,100% 0%)' };
    case 'pentagonBlock': return { clipPath: 'polygon(50% 0%,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%)' };
    case 'plus': return { clipPath: 'polygon(35% 0%,65% 0%,65% 35%,100% 35%,100% 65%,65% 65%,65% 100%,35% 100%,35% 65%,0% 65%,0% 35%,35% 35%)' };
    case 'heart': return { clipPath: 'polygon(50% 100%,2% 55%,2% 28%,25% 6%,50% 18%,75% 6%,98% 28%,98% 55%)' };
    case 'lightningBolt': return { clipPath: 'polygon(56% 0%,18% 56%,46% 56%,34% 100%,82% 38%,54% 38%)' };
    case 'pie': return { clipPath: 'polygon(50% 50%,50% 0%,100% 12%,100% 50%)' };
    case 'cube': return { clipPath: 'polygon(50% 0%,100% 25%,100% 75%,50% 100%,0% 75%,0% 25%)' };
    case 'funnel': return { clipPath: 'polygon(0% 0%,100% 0%,65% 56%,65% 100%,35% 100%,35% 56%)' };
    case 'foldedCorner': return { clipPath: 'polygon(0% 0%,68% 0%,100% 32%,100% 100%,0% 100%)' };
    case 'frame': return { border: '8px solid currentColor', color: '#9AA0A6' };
    default: return {};
  }
}

/* ======================= 页面 ======================= */
export function renderSlideInto(frame, slide, doc, opts = {}) {
  const theme = getTheme(doc.theme);
  frame.innerHTML = '';
  frame.style.width = `${doc.slideSize.width}px`;
  frame.style.height = `${doc.slideSize.height}px`;
  Object.assign(frame.style, backgroundStyle(slide.background, theme));
  frame.classList.toggle('show-grid', !!opts.grid);
  if (opts.scale && opts.scale !== 1) {
    frame.style.transform = `scale(${opts.scale})`;
    frame.style.transformOrigin = 'top left';
  } else {
    frame.style.transform = '';
  }
  for (const el of slide.elements || []) {
    if (el.hidden && !opts.showHidden) continue;
    frame.appendChild(renderElement(el, { theme: doc.theme, editing: opts.editingId === el.id }));
  }
}

/** 缩略图渲染：内部按幻灯片尺寸渲染后整体缩放 */
export function renderThumbInto(host, slide, doc) {
  const W = doc.slideSize.width, H = doc.slideSize.height;
  const boxW = host.clientWidth || 150;
  const scale = boxW / W;
  host.innerHTML = '';
  const inner = h('div', { class: 'thumb-inner' });
  inner.style.width = `${W}px`;
  inner.style.height = `${H}px`;
  inner.style.transform = `scale(${scale})`;
  inner.style.transformOrigin = 'top left';
  host.style.height = `${H * scale}px`;
  const theme = getTheme(doc.theme);
  Object.assign(inner.style, backgroundStyle(slide.background, theme));
  for (const el of slide.elements || []) {
    if (el.hidden) continue;
    inner.appendChild(renderElement(el, { theme: doc.theme }));
  }
  host.appendChild(inner);
  return scale;
}
