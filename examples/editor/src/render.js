/**
 * 幻灯片渲染：内部模型 → DOM（画布 / 缩略图 / 演示共用）
 */
import { h, ptToPx, normalizeColor, withAlpha, hexToRgb, clamp } from './util.js';
import { presetShapePath } from './preset-paths.js';
import { renderChartSVG } from './charts.js';
import { getTheme } from './model.js';
// 预览端（examples/index.html）使用的同一套 ECharts 图表渲染器：option 构建逻辑（含 3D）与预览完全一致
import { chartRenderer } from '../../chart-lib/chart-renderer.js';

const MEDIA_MIME = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp',
  svg: 'image/svg+xml', webp: 'image/webp', tiff: 'image/tiff',
  mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', webm: 'video/webm', avi: 'video/x-msvideo',
  mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', aac: 'audio/aac', ogg: 'audio/ogg', wma: 'audio/x-ms-wma'
};

/**
 * 媒体源规范化：parser 现在统一输出完整 dataURL，但为兼容旧版 JSON（媒体为裸 base64），
 * 当 data 不是 data:/http(s): 开头时，依据扩展名补全为 dataURL，避免被浏览器当作相对路径请求 404。
 */
function mediaSrc(value, ext, fallbackMime) {
  if (!value) return '';
  if (/^data:/i.test(value) || /^https?:/i.test(value)) return value;
  const mime = fallbackMime || (ext && MEDIA_MIME[String(ext).toLowerCase().replace(/^\./, '')]) || 'application/octet-stream';
  return `data:${mime};base64,${value}`;
}

/** 元素的实际矩形（组合：page 坐标取子元素并集；relative 坐标直接用声明矩形） */
export function elementRect(el) {
  if (el.type === 'group') {
    const kids = el.children || [];
    // 子元素为组合相对坐标时，其并集原点随内容浮动，不能代表组合位置，用声明的 x/y/w/h
    if (!kids.length || el.childrenCoordinates === 'relative') {
      return { x: el.x, y: el.y, width: el.width || 0, height: el.height || 0 };
    }
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

/** 为包含阴影/发光等效果的元素计算选择框外边距（左右/上下，单位 px） */
export function effectMargin(el) {
  let left = 0, top = 0, right = 0, bottom = 0;
  if (el.shadow) {
    const s = el.shadow;
    const rad = ((s.angle ?? 45) * Math.PI) / 180;
    const dist = ptToPx(s.distance ?? 4);
    const dx = dist * Math.cos(rad);
    const dy = dist * Math.sin(rad);
    const blur = ptToPx(s.blur ?? 8);
    const EPS = 1e-6;
    if (dx > EPS) right += dx + blur; else if (dx < -EPS) left += -dx + blur;
    else { right += blur; left += blur; }
    if (dy > EPS) bottom += dy + blur; else if (dy < -EPS) top += -dy + blur;
    else { bottom += blur; top += blur; }
  }
  if (el.glow && el.glow.blur) {
    const blur = ptToPx(el.glow.blur);
    left = Math.max(left, blur); right = Math.max(right, blur);
    top = Math.max(top, blur); bottom = Math.max(bottom, blur);
  }
  return { left, top, right, bottom };
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
  const effects = {};
  const fill = el.fill;
  if (fill && fill !== 'none') {
    if (typeof fill === 'string') {
      style.background = normalizeColor(fill) || 'transparent';
    } else if (fill.type === 'gradient') {
      const stops = (fill.stops || []).slice().sort((a, b) => a.position - b.position);
      const deg = fill.direction === 'vertical' ? 180 : fill.direction === 'diagonal' ? 135 : 90;
      style.background = `linear-gradient(${deg}deg, ${stops.map((s) => `${withAlpha(s.color, 0)} ${clamp(s.position, 0, 1) * 100}%`).join(',')})`;
    } else if (fill.type === 'image') {
      // 形状图片填充：铺满 / 平铺（tile）/ 平铺 + 源图裁剪（srcRect）
      const url = fill.data || fill.src || '';
      const sr = fill.srcRect || {};
      const cw = Math.max(0.05, 1 - (sr.l || 0) - (sr.r || 0));
      const ch = Math.max(0.05, 1 - (sr.t || 0) - (sr.b || 0));
      if (!url) {
        style.background = 'transparent';
      } else if (fill.tile) {
        const sx = (fill.tile.sx ?? 1) / cw, sy = (fill.tile.sy ?? 1) / ch;
        style.background = `url("${url}") repeat`;
        style.backgroundSize = `${(sx * 100).toFixed(2)}% ${(sy * 100).toFixed(2)}%`;
      } else {
        style.background = `url("${url}") center / cover no-repeat`;
      }
    } else if (fill.type === 'pattern') {
      // 图案填充：交给 renderElement 用 SVG <pattern> 渲染（与预览端一致，支持圆角/异形裁剪）
      style.background = 'transparent';
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
    // OOXML outerShdw 的 dir：0° 向右，顺时针增加；90° 为向下偏移。
    const rad = ((s.angle ?? 45) * Math.PI) / 180;
    const dist = ptToPx(s.distance ?? 4);
    const dx = dist * Math.cos(rad);
    const dy = dist * Math.sin(rad);
    const blur = ptToPx(s.blur ?? 8);
    const { r, g, b } = hexToRgb(s.color || '#000000');
    const alpha = clamp(1 - (s.transparency ?? 60) / 100, 0, 1);
    effects.shadow = {
      dx: Number(dx.toFixed(2)),
      dy: Number(dy.toFixed(2)),
      blur: Number(blur.toFixed(2)),
      color: `rgba(${r},${g},${b},${alpha.toFixed(2)})`
    };
  }
  // 发光：a:glow，颜色来自效果自身
  if (el.glow && el.glow.color) {
    const { r, g, b } = hexToRgb(el.glow.color);
    const blur = Math.max(2, ptToPx(el.glow.blur ?? 5));
    const alpha = clamp(1 - (el.glow.transparency ?? 25) / 100, 0, 1);
    effects.glow = {
      blur: Number(blur.toFixed(2)),
      color: `rgba(${r},${g},${b},${alpha.toFixed(2)})`
    };
  }
  return { style, effects };
}

let _patternUid = 0;
/** 构造 OOXML 图案预设对应的 SVG <pattern> 定义与填充 rect，与预览端 buildPatternTile 对齐 */
function createPatternFillLayer(el) {
  const fill = el.fill;
  if (!fill || fill.type !== 'pattern') return null;
  const id = `pat_${++_patternUid}_${Math.random().toString(36).slice(2, 6)}`;
  const fg = normalizeColor(fill.fg) || '#1A73E8';
  const bg = normalizeColor(fill.bg) || '#FFFFFF';
  const prst = fill.prst || 'pct10';
  const tile = buildPatternTile(prst, fg, bg);

  const svgNs = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(svgNs, 'svg');
  svg.setAttribute('width', '100%');
  svg.setAttribute('height', '100%');
  svg.style.position = 'absolute';
  svg.style.inset = '0';
  svg.style.display = 'block';
  svg.style.pointerEvents = 'none';

  const defs = document.createElementNS(svgNs, 'defs');
  const pattern = document.createElementNS(svgNs, 'pattern');
  pattern.setAttribute('id', id);
  pattern.setAttribute('width', String(tile.size));
  pattern.setAttribute('height', String(tile.size));
  pattern.setAttribute('patternUnits', 'userSpaceOnUse');
  pattern.innerHTML = `<rect width="${tile.size}" height="${tile.size}" fill="${bg}"/>${tile.body}`;
  defs.appendChild(pattern);
  svg.appendChild(defs);

  const rect = document.createElementNS(svgNs, 'rect');
  rect.setAttribute('width', '100%');
  rect.setAttribute('height', '100%');
  rect.setAttribute('fill', `url(#${id})`);
  svg.appendChild(rect);
  return svg;
}

/** 根据 prst 生成一块 tile 的 SVG body 与 size，与预览端 buildPatternTile 语义对齐 */
function buildPatternTile(prst, fg, bg) {
  const stroke = (d, width, dash) => `<path d="${d}" fill="none" stroke="${fg}" stroke-width="${width}"${dash ? ` stroke-dasharray="${dash}"` : ''}/>`;
  const fillPath = (d) => `<path d="${d}" fill="${fg}"/>`;
  const lineFamilies = {
    horz: ['h', 8, 1, ''], ltHorz: ['h', 8, 1, ''], dkHorz: ['h', 8, 3, ''], narHorz: ['h', 4, 1, ''], dashHorz: ['h', 8, 2, '4 4'],
    vert: ['v', 8, 1, ''], ltVert: ['v', 8, 1, ''], dkVert: ['v', 8, 3, ''], narVert: ['v', 4, 1, ''], dashVert: ['v', 8, 2, '4 4'],
    dnDiag: ['dn', 8, 1, ''], ltDnDiag: ['dn', 4, 1, ''], dkDnDiag: ['dn', 8, 3, ''], wdDnDiag: ['dn', 8, 4, ''], dashDnDiag: ['dn', 8, 2, '4 4'],
    upDiag: ['up', 8, 1, ''], ltUpDiag: ['up', 4, 1, ''], dkUpDiag: ['up', 8, 3, ''], wdUpDiag: ['up', 8, 4, ''], dashUpDiag: ['up', 8, 2, '4 4']
  };
  const family = lineFamilies[prst];
  if (family) {
    const [dir, size, width, dash] = family;
    const half = size / 2;
    let d = '';
    if (dir === 'h') d = `M0 ${half}H${size}`;
    else if (dir === 'v') d = `M${half} 0V${size}`;
    else if (dir === 'dn') d = `M0 0L${size} ${size}M${half} 0L${size} ${half}M0 ${half}L${half} ${size}`;
    else d = `M0 ${size}L${size} 0M0 ${half}L${half} 0M${half} ${size}L${size} ${half}`;
    return { size, body: stroke(d, width, dash || undefined) };
  }
  switch (prst) {
    case 'cross': return { size: 8, body: stroke('M0 4H8M4 0V8', 2) };
    case 'diagCross': return { size: 8, body: stroke('M0 0L8 8M8 0L0 8', 2) };
    case 'smGrid': return { size: 4, body: stroke('M0 0H4M0 0V4', 1) };
    case 'lgGrid': return { size: 8, body: stroke('M0 0H8M0 0V8', 2) };
    case 'dotGrid': return { size: 8, body: stroke('M0 4H8M4 0V8', 1, '2 2') };
    case 'smCheck': return { size: 4, body: fillPath('M0 0h2v2h-2zM2 2h2v2h-2z') };
    case 'lgCheck': return { size: 8, body: fillPath('M0 0h4v4h-4zM4 4h4v4h-4z') };
    case 'dotDmnd': return { size: 4, body: fillPath('M2 0L4 2L2 4L0 2Z') };
    case 'solidDmnd': return { size: 8, body: fillPath('M4 1L7 4L4 7L1 4Z') };
    case 'openDmnd': return { size: 8, body: stroke('M4 1L7 4L4 7L1 4Z', 1) };
    case 'smConfetti': return { size: 8, body: fillPath('M1 1h1v1h-1zM5 2h1v1h-1zM3 5h1v1h-1z') };
    case 'lgConfetti': return { size: 16, body: fillPath('M2 2h3v3h-3zM9 5h3v3h-3zM5 10h3v3h-3z') };
    case 'horzBrick': return { size: 8, body: stroke('M0 0H8M0 4H8M2 0V4M6 4V8', 1.5) };
    case 'diagBrick': return { size: 8, body: stroke('M0 4L4 0M4 8L8 4M0 4L4 8M4 0L8 4', 1.5) };
    case 'weave': return { size: 8, body: stroke('M0 2H8M0 6H8M2 0V8M6 0V8', 1) };
    case 'trellis': return { size: 8, body: stroke('M0 0H8M0 0V8M0 8L8 0', 1) };
    case 'plaid': return { size: 8, body: stroke('M0 4H8M4 0V8', 3) };
    case 'shingle':
    case 'wave': return { size: 8, body: stroke('M0 2Q2 0 4 2T8 2M0 6Q2 4 4 6T8 6', 1) };
    case 'zigZag': return { size: 8, body: stroke('M0 4L2 2L4 4L6 2L8 4', 1) };
    case 'sphere': return { size: 8, body: `<circle cx="4" cy="4" r="2.6" fill="${fg}"/><circle cx="3.2" cy="3.2" r="0.8" fill="#ffffff" fill-opacity="0.55"/>` };
    case 'divot': return { size: 4, body: `<circle cx="2" cy="2" r="1.1" fill="${fg}"/>` };
  }
  const pct = /^pct(\d+)$/.exec(prst);
  if (pct) {
    const size = 16;
    const n = Math.min(15, Math.max(1, Math.round(Math.sqrt(Number(pct[1]) / 100) * size)));
    const step = size / n;
    const dots = [];
    for (let r = 0; r < n; r++) {
      for (let c = 0; c < n; c++) {
        const x = (c * step + step / 2).toFixed(2);
        const y = (r * step + step / 2).toFixed(2);
        dots.push(`M${x} ${y}h0.9v0.9h-0.9z`);
      }
    }
    return { size, body: fillPath(dots.join('')) };
  }
  return { size: 8, body: stroke('M0 8L8 0', 1) };
}

/** 形状图片填充（平铺 / 裁剪）异步应用。先以拉伸图占位，图片加载后按 srcRect 裁切并平铺。 */
function applyShapeImageFill(inner, el) {
  const fill = el.fill;
  if (!fill || fill.type !== 'image') return;
  const url = fill.data || fill.src || '';
  if (!url) return;

  // 无平铺无裁剪：CSS 背景拉伸铺满即可
  const sr = fill.srcRect || {};
  const hasCrop = sr.l || sr.t || sr.r || sr.b;
  const tile = fill.tile || {};
  const hasTile = tile.sx != null || tile.sy != null;
  if (!hasCrop && !hasTile) {
    inner.style.background = `url("${url}") center / cover no-repeat`;
    return;
  }

  const img = new Image();
  img.crossOrigin = 'anonymous';
  img.onload = () => {
    const iw = img.naturalWidth || 1, ih = img.naturalHeight || 1;
    const l = Math.max(0, Math.min(1, sr.l || 0));
    const t = Math.max(0, Math.min(1, sr.t || 0));
    const r = Math.max(0, Math.min(1 - l, sr.r || 0));
    const b = Math.max(0, Math.min(1 - t, sr.b || 0));
    const cropX = l * iw, cropY = t * ih;
    const cropW = Math.max(1, (1 - l - r) * iw), cropH = Math.max(1, (1 - t - b) * ih);

    // 每格像素尺寸：sx/sy 是相对于裁剪后图片尺寸的比例（与 PowerPoint 语义一致）
    const sx = Math.max(0.01, tile.sx ?? 1);
    const sy = Math.max(0.01, tile.sy ?? 1);
    const tileW = cropW * sx;
    const tileH = cropH * sy;

    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(tileW));
    canvas.height = Math.max(1, Math.round(tileH));
    const ctx2d = canvas.getContext('2d');
    if (!ctx2d) return;
    ctx2d.drawImage(img, cropX, cropY, cropW, cropH, 0, 0, tileW, tileH);
    const dataUrl = canvas.toDataURL('image/png');

    inner.style.backgroundImage = `url("${dataUrl}")`;
    inner.style.backgroundRepeat = 'repeat';
    inner.style.backgroundPosition = '0 0';
    inner.style.backgroundSize = `${tileW.toFixed(1)}px ${tileH.toFixed(1)}px`;
    inner.style.backgroundColor = 'transparent';
  };
  // 加载失败则保持占位：拉伸铺满原图
  img.onerror = () => { inner.style.background = `url("${url}") center / cover no-repeat`; };
  img.src = url;
}

/* 自动编号格式化：支持常用 ST_TextAutonumberScheme */
const CN_DIGITS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
const CN_LEGAL = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
function cnNumber(n, legal) {
  const digits = legal ? CN_LEGAL : CN_DIGITS;
  const ten = legal ? '拾' : '十';
  if (n <= 0) return String(n);
  if (n < 10) return digits[n];
  if (n === 10) return ten;
  if (n < 20) return ten + digits[n % 10];
  if (n < 100) return digits[Math.floor(n / 10)] + ten + (n % 10 ? digits[n % 10] : '');
  return String(n);
}
function romanNumber(n, upper) {
  const map = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let out = '', rest = n;
  for (const [v, s] of map) { while (rest >= v) { out += s; rest -= v; } }
  return upper ? out : out.toLowerCase();
}
function alphaNumber(n, upper) {
  let out = '', rest = n;
  while (rest > 0) { rest -= 1; out = String.fromCharCode(97 + (rest % 26)) + out; rest = Math.floor(rest / 26); }
  return upper ? out.toUpperCase() : out;
}
const CIRCLED = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩'];
function formatAutoNum(fmt, n) {
  switch (fmt) {
    case 'arabicPeriod': return `${n}.`;
    case 'arabicParenR': return `${n})`;
    case 'arabicParenBoth': return `(${n})`;
    case 'arabicPlain': return `${n}`;
    case 'alphaLcPeriod': return `${alphaNumber(n, false)}.`;
    case 'alphaUcPeriod': return `${alphaNumber(n, true)}.`;
    case 'alphaLcParenR': return `${alphaNumber(n, false)})`;
    case 'alphaUcParenR': return `${alphaNumber(n, true)})`;
    case 'alphaLcParenBoth': return `(${alphaNumber(n, false)})`;
    case 'alphaUcParenBoth': return `(${alphaNumber(n, true)})`;
    case 'romanLcPeriod': return `${romanNumber(n, false)}.`;
    case 'romanUcPeriod': return `${romanNumber(n, true)}.`;
    case 'romanLcParenR': return `${romanNumber(n, false)})`;
    case 'romanUcParenR': return `${romanNumber(n, true)})`;
    case 'romanLcParenBoth': return `(${romanNumber(n, false)})`;
    case 'romanUcParenBoth': return `(${romanNumber(n, true)})`;
    case 'chineseCounting': case 'chineseCountingThousand': case 'ea1ChsPeriod': case 'ea1ChtPeriod':
      return `${cnNumber(n, false)}、`;
    case 'chineseLegalSimplified':
      return `${cnNumber(n, true)}、`;
    case 'ea1ChsPlain': case 'ea1ChtPlain':
      return cnNumber(n, false);
    case 'ideographDigital': case 'circleNumDbPlain': case 'circleNumWdWhitePlain': case 'circleNumWdBlackPlain':
      return CIRCLED[(n - 1) % CIRCLED.length];
    default: return `${n}.`;
  }
}
/** 段落是否为自动编号 */
function isNumberBullet(b) {
  return b === 'number' || (b && typeof b === 'object' && b.type === 'number');
}

function renderTextBody(el, ctx) {
  const body = h('div', { class: 'tb-body' });
  const vmap = { top: 'flex-start', middle: 'center', bottom: 'flex-end' };
  body.style.justifyContent = vmap[el.valign] || 'flex-start';
  body.style.whiteSpace = 'pre-wrap';
  body.style.counterReset = 'pnum 0';
  // 竖排文字（a:bodyPr/@vert）：eaVert/vert 等 → CSS writing-mode
  if (el.textDirection && el.textDirection !== 'horz') {
    body.style.writingMode = 'vertical-rl';
    body.style.textOrientation = 'upright';
  }
  // 文本框内边距（a:bodyPr/@lIns 等，px）
  if (el.inset) {
    body.style.padding = `${el.inset.t || 0}px ${el.inset.r || 0}px ${el.inset.b || 0}px ${el.inset.l || 0}px`;
  }
  // 编号状态机：连续编号段落共用一个计数器，遇到非编号段落则重置
  let numState = null;
  (el.paragraphs || []).forEach((p) => {
    const para = h('div', {});
    para.style.textAlign = p.align || el.align || 'left';
    if (p.lineSpacing) para.style.lineHeight = String(p.lineSpacing);
    const bullet = p.bullet ?? el.bullet;
    const firstRun = (p.runs || [])[0] || {};
    const bulletFontSize = ptToPx(firstRun.fontSize ?? el.fontSize ?? 18);
    const bulletColor = normalizeColor(firstRun.color) || normalizeColor(el.color) || '#202124';
    const bulletFont = quoteFont(firstRun.fontFace || el.fontFace || '微软雅黑');
    if (isNumberBullet(bullet)) {
      if (!numState) numState = { n: (typeof bullet === 'object' && bullet.start) || 1 };
      const fmt = (typeof bullet === 'object' && bullet.fmt) || 'arabicPeriod';
      para.classList.add('num-para');
      const b = document.createElement('span');
      b.className = 'bullet-mark';
      b.textContent = formatAutoNum(fmt, numState.n) + '\u00A0';
      b.style.fontSize = `${bulletFontSize}px`;
      b.style.color = bulletColor;
      b.style.fontFamily = bulletFont;
      b.contentEditable = 'false';
      para.appendChild(b);
      numState.n += 1;
    } else {
      numState = null;
      if (bullet) {
        para.classList.add('bullet-para');
        const b = document.createElement('span');
        b.className = 'bullet-mark';
        b.textContent = ((typeof bullet === 'object' && bullet.char) || '•') + '\u00A0';
        b.style.fontSize = `${bulletFontSize}px`;
        b.style.color = bulletColor;
        b.style.fontFamily = bulletFont;
        b.contentEditable = 'false';
        para.appendChild(b);
      }
    }
    const indent = p.indent != null ? p.indent : el.indent;
    if (indent) {
      if (indent < 0) {
        // 悬挂缩进（OOXML indent 为负 + marL 组合）：编号/首行不得越出文本框，
        // 与预览端一致地把列表起点收到框内，悬挂量转为正的 padding
        para.style.paddingLeft = `${-indent}px`;
      } else {
        para.style.marginLeft = `${indent}px`;
      }
    }
    if (p.spaceBefore) para.style.marginTop = `${p.spaceBefore}px`;
    if (p.spaceAfter) para.style.marginBottom = `${p.spaceAfter}px`;
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
      // 文字描边 + 外阴影：与预览端一致——描边用四方向 text-shadow 模拟，叠加外阴影
      const shadows = [];
      if (run.outline && run.outline !== 'none' && run.outline.color) {
        const oc = normalizeColor(run.outline.color);
        if (oc) {
          // 预览端规则：宽度 <1pt 按 4/3px、再取整
          const wpx = Math.max(1, Math.floor(run.outline.width && run.outline.width >= 1 ? run.outline.width : 4 / 3));
          shadows.push(`-${wpx}px 0 ${oc}, 0 ${wpx}px ${oc}, ${wpx}px 0 ${oc}, 0 -${wpx}px ${oc}`);
        }
      }
      if (run.shadow && run.shadow.color) {
        const sc = normalizeColor(run.shadow.color);
        if (sc) {
          const a = run.shadow.alpha != null ? run.shadow.alpha : 1;
          shadows.push(`${ptToPx(run.shadow.x || 0)}px ${ptToPx(run.shadow.y || 0)}px ${ptToPx(run.shadow.blur || 0)}px ${withAlpha(sc, a * 100)}`);
        }
      }
      if (shadows.length) st.textShadow = shadows.join(',');
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
  const defaultInset = el.inset || {};
  // 显式计算每个单元格的 grid-column/row：CSS grid auto-placement 遇到 rowspan/colspan
  // 组合时会把源顺序里缺失的占位单元格挤到下一列，导致错位（如 slide 5/6）。
  const occupied = new Map(); // key="r,c" -> true
  const isOcc = (r, c) => occupied.get(`${r},${c}`);
  const mark = (r, c, rs, cs) => {
    for (let i = 0; i < (rs || 1); i++) {
      for (let j = 0; j < (cs || 1); j++) {
        occupied.set(`${r + i},${c + j}`, true);
      }
    }
  };
  rows.forEach((row, ri) => {
    let runningCol = 0; // 按 colSpan 累加得到的期望列，用于跳过同 row 被覆盖的占位单元格
    (row.cells || []).forEach((cell, cj) => {
      // 源顺序里被前面单元格 colSpan 覆盖的占位格跳过
      if (cj < runningCol) return;
      // 找到当前行第一个空列（同时考虑上方 rowSpan 占位）
      let col = runningCol;
      while (col < cols && isOcc(ri, col)) col++;
      if (col >= cols) return;
      const d = h('div', { class: 'cell cell-' + (cell.valign || 'middle') });
      d.style.boxSizing = 'border-box';
      // 分边边框：单元格 borders 优先（'none' 显式无边框），缺省回退表格统一边框
      const b = cell.borders;
      const side = (k) => {
        const v = b && b[k];
        if (v === 'none') return 'none';
        if (v && typeof v === 'object') {
          return `${Math.max(0.5, ptToPx(v.width || 1))}px solid ${normalizeColor(v.color) || '#000000'}`;
        }
        return `${bw}px solid ${bc}`;
      };
      d.style.borderLeft = side('left');
      d.style.borderRight = side('right');
      d.style.borderTop = side('top');
      d.style.borderBottom = side('bottom');
      if (cell.fill) d.style.background = normalizeColor(cell.fill) || 'transparent';
      d.style.fontSize = `${ptToPx(cell.fontSize || 14)}px`;
      d.style.color = normalizeColor(cell.color) || '#202124';
      if (cell.bold) d.style.fontWeight = '700';
      d.style.textAlign = cell.align || 'left';
      d.style.justifyContent = cell.align === 'center' ? 'center' : cell.align === 'right' ? 'flex-end' : 'flex-start';
      // 单元格内边距：单元格级 > 表格级 > CSS 默认
      const inset = cell.inset || defaultInset;
      if (inset.l != null || inset.r != null || inset.t != null || inset.b != null) {
        d.style.paddingTop = `${inset.t != null ? inset.t : 4}px`;
        d.style.paddingRight = `${inset.r != null ? inset.r : 6}px`;
        d.style.paddingBottom = `${inset.b != null ? inset.b : 4}px`;
        d.style.paddingLeft = `${inset.l != null ? inset.l : 6}px`;
      }
      const cs = Math.max(1, Number(cell.colSpan) || 1);
      const rs = Math.max(1, Number(cell.rowSpan) || 1);
      d.style.gridColumn = `${col + 1} / span ${cs}`;
      d.style.gridRow = `${ri + 1} / span ${rs}`;
      mark(ri, col, rs, cs);
      runningCol = col + cs;
      d.textContent = cell.text || '';
      // 对角线：用绝对定位 SVG 覆盖
      const diag = (b && b.diagonal) || el.diagonal;
      if (diag === 'tlBr' || diag === 'blTr' || diag === 'both') {
        const lineColor = (b && b.top && typeof b.top === 'object' && normalizeColor(b.top.color))
          || (el.border && normalizeColor(el.border.color)) || '#000000';
        const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        svg.setAttribute('width', '100%');
        svg.setAttribute('height', '100%');
        svg.style.position = 'absolute';
        svg.style.inset = '0';
        svg.style.pointerEvents = 'none';
        svg.style.zIndex = '1';
        const mkLine = (x1, y1, x2, y2) => {
          const ln = document.createElementNS('http://www.w3.org/2000/svg', 'line');
          ln.setAttribute('x1', x1); ln.setAttribute('y1', y1);
          ln.setAttribute('x2', x2); ln.setAttribute('y2', y2);
          ln.setAttribute('stroke', lineColor);
          ln.setAttribute('stroke-width', '1');
          return ln;
        };
        if (diag === 'tlBr' || diag === 'both') svg.appendChild(mkLine('0', '0', '100%', '100%'));
        if (diag === 'blTr' || diag === 'both') svg.appendChild(mkLine('0', '100%', '100%', '0'));
        d.style.position = 'relative';
        d.appendChild(svg);
      }
      grid.appendChild(d);
    });
  });
  return grid;
}

/** 形状填充色（SVG path 用）：字符串/solid 色返回 #RRGGBB，渐变取首停色，图片/图案返回 null（回退 CSS 渲染） */
function presetFillColor(el) {
  const f = el.fill;
  if (f == null || f === 'none') return 'none';
  if (typeof f === 'string') return normalizeColor(f);
  if (f.type === 'solid' && f.color) return normalizeColor(f.color);
  if (f.type === 'gradient' && Array.isArray(f.stops) && f.stops.length) return normalizeColor(f.stops[0].color);
  return null; // image / pattern → 回退 CSS 渲染管线
}

/** 预设几何 SVG（与预览端同源公式），不支持或需 CSS 填充时返回 null */
function presetShapeSvg(el) {
  if (!el.shapeType || el.shapeType === 'rect') return null;
  const fill = presetFillColor(el);
  if (fill === null) return null;
  const geo = presetShapePath(el.shapeType, el.width || 100, el.height || 100, el.adjust || {});
  if (!geo) return null;
  const lineColor = el.line && el.line !== 'none' ? normalizeColor(el.line.color) : null;
  const lineW = el.line && el.line !== 'none' ? Math.max(0.5, ptToPx(el.line.width || 0.75)) : 0;
  const NS = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(NS, 'svg');
  const W = el.width || 100, H = el.height || 100;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  svg.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;overflow:visible';
  let tf = geo.transform || '';
  if (el.flipH || el.flipV) tf += ` scale(${el.flipH ? -1 : 1},${el.flipV ? -1 : 1})`;
  const main = document.createElementNS(NS, 'path');
  main.setAttribute('d', geo.d);
  if (tf.trim()) main.setAttribute('transform', tf.trim());
  main.setAttribute('fill', geo.noFill ? 'none' : (fill || 'none'));
  if (geo.fillRule) main.setAttribute('fill-rule', geo.fillRule);
  if (lineColor) {
    main.setAttribute('stroke', lineColor);
    main.setAttribute('stroke-width', String(lineW));
    if (el.line && el.line.dashType) main.setAttribute('stroke-dasharray', DASH_MAP[el.line.dashType] || 'none');
  }
  svg.appendChild(main);
  // 附加描边路径（callout 引线、笑脸嘴等）：颜色优先线条色，其次深化的填充色
  const accent = lineColor || shadeColor(fill, -0.25);
  (geo.strokes || []).forEach((s) => {
    const p = document.createElementNS(NS, 'path');
    p.setAttribute('d', s.d);
    p.setAttribute('fill', 'none');
    p.setAttribute('stroke', accent || '#5B9BD5');
    p.setAttribute('stroke-width', String(s.width || Math.max(1, W * 0.05)));
    p.setAttribute('stroke-linecap', 'round');
    if (tf.trim()) p.setAttribute('transform', tf.trim());
    svg.appendChild(p);
  });
  return svg;
}

/** 简单加深/减淡（amt -1..1） */
function shadeColor(hex, amt) {
  const c = normalizeColor(hex);
  if (!c) return null;
  const ch = (i) => clamp(Math.round(parseInt(c.slice(i, i + 2), 16) * (1 + amt)), 0, 255).toString(16).padStart(2, '0');
  return `#${ch(1)}${ch(3)}${ch(5)}`.toUpperCase();
}

/** 自定义几何（a:custGeom）SVG：路径坐标已归一化到形状 EMU 空间，viewBox 拉伸铺满元素框 */
function custGeomSvg(el) {
  const cg = el.custGeom;
  if (!cg || !Array.isArray(cg.paths) || !cg.paths.length) return null;
  const fill = presetFillColor(el);
  const NS = 'http://www.w3.org/2000/svg';
  let maxW = 0, maxH = 0;
  const pathD = cg.paths.map((p) => {
    const W = p.w || el.width || 100, H = p.h || el.height || 100;
    if (W > maxW) maxW = W;
    if (H > maxH) maxH = H;
    const cmd = (c) => {
      switch (c.type) {
        case 'moveTo': return `M${c.x} ${c.y}`;
        case 'lnTo': return `L${c.x} ${c.y}`;
        case 'cubicBezTo': return `C${c.x1} ${c.y1} ${c.x2} ${c.y2} ${c.x} ${c.y}`;
        case 'quadBezTo': return `Q${c.x1} ${c.y1} ${c.x} ${c.y}`;
        case 'close': return 'Z';
        default: return '';
      }
    };
    const d = (p.commands.map(cmd).join(' ') + (p.closed ? ' Z' : '')).trim();
    return d;
  }).filter(Boolean);
  const d = pathD.join(' ');
  if (!d) return null;
  const svg = document.createElementNS(NS, 'svg');
  svg.setAttribute('viewBox', `0 0 ${maxW || el.width || 100} ${maxH || el.height || 100}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  svg.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;overflow:visible';
  let tf = '';
  if (el.flipH || el.flipV) tf += ` scale(${el.flipH ? -1 : 1},${el.flipV ? -1 : 1})`;
  const main = document.createElementNS(NS, 'path');
  main.setAttribute('d', d);
  if (tf.trim()) main.setAttribute('transform', tf.trim());
  // 任一 path 闭合才填充（开放曲线如涂鸦仅描边）
  const anyClosed = cg.paths.some((p) => p.closed);
  main.setAttribute('fill', anyClosed && fill ? (fill || 'none') : 'none');
  const lineColor = el.line && el.line !== 'none' ? normalizeColor(el.line.color) : null;
  const lineW = el.line && el.line !== 'none' ? Math.max(0.5, ptToPx(el.line.width || 0.75)) : 0;
  if (lineColor) {
    main.setAttribute('stroke', lineColor);
    main.setAttribute('stroke-width', String(lineW));
    if (el.line && el.line.dashType) main.setAttribute('stroke-dasharray', DASH_MAP[el.line.dashType] || 'none');
  }
  svg.appendChild(main);
  return svg;
}

export function renderElement(el, ctx = {}) {
  const node = h('div', {
    class: `el el-${el.type}${el.locked ? ' locked' : ''}${el.hidden ? ' hidden-el' : ''}`,
    dataset: { id: el.id, type: el.type }
  });
  Object.assign(node.style, boxStyle(el));
  // 幻灯片内容为英文/数字时，将元素 lang 设为 en，使缺失的拉丁字体回退到衬线体，
  // 与预览端（lang=en）一致；否则继承外层 zh-CN 会导致数字/英文被中文无衬线替代而显示异常。
  node.lang = 'en';

  switch (el.type) {
    case 'text': {
      // 带文字的形状底（椭圆/饼图/弧线等）或带填充的文本框：先画形状/背景再叠文字
      const needsShapeBg = el.shapeType && el.shapeType !== 'rect';
      const needsTextBoxBg = el.fill || (el.line && el.line !== 'none');
      if (needsShapeBg || needsTextBoxBg) {
        const presetSvg = needsShapeBg ? presetShapeSvg(el) : null;
        if (presetSvg) {
          node.appendChild(presetSvg);
        } else {
          const { style } = shapeVisual(el);
          const geo = needsShapeBg ? shapeGeometry(el) : {};
          const bg = h('div', { style: { position: 'absolute', inset: '0' } });
          Object.assign(bg.style, style, geo);
          node.appendChild(bg);
          if (el.fill && el.fill.type === 'image') applyShapeImageFill(bg, el);
        }
      }
      node.appendChild(renderTextBody(el, ctx));
      break;
    }
    case 'shape': {
      if (/^(curvedConnector|bentConnector|straightConnector)/.test(el.shapeType || '')) {
        // 连接符：按 OOXML 预设几何的典型路径绘制，避免退化成直线
        const lc = normalizeColor((el.line && el.line.color) || '#5B9BD5') || '#5B9BD5';
        const lw = Math.max(1, ptToPx((el.line && el.line.width) || 1));
        const W = el.width || 100, H = el.height || 100;
        const NS = 'http://www.w3.org/2000/svg';
        const svg = document.createElementNS(NS, 'svg');
        svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
        svg.setAttribute('preserveAspectRatio', 'none');
        svg.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;overflow:visible';
        const fx = el.flipH ? -1 : 1, fy = el.flipV ? -1 : 1;
        if (el.flipH || el.flipV) svg.style.transform = `scaleX(${fx}) scaleY(${fy})`;
        const path = document.createElementNS(NS, 'path');
        let d;
        const st = el.shapeType || '';
        if (st === 'straightConnector1') {
          d = `M 0,0 L ${W},${H}`;
        } else if (st.startsWith('bentConnector')) {
          d = `M 0,0 L ${W * 0.5},0 L ${W * 0.5},${H} L ${W},${H}`;
        } else {
          // curvedConnector2/3/4/5：与预览端同款二次贝塞尔 S 曲线（控制点由 adj1 决定）
          const r = Math.min(Math.max(((el.adjust && el.adjust.adj1) != null ? el.adjust.adj1 : 50000) / 100000, 0), 1);
          d = `M 0,0 Q ${W * r},0 ${W / 2},${H / 2} Q ${W * (1 - r)},${H} ${W},${H}`;
        }
        path.setAttribute('d', d);
        path.setAttribute('fill', 'none');
        path.setAttribute('stroke', lc);
        path.setAttribute('stroke-width', String(lw));
        path.setAttribute('stroke-linecap', 'round');
        svg.appendChild(path);
        node.appendChild(svg);
        break;
      }
      if (el.shapeType === 'line') {
        const ln = h('div', { style: { position: 'absolute', left: '0', top: '50%', width: '100%' } });
        const w = Math.max(1, ptToPx((el.line && el.line.width) || 2));
        ln.style.height = `${w}px`;
        ln.style.background = normalizeColor((el.line && el.line.color) || '#000000') || '#000';
        ln.style.transform = 'translateY(-50%)';
        node.appendChild(ln);
      } else {
        const { style, effects } = shapeVisual(el);
        const geo = shapeGeometry(el);
        // 水平/垂直翻转（a:xfrm/@flipH/@flipV）
        const fx = el.flipH ? 'scaleX(-1)' : '';
        const fy = el.flipV ? 'scaleY(-1)' : '';
        const flip = [fx, fy].filter(Boolean).join(' ');

        // 阴影 / 发光：在形状背后放一层 SVG 滤镜（与预览端 pptxToHtml 同算法，不被 clip-path 裁掉）
        if (effects && (effects.shadow || effects.glow)) {
          const fx = shapeEffectSvg(el, effects);
          if (fx) {
            if (flip) fx.style.transform = flip;
            node.appendChild(fx);
          }
        }

        // 自定义自由曲线/任意多边形（a:custGeom）：优先用 custGeom 路径渲染
        const custSvg = custGeomSvg(el);
        if (custSvg) {
          node.appendChild(custSvg);
          break;
        }

        // 预设几何 SVG（与预览端同源公式）；图片/图案填充或未实现 preset 时回退 CSS 渲染
        const presetSvg = presetShapeSvg(el);
        if (presetSvg) {
          node.appendChild(presetSvg);
          break;
        }

        const inner = h('div', { style: { position: 'absolute', inset: '0' } });
        Object.assign(inner.style, style, geo);
        if (flip) inner.style.transform = flip;
        // 图案 / 图片填充需要子层溢出被裁剪（圆角/异形）
        const fill = el.fill;
        if (fill && (fill.type === 'pattern' || fill.type === 'image')) {
          inner.style.overflow = 'hidden';
        }
        node.appendChild(inner);

        if (fill && fill.type === 'pattern') {
          const patSvg = createPatternFillLayer(el);
          if (patSvg) inner.appendChild(patSvg);
        } else if (fill && fill.type === 'image') {
          applyShapeImageFill(inner, el);
        }
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
      node.appendChild(renderChartEl(el, ctx));
      break;
    }
    case 'video': {
      // 与预览端一致：用 poster 作为视觉主体（视频在 headless/编辑态下真实视频多为黑屏，
      // 海报图才是设计稿想要呈现的内容），叠加一个小的播放图标。
      const box = h('div', { style: { width: '100%', height: '100%', position: 'relative', overflow: 'hidden' } });
      const img = document.createElement('img');
      const posterData = (el.poster && (el.poster.data || el.poster.src)) || el.data || el.src || '';
      img.src = mediaSrc(posterData, el.poster && el.poster.extension || el.extension, 'image/png');
      img.draggable = false;
      img.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;object-fit:contain;display:block';
      box.appendChild(img);
      const icon = h('div', { style: {
        position: 'absolute', left: '50%', top: '50%', transform: 'translate(-50%, -50%)',
        width: '40px', height: '40px', borderRadius: '50%', background: 'rgba(0,0,0,0.55)',
        display: 'flex', alignItems: 'center', justifyContent: 'center', pointerEvents: 'none'
      }});
      icon.innerHTML = '<svg width="20" height="20" viewBox="0 0 20 20"><polygon fill="#fff" points="6,4 16,10 6,16"/></svg>';
      box.appendChild(icon);
      node.appendChild(box);
      break;
    }
    case 'audio': {
      const box = h('div', { style: { width: '100%', height: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#F1F3F4', borderRadius: '8px' } });
      const a = document.createElement('audio');
      a.src = mediaSrc(el.data || el.src || '', el.extension);
      a.controls = true;
      box.appendChild(a);
      node.appendChild(box);
      break;
    }
    case 'diagram': {
      node.appendChild(renderDiagramEl(el, ctx));
      break;
    }
    case 'group': {
      // relative：子元素坐标已相对组合原点；其余（page/绝对坐标）需减去组合原点
      const rel = el.childrenCoordinates === 'relative';
      const gx = rel ? 0 : (el.x || 0);
      const gy = rel ? 0 : (el.y || 0);
      for (const child of el.children || []) {
        const cn = renderElement(child, ctx);
        cn.style.left = `${(child.x || 0) - gx}px`;
        cn.style.top = `${(child.y || 0) - gy}px`;
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

/* ======================= 图表（复用预览端 chart-renderer 的 option 构建，含 3D） ======================= */
// 自管理 ECharts 实例，按 chartId 缓存并随重新渲染释放，避免 chart-renderer 单例重复绑定 resize 监听器
const _echartsMap = new Map();

function buildChartInfo(el, ctx) {
  const theme = getTheme(ctx.theme);
  const type = el.chartType3D || el.chartType || 'barChart';
  const isScatter = type === 'scatterChart' || type === 'bubbleChart';
  const isStock = type === 'stockChart';
  const data = (el.series || []).map((s, i) => {
    let values;
    if (isScatter) {
      // 散点/气泡保留 {x, y, size} 结构
      values = (s.values || []).map((v) => ({ x: v.x, y: v.y, size: v.size }));
    } else if (isStock) {
      // 股票图保留 [open, close, low, high]
      values = (s.values || []).map((v) => Array.isArray(v) ? v.slice(0, 4) : [0, 0, 0, 0]);
    } else {
      values = (s.values || []).map((y) => ({ y: Number(y) || 0 }));
    }
    return {
      key: s.name || `Series ${i + 1}`,
      xlabels: el.categories || [],
      values,
      style: s.color ? { fillColor: s.color } : {}
    };
  });
  const style = {
    legend: el.legend ? { position: 'right' } : { position: 'none' },
    title: el.title || '',
    grouping: el.grouping,
    holeSize: el.holeSize,
    smooth: el.smooth,
    marker: el.marker,
    view3D: el.view3D || undefined
  };
  return { chartId: 'chart_' + el.id, type, data, style, title: el.title || '', theme };
}

/**
 * SmartArt 图示渲染：
 * 优先用缓存绘图形状（drawingN.xml 提取的树形布局，与预览端同源）；
 * 无 shapes 时回退为扁平文本列表。
 */
function renderDiagramEl(el) {
  const shapes = Array.isArray(el.shapes) ? el.shapes : [];
  const W = el.width || 400, H = el.height || 300;
  const wrap = h('div', { style: { width: '100%', height: '100%', position: 'relative', overflow: 'hidden' } });

  if (shapes.length) {
    // 旧数据可能残留未解析的 'scheme:<name>' 引用，渲染前兜底
    const safeColor = (c, fallback) => (c && !/^scheme:/i.test(c) ? c : fallback);
    // 连接线层（SVG，viewBox 随容器缩放）
    const NS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
    svg.setAttribute('preserveAspectRatio', 'none');
    svg.style.cssText = 'position:absolute;inset:0;width:100%;height:100%';
    shapes.forEach((s) => {
      if (!s.connector) return;
      const line = document.createElementNS(NS, 'line');
      line.setAttribute('x1', String(s.flipH ? s.x + s.width : s.x));
      line.setAttribute('y1', String(s.flipV ? s.y + s.height : s.y));
      line.setAttribute('x2', String(s.flipH ? s.x : s.x + s.width));
      line.setAttribute('y2', String(s.flipV ? s.y : s.y + s.height));
      line.setAttribute('stroke', safeColor(s.lineColor, '#B0B4BA'));
      line.setAttribute('stroke-width', String(Math.max(0.75, s.lineWidth || 1)));
      svg.appendChild(line);
    });
    wrap.appendChild(svg);

    // 节点框
    const radiusOf = (s) => {
      if (s.prst === 'ellipse') return '50%';
      if (!s.prst || s.prst === 'rect') return '2px';
      return Math.round(Math.min(s.width, s.height) * 0.18) + 'px';
    };
    shapes.forEach((s) => {
      if (s.connector) return;
      const cell = h('div', {
        style: {
          position: 'absolute',
          left: `${(s.x / W) * 100}%`, top: `${(s.y / H) * 100}%`,
          width: `${(s.width / W) * 100}%`, height: `${(s.height / H) * 100}%`,
          background: s.fill && s.fill !== 'none' ? safeColor(s.fill, 'transparent') : 'transparent',
          border: safeColor(s.lineColor, null) ? `${Math.max(0.75, s.lineWidth || 1)}px solid ${safeColor(s.lineColor, '')}` : 'none',
          borderRadius: radiusOf(s),
          display: 'flex',
          alignItems: s.anchor === 't' ? 'flex-start' : s.anchor === 'b' ? 'flex-end' : 'center',
          justifyContent: s.align === 'l' ? 'flex-start' : s.align === 'r' ? 'flex-end' : 'center',
          padding: '2px 8px', boxSizing: 'border-box', overflow: 'hidden'
        }
      });
      if (s.text) {
        const span = h('span', {
          style: {
            fontSize: `${s.fontSize || 12}pt`, color: safeColor(s.color, '#FFFFFF'),
            fontWeight: s.bold ? 600 : 400, lineHeight: 1.2,
            width: '100%', textAlign: s.align === 'l' ? 'left' : s.align === 'r' ? 'right' : 'center',
            whiteSpace: 'pre-wrap', wordBreak: 'break-word'
          }
        });
        span.textContent = s.text;
        cell.appendChild(span);
      }
      wrap.appendChild(cell);
    });
    return wrap;
  }

  // 兜底：无 shapes（旧数据/创作生成）——扁平文本列表
  const box = h('div', { style: { width: '100%', height: '100%', display: 'flex', flexDirection: 'column', gap: '8px', padding: '12px', overflow: 'hidden' } });
  const accents = getTheme('blue').accents;
  (el.texts || []).forEach((t, i) => {
    const item = h('div', {
      style: {
        padding: '8px 14px', borderRadius: '8px', fontSize: '14px', color: '#fff',
        background: accents[i % accents.length] || '#1A73E8',
        marginLeft: `${Math.min(i, 3) * 18}px`,
        whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis'
      }
    });
    item.textContent = t;
    box.appendChild(item);
  });
  wrap.appendChild(box);
  return wrap;
}

function renderChartEl(el, ctx) {
  const wrap = h('div', { class: 'el-chart', style: { width: '100%', height: '100%' } });
  wrap.id = 'chart_' + (ctx.chartScope || 'canvas') + '_' + el.id;
  const theme = getTheme(ctx.theme);

  // 优先用 ECharts（与预览端一致的 option 构建，支持 3D）；缺库时回退到内置 SVG
  if (typeof window !== 'undefined' && window.echarts && chartRenderer && chartRenderer.prepareEChartsOption) {
    try {
      const info = buildChartInfo(el, ctx);
      const option = chartRenderer.prepareEChartsOption(info);
      if (option) {
        if (!el.legend && option.legend) delete option.legend;
        enqueueEchart(wrap, option);
        return wrap;
      }
    } catch (err) {
      console.error('ECharts 图表 option 构建失败，回退 SVG：', err);
    }
  }
  wrap.innerHTML = renderChartSVG(el, {
    width: el.width, height: el.height,
    palette: theme.accents,
    textColor: theme.text, titleColor: theme.title
  });
  return wrap;
}

function enqueueEchart(wrap, option) {
  // 节点此时尚未挂载到文档，延迟到下一帧（已挂载且有尺寸）再 init
  requestAnimationFrame(() => {
    const w = wrap.clientWidth, h = wrap.clientHeight;
    const hasSize = w > 0 && h > 0;
    const inst = window.echarts.init(
      wrap,
      null,
      hasSize ? undefined : { width: Math.max(w, 320), height: Math.max(h, 220) }
    );
    inst.setOption(option);
    const prev = _echartsMap.get(wrap.id);
    if (prev) { window.removeEventListener('resize', prev.onResize); prev.chart.dispose(); }
    const onResize = () => inst.resize();
    window.addEventListener('resize', onResize);
    _echartsMap.set(wrap.id, { chart: inst, onResize });
    if (!hasSize) {
      requestAnimationFrame(() => { if (wrap.clientWidth > 0 && wrap.clientHeight > 0) inst.resize(); });
    }
  });
}

/** 释放所有 ECharts 实例（切换幻灯片 / 卸载画布前调用，避免内存泄漏） */
export function disposeAllCharts() {
  for (const { chart, onResize } of _echartsMap.values()) {
    window.removeEventListener('resize', onResize);
    try { chart.dispose(); } catch { /* ignore */ }
  }
  _echartsMap.clear();
}

function shapeGeometry(el) {
  const r = Math.min(el.width, el.height) * 0.16;
  const adj = el.adjust || {};
  switch (el.shapeType) {
    case 'ellipse': case 'moon': case 'sun': case 'donut': case 'arc': case 'smileyFace':
      return { borderRadius: '50%' };
    case 'roundRect': {
      // adj = 圆角半径占短边比例（OOXML 原生千分比，默认 16667）
      const rad = ((adj.adj != null ? adj.adj : 16667) / 100000) * Math.min(el.width, el.height);
      return { borderRadius: `${Math.min(rad, 60)}px` };
    }
    case 'gear6': return { borderRadius: '18%' };
    case 'can': return { borderRadius: '10% / 18%' };
    case 'cloud': return { borderRadius: '44% 44% 38% 38% / 50% 50% 50% 50%' };
    case 'noSmoking': {
      // 禁止标识：圆环 + 45° 斜杠（填充色仅作描边兜底）
      const c = normalizeColor((el.line && el.line.color) || (typeof el.fill === 'string' ? el.fill : '')) || '#5B9BD5';
      return {
        borderRadius: '50%', background: 'transparent',
        border: `3px solid ${c}`,
        backgroundImage: `linear-gradient(45deg, transparent 42%, ${c} 42%, ${c} 58%, transparent 58%)`
      };
    }
    case 'uturnArrow':
      return { clipPath: 'polygon(0% 100%,0% 40%,25% 12%,55% 12%,55% 0%,100% 22%,55% 45%,55% 32%,40% 32%,22% 52%,22% 100%)' };
    case 'quadArrow':
      return { clipPath: 'polygon(50% 0%,64% 20%,56% 20%,56% 44%,80% 44%,80% 33%,100% 50%,80% 67%,80% 56%,56% 56%,56% 80%,64% 80%,50% 100%,36% 80%,44% 80%,44% 56%,20% 56%,20% 67%,0% 50%,20% 33%,20% 44%,44% 44%,44% 20%,36% 20%)' };
    case 'plaque': return { borderRadius: '10% / 16%' };
    case 'flowChartMagneticDisk': return { borderRadius: '50% 50% 8% 8% / 20% 20% 5% 5%' };
    case 'flowChartMultidocument': return { borderRadius: '12% 12% 0 0' };
    case 'wedgeRectCallout':
      return { clipPath: 'polygon(0% 0%,100% 0%,100% 100%,26% 100%,8% 118%,16% 100%,0% 100%)' };
    case 'ellipseRibbon':
      return { clipPath: 'polygon(0% 35%,15% 20%,50% 35%,85% 20%,100% 35%,100% 65%,85% 80%,50% 65%,15% 80%,0% 65%)' };
    case 'triangle': return { clipPath: 'polygon(50% 0%,100% 100%,0% 100%)' };
    case 'rtTriangle': return { clipPath: 'polygon(0% 0%,100% 100%,0% 100%)' };
    case 'diamond': return { clipPath: 'polygon(50% 0%,100% 50%,50% 100%,0% 50%)' };
    case 'parallelogram': return { clipPath: 'polygon(20% 0%,100% 0%,80% 100%,0% 100%)' };
    case 'trapezoid': return { clipPath: 'polygon(20% 0%,80% 0%,100% 100%,0% 100%)' };
    case 'pentagon': return { clipPath: 'polygon(50% 0%,100% 38%,82% 100%,18% 100%,0% 38%)' };
    case 'hexagon': return { clipPath: 'polygon(25% 0%,75% 0%,100% 50%,75% 100%,25% 100%,0% 50%)' };
    case 'octagon': return { clipPath: 'polygon(30% 0%,70% 0%,100% 30%,100% 70%,70% 100%,30% 100%,0% 70%,0% 30%)' };
    case 'chevron': return { clipPath: 'polygon(0% 0%,62% 0%,100% 50%,62% 100%,0% 100%,38% 50%)' };
    // 箭头：与预览端 /src/shape/arrow-shapes.ts 保持一致。
    // adj1 = 箭身厚度控制（千分比，默认 50000），adj2 = 箭头长度控制（千分比，默认 50000）。
    // 渲染公式：半箭身厚度 = H*adj1/200000，箭头长度 = min(W,H)*adj2/100000。
    case 'rightArrow': {
      const a1 = Math.min(Math.max(adj.adj1 != null ? adj.adj1 : 50000, 0), 100000);
      const maxA2 = Math.min(adj.adj2 != null ? adj.adj2 : 50000, 100000 * el.width / Math.min(el.width, el.height));
      const a2 = Math.max(0, maxA2);
      const y1 = 50 - a1 / 2000;
      const y2 = 50 + a1 / 2000;
      const x1 = 100 - (Math.min(el.width, el.height) / el.width) * (a2 / 100000) * 100;
      return { clipPath: `polygon(0% ${y1}%,${x1}% ${y1}%,${x1}% 0%,100% 50%,${x1}% 100%,${x1}% ${y2}%,0% ${y2}%)` };
    }
    case 'leftArrow': {
      const a1 = Math.min(Math.max(adj.adj1 != null ? adj.adj1 : 50000, 0), 100000);
      const maxA2 = Math.min(adj.adj2 != null ? adj.adj2 : 50000, 100000 * el.width / Math.min(el.width, el.height));
      const a2 = Math.max(0, maxA2);
      const y1 = 50 - a1 / 2000;
      const y2 = 50 + a1 / 2000;
      const x1 = (Math.min(el.width, el.height) / el.width) * (a2 / 100000) * 100;
      return { clipPath: `polygon(100% ${y1}%,${x1}% ${y1}%,${x1}% 0%,0% 50%,${x1}% 100%,${x1}% ${y2}%,100% ${y2}%)` };
    }
    case 'upArrow': {
      const a1 = Math.min(Math.max(adj.adj1 != null ? adj.adj1 : 50000, 0), 100000);
      const maxA2 = Math.min(adj.adj2 != null ? adj.adj2 : 50000, 100000 * el.height / Math.min(el.width, el.height));
      const a2 = Math.max(0, maxA2);
      const x1 = 50 - a1 / 2000;
      const x2 = 50 + a1 / 2000;
      const y1 = (Math.min(el.width, el.height) / el.height) * (a2 / 100000) * 100;
      return { clipPath: `polygon(${x1}% 100%,${x1}% ${y1}%,0% ${y1}%,50% 0%,100% ${y1}%,${x2}% ${y1}%,${x2}% 100%)` };
    }
    case 'downArrow': {
      const a1 = Math.min(Math.max(adj.adj1 != null ? adj.adj1 : 50000, 0), 100000);
      const maxA2 = Math.min(adj.adj2 != null ? adj.adj2 : 50000, 100000 * el.height / Math.min(el.width, el.height));
      const a2 = Math.max(0, maxA2);
      const x1 = 50 - a1 / 2000;
      const x2 = 50 + a1 / 2000;
      const y1 = 100 - (Math.min(el.width, el.height) / el.height) * (a2 / 100000) * 100;
      return { clipPath: `polygon(${x1}% 0%,${x1}% ${y1}%,0% ${y1}%,50% 100%,100% ${y1}%,${x2}% ${y1}%,${x2}% 0%)` };
    }
    case 'pentagonBlock': return { clipPath: 'polygon(50% 0%,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%)' };
    case 'plus': return { clipPath: 'polygon(35% 0%,65% 0%,65% 35%,100% 35%,100% 65%,65% 65%,65% 100%,35% 100%,35% 65%,0% 65%,0% 35%,35% 35%)' };
    case 'heart': return { clipPath: 'polygon(50% 100%,2% 55%,2% 28%,25% 6%,50% 18%,75% 6%,98% 28%,98% 55%)' };
    case 'lightningBolt': return { clipPath: 'polygon(56% 0%,18% 56%,46% 56%,34% 100%,82% 38%,54% 38%)' };
    case 'pie': {
      // OOXML pie：adj1/adj2 为起止角（1/60000 度，0°=3 点钟顺时针），默认 0°~270°（16200000）
      const a1 = (adj.adj1 != null ? Number(adj.adj1) : 0) / 60000;
      const a2 = (adj.adj2 != null ? Number(adj.adj2) : 16200000) / 60000;
      const pts = ['50% 50%'];
      const span = ((a2 - a1) % 360 + 360) % 360 || 360;
      const steps = Math.max(2, Math.ceil(span / 10));
      for (let i = 0; i <= steps; i++) {
        const a = (a1 + span * i / steps) * Math.PI / 180;
        pts.push(`${(50 + 50 * Math.cos(a)).toFixed(2)}% ${(50 + 50 * Math.sin(a)).toFixed(2)}%`);
      }
      return { clipPath: `polygon(${pts.join(',')})` };
    }
    case 'arc': case 'blockArc': {
      // 弧线：椭圆描边环近似（transparent 填充 + line 色边框）
      const c = normalizeColor((el.line && el.line.color) || '') || '#5B9BD5';
      const w = Math.max(1.5, ptToPx((el.line && el.line.width) || 1.5));
      return { borderRadius: '50%', background: 'transparent', border: `${w}px solid ${c}` };
    }
    case 'cube': return { clipPath: 'polygon(50% 0%,100% 25%,100% 75%,50% 100%,0% 75%,0% 25%)' };
    case 'funnel': return { clipPath: 'polygon(0% 0%,100% 0%,65% 56%,65% 100%,35% 100%,35% 56%)' };
    case 'foldedCorner': {
      // OOXML foldedCorner：矩形右下角向内折，折角边长 = 短边 × adj/100000（默认 16667）
      const a = Math.min(Math.max(adj.adj != null ? adj.adj : 16667, 0), 100000);
      const f = Math.min(el.width || 1, el.height || 1) * a / 100000;
      const fx = (f / (el.width || 1)) * 100;
      const fy = (f / (el.height || 1)) * 100;
      return { clipPath: `polygon(0% 0%,100% 0%,100% ${100 - fy}%,${100 - fx}% 100%,0% 100%)` };
    }
    case 'frame': return { border: '8px solid currentColor', color: '#9AA0A6' };
    default: return {};
  }
}

let _fxUid = 0;
/**
 * 阴影/发光用内联 SVG 滤镜层渲染（与 pptxToHtml 同一套 feGaussianBlur + feFlood + feComposite 算法）。
 * 关键：CSS filter 会被同元素的 clip-path 裁掉（CSS 渲染顺序 filter → clip-path → mask），
 * 而把图形画成 SVG（无 clip-path）再挂滤镜就不会被裁，因此对菱形（clip-path）和椭圆（border-radius）都生效。
 * 该层放在形状本体背后，仅露出外延的光晕/阴影。
 */
function shapeEffectSvg(el, effects) {
  if (!effects || (!effects.shadow && !effects.glow)) return null;
  const W = el.width || 100, H = el.height || 100;
  const geo = shapeGeometry(el);
  let geom;
  if (geo.clipPath) {
    const m = geo.clipPath.match(/polygon\(([^)]+)\)/);
    if (m) {
      const pts = m[1].trim().split(/\s*,\s*/).map((p) => p.trim().split(/\s+/).map((v) => parseFloat(v)));
      geom = `<polygon points="${pts.map(([x, y]) => `${(x / 100 * W).toFixed(2)},${(y / 100 * H).toFixed(2)}`).join(' ')}"/>`;
    } else {
      geom = `<rect x="0" y="0" width="${W}" height="${H}"/>`;
    }
  } else if (geo.borderRadius) {
    const br = geo.borderRadius;
    if (br === '50%') {
      geom = `<ellipse cx="${(W / 2).toFixed(2)}" cy="${(H / 2).toFixed(2)}" rx="${(W / 2).toFixed(2)}" ry="${(H / 2).toFixed(2)}"/>`;
    } else {
      const num = parseFloat(br);
      const rx = br.endsWith('%') ? (num / 100 * W) : num;
      const ry = br.endsWith('%') ? (num / 100 * H) : num;
      geom = `<rect x="0" y="0" width="${W}" height="${H}" rx="${rx.toFixed(2)}" ry="${ry.toFixed(2)}"/>`;
    }
  } else {
    geom = `<rect x="0" y="0" width="${W}" height="${H}"/>`;
  }

  const SHADOW_SIGMA_RATIO = 0.29, GLOW_DILATE_RATIO = 0.38, GLOW_SIGMA_RATIO = 0.17;
  const defs = [];
  const layers = [];
  const uid = 'efx' + (_fxUid++);

  if (effects.glow) {
    const g = effects.glow;
    const rad = g.blur;
    const dilate = Math.max(0.5, rad * GLOW_DILATE_RATIO);
    const sigma = Math.max(0.5, rad * GLOW_SIGMA_RATIO);
    const margin = Math.ceil(dilate + sigma * 3) + 2;
    const fid = uid + '_g';
    defs.push(`<filter id="${fid}" filterUnits="userSpaceOnUse" x="${-margin}" y="${-margin}" width="${W + margin * 2}" height="${H + margin * 2}" color-interpolation-filters="sRGB"><feMorphology in="SourceAlpha" operator="dilate" radius="${dilate.toFixed(2)}" result="d"/><feGaussianBlur in="d" stdDeviation="${sigma.toFixed(2)}" result="b"/><feFlood flood-color="${g.color}" result="c"/><feComposite in="c" in2="b" operator="in"/></filter>`);
    layers.push(`<g filter="url(#${fid})">${geom}</g>`);
  }
  if (effects.shadow) {
    const s = effects.shadow;
    const sigma = Math.max(0.5, s.blur * SHADOW_SIGMA_RATIO);
    const margin = Math.ceil(Math.max(Math.abs(s.dx), Math.abs(s.dy)) + sigma * 3) + 2;
    const fid = uid + '_s';
    defs.push(`<filter id="${fid}" filterUnits="userSpaceOnUse" x="${-margin}" y="${-margin}" width="${W + margin * 2}" height="${H + margin * 2}" color-interpolation-filters="sRGB"><feGaussianBlur in="SourceAlpha" stdDeviation="${sigma.toFixed(2)}" result="b"/><feOffset in="b" dx="${s.dx}" dy="${s.dy}" result="o"/><feFlood flood-color="${s.color}" result="c"/><feComposite in="c" in2="o" operator="in"/></filter>`);
    layers.push(`<g filter="url(#${fid})">${geom}</g>`);
  }

  const wrap = h('div', { style: { position: 'absolute', inset: '0', pointerEvents: 'none' } });
  wrap.innerHTML = `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="position:absolute;inset:0;overflow:visible"><defs>${defs.join('')}</defs>${layers.join('')}</svg>`;
  return wrap;
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
    frame.appendChild(renderElement(el, { theme: doc.theme, editing: opts.editingId === el.id, chartScope: 'canvas' }));
  }
  // 批注卡片（仅画布层展示；pos 为 px，缺省锚右上区域）
  (slide.comments || []).forEach((c, i) => {
    const card = h('div', { class: 'slide-comment' });
    card.style.left = `${(c.pos && c.pos.x != null) ? c.pos.x : Math.max(8, doc.slideSize.width - 228)}px`;
    card.style.top = `${(c.pos && c.pos.y != null) ? c.pos.y : 8 + i * 76}px`;
    const author = h('div', { class: 'sc-author', text: c.author || 'Author' });
    const body = h('div', { class: 'sc-text' });
    body.textContent = c.text || '';
    card.append(author, body);
    if (c.dt) {
      const d = h('div', { class: 'sc-date', text: String(c.dt).slice(0, 10) });
      card.appendChild(d);
    }
    frame.appendChild(card);
  });
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
    inner.appendChild(renderElement(el, { theme: doc.theme, chartScope: 'thumb' }));
  }
  host.appendChild(inner);
  return scale;
}
