/**
 * 幻灯片渲染：内部模型 → DOM（画布 / 缩略图 / 演示共用）
 */
import { h, ptToPx, normalizeColor, withAlpha, hexToRgb, clamp } from './util.js';
import { presetShapePath, elementRect, rotatedRect, effectMargin, absoluteElementRect } from '../../../dist/ppt-parser.browser.js';
// 元素级几何已下沉到库（src/editor/geometry.ts），此处 re-export 以兼容既有 import
export { elementRect, rotatedRect, effectMargin, absoluteElementRect };
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

/* ======================= 背景 ======================= */
export function backgroundStyle(bg, theme) {
  if (!bg || bg === 'none') return { background: '#FFFFFF' };
  if (typeof bg === 'string') return { background: normalizeColor(bg) || '#FFFFFF' };
  if (bg.type === 'image') {
    // 逐属性设置而非 `background` 简写：背景图常是数百 KB 的 data URL，
    // 简写里 position/size 的 `/` 分隔语法一旦被解析失败会整条声明失效（连底色一起丢），
    // 表现为「整页背景消失」；分开设置各条互不牵连。
    const url = bg.data || bg.src || '';
    return {
      backgroundColor: normalizeColor(theme?.bg) || '#FFFFFF',
      // OOXML 背景图 blipFill（a:stretch）：整图拉伸铺满画布（非 cover 裁剪）
      backgroundImage: url ? `url("${url}")` : 'none',
      backgroundPosition: 'center',
      backgroundSize: '100% 100%',
      backgroundRepeat: 'no-repeat'
    };
  }
  if (bg.type === 'gradient') {
    const stops = (bg.stops || []).slice().sort((a, b) => a.position - b.position);
    const list = stops.length >= 2
      ? stops.map((s) => `${withAlpha(s.color, s.transparency || 0)} ${clamp(s.position, 0, 1) * 100}%`).join(',')
      : '#FFFFFF,#FFFFFF';
    if (bg.gradientType === 'radial') return { background: `radial-gradient(circle at 50% 50%, ${list})` };
    // OOXML a:lin@ang：0°=向右，顺时针；CSS linear-gradient：0°=向上，顺时针。
    // 换算：CSS deg = (OOXML deg + 90) % 360；fallback 到旧 direction 枚举。
    let deg;
    if (typeof bg.angle === 'number') {
      deg = Math.round(((bg.angle + 90) % 360) * 10) / 10;
    } else {
      deg = bg.direction === 'vertical' ? 180 : bg.direction === 'diagonal' ? 135 : 90;
    }
    return { background: `linear-gradient(${deg}deg, ${list})` };
  }
  return { background: normalizeColor(bg.color) || '#FFFFFF' };
}

/* ======================= 元素渲染 ======================= */
const DASH_MAP = { solid: 'solid', dash: 'dashed', dashDot: 'dashed', dotted: 'dotted', lgDash: 'dashed', sysDot: 'dotted' };
/** stroke-dasharray 用的是虚线长度列表，不能复用 DASH_MAP 的 CSS border-style 关键字；取值与预览端 getBorder 对齐 */
const SVG_DASH_MAP = {
  solid: null, dash: '5', dashDot: '5, 5, 1, 5', dot: '1, 5', dbl: null,
  lgDash: '10, 5', lgDashDot: '10, 5, 1, 5', lgDashDotDot: '10, 5, 1, 5, 1, 5',
  sysDash: '5, 2', sysDot: '2, 5', sysDashDot: '5, 2, 1, 5', sysDashDotDot: '5, 2, 1, 5, 1, 5'
};

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
      const list = stops.map((s) => `${withAlpha(s.color, s.transparency || 0)} ${clamp(s.position, 0, 1) * 100}%`).join(',');
      // 径向渐变（a:path）：CSS 用 radial-gradient 近似（OOXML 的 path="rect"/"shape" 统一按椭圆近似）
      if (fill.gradientType === 'radial') {
        style.background = `radial-gradient(circle at 50% 50%, ${list})`;
      } else {
        // OOXML a:lin@ang → CSS deg：0°向右 → 90°向上，顺时针递增
        let deg;
        if (typeof fill.angle === 'number') {
          deg = Math.round(((fill.angle + 90) % 360) * 10) / 10;
        } else {
          deg = fill.direction === 'vertical' ? 180 : fill.direction === 'diagonal' ? 135 : 90;
        }
        style.background = `linear-gradient(${deg}deg, ${list})`;
      }
    } else if (fill.type === 'image') {
      // 形状图片填充：铺满 / 平铺（tile）/ 平铺 + 源图裁剪（srcRect）
      const url = fill.data || fill.src || '';
      if (!url) {
        style.background = 'transparent';
      } else if (hasTileOrCrop(fill)) {
        // 平铺 / 裁剪需要按原图像素换算每格尺寸（a:tile 的 sx/sy 是「每格占原图比例」，
        // 与裁剪窗口无关），交给 applyShapeImageFill 用 canvas 按原图精确绘制。
        // 此处不预设 background：否则会先画出格子密度错误的图案再被覆盖（可见闪烁）。
        style.background = 'transparent';
      } else {
        // OOXML a:stretch/fillRect：整图拉伸铺满形状（非 cover 裁剪）
        style.background = `url("${url}") center / 100% 100% no-repeat`;
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
      color: `rgba(${r},${g},${b},${alpha.toFixed(2)})`,
      // 把透明度透传给 shapeEffectSvg：外层滤镜的 flood-opacity 以 transparency 为准，
      // 否则一律按 1.0 渲染（与预览端 pptxToHtml 依 PPTX alpha 渲染不一致）。
      transparency: s.transparency != null ? s.transparency : undefined
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

/** 图片填充是否需要 canvas 处理（平铺或源图裁剪）；拉伸铺满可纯 CSS 完成。 */
function hasTileOrCrop(fill) {
  if (!fill || fill.type !== 'image') return false;
  const EPS = 1e-4;
  const sr = fill.srcRect || {};
  if (Math.abs(sr.l || 0) > EPS || Math.abs(sr.t || 0) > EPS || Math.abs(sr.r || 0) > EPS || Math.abs(sr.b || 0) > EPS) {
    return true;
  }
  const tile = fill.tile || {};
  return tile.sx != null || tile.sy != null;
}

/** 形状图片填充（平铺 / 裁剪）异步应用。先以拉伸图占位，图片加载后按 srcRect 裁切并平铺。 */
function applyShapeImageFill(inner, el) {
  const fill = el.fill;
  if (!fill || fill.type !== 'image') return;
  const url = fill.data || fill.src || '';
  if (!url) return;

  // 无平铺无裁剪：CSS 背景拉伸铺满即可
  const tile = fill.tile || {};
  if (!hasTileOrCrop(fill)) {
    // OOXML a:stretch/fillRect：整图拉伸铺满形状（非 cover 裁剪）
    inner.style.background = `url("${url}") center / 100% 100% no-repeat`;
    return;
  }

  const img = new Image();
  img.crossOrigin = 'anonymous';
  img.onload = () => {
    const iw = img.naturalWidth || 1, ih = img.naturalHeight || 1;
    const sr = fill.srcRect || {};
    const l = Math.max(0, Math.min(1, sr.l || 0));
    const t = Math.max(0, Math.min(1, sr.t || 0));
    const r = Math.max(0, Math.min(1 - l, sr.r || 0));
    const b = Math.max(0, Math.min(1 - t, sr.b || 0));
    const cropX = l * iw, cropY = t * ih;
    const cropW = Math.max(1, (1 - l - r) * iw), cropH = Math.max(1, (1 - t - b) * ih);

    // 每格像素尺寸：OOXML a:tile 的 sx/sy 是「每格占原图尺寸的比例」，与 srcRect 裁剪无关
    // （解析端 json-from-pptx 同样按「每格占原图比例」写入 model.tile）。
    // 预览端 SVG <pattern> 用 sx * 原图宽 作 tile 宽，这里保持同一换算，
    // 否则裁剪过的形状在编辑器里每格会明显小于预览端（表现为格子过密）。
    const sx = Math.max(0.01, tile.sx ?? 1);
    const sy = Math.max(0.01, tile.sy ?? 1);
    const tileW = Math.max(1, Math.round(iw * sx));
    const tileH = Math.max(1, Math.round(ih * sy));

    const canvas = document.createElement('canvas');
    canvas.width = tileW;
    canvas.height = tileH;
    const ctx2d = canvas.getContext('2d');
    if (!ctx2d) return;
    // 源为裁剪窗口（图片像素坐标），目标为整格：与预览端 pattern 的 viewBox + width 同构
    ctx2d.drawImage(img, cropX, cropY, cropW, cropH, 0, 0, tileW, tileH);
    const dataUrl = canvas.toDataURL('image/png');

    inner.style.backgroundImage = `url("${dataUrl}")`;
    inner.style.backgroundRepeat = 'repeat';
    // a:tile 的 tx/ty 是相对原图尺寸的偏移比例
    inner.style.backgroundPosition = `${((tile.tx ?? 0) * iw).toFixed(1)}px ${((tile.ty ?? 0) * ih).toFixed(1)}px`;
    inner.style.backgroundSize = `${tileW}px ${tileH}px`;
    inner.style.backgroundColor = 'transparent';
  };
  // 加载失败则保持占位：拉伸铺满原图
  img.onerror = () => { inner.style.background = `url("${url}") center / 100% 100% no-repeat`; };
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
const HEBREW_LETTERS = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט', 'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ', 'ק', 'ר', 'ש', 'ת'];
function hebrewNumber(n) {
  // 简单 22 字母表循环（与 PowerPoint 的 hebrew 编号近似）
  let out = '', rest = Math.max(0, n - 1);
  do { out = HEBREW_LETTERS[rest % HEBREW_LETTERS.length] + out; rest = Math.floor(rest / HEBREW_LETTERS.length) - 1; } while (rest >= 0);
  return out || HEBREW_LETTERS[0];
}
function formatAutoNum(fmt, n) {
  switch (fmt) {
    case 'hebrew1Minus': return `${hebrewNumber(n)}-`;
    case 'hebrew2Minus': return `${hebrewNumber(n)}-`;
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

/** 圆形/扇形类几何：PowerPoint 与预览端会把文本强制水平和垂直居中。 */
function isCircularShape(el) {
  return ['ellipse', 'ovalCallout', 'wedgeEllipseCallout', 'pie', 'pieWedge', 'chord', 'sector', 'arc', 'blockArc'].includes(el.shapeType || '');
}

function renderTextBody(el, ctx) {
  const body = h('div', { class: 'tb-body' });
  // 自动适配（a:normAutofit@fontScale）：PowerPoint 按此比例整体缩放字号以贴合文本框，
  // 不应用会导致大字号标题溢出折行（如 72pt 标题），与预览端不一致。
  const afs = (el.fontScale && el.fontScale > 0) ? el.fontScale : 1;
  const effFont = (r) => ((((r && r.fontSize != null) ? r.fontSize : (el.fontSize != null ? el.fontSize : 18))) * afs);
  const circular = isCircularShape(el);
  const vmap = { top: 'flex-start', middle: 'center', bottom: 'flex-end' };
  // 圆形/椭圆类形状：OOXML 常写 anchor="t"，但视觉上应居中，与预览端 getVerticalAlign 一致。
  let effectiveVAlign = el.valign;
  if (circular && effectiveVAlign === 'top') effectiveVAlign = 'middle';
  body.style.justifyContent = vmap[effectiveVAlign] || 'flex-start';
  body.style.whiteSpace = el.noWrap ? 'nowrap' : 'pre-wrap';
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
  let numState = null;  (el.paragraphs || []).forEach((p) => {
    const para = h('div', {});
    // 从右到左段落（a:pPr@rtl）：dir=rtl 且缺省右对齐（OOXML 语义）
    // 圆形/椭圆类形状：预览端强制水平居中，编辑器同步处理。
    if (circular) {
      para.style.textAlign = 'center';
    } else if (p.rtl) {
      para.dir = 'rtl';
      para.style.textAlign = p.align || el.align || 'right';
    } else {
      para.style.textAlign = p.align || el.align || 'left';
    }
    // 行距：区分「显式值」与「继承值」。编辑器为方便面板编辑会给每个段落填默认
    // lineSpacing(=1.15)，但那是面板默认值而非源文件值；直接用会让继承排版的行距偏离预览端
    //（预览端对未显式 a:lnSpc 的段落不设 line-height，走 CSS normal）。
    const pLsSet = p.lineSpacingSet !== undefined ? p.lineSpacingSet : true;
    const eLsSet = el.propsSet && el.propsSet.lineSpacing !== undefined ? el.propsSet.lineSpacing : !!el.lineSpacing;
    const ls = pLsSet && p.lineSpacing != null ? p.lineSpacing : (eLsSet ? el.lineSpacing : null);
    if (ls != null) {
      // 编辑器内部用两种约定混存：>=10 表示 OOXML 百分值（100=100%），<5 表示倍数（1.5）
      const ratio = typeof ls === 'number' && ls >= 10 ? ls / 100 : ls;
      para.style.lineHeight = String(ratio);
    }
    const bulletRaw = p.bullet ?? el.bullet;
    // 空段落（无文本、无软换行）不画项目符号/编号：PowerPoint 放映时空段落只占一行空白，
    // 预览端同样不显示；画出来会出现「孤立的 ● / ➢」挂在文本框末尾（如 slide9 的空尾段）
    const paraHasContent = (p.runs || []).some((r) => r.text || r.break);
    const bullet = paraHasContent ? bulletRaw : null;
    const firstRun = (p.runs || [])[0] || {};
    const bulletFontSize = ptToPx(effFont(firstRun));
    const bulletColor = normalizeColor((bullet && bullet.color) || firstRun.color) || normalizeColor(el.color) || '#202124';
    const bulletFont = fontFamilyOf(
      firstRun.fontFace, firstRun.fontFaceEa, el.fontFace, el.fontFaceEa, '微软雅黑'
    ) || quoteFont('微软雅黑');
    if (isNumberBullet(bullet)) {
      if (!numState) numState = { n: (typeof bullet === 'object' && bullet.start) || 1 };
      const fmt = (typeof bullet === 'object' && bullet.fmt) || 'arabicPeriod';
      para.classList.add('num-para');
      const b = document.createElement('span');
      b.className = 'bullet-mark';
      // 间距用 CSS margin 而非 \u00A0：nbsp 放在符号字体（Wingdings 等）的 span 里，
      // 符号字体缺失走 fallback 时 nbsp 会渲染成「·」点，紧跟在项目符号后（预览端用普通空格无此问题）
      b.textContent = formatAutoNum(fmt, numState.n);
      b.style.marginRight = '0.3em';
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
        // 图片项目符号（a:buBlp）：用 <img> 渲染，尺寸跟随符号字号
        if (typeof bullet === 'object' && bullet.type === 'picture' && bullet.data) {
          const b = document.createElement('span');
          b.className = 'bullet-mark bullet-img';
          const img = document.createElement('img');
          img.src = bullet.data;
          img.alt = '';
          img.draggable = false;
          const sz = bullet.sizePct ? bulletFontSize * (bullet.sizePct / 100) : bulletFontSize;
          img.style.width = `${sz.toFixed(2)}px`;
          img.style.height = `${sz.toFixed(2)}px`;
          img.style.verticalAlign = '-0.15em';
          b.appendChild(img);
          b.style.marginRight = '0.3em';
          b.contentEditable = 'false';
          para.appendChild(b);
        } else {
          const b = document.createElement('span');
          b.className = 'bullet-mark';
          // 间距用 CSS margin 而非 \u00A0（理由同上：符号字体 fallback 时 nbsp 渲染成点）
          b.textContent = (typeof bullet === 'object' && bullet.char) || '•';
          b.style.marginRight = '0.3em';
          b.style.fontSize = `${bulletFontSize}px`;
          b.style.color = bulletColor;
          // 符号字体（a:buFont，如 Wingdings / Wingdings 3）：缺了它，
          // U+F0AD 之类的私用区字符会退化成豆腐块，画不出图标
          const bf = (typeof bullet === 'object' && bullet.font) || '';
          b.style.fontFamily = bf ? quoteFont(bf) : bulletFont;
          // a:buSzPct：符号相对正文的字号比例
          if (typeof bullet === 'object' && bullet.sizePct) {
            b.style.fontSize = `${(bulletFontSize * bullet.sizePct / 100).toFixed(2)}px`;
          }
          b.contentEditable = 'false';
          para.appendChild(b);
        }
      }
    }
    // 段落缩进按 OOXML 悬挂缩进语义实现：
    //   marL（p.indentLeft）= 正文左边界（折行也停在这里）
    //   indent（p.indent）  = 首行相对 marL 的偏移，项目符号通常为负
    // 合起来即：项目符号起点 = marL + indent，折行起点 = marL。
    // 修复前只读 indent 并用行内 padding-left 覆盖了样式表 .bullet-para/.num-para 的
    // em 级 text-indent，导致首行符号被再按字号左拉一次（各层级字号不同→错位不同），
    // 且完全忽略 marL，导致二级列表的符号反而比一级更靠左。
    // 有源数据时用行内 padding-left + text-indent 精确表达并覆盖 em 兜底；
    // 无源数据（编辑器新建的列表）保持样式表的 1.3em/1.8em 悬挂缩进。
    const marLSrc = p.indentLeft != null ? p.indentLeft : (el.indent ? el.indent : null);
    const indentSrc = p.indent != null ? p.indent : null;
    if (marLSrc != null || indentSrc != null) {
      const marL = Math.max(0, marLSrc || 0);
      // 首行（项目符号/编号）不得越出文本框，与预览端的 clamp 一致
      const firstLine = (marL + (indentSrc || 0) < 0) ? -marL : (indentSrc || 0);
      para.style.paddingLeft = `${marL}px`;
      para.style.textIndent = `${firstLine}px`;
    }
    if (p.spaceBefore) para.style.marginTop = `${p.spaceBefore}px`;
    if (p.spaceAfter) para.style.marginBottom = `${p.spaceAfter}px`;
    (p.runs || []).forEach((run) => {
      if (run.break) {
        // 软换行（a:br）：空文本 run 渲染为 <br>
        para.appendChild(document.createElement('br'));
        return;
      }
      const span = document.createElement('span');
      const st = span.style;
      st.fontSize = `${ptToPx(effFont(run))}px`;
      // 上下标（a:rPr@baseline）：>0 上标、<0 下标，浏览器 super/sub 自动缩小并抬高基线
      if (run.baseline != null && run.baseline !== 0) {
        st.verticalAlign = run.baseline > 0 ? 'super' : 'sub';
      }
      // 字符间距（字距，a:rPr/a:spc，单位 px）
      if (run.spacing) st.letterSpacing = `${run.spacing}px`;
      st.fontFamily = fontFamilyOf(
        run.fontFace, run.fontFaceEa, el.fontFace, el.fontFaceEa, '微软雅黑'
      ) || quoteFont('微软雅黑');
      // 超链接（a:hlinkClick）：run 无显式色时用主题 hlink 色（OOXML 语义，与预览端一致）；
      // 有显式色则保持原色（PowerPoint 行为：显式色优先于主题链接色）
      if (run.href) {
        st.color = normalizeColor(run.color) || normalizeColor(el.color)
          || getTheme(ctx.theme).accents[0] || '#1A73E8';
      } else {
        st.color = normalizeColor(run.color) || normalizeColor(el.color) || '#202124';
      }
      if (run.bold ?? el.bold) st.fontWeight = '700';
      if (run.italic ?? el.italic) st.fontStyle = 'italic';
      if (run.underline ?? el.underline) st.textDecoration = 'underline';
      // 删除线（a:rPr@strike / a:strike）：与下划线可叠加
      if (run.strike ?? el.strike) st.textDecoration = (st.textDecoration ? st.textDecoration + ' ' : '') + 'line-through';
      // 文本高亮（a:rPr/a:highlight）：CSS 背景色块（非半透明）
      if (run.highlight || el.highlight) st.backgroundColor = normalizeColor(run.highlight || el.highlight);
      // 小型大写（a:rPr@cap="small"）
      if (run.smallCaps ?? el.smallCaps) st.fontVariant = 'small-caps';
      // 着重号（a:rPr/a:em）：CSS text-emphasis 模拟（位置取下，贴近东亚下标着重习惯）
      if (run.emphasisMark ?? el.emphasisMark) {
        const emType = run.emphasisMark || el.emphasisMark;
        const cssEm = emType === 'circle' ? 'circle' : emType === 'comma' ? 'comma' : 'dot';
        st.textEmphasisStyle = cssEm;
        st.textEmphasisPosition = 'under';
      }
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
          // run.shadow.alpha 是 0..1 不透明度，withAlpha 收的是 0..100 透明度，需换算
          const a = run.shadow.alpha != null ? run.shadow.alpha : 1;
          // blurRad 在 OOXML 中缺省即 0 = 硬边投影，预览端据此生成 3 段式 text-shadow，
          // 这里不做额外柔化（曾按最小 2pt 渲染会把投影糊成一片蓝光）
          const blur = run.shadow.blur != null ? run.shadow.blur : 0;
          shadows.push(`${ptToPx(run.shadow.x || 0)}px ${ptToPx(run.shadow.y || 0)}px ${ptToPx(blur)}px ${withAlpha(sc, (1 - a) * 100)}`);
        }
      }
      if (shadows.length) st.textShadow = shadows.join(',');
      if (run.href) {
        // 链接 run 渲染为 <a>：浏览器默认下划线 + 可点击；内部跳转 '#N' 在放映态定位页
        const a = document.createElement('a');
        a.textContent = run.text == null ? '' : String(run.text);
        a.href = run.href;
        if (run.hrefTooltip) a.title = run.hrefTooltip;
        if (String(run.href).startsWith('#')) {
          a.dataset.slideJump = String(run.href).slice(1);
        } else {
          a.target = '_blank';
          a.rel = 'noopener noreferrer';
        }
        a.style.textDecoration = 'underline';
        // 颜色必须设在 <a> 上：UA 样式表的 a:link{color:blue} 会盖过 span 的继承色
        a.style.color = st.color;
        // 编辑态点击只做选中/拖拽，不导航（放映态由 present 层放行）
        if (ctx.editing) a.addEventListener('click', (e) => e.preventDefault());
        span.textContent = '';
        span.appendChild(a);
      } else {
        span.textContent = run.text == null ? '' : String(run.text);
      }
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

/**
 * 拼 font-family：西文字体（OOXML 的 a:latin）在前、东亚字体（a:ea）在后，
 * 浏览器按字符集自动回退 —— 中文 PPT 常见 `a:latin="Arial"` + `a:ea="微软雅黑"`，
 * 只写 latin 会让中文回退到系统默认字形（宋体/黑体），与预览端不一致。
 * 去重并忽略空值；全部为空时返回 null（交由调用方决定兜底）。
 */
function fontFamilyOf(...faces) {
  const out = [];
  for (const f of faces) {
    if (!f) continue;
    const q = quoteFont(f);
    if (!out.includes(q)) out.push(q);
  }
  return out.length ? out.join(', ') : null;
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
  // OOXML trHeight 是最小行高：内容更高时行自动撑高（与 PowerPoint/预览端一致）
  grid.style.gridTemplateRows = rows.map((r) => `minmax(${((Number(r.height) || 40) * hScale).toFixed(1)}px, auto)`).join(' ');

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
  // transparency 是 0..100 透明度，需转 rgba（与 shapeVisual 的 CSS 路径保持一致）
  if (f.type === 'solid' && f.color) return withAlpha(f.color, f.transparency || 0);
  if (f.type === 'gradient' && Array.isArray(f.stops) && f.stops.length) return withAlpha(f.stops[0].color, f.stops[0].transparency || 0);
  return null; // image / pattern → 回退 CSS 渲染管线
}

/** 预设几何 SVG（与预览端同源公式），不支持或需 CSS 填充时返回 null */
function presetShapeSvg(el) {
  if (!el.shapeType) return null;
  const fill = presetFillColor(el);
  if (fill === null) return null;
  // rect 默认走 CSS（渐变/圆角渲染精度更高）；仅纯色/无填充时改用 SVG，
  // 使虚线边框的 dasharray 与预览端一致（CSS border-style: dashed 的 dash 长度≈3×线宽，
  // 无法表达预览端 lgDash='10, 5' 这类固定图案）
  if (el.shapeType === 'rect') {
    const f = el.fill;
    if (!(f == null || f === 'none' || f.type === 'solid')) return null;
  }
  const geo = presetShapePath(el.shapeType, el.width || 100, el.height || 100, el.adjust || {});
  if (!geo) return null;
  const lineColor = el.line && el.line !== 'none' ? withAlpha(el.line.color, el.line.transparency || 0) : null;
  const lineW = el.line && el.line !== 'none' ? Math.max(0.5, ptToPx(el.line.width || 0.75)) : 0;
  const NS = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(NS, 'svg');
  const W = el.width || 100, H = el.height || 100;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  svg.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;overflow:visible';
  let tf = geo.transform || '';
  const flipTf = svgFlipTransform(el.flipH, el.flipV, W, H);
  if (flipTf) tf += (tf ? ' ' : '') + flipTf;
  const main = document.createElementNS(NS, 'path');
  main.setAttribute('d', geo.d);
  if (tf.trim()) main.setAttribute('transform', tf.trim());
  main.setAttribute('fill', geo.noFill ? 'none' : (fill || 'none'));
  if (geo.fillRule) main.setAttribute('fill-rule', geo.fillRule);
  if (lineColor) {
    main.setAttribute('stroke', lineColor);
    main.setAttribute('stroke-width', String(lineW));
    if (el.line && el.line.dashType) main.setAttribute('stroke-dasharray', SVG_DASH_MAP[el.line.dashType] || 'none');
    // 线帽（a:ln@cap：rnd→round / flat→butt / sq→square）：开放路径（连接线/弧线）端点样式
    if (el.line && el.line.cap) {
      const capMap = { rnd: 'round', flat: 'butt', sq: 'square' };
      main.setAttribute('stroke-linecap', capMap[el.line.cap] || 'butt');
    }
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

/** SVG path 水平/垂直翻转：scale(-1,…) 默认绕原点，会把路径翻转到 viewBox 外，
 *  需先 translate(W,0)/translate(0,H) 使其保持在元素框内（与 PowerPoint 先 flip 后 rot 一致）。
 */
function svgFlipTransform(flipH, flipV, W, H) {
  if (!flipH && !flipV) return '';
  const sx = flipH ? -1 : 1;
  const sy = flipV ? -1 : 1;
  const tx = flipH ? W : 0;
  const ty = flipV ? H : 0;
  return `translate(${tx},${ty}) scale(${sx},${sy})`;
}

/**
 * 自定义几何（a:custGeom）路径数据：路径自带 EMU 坐标系，按各 path 的 w/h
 * 归一化到 (W,H) 元素框（与预览端同空间）。
 * 必须归一化而不是把 EMU 坐标直接塞进 viewBox——viewBox 缩放会同步缩小
 * stroke-width（用户单位），1.25pt 的笔宽会被压到 1e-4 倍而"消失"。
 */
function custGeomPathD(cg, W, H) {
  if (!cg || !Array.isArray(cg.paths) || !cg.paths.length) return '';
  const n2 = (v) => Math.round(v * 100) / 100;
  const ds = [];
  cg.paths.forEach((p) => {
    const kx = W / (p.w || W), ky = H / (p.h || H);
    const X = (v) => n2((Number(v) || 0) * kx);
    const Y = (v) => n2((Number(v) || 0) * ky);
    const cmd = (c) => {
      switch (c.type) {
        case 'moveTo': return `M${X(c.x)} ${Y(c.y)}`;
        case 'lnTo': return `L${X(c.x)} ${Y(c.y)}`;
        case 'cubicBezTo': return `C${X(c.x1)} ${Y(c.y1)} ${X(c.x2)} ${Y(c.y2)} ${X(c.x)} ${Y(c.y)}`;
        case 'quadBezTo': return `Q${X(c.x1)} ${Y(c.y1)} ${X(c.x)} ${Y(c.y)}`;
        case 'close': return 'Z';
        default: return '';
      }
    };
    const d = (p.commands.map(cmd).filter(Boolean).join(' ') + (p.closed ? ' Z' : '')).trim();
    if (d) ds.push(d);
  });
  return ds.join(' ');
}

/** 自定义几何（a:custGeom）SVG */
function custGeomSvg(el) {
  const cg = el.custGeom;
  if (!cg || !Array.isArray(cg.paths) || !cg.paths.length) return null;
  const fill = presetFillColor(el);
  const NS = 'http://www.w3.org/2000/svg';
  const EW = el.width || 100, EH = el.height || 100;
  const d = custGeomPathD(cg, EW, EH);
  if (!d) return null;
  const svg = document.createElementNS(NS, 'svg');
  svg.setAttribute('viewBox', `0 0 ${EW} ${EH}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  svg.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;overflow:visible';
  const tf = svgFlipTransform(el.flipH, el.flipV, EW, EH);
  const main = document.createElementNS(NS, 'path');
  main.setAttribute('d', d);
  if (tf.trim()) main.setAttribute('transform', tf.trim());
  // 任一 path 闭合才填充（开放曲线如涂鸦仅描边）
  const anyClosed = cg.paths.some((p) => p.closed);
  main.setAttribute('fill', anyClosed && fill ? (fill || 'none') : 'none');
  const lineColor = el.line && el.line !== 'none' ? withAlpha(el.line.color, el.line.transparency || 0) : null;
  const lineW = el.line && el.line !== 'none' ? Math.max(0.5, ptToPx(el.line.width || 0.75)) : 0;
  if (lineColor) {
    main.setAttribute('stroke', lineColor);
    main.setAttribute('stroke-width', String(lineW));
    // stroke-dasharray 收虚线长度列表（不是 CSS border-style 关键字），solid 不输出
    const dash = el.line && el.line.dashType ? SVG_DASH_MAP[el.line.dashType] : null;
    if (dash) main.setAttribute('stroke-dasharray', dash);
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
        const { effects } = shapeVisual(el);
        // 阴影 / 发光：带形状底（弧线/饼图/椭圆等）的描边/填充层背后放一层 SVG 滤镜
        if (effects && (effects.shadow || effects.glow)) {
          const fx = shapeEffectSvg(el, effects);
          if (fx) node.appendChild(fx);
        }
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
        // 连接符：与预览端 shape.ts 完全对齐
        // - stroke-width 用 pt 值直接当 px（预览端 getBorder 返回 pt 作 px）
        // - 不设 stroke-linecap（预览端默认 butt）
        // - w/h 为 0 时设最小 SVG 尺寸保证可见（预览端同款处理）
        // - 箭头用 SVG <marker>（与预览端一致，消除手动 polygon 的反锯齿差异）
        const lc = normalizeColor((el.line && el.line.color) || '#5B9BD5') || '#5B9BD5';
        const lw = (el.line && el.line.width) || 1;
        const W = el.width || 0, H = el.height || 0;
        const NS = 'http://www.w3.org/2000/svg';
        const svg = document.createElementNS(NS, 'svg');
        const minSize = Math.max(lw * 2, 4);
        const svgW = Math.max(W, minSize), svgH = Math.max(H, minSize);
        svg.style.cssText = `position:absolute;inset:0;width:${svgW}px;height:${svgH}px;overflow:visible`;
        const fx = el.flipH ? -1 : 1, fy = el.flipV ? -1 : 1;
        if (el.flipH || el.flipV) svg.style.transform = `scaleX(${fx}) scaleY(${fy})`;
        const st = el.shapeType || '';
        const sid = el.id || Math.random().toString(36).slice(2);
        // 箭头 marker（与预览端 shape.ts markerTriangle 完全一致）
        const sArrow = el.line && el.line.startArrow;
        const eArrow = el.line && el.line.endArrow;
        if (sArrow || eArrow) {
          const defs = document.createElementNS(NS, 'defs');
          const mk = (id, type) => {
            const m = document.createElementNS(NS, 'marker');
            m.setAttribute('id', id);
            m.setAttribute('viewBox', '0 0 10 10');
            m.setAttribute('refX', '10');
            m.setAttribute('refY', '5');
            m.setAttribute('markerWidth', '5');
            m.setAttribute('markerHeight', '5');
            m.setAttribute('stroke', lc);
            m.setAttribute('fill', lc);
            m.setAttribute('orient', 'auto-start-reverse');
            m.setAttribute('markerUnits', 'strokeWidth');
            const p = document.createElementNS(NS, 'path');
            p.setAttribute('d', 'M 0 0 L 10 5 L 0 10 z');
            m.appendChild(p);
            return m;
          };
          if (sArrow) defs.appendChild(mk(`mkS_${sid}`, sArrow));
          if (eArrow) defs.appendChild(mk(`mkE_${sid}`, eArrow));
          svg.appendChild(defs);
        }
        // 线条：straightConnector1 用 <line>（与预览端一致），其他用 <path>
        let lineEl;
        if (st === 'straightConnector1') {
          lineEl = document.createElementNS(NS, 'line');
          lineEl.setAttribute('x1', '0');
          lineEl.setAttribute('y1', '0');
          lineEl.setAttribute('x2', String(W));
          lineEl.setAttribute('y2', String(H));
        } else {
          lineEl = document.createElementNS(NS, 'path');
          let d;
          if (st.startsWith('bentConnector')) {
            d = `M 0,0 L ${W * 0.5},0 L ${W * 0.5},${H} L ${W},${H}`;
          } else {
            const r = Math.min(Math.max(((el.adjust && el.adjust.adj1) != null ? el.adjust.adj1 : 50000) / 100000, 0), 1);
            d = `M 0,0 Q ${W * r},0 ${W / 2},${H / 2} Q ${W * (1 - r)},${H} ${W},${H}`;
          }
          lineEl.setAttribute('d', d);
          lineEl.setAttribute('fill', 'none');
        }
        lineEl.setAttribute('stroke', lc);
        lineEl.setAttribute('stroke-width', String(lw));
        const dash = el.line && el.line.dashType ? SVG_DASH_MAP[el.line.dashType] : null;
        if (dash) lineEl.setAttribute('stroke-dasharray', dash);
        if (sArrow) lineEl.setAttribute('marker-start', `url(#mkS_${sid})`);
        if (eArrow) lineEl.setAttribute('marker-end', `url(#mkE_${sid})`);
        svg.appendChild(lineEl);
        node.appendChild(svg);
        break;
      }
      if (el.shapeType === 'line') {
        const lc = normalizeColor((el.line && el.line.color) || '#000000') || '#000';
        const lw = Math.max(1, ptToPx((el.line && el.line.width) || 2));
        const W = el.width || 100, H = el.height || 100;
        const NS = 'http://www.w3.org/2000/svg';
        const svg = document.createElementNS(NS, 'svg');
        svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
        svg.setAttribute('preserveAspectRatio', 'none');
        svg.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;overflow:visible';
        const fx = el.flipH ? -1 : 1, fy = el.flipV ? -1 : 1;
        if (el.flipH || el.flipV) svg.style.transform = `scaleX(${fx}) scaleY(${fy})`;
        const line = document.createElementNS(NS, 'line');
        line.setAttribute('x1', '0'); line.setAttribute('y1', '0');
        line.setAttribute('x2', String(W)); line.setAttribute('y2', String(H));
        line.setAttribute('stroke', lc);
        line.setAttribute('stroke-width', String(lw));
        const lnDash = el.line && el.line.dashType ? SVG_DASH_MAP[el.line.dashType] : null;
        if (lnDash) line.setAttribute('stroke-dasharray', lnDash);
        svg.appendChild(line);
        node.appendChild(svg);
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
      // 裁剪：a:srcRect 各边为 0~1 比例，容器裁掉四周后图片放大填满（与预览端一致）
      const crop = el.crop;
      if (crop && (crop.l || crop.t || crop.r || crop.b)) {
        const cropW = Math.max(0.01, 1 - (crop.l || 0) - (crop.r || 0));
        const cropH = Math.max(0.01, 1 - (crop.t || 0) - (crop.b || 0));
        img.style.position = 'absolute';
        img.style.left = `${(-(crop.l || 0) / cropW * 100).toFixed(4)}%`;
        img.style.top = `${(-(crop.t || 0) / cropH * 100).toFixed(4)}%`;
        img.style.width = `${(100 / cropW).toFixed(4)}%`;
        img.style.height = `${(100 / cropH).toFixed(4)}%`;
        img.style.maxWidth = 'none';
        node.style.overflow = 'hidden';
      }
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
    case 'video':
    case 'audio': {
      renderMediaEl(el, node, ctx);
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
      // 组合级旋转/翻转：作用于整个组容器，子元素坐标保持相对偏移（与 PowerPoint 一致：
      // 旋转中心为组边界框中心；OOXML 先 flip 后 rot，故 CSS transform 写 rotate(rot) scale(flip)）。
      let tf = '';
      if (el.rotation) tf += ` rotate(${el.rotation}deg)`;
      if (el.flipH) tf += ' scaleX(-1)';
      if (el.flipV) tf += ' scaleY(-1)';
      if (tf) node.style.transform = tf.trim();
      break;
    }
    default: {
      node.innerHTML = `<div class="el-raw" style="width:100%;height:100%">${el.label || el.type || '未支持元素'}</div>`;
    }
  }
  // 形状特效：反射/柔化边缘/模糊（a:effectLst/a:reflection/a:softEdge/a:blur）——近似渲染
  applyShapeEffects(node, el);
  if (ctx.editing) node.dataset.editing = '1';
  return node;
}

// 形状特效近似渲染：反射（镜像）/ 柔化边缘 / 模糊（CSS filter）。
// 仅作用于编辑器/放映预览，不影响导出（导出由生成端按 OOXML effectLst 精确还原）。
function applyShapeEffects(node, el) {
  const fx = el.effects;
  if (!fx) return;
  const filters = [];
  if (fx.softEdge && fx.softEdge.radius) filters.push(`blur(${(fx.softEdge.radius * 0.6).toFixed(1)}px)`);
  if (fx.blur && fx.blur.radius) filters.push(`blur(${fx.blur.radius.toFixed(1)}px)`);
  if (filters.length) node.style.filter = filters.join(' ');
  if (fx.reflection) {
    // 水平镜像倒影：below 表示从元素底部向下镜像，渐变实现透明度淡出
    const dist = fx.reflection.distance ? fx.reflection.distance : 0;
    node.style.webkitBoxReflect = `below ${dist}px linear-gradient(transparent 55%, rgba(0,0,0,0.45))`;
  }
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
  return { chartId: 'chart_' + el.id, type, data, style, title: el.title || '', theme, dataLabels: !!el.dataLabels };
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
    // 旧数据可能残留未解析的 'scheme:<name>' 引用；解析器给出的色值不带 '#'（如 'FFFFFF'），
    // 需经 normalizeColor 归一化，否则 CSS 视为非法而回退成继承色
    const safeColor = (c, fallback) => {
      if (!c || /^scheme:/i.test(c)) return fallback;
      return normalizeColor(c) || fallback;
    };
    // 连接线层（SVG，viewBox 随容器缩放）
    const NS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
    svg.setAttribute('preserveAspectRatio', 'none');
    svg.style.cssText = 'position:absolute;inset:0;width:100%;height:100%';
    shapes.forEach((s) => {
      if (!s.connector) return;
      const x1 = s.flipH ? s.x + s.width : s.x;
      const y1 = s.flipV ? s.y + s.height : s.y;
      const x2 = s.flipH ? s.x : s.x + s.width;
      const y2 = s.flipV ? s.y : s.y + s.height;
      const line = document.createElementNS(NS, 'line');
      line.setAttribute('x1', String(x1));
      line.setAttribute('y1', String(y1));
      line.setAttribute('x2', String(x2));
      line.setAttribute('y2', String(y2));
      const stroke = safeColor(s.lineColor, '#B0B4BA');
      const lw = Math.max(0.75, s.lineWidth || 1);
      line.setAttribute('stroke', stroke);
      line.setAttribute('stroke-width', String(lw));
      svg.appendChild(line);
      const arrowLen = Math.max(7, lw * 3.2);
      const arrowW = Math.max(4, lw * 2.2);
      const addArrow = (px, py, dx, dy) => {
        const len = Math.hypot(dx, dy) || 1;
        const ux = dx / len, uy = dy / len;
        const bx = px - ux * arrowLen, by = py - uy * arrowLen;
        const px2 = -uy, py2 = ux; // 垂直方向
        const poly = document.createElementNS(NS, 'polygon');
        poly.setAttribute('points', `${px},${py} ${bx + px2 * arrowW / 2},${by + py2 * arrowW / 2} ${bx - px2 * arrowW / 2},${by - py2 * arrowW / 2}`);
        poly.setAttribute('fill', stroke);
        svg.appendChild(poly);
      };
      // OOXML：headEnd 在起点，tailEnd 在终点
      if (s.startArrow) addArrow(x1, y1, x1 - x2, y1 - y2);
      if (s.endArrow) addArrow(x2, y2, x2 - x1, y2 - y1);
    });
    wrap.appendChild(svg);

    // 节点框：能由预设几何生成路径的（arc/pie/star 等）走 SVG，其余回退 CSS 盒
    const radiusOf = (s) => {
      if (s.prst === 'ellipse') return '50%';
      if (!s.prst || s.prst === 'rect') return '2px';
      if (s.prst === 'roundRect') return Math.round(Math.min(s.width, s.height) * 0.18) + 'px';
      return '2px';
    };
    shapes.forEach((s) => {
      if (s.connector) return;
      const geo = s.prst ? presetShapePath(s.prst, s.width, s.height, s.adjust || {}) : null;
      let frame;
      if (geo && geo.d) {
        // SVG 形状：viewBox 与形状框 1:1，stroke-width 可直接用 CSS px
        const svg = document.createElementNS(NS, 'svg');
        svg.setAttribute('viewBox', `0 0 ${s.width} ${s.height}`);
        svg.setAttribute('preserveAspectRatio', 'none');
        svg.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;overflow:visible';
        let tf = geo.transform || '';
        if (s.flipH || s.flipV) tf += ` scale(${s.flipH ? -1 : 1},${s.flipV ? -1 : 1})`;
        const p = document.createElementNS(NS, 'path');
        p.setAttribute('d', geo.d);
        if (tf.trim()) p.setAttribute('transform', tf.trim());
        const fc = s.fill && s.fill !== 'none' ? safeColor(s.fill, 'none') : 'none';
        p.setAttribute('fill', geo.noFill ? 'none' : fc);
        if (geo.fillRule) p.setAttribute('fill-rule', geo.fillRule);
        const lc = safeColor(s.lineColor, null);
        if (lc) {
          p.setAttribute('stroke', lc);
          p.setAttribute('stroke-width', String(Math.max(0.5, s.lineWidth || 0.75)));
          p.setAttribute('stroke-linecap', 'round');
        }
        svg.appendChild(p);
        frame = svg;
      } else {
        frame = h('div', {
          style: {
            position: 'absolute', inset: '0',
            background: s.fill && s.fill !== 'none' ? safeColor(s.fill, 'transparent') : 'transparent',
            border: safeColor(s.lineColor, null) ? `${Math.max(0.5, s.lineWidth || 0.75)}px solid ${safeColor(s.lineColor, '')}` : 'none',
            borderRadius: radiusOf(s),
            display: 'flex',
            alignItems: s.anchor === 't' ? 'flex-start' : s.anchor === 'b' ? 'flex-end' : 'center',
            justifyContent: s.align === 'l' ? 'flex-start' : s.align === 'r' ? 'flex-end' : 'center',
            padding: '2px 8px', boxSizing: 'border-box', overflow: 'hidden'
          }
        });
      }
      const cell = h('div', {
        style: {
          position: 'absolute',
          left: `${(s.x / W) * 100}%`, top: `${(s.y / H) * 100}%`,
          width: `${(s.width / W) * 100}%`, height: `${(s.height / H) * 100}%`
        }
      });
      cell.appendChild(frame);
      if (s.text && s.text.trim()) {
        const span = h('span', {
          style: {
            position: 'absolute', inset: '0',
            display: 'flex',
            flexDirection: 'column',
            alignItems: s.align === 'l' ? 'flex-start' : s.align === 'r' ? 'flex-end' : 'center',
            justifyContent: s.anchor === 't' ? 'flex-start' : s.anchor === 'b' ? 'flex-end' : 'center',
            fontSize: `${s.fontSize || 12}pt`, color: safeColor(s.color, '#FFFFFF'),
            fontWeight: s.bold ? 600 : 400, lineHeight: 1.2,
            textAlign: s.align === 'l' ? 'left' : s.align === 'r' ? 'right' : 'center',
            whiteSpace: 'pre-wrap', wordBreak: 'break-word',
            padding: '2px 8px', boxSizing: 'border-box', pointerEvents: 'none'
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

/**
 * 音视频元素（p:pic + p:nvPr/a:audioFile | a:videoFile）渲染。
 *
 * OOXML 里这类形状的 a:blip 指向的是「海报/占位图」（音频是喇叭图标，视频是首帧封面），
 * PowerPoint 播放前展示的就是它，所以视觉主体始终用 poster：
 *  - 视频：<video poster> 未播放时显示封面，播放后显示画面；原生控件仅在选中时可交互
 *  - 音频：<img> 封面 + 隐藏的 <audio>。原先把原生 <audio controls> 直接塞进元素框，
 *    在 21×21px 的喇叭图标尺寸下会被压成一条竖线，完全不可用
 *
 * 播放/暂停徽标按元素尺寸缩放（小到 21px 也能看清），点击时 stopPropagation，
 * 避免触发画布的拖拽；原生 controls 需要 pointer-events，故由 CSS 依据 .is-sel 放开。
 */
function renderMediaEl(el, host, ctx = {}) {
  const isVideo = el.type === 'video';
  const poster = el.poster || {};
  const posterSrc = mediaSrc(poster.data || poster.src || '', poster.extension || 'png', 'image/png');
  const box = h('div', { class: 'el-media', style: { width: '100%', height: '100%', position: 'relative', overflow: 'hidden' } });

  // 缩略图：只画海报，不建媒体元素（否则每页缩略图都会拉一次媒体 metadata）
  if (ctx.chartScope === 'thumb') {
    if (posterSrc) {
      const img = document.createElement('img');
      img.src = posterSrc;
      img.alt = '';
      img.draggable = false;
      img.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;object-fit:contain;display:block';
      box.appendChild(img);
    }
    host.appendChild(box);
    return;
  }

  const src = mediaSrc(el.data || el.src || '', el.extension, isVideo ? 'video/mp4' : 'audio/mpeg');
  const media = document.createElement(isVideo ? 'video' : 'audio');
  if (src) media.src = src;
  media.preload = 'metadata';
  media.setAttribute('playsinline', '');
  if (isVideo) {
    if (posterSrc) media.poster = posterSrc;
    media.controls = true;
    media.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;object-fit:contain;display:block;background:#000';
  } else {
    // 音频：封面图作视觉主体，媒体元素只负责播放（隐藏但仍可 load/play）
    const img = document.createElement('img');
    img.src = posterSrc;
    img.alt = '';
    img.draggable = false;
    img.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;object-fit:contain;display:block';
    box.appendChild(img);
    media.style.cssText = 'position:absolute;width:1px;height:1px;opacity:0;left:-9999px;top:0;pointer-events:none';
  }
  box.appendChild(media);

  // 播放/暂停徽标（尺寸跟随元素，小图标不至于被撑破）
  // 注意：display/对齐放在 CSS（.media-badge），选中态才能用 `.is-sel .media-badge{display:none}` 覆盖
  const boxW = el.width || 40, boxH = el.height || 40;
  const badge = Math.max(10, Math.min(44, Math.min(boxW, boxH) * 0.62));
  const icon = h('div', {
    class: 'media-badge',
    title: '播放/暂停',
    style: {
      position: 'absolute', left: '50%', top: '50%', transform: 'translate(-50%, -50%)',
      width: `${badge}px`, height: `${badge}px`, borderRadius: '50%',
      background: 'rgba(0,0,0,0.55)', cursor: 'pointer'
    }
  });
  const paint = (playing) => {
    const s = Math.round(badge * 0.5);
    icon.innerHTML = playing
      ? `<svg width="${s}" height="${s}" viewBox="0 0 20 20"><rect fill="#fff" x="5" y="4" width="3.6" height="12" rx="1"/><rect fill="#fff" x="11.4" y="4" width="3.6" height="12" rx="1"/></svg>`
      : `<svg width="${s}" height="${s}" viewBox="0 0 20 20"><polygon fill="#fff" points="6,4 16,10 6,16"/></svg>`;
  };
  paint(false);
  const toggle = (e) => {
    // 不阻断冒泡：点击徽标既要播放，也要让画布收到 pointerdown 完成选中
    // （小尺寸音频元素几乎被徽标占满，若吞掉事件就无法选中）
    if (media.paused || media.ended) { const r = media.play(); if (r && r.catch) r.catch(() => {}); }
    else media.pause();
  };
  icon.addEventListener('pointerdown', toggle);
  media.addEventListener('play', () => paint(true));
  media.addEventListener('pause', () => paint(false));
  media.addEventListener('ended', () => paint(false));
  if (isVideo) media.addEventListener('playing', () => paint(true));
  box.appendChild(icon);

  // 无媒体源时给出可编辑的占位提示（双击可打开媒体对话框）
  if (!src) {
    box.appendChild(h('div', {
      style: {
        position: 'absolute', inset: '0', display: 'flex', alignItems: 'center', justifyContent: 'center',
        color: '#64748b', fontSize: '11px', textAlign: 'center', padding: '2px', pointerEvents: 'none'
      },
      text: isVideo ? '视频占位' : '音频占位'
    }));
    icon.style.display = 'none';
  }
  host.appendChild(box);
}

function renderChartEl(el, ctx) {
  const id = 'chart_' + (ctx.chartScope || 'canvas') + '_' + el.id;
  const theme = getTheme(ctx.theme);
  const wrap = h('div', { class: 'el-chart', style: { width: '100%', height: '100%' } });
  wrap.id = id;
  // 图表区填充（c:chartSpace/c:spPr）：'none' 透明，缺省透明，色值铺满
  if (el.spaceFill === 'none') {
    wrap.style.background = 'transparent';
  } else if (el.spaceFill && el.spaceFill !== 'none') {
    wrap.style.background = normalizeColor(el.spaceFill) || 'transparent';
  }

  // 优先用 ECharts（与预览端一致的 option 构建，支持 3D）；缺库时回退到内置 SVG
  if (typeof window !== 'undefined' && window.echarts && chartRenderer && chartRenderer.prepareEChartsOption) {
    try {
      const info = buildChartInfo(el, ctx);
      const option = chartRenderer.prepareEChartsOption(info);
      if (option) {
        if (!el.legend && option.legend) delete option.legend;
        const cached = _echartsMap.get(id);
        // 复用已有实例：仅更新 option 与尺寸，避免整页重绘时反复销毁重建 ECharts（造成闪烁/卡顿）
        if (cached && cached.dom && cached.type === info.type) {
          cached.dom.style.background = wrap.style.background;
          enqueueChartUpdate(cached, option);
          return cached.dom;
        }
        enqueueEchart(wrap, option, info.type);
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

// 复用实例：仅 setOption + resize，不销毁重建
function enqueueChartUpdate(cached, option) {
  requestAnimationFrame(() => {
    try { cached.chart.setOption(option); } catch (e) { /* 极端情况下忽略 */ }
    const w = cached.dom.clientWidth, h = cached.dom.clientHeight;
    if (w > 0 && h > 0) cached.chart.resize();
  });
}

function enqueueEchart(wrap, option, type) {
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
    if (prev && prev !== inst) {
      window.removeEventListener('resize', prev.onResize);
      try { prev.chart.dispose(); } catch { /* ignore */ }
    }
    const onResize = () => inst.resize();
    window.addEventListener('resize', onResize);
    _echartsMap.set(wrap.id, { chart: inst, onResize, dom: wrap, type });
    if (!hasSize) {
      requestAnimationFrame(() => { if (wrap.clientWidth > 0 && wrap.clientHeight > 0) inst.resize(); });
    }
  });
}

/** 仅释放已脱离文档的 ECharts 实例（图表被删除 / 切换幻灯片等场景） */
export function disposeDetachedCharts() {
  for (const [id, entry] of _echartsMap) {
    if (!entry.dom || !entry.dom.isConnected) {
      window.removeEventListener('resize', entry.onResize);
      try { entry.chart.dispose(); } catch { /* ignore */ }
      _echartsMap.delete(id);
    }
  }
}

/** 释放所有 ECharts 实例（卸载画布 / 关闭文档前调用，避免内存泄漏） */
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
      // adj = 圆角半径占短边比例（OOXML 原生千分比，默认 16667）。
      // 不设上限：预览端按同一公式换算并允许半径达到短边（此时视觉上即胶囊/椭圆），
      // 此前额外 min(rad, 60) 会在大尺寸形状上把圆角压平，与预览端不符。
      const rad = ((adj.adj != null ? adj.adj : 16667) / 100000) * Math.min(el.width, el.height);
      return { borderRadius: `${rad}px` };
    }
    // 圆角变体：与预览端 src/shape/shape.ts 的 round + cornr1/cornr2 分支同源。
    // OOXML 用 adj1/adj2 指定被圆角的两角，缺省 adj1=33333（半径占短边 1/3）、adj2=0。
    case 'round1Rect':
    case 'round2SameRect':
    case 'round2DiagRect': {
      const a1 = (adj.adj1 != null ? adj.adj1 : 33333) / 100000;
      const a2 = (adj.adj2 != null ? adj.adj2 : 0) / 100000;
      const r1 = a1 * Math.min(el.width, el.height);
      const r2 = a2 * Math.min(el.width, el.height);
      const B = '0px';
      const R = (n) => `${n}px`;
      // adj1/adj2 落在哪个角由 OOXML 的 gt/avLst 语义固定：round1Rect 只第一角，
      // round2SameRect 为上方两角，round2DiagRect 为对角两角。
      if (el.shapeType === 'round1Rect') return { borderRadius: `${R(r1)} ${B} ${B} ${B}` };
      if (el.shapeType === 'round2SameRect') return { borderRadius: `${R(r1)} ${R(r1)} ${B} ${B}` };
      return { borderRadius: `${R(r1)} ${R(r2)} ${R(r1)} ${R(r2)}` };
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
    case 'plus': {
      // 与 presetShapePath('plus') 同源：adj 缺省 25000 → 竖/横臂占 50%（25%/75%）。
      // 原先硬编码 35%/65%，使图片填充的十字比形状本体细一圈，与预览端不符。
      const a1 = Math.min(Math.max(adj.adj != null ? adj.adj : 25000, 0), 50000) / 100000;
      const a2 = 1 - a1;
      const p1 = `${(a1 * 100).toFixed(2)}%`, p2 = `${(a2 * 100).toFixed(2)}%`;
      return { clipPath: `polygon(${p1} 0%,${p1} ${p1},0% ${p1},0% ${p2},${p1} ${p2},${p1} 100%,${p2} 100%,${p2} ${p2},100% ${p2},100% ${p1},${p2} ${p1},${p2} 0%)` };
    }
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

/**
 * 阴影/发光层的轮廓几何：优先用与本体同源的预设路径（presetShapePath），
 * 保证 arc/pie/箭头等异形的阴影跟随真实轮廓；无预设路径时回退 CSS 几何
 * （clip-path 多边形 / border-radius 圆角矩形或椭圆）。
 * 无填充 + 有描边的形状（arc 弧线、noSmoking 等）用描边轮廓，避免阴影糊成实心块。
 */
function effectGeometry(el, W, H) {
  const fill = presetFillColor(el);
  const strokeOnly = fill === 'none' || fill == null;
  const lineColor = el.line && el.line !== 'none' ? withAlpha(el.line.color, el.line.transparency || 0) : null;
  const lineW = el.line && el.line !== 'none' ? Math.max(0.5, ptToPx(el.line.width || 0.75)) : 0;
  // 0) 自定义几何（a:custGeom）：与本体 custGeomSvg 同源，复用其归一化路径
  if (el.custGeom && Array.isArray(el.custGeom.paths) && el.custGeom.paths.length) {
    const d = custGeomPathD(el.custGeom, W, H);
    if (d) {
      const tf = svgFlipTransform(el.flipH, el.flipV, W, H);
      const tfAttr = tf.trim() ? ` transform="${tf.trim()}"` : '';
      // 阴影/发光需要 SourceAlpha：有填充时必须保留填充（颜色值本身会被 filter 替换），
      // 仅描边或开放曲线才 fill="none"，并用描边提供 alpha。
      const strokeAttrs = (strokeOnly && lineColor)
        ? ` stroke="${lineColor}" stroke-width="${lineW.toFixed(2)}" stroke-linecap="round"`
        : '';
      const shadowFill = strokeOnly ? 'none' : (fill || '#000000');
      return `<path d="${d}"${tfAttr} fill="${shadowFill}"${strokeAttrs}/>`;
    }
  }
  // 1) 预设几何（本体就用 presetShapeSvg 渲染，阴影同源）
  if (el.shapeType && el.shapeType !== 'rect' && fill !== null) {
    const geo = presetShapePath(el.shapeType, W, H, el.adjust || {});
    if (geo && geo.d) {
      let tf = geo.transform || '';
      const flipTf = svgFlipTransform(el.flipH, el.flipV, W, H);
      if (flipTf) tf += (tf ? ' ' : '') + flipTf;
      const tfAttr = tf.trim() ? ` transform="${tf.trim()}"` : '';
      const useStroke = strokeOnly && lineColor;
      const parts = [`<path d="${geo.d}"${tfAttr} fill="${useStroke || geo.noFill || !fill ? 'none' : fill}"${geo.fillRule ? ` fill-rule="${geo.fillRule}"` : ''}${useStroke ? ` stroke="${lineColor}" stroke-width="${lineW.toFixed(2)}" stroke-linecap="round"` : ''}/>`];
      (geo.strokes || []).forEach((s) => {
        parts.push(`<path d="${s.d}"${tfAttr} fill="none" stroke="#000" stroke-width="${(s.width || Math.max(1, W * 0.05)).toFixed(2)}" stroke-linecap="round"/>`);
      });
      return parts.join('');
    }
  }
  // 2) CSS 几何回退
  const geo = shapeGeometry(el);
  const cssFill = strokeOnly ? 'none' : (fill || '#000000');
  if (geo.clipPath) {
    const m = geo.clipPath.match(/polygon\(([^)]+)\)/);
    if (m) {
      const pts = m[1].trim().split(/\s*,\s*/).map((p) => p.trim().split(/\s+/).map((v) => parseFloat(v)));
      return `<polygon points="${pts.map(([x, y]) => `${(x / 100 * W).toFixed(2)},${(y / 100 * H).toFixed(2)}`).join(' ')}" fill="${cssFill}"/>`;
    }
    return `<rect x="0" y="0" width="${W}" height="${H}" fill="${cssFill}"/>`;
  }
  if (geo.borderRadius) {
    const br = geo.borderRadius;
    const strokeAttrs = strokeOnly && lineColor ? ` fill="none" stroke="${lineColor}" stroke-width="${lineW.toFixed(2)}"` : ` fill="${cssFill}"`;
    if (br === '50%') {
      return `<ellipse cx="${(W / 2).toFixed(2)}" cy="${(H / 2).toFixed(2)}" rx="${(W / 2).toFixed(2)}" ry="${(H / 2).toFixed(2)}"${strokeAttrs}/>`;
    }
    const num = parseFloat(br);
    const rx = br.endsWith('%') ? (num / 100 * W) : num;
    const ry = br.endsWith('%') ? (num / 100 * H) : num;
    return `<rect x="0" y="0" width="${W}" height="${H}" rx="${rx.toFixed(2)}" ry="${ry.toFixed(2)}"${strokeAttrs}/>`;
  }
  return `<rect x="0" y="0" width="${W}" height="${H}" fill="${cssFill}"/>`;
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
  const geom = effectGeometry(el, W, H);

  // blurRad 是「半径」，与 CSS 模糊半径换算到高斯标准差约为其一半（σ ≈ R / 2）
  const SHADOW_SIGMA_RATIO = 0.5, GLOW_DILATE_RATIO = 0.38, GLOW_SIGMA_RATIO = 0.17;
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
    const sc = normalizeColor(s.color) || '#000000';
    // transparency=0..100：0 表示完全不透明，100 表示完全透明
    const opacity = s.transparency != null ? Math.max(0, Math.min(1, (100 - s.transparency) / 100)) : 1;
    defs.push(`<filter id="${fid}" filterUnits="userSpaceOnUse" x="${-margin}" y="${-margin}" width="${W + margin * 2}" height="${H + margin * 2}" color-interpolation-filters="sRGB"><feGaussianBlur in="SourceAlpha" stdDeviation="${sigma.toFixed(2)}" result="b"/><feOffset in="b" dx="${s.dx}" dy="${s.dy}" result="o"/><feFlood flood-color="${sc}" flood-opacity="${opacity.toFixed(2)}" result="c"/><feComposite in="c" in2="o" operator="in"/></filter>`);
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
