/**
 * 通用工具：DOM 构建、颜色、几何、提示等
 */
export const uid = (p = 'e') => `${p}_${Math.random().toString(36).slice(2, 9)}${(Date.now() % 46656).toString(36)}`;

export const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
export const round = (v, d = 2) => { const m = 10 ** d; return Math.round(v * m) / m; };
export const clone = (o) => (o == null ? o : JSON.parse(JSON.stringify(o)));

/** pt → px（96dpi） */
export const ptToPx = (pt) => (Number(pt) || 0) * 96 / 72;
/** px → pt */
export const pxToPt = (px) => (Number(px) || 0) * 72 / 96;
/** in → px */
export const inToPx = (i) => (Number(i) || 0) * 96;

/** DOM 构建：h('div', {class:'x', onclick}, child, 'text') */
export function h(tag, props, ...children) {
  const node = document.createElement(tag);
  if (props) {
    for (const [k, v] of Object.entries(props)) {
      if (v == null || v === false) continue;
      if (k === 'class') node.className = v;
      else if (k === 'style' && typeof v === 'object') Object.assign(node.style, v);
      else if (k === 'html') node.innerHTML = v;
      else if (k === 'text') node.textContent = v;
      else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2).toLowerCase(), v);
      else if (k === 'dataset') Object.assign(node.dataset, v);
      else node.setAttribute(k, v === true ? '' : String(v));
    }
  }
  for (const c of children.flat()) {
    if (c == null || c === false) continue;
    node.appendChild(typeof c === 'string' || typeof c === 'number' ? document.createTextNode(String(c)) : c);
  }
  return node;
}

export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

export function svgFromPaths(paths, size = 24) {
  return `<svg viewBox="0 0 ${size} ${size}" width="100%" height="100%">${paths}</svg>`;
}

/* ======================= 颜色 ======================= */
const NAMED = {
  white: '#ffffff', black: '#000000', red: '#d93025', green: '#1e8e3e', blue: '#1a73e8',
  yellow: '#f9ab00', gray: '#80868b', grey: '#80868b', orange: '#f29900', purple: '#8430ce',
  pink: '#ff6d9e', cyan: '#12b5cb', transparent: null
};

/** 任意颜色写法 → '#RRGGBB'，无法识别返回 null */
export function normalizeColor(input) {
  if (input == null) return null;
  let c = String(input).trim().toLowerCase();
  if (!c) return null;
  if (NAMED[c] !== undefined) return NAMED[c];
  if (c === 'none' || c === 'transparent') return null;
  if (c[0] === '#') {
    c = c.slice(1);
    if (c.length === 3) c = c.split('').map((x) => x + x).join('');
    if (c.length === 8) c = c.slice(0, 6);
    if (/^[0-9a-f]{6}$/.test(c)) return '#' + c.toUpperCase();
    return null;
  }
  if (c.startsWith('rgb')) {
    const nums = c.replace(/[^0-9.,]/g, '').split(',').map(Number);
    if (nums.length >= 3) return rgbToHex(nums[0], nums[1], nums[2]);
  }
  return null;
}
export function rgbToHex(r, g, b) {
  const f = (v) => clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0');
  return `#${f(r)}${f(g)}${f(b)}`.toUpperCase();
}
export function hexToRgb(hex) {
  const c = normalizeColor(hex) || '#000000';
  return { r: parseInt(c.slice(1, 3), 16), g: parseInt(c.slice(3, 5), 16), b: parseInt(c.slice(5, 7), 16) };
}
/** 亮度 0(黑)~1(白) */
export function luminance(hex) {
  const { r, g, b } = hexToRgb(hex);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
}
/** 在其上叠加白/黑得到新的明度（用于生成渐变第二色） */
export function shade(hex, amount = 0.2) {
  const { r, g, b } = hexToRgb(hex);
  const t = amount < 0 ? 0 : 255;
  const p = Math.abs(amount);
  return rgbToHex(r + (t - r) * p, g + (t - g) * p, b + (t - b) * p);
}
/** 带透明度的 css 颜色 */
export function withAlpha(hex, alphaPct) {
  const c = normalizeColor(hex);
  if (!c) return 'transparent';
  if (!alphaPct) return c;
  const { r, g, b } = hexToRgb(c);
  return `rgba(${r},${g},${b},${clamp(1 - alphaPct / 100, 0, 1)})`;
}
export function mixHex(a, b, t = 0.5) {
  const A = hexToRgb(a), B = hexToRgb(b);
  return rgbToHex(A.r + (B.r - A.r) * t, A.g + (B.g - A.g) * t, A.b + (B.b - A.b) * t);
}

/* ======================= 几何 ======================= */
export function bboxOf(el) {
  const w = el.width || 0, h = el.height || 0;
  if (!el.rotation) return { x: el.x, y: el.y, width: w, height: h };
  const cx = el.x + w / 2, cy = el.y + h / 2;
  const rad = (el.rotation * Math.PI) / 180;
  const cos = Math.abs(Math.cos(rad)), sin = Math.abs(Math.sin(rad));
  const nw = w * cos + h * sin, nh = w * sin + h * cos;
  return { x: cx - nw / 2, y: cy - nh / 2, width: nw, height: nh };
}
export function unionBBox(els) {
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
export function rotatePoint(px, py, cx, cy, deg) {
  const r = (deg * Math.PI) / 180, cos = Math.cos(r), sin = Math.sin(r);
  const dx = px - cx, dy = py - cy;
  return { x: cx + dx * cos - dy * sin, y: cy + dx * sin + dy * cos };
}

/* ======================= 其它 ======================= */
export function debounce(fn, wait = 300) {
  let t = 0;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), wait); };
}

export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
}

export function downloadText(text, filename, mime = 'application/json') {
  downloadBlob(new Blob([text], { type: mime }), filename);
}

let toastTimer = [];
export function toast(msg, type = '') {
  const root = document.getElementById('toastRoot');
  if (!root) return;
  const node = h('div', { class: 'toast ' + type, text: msg });
  root.appendChild(node);
  const t = setTimeout(() => node.remove(), 2400);
  toastTimer.push(t);
  if (root.children.length > 4) root.firstChild.remove();
}

/** 读取文件为 ArrayBuffer */
export function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = () => reject(fr.error || new Error('读取文件失败'));
    fr.readAsArrayBuffer(file);
  });
}
export function readFileAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = () => reject(fr.error || new Error('读取文件失败'));
    fr.readAsDataURL(file);
  });
}
export function pickFile(accept) {
  return new Promise((resolve) => {
    const input = document.getElementById('hiddenFile') || document.createElement('input');
    input.type = 'file';
    input.accept = accept || '';
    input.value = '';
    input.onchange = () => resolve(input.files && input.files[0]);
    input.click();
  });
}
/** dataURL → 裸 base64 */
export function stripDataUrl(dataUrl) {
  if (typeof dataUrl !== 'string') return '';
  const i = dataUrl.indexOf(',');
  return i >= 0 ? dataUrl.slice(i + 1) : dataUrl;
}
export function extOfDataUrl(dataUrl, fallback = 'png') {
  const m = /^data:image\/([a-zA-Z0-9.+-]+)/.exec(String(dataUrl || ''));
  if (!m) return fallback;
  const e = m[1].toLowerCase();
  return e === 'jpeg' ? 'jpg' : e === 'svg+xml' ? 'svg' : e;
}

export function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
