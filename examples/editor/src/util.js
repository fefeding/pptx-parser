/**
 * 通用工具：DOM 构建、颜色、几何、提示等。
 *
 * 纯逻辑（颜色/几何/单位/字符串）已下沉到库（src/utils/*），此处从 dist 引入并
 * re-export，使其它 editor 模块的 `import ... from './util.js'` 无需改动；
 * 本文件仅保留编辑器特有的 DOM / 浏览器胶水。
 */
import {
  normalizeColor, rgbToHex, hexToRgb, luminance, shade, withAlpha, mixHex,
  bboxOf, unionBBox, rotatePoint,
  ptToPx, pxToPt, inToPx,
  escapeHtml, stripDataUrl, extOfDataUrl
} from '../../../dist/ppt-parser.browser.js';

export {
  normalizeColor, rgbToHex, hexToRgb, luminance, shade, withAlpha, mixHex,
  bboxOf, unionBBox, rotatePoint,
  ptToPx, pxToPt, inToPx,
  escapeHtml, stripDataUrl, extOfDataUrl
};

export const uid = (p = 'e') => `${p}_${Math.random().toString(36).slice(2, 9)}${(Date.now() % 46656).toString(36)}`;

export const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
export const round = (v, d = 2) => { const m = 10 ** d; return Math.round(v * m) / m; };
export const clone = (o) => (o == null ? o : JSON.parse(JSON.stringify(o)));

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
