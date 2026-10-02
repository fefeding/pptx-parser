#!/usr/bin/env node
/**
 * 把 PPTX 渲染为单个自包含 HTML 文件（含全局 CSS），用于浏览器预览或无头截图核对。
 *
 * 用法：
 *   node skills/pptx-parser/scripts/pptx-to-html.mjs <file.pptx> [--out out.html] [--page 10]
 *     --out   输出路径（默认 <file>.html）
 *     --page  只渲染指定页（1 起，可重复指定）
 *
 * 截图核对（渲染效果与 WPS/PowerPoint 对比时很有用）：
 *   "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --disable-gpu \
 *     --hide-scrollbars --screenshot=/tmp/p10.png --window-size=1280,720 \
 *     --virtual-time-budget=3000 file:///tmp/p10.html
 */

import { readFile, writeFile } from 'node:fs/promises';
import { resolve, basename, extname } from 'node:path';

const lib = await loadLib();
const { pptxToHtml } = lib;

const args = process.argv.slice(2);
const file = args.find((a) => !a.startsWith('--'));
const pick = (name) => {
  const i = args.indexOf(name);
  return i >= 0 ? args[i + 1] : undefined;
};
const out = pick('--out');
const pages = args.reduce((acc, a, i) => {
  if (a === '--page') acc.push(Number(args[i + 1]));
  return acc;
}, []);

if (!file) {
  console.error('用法: node pptx-to-html.mjs <file.pptx> [--out out.html] [--page 10]');
  process.exit(1);
}

// pptxToHtml 内部会用 document 测量文字宽度，Node 下必须先装一个 DOM 环境
await ensureDom();

const buf = await readFile(resolve(file));
const result = await pptxToHtml(toArrayBuffer(buf), { mediaProcess: true, themeProcess: true });

const chosen = pages.length
  ? pages.map((n) => result.slides[n - 1]).filter(Boolean)
  : result.slides;

const width = result.slideSize?.width || 1280;
const body = chosen.map((s) => s.html).join('\n');

const html = `<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>${escapeHtml(basename(file, extname(file)))}</title>
<style>
html,body{margin:0;padding:0;background:#e5e7eb;}
.slide{position:relative;margin:16px auto;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.2);}
${result.styles.global || ''}
</style></head>
<body>
${body}
</body></html>`;

const outPath = resolve(out || `${basename(file, extname(file))}.html`);
await writeFile(outPath, html);
console.log(`已输出: ${outPath}（${chosen.length}/${result.slides.length} 页，画布宽 ${width}px）`);
if (result.charts?.length) {
  console.log(`提示: 含 ${result.charts.length} 个图表，HTML 中只有占位；需自行用 examples/chart-lib/chart-renderer.js + echarts 渲染。`);
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Node 下没有 document：尝试加载 jsdom 搭一个最小 DOM（浏览器中可跳过） */
async function ensureDom() {
  if (typeof globalThis.document !== 'undefined') return;
  let JSDOM;
  try {
    ({ JSDOM } = await import('jsdom'));
  } catch {
    console.error('pptxToHtml 需要 DOM 环境。请安装 jsdom 后重试：npm i -D jsdom（或在浏览器中调用）。');
    process.exit(1);
  }
  const dom = new JSDOM('<!doctype html><html><body></body></html>', { pretendToBeVisual: true });
  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  Object.defineProperty(globalThis, 'navigator', { value: dom.window.navigator, configurable: true });
  for (const key of ['HTMLElement', 'Image', 'Node', 'DOMParser', 'XMLSerializer', 'getComputedStyle']) {
    if (dom.window[key] !== undefined) globalThis[key] = dom.window[key];
  }
}

/** Buffer → 独立 ArrayBuffer */
function toArrayBuffer(b) {
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
}

/** 优先用已安装的包，否则回退到仓库构建产物 */
async function loadLib() {
  try {
    return await import('@fefeding/ppt-parser');
  } catch {
    return await import('../../../dist/ppt-parser.esm.js');
  }
}
