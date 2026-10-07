/**
 * 计算编辑器/预览 PNG 的差异密度热力图（ASCII），定位差异集中在哪些区域。
 * 用法：node test/analyze-diff.mjs <slideIndex 0-based> [pptx路径]
 */
import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const ROOT = '/Users/jiamao/project/github/pptx-parser';
const OUT = path.join(ROOT, 'test', 'evp');
const TOL = 24;
const CHROME = '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell';
const slideIdx = parseInt(process.argv[2] || '6', 10);
const arg = process.argv[3] || 'examples/fefeding金腾科技前端开发通道答辨.pptx';
const file = path.resolve(ROOT, arg);

const server = http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/') p = '/examples/index.html';
  const fp = path.join(ROOT, p);
  if (!fp.startsWith(ROOT)) { res.writeHead(403); res.end(); return; }
  fs.readFile(fp, (err, data) => {
    if (err) { res.writeHead(404); res.end(); return; }
    res.writeHead(200, { 'Content-Type': fp.endsWith('.js') ? 'application/javascript' : fp.endsWith('.css') ? 'text/css' : 'text/html' });
    res.end(data);
  });
});
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const PORT = server.address().port;
const BASE = `http://127.0.0.1:${PORT}`;
const browser = await pw.chromium.launch({ executablePath: CHROME });
const buf = fs.readFileSync(file);

async function renderPreview() {
  const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
  await page.route('**/pptx.js.org/**', (r) => r.abort());
  await page.goto(`${BASE}/examples/index.html`);
  await page.waitForTimeout(600);
  await (await page.$('#uploadFileInput')).setInputFiles({ name: path.basename(file), mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer: buf });
  await page.waitForFunction(() => document.querySelectorAll('#result .slide').length >= 1, null, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(6000);
  await page.evaluate(() => {
    document.querySelectorAll('#result .slide-scaler').forEach((sc) => {
      const inner = sc.firstElementChild;
      if (!inner) return;
      sc.style.width = inner.offsetWidth + 'px';
      sc.style.height = inner.offsetHeight + 'px';
      inner.style.transform = 'scale(1)';
    });
  });
  const slides = await page.$$('#result .slide');
  const urls = [];
  for (let i = 0; i < slides.length; i++) urls.push('data:image/png;base64,' + (await slides[i].screenshot()).toString('base64'));
  await page.close();
  return urls;
}
async function renderEditor() {
  const page = await browser.newPage({ viewport: { width: 2200, height: 1200 } });
  await page.route('**/pptx.js.org/**', (r) => r.abort());
  const fileUrl = '/' + arg.split('\\').join('/');
  await page.goto(`${BASE}/examples/editor/_harness.html?file=${encodeURIComponent(fileUrl)}`);
  await page.waitForFunction(() => document.querySelectorAll('#root .slide-frame').length >= 1, null, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(4000);
  const frames = await page.$$('#root .slide-frame');
  const urls = [];
  for (let i = 0; i < frames.length; i++) urls.push('data:image/png;base64,' + (await frames[i].screenshot()).toString('base64'));
  await page.close();
  return urls;
}

const p = await renderPreview();
const e = await renderEditor();
const diffPage = await browser.newPage();
await diffPage.goto('about:blank');
const res = await diffPage.evaluate(async ({ ua, ub, tol, COLS, ROWS }) => {
  const load = (u) => new Promise((r) => { const i = new Image(); i.onload = () => r(i); i.src = u; });
  const [ia, ib] = await Promise.all([load(ua), load(ub)]);
  const w = ib.width, h = ib.height;
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  const ctx = c.getContext('2d');
  ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(ia, 0, 0, w, h); const da = ctx.getImageData(0, 0, w, h).data;
  ctx.drawImage(ib, 0, 0); const db = ctx.getImageData(0, 0, w, h).data;
  const grid = Array.from({ length: ROWS }, () => new Array(COLS).fill(0));
  const cellW = w / COLS, cellH = h / ROWS;
  let total = 0;
  let minx = w, miny = h, maxx = 0, maxy = 0;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = (y * w + x) * 4;
      if (Math.abs(da[i] - db[i]) > tol || Math.abs(da[i + 1] - db[i + 1]) > tol || Math.abs(da[i + 2] - db[i + 2]) > tol) {
        total++;
        const cx = Math.min(COLS - 1, (x / cellW) | 0), cy = Math.min(ROWS - 1, (y / cellH) | 0);
        grid[cy][cx]++;
        if (x < minx) minx = x; if (x > maxx) maxx = x;
        if (y < miny) miny = y; if (y > maxy) maxy = y;
      }
    }
  }
  const mx = Math.max(...grid.flat());
  const rows = grid.map((r) => r.map((v) => v === 0 ? ' ' : v / mx > 0.66 ? '#' : v / mx > 0.33 ? '+' : '.').join('')).join('\n');
  return { w, h, total, ratio: total / (w * h), minx, miny, maxx, maxy, rows };
}, { ua: e[slideIdx], ub: p[slideIdx], tol: TOL, COLS: 40, ROWS: 22 });

console.log(`slide ${slideIdx + 1}: ${res.w}x${res.h}, diff ${(res.ratio * 100).toFixed(1)}%`);
console.log(`diff bbox: x[${res.minx}-${res.maxx}] y[${res.miny}-${res.maxy}]`);
console.log(res.rows);
await browser.close();
server.close();
