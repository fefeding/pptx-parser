import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
import fs from 'node:fs';
import path from 'path';
const ROOT = '/Users/jiamao/project/github/pptx-parser';
const CHROME = '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell';
const arg = process.argv[2] || 'examples/fefeding金腾科技前端开发通道答辨.pptx';
const file = path.resolve(ROOT, arg);
const server = http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/') p = '/examples/index.html';
  const fp = path.join(ROOT, p);
  if (!fp.startsWith(ROOT)) { res.writeHead(403); res.end(); return; }
  fs.readFile(fp, (err, data) => { if (err) { res.writeHead(404); res.end(); return; } res.writeHead(200, { 'Content-Type': fp.endsWith('.js') ? 'application/javascript' : fp.endsWith('.css') ? 'text/css' : 'text/html' }); res.end(data); });
});
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const PORT = server.address().port; const BASE = `http://127.0.0.1:${PORT}`;
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
  await page.evaluate(() => { document.querySelectorAll('#result .slide-scaler').forEach((sc) => { const inner = sc.firstElementChild; if (!inner) return; sc.style.width = inner.offsetWidth + 'px'; sc.style.height = inner.offsetHeight + 'px'; inner.style.transform = 'scale(1)'; }); });
  const slide = (await page.$$('#result .slide'))[6];
  const data = await page.evaluate((s) => {
    const sb = s.getBoundingClientRect();
    // 预览：形状容器 .block（带 _type），取其在 slide 内的 bbox 与 fill
    const els = [...s.querySelectorAll('.block')].map((e) => {
      const r = e.getBoundingClientRect();
      const svg = e.previousElementSibling;
      const rect = svg && svg.tagName === 'svg' ? svg.querySelector('rect') : null;
      const fill = svg && svg.tagName === 'svg' && rect ? rect.getAttribute('fill') : (getComputedStyle(e).backgroundColor || '');
      return { x: +(r.left - sb.left).toFixed(0), y: +(r.top - sb.top).toFixed(0), w: +r.width.toFixed(0), h: +r.height.toFixed(0), fill, type: e.getAttribute('_type') };
    }).filter((e) => e.w > 5 && e.h > 5);
    return els;
  }, slide);
  await page.close(); return data;
}
async function renderEditor() {
  const page = await browser.newPage({ viewport: { width: 2200, height: 1200 } });
  await page.route('**/pptx.js.org/**', (r) => r.abort());
  const fileUrl = '/' + arg.split('\\').join('/');
  await page.goto(`${BASE}/examples/editor/_harness.html?file=${encodeURIComponent(fileUrl)}`);
  await page.waitForFunction(() => document.querySelectorAll('#root .slide-frame').length >= 1, null, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(4000);
  const frame = (await page.$$('#root .slide-frame'))[6];
  const data = await page.evaluate((fr) => {
    const sb = fr.getBoundingClientRect();
    // 编辑器：.el 元素（shape/text），取其 bbox 与背景
    const els = [...fr.querySelectorAll('.el')].map((e) => {
      const r = e.getBoundingClientRect();
      const bg = getComputedStyle(e).backgroundColor;
      return { x: +(r.left - sb.left).toFixed(0), y: +(r.top - sb.top).toFixed(0), w: +r.width.toFixed(0), h: +r.height.toFixed(0), fill: bg, type: e.getAttribute('data-type') || e.className };
    }).filter((e) => e.w > 5 && e.h > 5);
    return els;
  }, frame);
  await page.close(); return data;
}
const p = await renderPreview();
const e = await renderEditor();
console.log('=== PREVIEW shapes (slide7) ===');
p.sort((a, b) => a.y - b.y || a.x - b.x).forEach((d) => console.log(`(${d.x},${d.y}) ${d.w}x${d.h} fill=${d.fill} type=${d.type}`));
console.log('=== EDITOR shapes (slide7) ===');
e.sort((a, b) => a.y - b.y || a.x - b.x).forEach((d) => console.log(`(${d.x},${d.y}) ${d.w}x${d.h} fill=${d.fill} type=${d.type}`));
await browser.close(); server.close();
