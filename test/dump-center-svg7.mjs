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
async function preview() {
  const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
  await page.route('**/pptx.js.org/**', (r) => r.abort());
  await page.goto(`${BASE}/examples/index.html`);
  await page.waitForTimeout(600);
  await (await page.$('#uploadFileInput')).setInputFiles({ name: path.basename(file), mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer: buf });
  await page.waitForFunction(() => document.querySelectorAll('#result .slide').length >= 1, null, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(6000);
  await page.evaluate(() => { document.querySelectorAll('#result .slide-scaler').forEach((sc) => { const inner = sc.firstElementChild; if (!inner) return; sc.style.width = inner.offsetWidth + 'px'; sc.style.height = inner.offsetHeight + 'px'; inner.style.transform = 'scale(1)'; }); });
  await page.waitForTimeout(400);
  const slide = (await page.$$('#result .slide'))[6];
  const data = await page.evaluate((s) => {
    const sb = s.getBoundingClientRect();
    const out = [];
    s.querySelectorAll('.block').forEach((e) => {
      const r = e.getBoundingClientRect();
      const bx = +(r.left - sb.left).toFixed(0), by = +(r.top - sb.top).toFixed(0);
      // 收集该 block 内（含前面兄弟 svg）所有 SVG 形状
      const svgs = [e.previousElementSibling, ...e.querySelectorAll('svg')].filter((x) => x && x.tagName === 'svg');
      svgs.forEach((svg) => {
        svg.querySelectorAll('rect, path, polygon, ellipse, circle, line, polyline').forEach((sh) => {
          const sr = sh.getBoundingClientRect();
          const fill = sh.getAttribute('fill') || getComputedStyle(sh).fill;
          const x = +(sr.left - sb.left).toFixed(0), y = +(sr.top - sb.top).toFixed(0);
          if (x > 250 && x < 950 && y > 150 && y < 560) {
            const inside = x <= 660 && x >= 620 && y <= 380 && y >= 340;
            out.push({ tag: sh.tagName, x, y, w: +sr.width.toFixed(0), h: +sr.height.toFixed(0), fill, insideCenter: inside });
          }
        });
      });
    });
    return out;
  }, slide);
  await page.close(); return data;
}
async function editor() {
  const page = await browser.newPage({ viewport: { width: 2200, height: 1200 } });
  await page.route('**/pptx.js.org/**', (r) => r.abort());
  const fileUrl = '/' + arg.split('\\').join('/');
  await page.goto(`${BASE}/examples/editor/_harness.html?file=${encodeURIComponent(fileUrl)}`);
  await page.waitForFunction(() => document.querySelectorAll('#root .slide-frame').length >= 1, null, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(4000);
  const frame = (await page.$$('#root .slide-frame'))[6];
  const data = await page.evaluate((fr) => {
    const sb = fr.getBoundingClientRect();
    const out = [];
    fr.querySelectorAll('.el').forEach((e) => {
      const r = e.getBoundingClientRect();
      const bx = +(r.left - sb.left).toFixed(0), by = +(r.top - sb.top).toFixed(0);
      const svgs = [...e.querySelectorAll('svg')];
      svgs.forEach((svg) => {
        svg.querySelectorAll('rect, path, polygon, ellipse, circle, line, polyline').forEach((sh) => {
          const sr = sh.getBoundingClientRect();
          const fill = sh.getAttribute('fill') || getComputedStyle(sh).fill;
          const x = +(sr.left - sb.left).toFixed(0), y = +(sr.top - sb.top).toFixed(0);
          if (x > 250 && x < 950 && y > 150 && y < 560) {
            const inside = x <= 660 && x >= 620 && y <= 380 && y >= 340;
            out.push({ tag: sh.tagName, x, y, w: +sr.width.toFixed(0), h: +sr.height.toFixed(0), fill, insideCenter: inside });
          }
        });
      });
    });
    return out;
  }, frame);
  await page.close(); return data;
}
const p = await preview();
const e = await editor();
console.log('=== PREVIEW center SVG shapes (x250-950, y150-560) ===');
p.sort((a,b)=>a.y-b.y||a.x-b.x).forEach((d)=>console.log(`${d.tag} (${d.x},${d.y}) ${d.w}x${d.h} fill=${d.fill}${d.insideCenter?'  <== CENTER':''}`));
console.log('=== EDITOR center SVG shapes ===');
e.sort((a,b)=>a.y-b.y||a.x-b.x).forEach((d)=>console.log(`${d.tag} (${d.x},${d.y}) ${d.w}x${d.h} fill=${d.fill}${d.insideCenter?'  <== CENTER':''}`));
await browser.close(); server.close();
