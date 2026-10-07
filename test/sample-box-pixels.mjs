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
  const pts = await page.evaluate((s) => {
    const c = document.createElement('canvas');
    c.width = s.offsetWidth; c.height = s.offsetHeight;
    const ctx = c.getContext('2d');
    // 用 html2canvas 思路：把 slide 内容画到 canvas 上
    // 更简单：对几个点用 elementFromPoint 不行，直接截 slide 像素
    return new Promise((resolve) => {
      s.scrollIntoView();
      requestAnimationFrame(() => {
        const r = s.getBoundingClientRect();
        const samples = [[200,300],[250,350],[180,250],[300,450],[350,500]].map(([x,y]) => {
          const el = document.elementFromPoint(r.left + x, r.top + y);
          return { x, y, tag: el ? el.tagName + '.' + (el.className||'') : null };
        });
        resolve(samples);
      });
    });
  }, slide);
  // 截 slide 小区域像素
  const shot = await slide.screenshot();
  const c = await page.evaluate((data) => {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const c = document.createElement('canvas');
        c.width = img.width; c.height = img.height;
        const ctx = c.getContext('2d');
        ctx.drawImage(img, 0, 0);
        const out = [];
        for (const [x,y] of [[200,300],[250,350],[180,250],[300,450],[350,500]]) {
          const d = ctx.getImageData(x, y, 1, 1).data;
          out.push({x,y,r:d[0],g:d[1],b:d[2],a:d[3]});
        }
        resolve(out);
      };
      img.src = 'data:image/png;base64,' + data;
    });
  }, shot.toString('base64'));
  await page.close(); return { pts, c };
}
async function editor() {
  const page = await browser.newPage({ viewport: { width: 2200, height: 1200 } });
  await page.route('**/pptx.js.org/**', (r) => r.abort());
  const fileUrl = '/' + arg.split('\\').join('/');
  await page.goto(`${BASE}/examples/editor/_harness.html?file=${encodeURIComponent(fileUrl)}`);
  await page.waitForFunction(() => document.querySelectorAll('#root .slide-frame').length >= 1, null, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(4000);
  const frame = (await page.$$('#root .slide-frame'))[6];
  const shot = await frame.screenshot();
  const c = await page.evaluate((data) => {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const c = document.createElement('canvas');
        c.width = img.width; c.height = img.height;
        const ctx = c.getContext('2d');
        ctx.drawImage(img, 0, 0);
        const out = [];
        for (const [x,y] of [[200,300],[250,350],[180,250],[300,450],[350,500]]) {
          const d = ctx.getImageData(x, y, 1, 1).data;
          out.push({x,y,r:d[0],g:d[1],b:d[2],a:d[3]});
        }
        resolve(out);
      };
      img.src = 'data:image/png;base64,' + data;
    });
  }, shot.toString('base64'));
  await page.close(); return { c };
}
const p = await preview();
const e = await editor();
console.log('PREVIEW pixels:', p.c);
console.log('EDITOR pixels:', e.c);
await browser.close(); server.close();
