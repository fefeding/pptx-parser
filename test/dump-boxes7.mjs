import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
import fs from 'node:fs';
import path from 'path';
const ROOT = '/Users/jiamao/project/github/pptx-parser';
const CHROME = '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell';
const arg = process.argv[2] || 'examples/fefeding金腾科技前端开发通道答辨.pptx';
const file = path.resolve(ROOT, arg);
const labels = ['DATA','MIXIN','COMPONENTS','TEMPLATE'];
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
  const data = await page.evaluate(({s, labels}) => {
    const sb = s.getBoundingClientRect();
    const out = [];
    labels.forEach((lab) => {
      const txt = [...s.querySelectorAll('*')].find((e) => (e.textContent||'').trim() === lab);
      if (!txt) { out.push({label:lab, found:false}); return; }
      let el = txt;
      // 向上找最近的 .block 祖先
      while (el && !el.classList.contains('block')) el = el.parentElement;
      if (!el) el = txt.parentElement;
      const r = el.getBoundingClientRect();
      const cs = getComputedStyle(el);
      const svg = el.previousElementSibling;
      const rect = svg && svg.tagName === 'svg' ? svg.querySelector('rect') : null;
      out.push({
        label: lab,
        tag: el.tagName + '.' + el.className,
        x: +(r.left - sb.left).toFixed(1), y: +(r.top - sb.top).toFixed(1), w: +r.width.toFixed(1), h: +r.height.toFixed(1),
        transform: cs.transform, opacity: cs.opacity, zIndex: cs.zIndex,
        bg: cs.backgroundColor,
        svgFill: rect ? rect.getAttribute('fill') : null,
        svgRect: rect ? `${rect.getAttribute('width')}x${rect.getAttribute('height')}` : null,
        inline: el.getAttribute('style') || ''
      });
    });
    return out;
  }, {s: slide, labels});
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
  const data = await page.evaluate(({fr, labels}) => {
    const sb = fr.getBoundingClientRect();
    const out = [];
    labels.forEach((lab) => {
      const txt = [...fr.querySelectorAll('*')].find((e) => (e.textContent||'').trim() === lab);
      if (!txt) { out.push({label:lab, found:false}); return; }
      let el = txt;
      while (el && !el.classList.contains('el')) el = el.parentElement;
      if (!el) el = txt.parentElement;
      const r = el.getBoundingClientRect();
      const cs = getComputedStyle(el);
      const childBg = [...el.querySelectorAll('div')].map((d) => getComputedStyle(d).backgroundColor).find((c) => c && c !== 'rgba(0, 0, 0, 0)');
      out.push({
        label: lab,
        tag: el.tagName + '.' + el.className,
        x: +(r.left - sb.left).toFixed(1), y: +(r.top - sb.top).toFixed(1), w: +r.width.toFixed(1), h: +r.height.toFixed(1),
        transform: cs.transform, opacity: cs.opacity, zIndex: cs.zIndex,
        bg: childBg || cs.backgroundColor,
        inline: el.getAttribute('style') || ''
      });
    });
    return out;
  }, {fr: frame, labels});
  await page.close(); return data;
}
const p = await preview();
const e = await editor();
console.log('=== PREVIEW boxes ===');
p.forEach((d)=>console.log(d));
console.log('=== EDITOR boxes ===');
e.forEach((d)=>console.log(d));
await browser.close(); server.close();
