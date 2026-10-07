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
  const labels = ['DATA','MIXIN','COMPONENTS','TEMPLATE'];
  return labels.map((lab) => {
    const txt = [...s.querySelectorAll('*')].find((e) => (e.textContent||'').trim() === lab);
    if (!txt) return {label: lab, found:false};
    let el = txt;
    while (el && !el.classList.contains('block')) el = el.parentElement;
    const svg = el.previousElementSibling;
    const rect = svg && svg.tagName === 'svg' ? svg.querySelector('rect') : null;
    return {
      label: lab,
      blockBg: getComputedStyle(el).backgroundColor,
      blockInline: el.getAttribute('style'),
      rectFill: rect ? rect.getAttribute('fill') : null,
      rectInline: rect ? rect.getAttribute('style') : null,
      svgClass: svg ? svg.className : null
    };
  });
}, slide);
console.log(JSON.stringify(data, null, 2));
await page.close(); await browser.close(); server.close();
