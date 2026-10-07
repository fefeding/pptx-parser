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
    const cs = getComputedStyle(s);
    const title = [...s.querySelectorAll('.block')].find((e) => (e.textContent||'').includes('C端难点之一'));
    const t = title || s.firstElementChild;
    return {
      slide: { ow: s.offsetWidth, oh: s.offsetHeight, cw: s.clientWidth, ch: s.clientHeight,
        border: `${cs.borderLeftWidth} ${cs.borderTopWidth}`, padding: `${cs.paddingLeft} ${cs.paddingTop}`, margin: `${cs.marginLeft} ${cs.marginTop}` },
      title: t ? { inline: t.getAttribute('style'), offsetLeft: t.offsetLeft, offsetTop: t.offsetTop,
        rectX: t.getBoundingClientRect().left - s.getBoundingClientRect().left,
        rectY: t.getBoundingClientRect().top - s.getBoundingClientRect().top } : null
    };
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
    const cs = getComputedStyle(fr);
    const title = [...fr.querySelectorAll('.el')].find((e) => (e.textContent||'').includes('C端难点之一'));
    const t = title || fr.firstElementChild;
    return {
      frame: { ow: fr.offsetWidth, oh: fr.offsetHeight, cw: fr.clientWidth, ch: fr.clientHeight,
        border: `${cs.borderLeftWidth} ${cs.borderTopWidth}`, padding: `${cs.paddingLeft} ${cs.paddingTop}`, margin: `${cs.marginLeft} ${cs.marginTop}` },
      title: t ? { inline: t.getAttribute('style'), offsetLeft: t.offsetLeft, offsetTop: t.offsetTop,
        rectX: t.getBoundingClientRect().left - fr.getBoundingClientRect().left,
        rectY: t.getBoundingClientRect().top - fr.getBoundingClientRect().top } : null
    };
  }, frame);
  await page.close(); return data;
}
const p = await preview();
const e = await editor();
console.log('PREVIEW:', JSON.stringify(p, null, 2));
console.log('EDITOR:', JSON.stringify(e, null, 2));
await browser.close(); server.close();
