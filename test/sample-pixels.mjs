import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
const ROOT = '/Users/jiamao/project/github/pptx-parser';
const OUT = path.join(ROOT, 'test', 'evp');
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
  const slides = await page.$$('#result .slide');
  const urls = []; for (let i = 0; i < slides.length; i++) urls.push('data:image/png;base64,' + (await slides[i].screenshot()).toString('base64'));
  await page.close(); return urls;
}
async function renderEditor() {
  const page = await browser.newPage({ viewport: { width: 2200, height: 1200 } });
  await page.route('**/pptx.js.org/**', (r) => r.abort());
  const fileUrl = '/' + arg.split('\\').join('/');
  await page.goto(`${BASE}/examples/editor/_harness.html?file=${encodeURIComponent(fileUrl)}`);
  await page.waitForFunction(() => document.querySelectorAll('#root .slide-frame').length >= 1, null, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(4000);
  const frames = await page.$$('#root .slide-frame');
  const urls = []; for (let i = 0; i < frames.length; i++) urls.push('data:image/png;base64,' + (await frames[i].screenshot()).toString('base64'));
  await page.close(); return urls;
}
const p = await renderPreview(); const e = await renderEditor();
const page = await browser.newPage(); await page.goto('about:blank');
const pts = [[90, 200], [640, 360], [430, 120], [430, 600], [200, 600]];
const out = await page.evaluate(async ({ ue, up, pts }) => {
  const load = (u) => new Promise((r) => { const i = new Image(); i.onload = () => r(i); i.src = u; });
  const [ie, ip] = await Promise.all([load(ue), load(up)]);
  const ce = document.createElement('canvas'); ce.width = ie.width; ce.height = ie.height; const xe = ce.getContext('2d'); xe.drawImage(ie, 0, 0);
  const cp = document.createElement('canvas'); cp.width = ip.width; cp.height = ip.height; const xp = cp.getContext('2d'); xp.drawImage(ip, 0, 0);
  return pts.map(([x, y]) => ({ x, y, e: [...xe.getImageData(x, y, 1, 1).data].slice(0, 3), p: [...xp.getImageData(x, y, 1, 1).data].slice(0, 3) }));
}, { ue: e[6], up: p[6], pts });
console.log('point -> [editorRGB, previewRGB]');
for (const o of out) console.log(`(${o.x},${o.y}) e=${JSON.stringify(o.e)} p=${JSON.stringify(o.p)}`);
await browser.close(); server.close();
