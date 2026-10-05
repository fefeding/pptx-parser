import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
const { chromium } = pw;
const ROOT = '/Users/jiamao/project/github/pptx-parser';
const fileBuffer = fs.readFileSync(path.join(ROOT, 'examples/Sample_12.pptx'));
const MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

const server = http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/') p = '/examples/index.html';
  const fp = path.join(ROOT, p);
  if (!fp.startsWith(ROOT)) { res.writeHead(403); res.end(); return; }
  fs.readFile(fp, (err, data) => {
    if (err) { res.writeHead(404); res.end(); return; }
    res.writeHead(200, { 'Content-Type': fp.endsWith('.js') ? 'application/javascript' : fp.endsWith('.css') ? 'text/css' : 'text/html', 'Access-Control-Allow-Origin': '*' }); res.end(data);
  });
}).listen(8771, '127.0.0.1');

const browser = await chromium.launch({
  executablePath: '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell'
});
const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
await page.route('**/pptx.js.org/**', (route) => route.abort());
await page.goto('http://127.0.0.1:8771/examples/index.html');
await page.waitForTimeout(800);
const input = await page.$('#uploadFileInput');
await input.setInputFiles({ name: 'Sample_12.pptx', mimeType: MIME, buffer: fileBuffer });
await page.waitForFunction(() => document.querySelectorAll('#result .slide').length >= 12, null, { timeout: 15000 }).catch(() => {});
const info = await page.evaluate(() => {
  const slides = document.querySelectorAll('#result .slide');
  const s = slides[4];
  const out = [];
  s.querySelectorAll('img, .block').forEach((el) => {
    const r = el.getBoundingClientRect();
    const sr = s.getBoundingClientRect();
    if (el.tagName === 'IMG' && r.width < 200 && r.width > 30) {
      out.push({
        src: (el.src || '').slice(0, 40),
        rel: { x: Math.round(r.left - sr.left), y: Math.round(r.top - sr.top), w: Math.round(r.width), h: Math.round(r.height) },
        style: el.getAttribute('style'),
        parentStyle: el.parentElement ? el.parentElement.getAttribute('style') : null
      });
    }
  });
  return out;
});
console.log(JSON.stringify(info, null, 1));
await browser.close();
server.close();
