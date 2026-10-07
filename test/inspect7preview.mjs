import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';

const ROOT = '/Users/jiamao/project/github/pptx-parser';
const MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const CHROME = '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell';
const file = path.resolve(ROOT, 'examples/fefeding金腾科技前端开发通道答辨.pptx');
const buf = fs.readFileSync(file);

const server = http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/') p = '/examples/index.html';
  const fp = path.join(ROOT, p);
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
const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
await page.route('**/pptx.js.org/**', (r) => r.abort());
const logs = [];
page.on('console', (m) => { if (m.text().includes('DBG')) logs.push(m.text()); });
await page.goto(`${BASE}/examples/index.html`);
await page.waitForTimeout(600);
await (await page.$('#uploadFileInput')).setInputFiles({ name: 'x.pptx', mimeType: MIME, buffer: buf });
await page.waitForFunction(() => document.querySelectorAll('#result .slide').length >= 21, null, { timeout: 30000 }).catch(() => {});
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
const info = await page.evaluate(() => {
  const s = document.querySelectorAll('#result .slide')[6];
  const sb = s.getBoundingClientRect();
  // 找带 'block' 类（形状容器）且文本含 MIXIN 的元素
  const blocks = [...s.querySelectorAll('.block')].filter((e) => (e.textContent || '').includes('MIXIN'));
  return blocks.slice(0, 1).map((e) => {
    const r = e.getBoundingClientRect();
    const svg = e.previousElementSibling;
    const rect = svg ? svg.querySelector('rect') : null;
    // 找含 MIXIN 且后代最少的叶子元素（真正的文本 run span）
    let leaf = e;
    let cur = e;
    while (cur && cur.children.length) {
      const child = [...cur.children].find((c) => (c.textContent || '').includes('MIXIN'));
      if (!child) break;
      leaf = child; cur = child;
    }
    const lcs = getComputedStyle(leaf);
    return {
      cls: e.className,
      x: +(r.left - sb.left).toFixed(1), y: +(r.top - sb.top).toFixed(1),
      w: +r.width.toFixed(1), h: +r.height.toFixed(1),
      blockBg: getComputedStyle(e).background,
      svgRect: rect ? (rect.getAttribute('fill')+' '+rect.getAttribute('width')+'x'+rect.getAttribute('height')) : '(no rect)',
      leafTag: leaf.tagName + '.' + (leaf.className||''),
      leafFontSize: lcs.fontSize, leafLineHeight: lcs.lineHeight, leafFontFamily: (lcs.fontFamily||'').slice(0,40),
      leafInline: leaf.getAttribute('style')
    };
  });
});
console.log(JSON.stringify(info, null, 1));
console.log('DBG LOGS:', logs.join(' | '));
await browser.close();
server.close();
