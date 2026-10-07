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
const page = await browser.newPage({ viewport: { width: 2200, height: 1200 } });
await page.route('**/pptx.js.org/**', (r) => r.abort());
const fileUrl = '/' + arg.split('\\').join('/');
await page.goto(`${BASE}/examples/editor/_harness.html?file=${encodeURIComponent(fileUrl)}`);
await page.waitForFunction(() => document.querySelectorAll('#root .slide-frame').length >= 1, null, { timeout: 30000 }).catch(() => {});
await page.waitForTimeout(4000);
const frame = (await page.$$('#root .slide-frame'))[6];
const data = await page.evaluate((fr) => {
  const txt = [...fr.querySelectorAll('.el-text')].find((e) => (e.textContent||'').trim() === 'DATA');
  if (!txt) return null;
  // 收集所有带背景色的元素
  const bgs = [];
  const collect = (e, depth) => {
    const cs = getComputedStyle(e);
    const bg = cs.backgroundColor;
    if (bg && bg !== 'rgba(0, 0, 0, 0)') {
      const r = e.getBoundingClientRect();
      bgs.push({ depth, tag: e.tagName + '.' + (e.className||''), bg, x: r.left, y: r.top, w: r.width, h: r.height });
    }
    for (const c of e.children) collect(c, depth+1);
  };
  collect(txt, 0);
  return { html: txt.outerHTML.slice(0,1500), bgs };
}, frame);
console.log(JSON.stringify(data, null, 2));
await page.close(); await browser.close(); server.close();
