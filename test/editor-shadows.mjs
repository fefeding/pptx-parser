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
  fs.readFileSync;
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
  const names = ['DATA', 'MIXIN', 'COMPONENTS', 'TEMPLATE'];
  const out = {};
  for (const n of names) {
    const el = [...fr.querySelectorAll('.el-text')].find((e) => (e.textContent||'').trim() === n);
    if (!el) { out[n] = 'NOT FOUND'; continue; }
    // 找这个 el 的祖先中带 svg filter 的部分
    // 编辑器结构: 祖先 .el 内第一个绝对定位 div 里有 svg>defs>filter
    const ancestor = el.closest('.el') || el.parentElement;
    const svg = ancestor.querySelector('svg');
    const filter = svg ? svg.querySelector('filter') : null;
    const info = { filter: null, rect: null };
    if (filter) {
      const f = {};
      for (const fe of filter.children) {
        const attrs = {};
        for (const a of fe.attributes) attrs[a.name] = a.value;
        f[fe.tagName] = attrs;
      }
      info.filter = f;
    }
    const rect = svg ? svg.querySelector('rect') : null;
    if (rect) info.rect = rect.getAttribute('width') + 'x' + rect.getAttribute('height');
    // 背景 div 的 background 与 border
    const bgDiv = [...ancestor.children].find((c) => {
      const s = getComputedStyle(c);
      return s.backgroundColor !== 'rgba(0, 0, 0, 0)' || s.borderStyle !== 'none';
    });
    if (bgDiv) {
      const s = getComputedStyle(bgDiv);
      info.bg = s.backgroundColor;
      info.border = s.borderTopWidth + ' ' + s.borderTopStyle + ' ' + s.borderTopColor;
    }
    out[n] = info;
  }
  return out;
}, frame);
console.log(JSON.stringify(data, null, 2));
await page.close(); await browser.close(); server.close();
