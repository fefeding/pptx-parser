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
const slide = (await page.$$('#result .slide'))[6];
const data = await page.evaluate((sl) => {
  // 所有 filter
  const filters = {};
  sl.querySelectorAll('filter').forEach((f) => {
    const o = {};
    for (const fe of f.children) {
      const a = {}; for (const at of fe.attributes) a[at.name] = at.value;
      o[fe.tagName] = a;
    }
    filters[f.id] = o;
  });
  // 所有带 filter 的元素及引用关系
  const refs = [];
  sl.querySelectorAll('[filter]').forEach((el) => {
    const fid = el.getAttribute('filter').replace(/url\(#|\)/g, '');
    const txt = el.closest('.block')?.textContent?.slice(0,30) || '';
    refs.push({ tag: el.tagName, filterId: fid, text: txt, attrs: [...el.attributes].map(a=>a.name+'='+a.value.slice(0,30)) });
  });
  // 每个盒子的内部 SVG 结构（是否含 filter）
  const names = ['DATA','MIXIN','COMPONENTS','TEMPLATE'];
  const boxes = {};
  for (const n of names) {
    const el = [...sl.querySelectorAll('.block')].find(e => (e.textContent||'').includes(n));
    if (!el) continue;
    const html = el.outerHTML.slice(0,600);
    const fids = [...el.querySelectorAll('[filter]')].map(e => e.getAttribute('filter').replace(/url\(#|\)/g,''));
    boxes[n] = { html, filterIds: fids, filterDetails: fids.map(id => filters[id]) };
  }
  return { filterCount: Object.keys(filters).length, filters, refs, boxes };
}, slide);
console.log(JSON.stringify(data, null, 2));
await page.close(); await browser.close(); server.close();
