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
  if (p === '/') p = '/examples/editor/_harness.html';
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
const page = await browser.newPage();
const fileParam = encodeURIComponent('/examples/fefeding金腾科技前端开发通道答辨.pptx');
await page.goto(`${BASE}/examples/editor/_harness.html?file=${fileParam}`);
await page.waitForFunction(() => document.querySelectorAll('.slide-frame').length >= 21, null, { timeout: 30000 }).catch(() => {});
await page.waitForTimeout(1500);

const info = await page.evaluate(() => {
  const frames = document.querySelectorAll('#root .slide-frame');
  const n = frames.length;
  const f = frames[6]; // slide 7
  if (!f) return { error: 'no frame', n };
  const boxes = [...f.querySelectorAll('.el-text')].filter((e) => (e.textContent || '').includes('MIXIN'));
  return { n, boxes: boxes.slice(0, 1).map((e) => {
    const sb = f.getBoundingClientRect();
    const r = e.getBoundingClientRect();
    // 找含 MIXIN 且后代最少的叶子元素（真正的文本 run span）
    let leaf = e;
    let cur = e;
    while (cur && cur.children.length) {
      const child = [...cur.children].find((c) => (c.textContent || '').includes('MIXIN'));
      if (!child) break;
      leaf = child; cur = child;
    }
    const lcs = getComputedStyle(leaf);
    const pcs = leaf.parentElement ? getComputedStyle(leaf.parentElement) : null;
    return {
      text: (e.textContent || '').slice(0, 20),
      x: +(r.left - sb.left).toFixed(1), y: +(r.top - sb.top).toFixed(1),
      w: +r.width.toFixed(1), h: +r.height.toFixed(1),
      childCount: e.children.length,
      divBackgrounds: [...e.querySelectorAll('div')].map((d) => getComputedStyle(d).background),
      leafTag: leaf.tagName + '.' + (leaf.className||''),
      leafFontSize: lcs.fontSize, leafLineHeight: lcs.lineHeight, leafFontFamily: lcs.fontFamily.slice(0,40),
      leafInline: leaf.getAttribute('style'),
      parentFontSize: pcs ? pcs.fontSize : '(none)'
    };
  }) };
});
console.log(JSON.stringify(info, null, 1));
await browser.close();
server.close();
