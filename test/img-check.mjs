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
await page.goto(`${BASE}/examples/index.html`);
await page.waitForTimeout(600);
await (await page.$('#uploadFileInput')).setInputFiles({ name: 'x.pptx', mimeType: MIME, buffer: buf });
await page.waitForFunction(() => document.querySelectorAll('#result .slide').length >= 1, null, { timeout: 30000 }).catch(() => {});
await page.waitForTimeout(6000);
// 不强制缩放：在预览端自然 fit 状态下读取
const preview = await page.evaluate(() => {
  const slides = [...document.querySelectorAll('#result .slide')];
  const s = slides[16];
  const sb = s.getBoundingClientRect();
  const sw = s.offsetWidth, sh = s.offsetHeight;
  const out = { sw, sh, imgs: [] };
  s.querySelectorAll('img').forEach((img) => {
    const r = img.getBoundingClientRect();
    const wrap = img.parentElement;
    out.imgs.push({
      x: +(r.left - sb.left).toFixed(1), y: +(r.top - sb.top).toFixed(1),
      w: +r.width.toFixed(1), h: +r.height.toFixed(1),
      style: (wrap && wrap.getAttribute('style') || '').replace(/\s+/g, ' ')
    });
  });
  return out;
});
console.log('PREVIEW slide17 native', preview.sw, preview.sh);
console.log('PREVIEW imgs:', JSON.stringify(preview.imgs, null, 0));

// 编辑器端（标准坐标，EMU→px）
const { pptxToStandard } = await import(path.join(ROOT, 'dist/ppt-parser.cjs'));
const doc = await pptxToStandard(buf);
const ed = doc.slides[16].elements.filter((e) => e.type === 'image')
  .map((e) => ({ x: Math.round(e.x), y: Math.round(e.y), w: Math.round(e.width), h: Math.round(e.height) }));
console.log('EDITOR slide17 native', doc.slideSize.width, doc.slideSize.height);
console.log('EDITOR imgs:', JSON.stringify(ed));
await browser.close();
server.close();
