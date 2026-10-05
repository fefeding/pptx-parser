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
  if (p === '/') p = '/examples/editor/index.html';
  const fp = path.join(ROOT, p);
  if (!fp.startsWith(ROOT)) { res.writeHead(403); res.end(); return; }
  fs.readFile(fp, (err, data) => {
    if (err) { res.writeHead(404); res.end(); return; }
    res.writeHead(200, { 'Content-Type': fp.endsWith('.js') ? 'application/javascript' : fp.endsWith('.css') ? 'text/css' : 'text/html', 'Access-Control-Allow-Origin': '*' }); res.end(data);
  });
}).listen(8772, '127.0.0.1');
const browser = await chromium.launch({ executablePath: '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell' });
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
await page.goto('http://127.0.0.1:8772/examples/editor/index.html');
await page.waitForTimeout(500);
await page.evaluate(({ bufArr, mime }) => {
  return new Promise((resolve, reject) => {
    const input = document.getElementById('hiddenFile');
    input.onchange = async () => {
      try { const { importPptxFile } = await import('./src/io.js'); await importPptxFile(input.files[0]); resolve('ok'); } catch (e) { reject(e); }
    };
    const blob = new Blob([new Uint8Array(bufArr)], { type: mime });
    const dt = new DataTransfer(); dt.items.add(new File([blob], 'Sample_12.pptx'));
    input.files = dt.files; input.dispatchEvent(new Event('change', { bubbles: true }));
  });
}, { bufArr: Array.from(fileBuffer), mime: MIME });
await page.waitForTimeout(1500);
const dump = await page.evaluate(async () => {
  const { store } = await import('./src/store.js');
  const { renderCanvas } = await import('./src/interact.js');
  store.setSlide(4, { force: true });
  renderCanvas();
  await new Promise((r) => setTimeout(r, 200));
  const out = [];
  document.querySelectorAll('#stage .el-text span').forEach((sp) => {
    if (sp.textContent.trim()) out.push({ t: sp.textContent.slice(0, 18), ff: getComputedStyle(sp).fontFamily, size: getComputedStyle(sp).fontSize });
  });
  return out.slice(0, 8);
});
console.log(JSON.stringify(dump, null, 1));
const frameEl = await page.$('#stage .slide-frame');
await frameEl.screenshot({ path: `${ROOT}/test/cmp/e4-check.png` });
await browser.close(); server.close();
