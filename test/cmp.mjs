import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
const { chromium } = pw;

const ROOT = '/Users/jiamao/project/github/pptx-parser';
const OUT = path.join(ROOT, 'test', 'cmp');
fs.mkdirSync(OUT, { recursive: true });
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
}).listen(8770, '127.0.0.1');

const browser = await chromium.launch({
  executablePath: '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell'
});

/* ================= 预览端 ================= */
{
  const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
  // 拦截示例首页默认异步加载的远程示例 PPTX，避免与上传文件竞态导致重复渲染
  await page.route('**/pptx.js.org/**', (route) => route.abort());
  await page.goto('http://127.0.0.1:8770/examples/index.html');
  await page.waitForTimeout(800);
  const input = await page.$('#uploadFileInput');
  await input.setInputFiles({ name: 'Sample_12.pptx', mimeType: MIME, buffer: fileBuffer });
  await page.waitForTimeout(6000);
  // 等待上传渲染完成（默认示例已被拦截，应恰好 12 张）
  await page.waitForFunction(() => document.querySelectorAll('#result .slide').length >= 12, null, { timeout: 15000 }).catch(() => {});
  const slides = await page.$$('.slide');
  console.log('preview slides:', slides.length);
  // 等待预览端图表（ECharts canvas）渲染完成
  await page.waitForFunction(() => {
    const frames = document.querySelectorAll('#result .slide');
    return Array.from(frames).every((f) => {
      const charts = f.querySelectorAll('.el-chart, canvas');
      if (!charts.length) return true;
      return Array.from(charts).every((c) => c.querySelector('canvas') || c.querySelector('svg') || c.tagName === 'CANVAS');
    });
  }, null, { timeout: 8000 }).catch(() => {});
  await page.waitForTimeout(500);
  for (let i = 0; i < slides.length; i++) {
    await slides[i].screenshot({ path: path.join(OUT, `p${i}.png`) });
  }
  await page.close();
}

/* ================= 编辑器 ================= */
{
  const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
  page.on('pageerror', (e) => console.log('EDITOR ERROR:', e.message));
  await page.goto('http://127.0.0.1:8770/examples/editor/index.html');
  await page.waitForTimeout(500);
  await page.evaluate(({ bufArr, mime }) => {
    return new Promise((resolve, reject) => {
      const input = document.getElementById('hiddenFile');
      input.onchange = async () => {
        try {
          const { importPptxFile } = await import('./src/io.js');
          await importPptxFile(input.files[0]);
          resolve('ok');
        } catch (e) { reject(e); }
      };
      const blob = new Blob([new Uint8Array(bufArr)], { type: mime });
      const dt = new DataTransfer();
      dt.items.add(new File([blob], 'Sample_12.pptx'));
      input.files = dt.files;
      input.dispatchEvent(new Event('change', { bubbles: true }));
    });
  }, { bufArr: Array.from(fileBuffer), mime: MIME });
  await page.waitForTimeout(2000);
  const n = await page.evaluate(async () => {
    const { store } = await import('./src/store.js');
    return store.doc.slides.length;
  });
  console.log('editor slides:', n);
  for (let i = 0; i < n; i++) {
    await page.evaluate(async (idx) => {
      const { store } = await import('./src/store.js');
      store.setSlide(idx, { force: true });
    }, i);
    // 等字体就绪 + ECharts 图表异步 init（rAF）完成，避免截图截到未渲染态
    await page.evaluate(() => document.fonts ? document.fonts.ready : Promise.resolve());
    await page.waitForFunction(() => {
      const frame = document.querySelector('#stage .slide-frame');
      if (!frame) return false;
      const charts = frame.querySelectorAll('.el-chart');
      if (!charts.length) return true;
      return Array.from(charts).every((c) => c.querySelector('canvas') || c.querySelector('svg'));
    }, null, { timeout: 5000 }).catch(() => {});
    await page.waitForTimeout(500);
    const frame = await page.$('#stage .slide-frame');
    if (frame) await frame.screenshot({ path: path.join(OUT, `e${i}.png`) });
  }
  await page.close();
}

await browser.close();
server.close();
console.log('done');
