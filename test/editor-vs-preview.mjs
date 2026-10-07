/**
 * 编辑器画布 vs 预览端 逐页像素对比。
 *
 * 用法：node test/editor-vs-preview.mjs [pptx相对仓库根的路径]
 * 默认：examples/fefeding金腾科技前端开发通道答辨.pptx
 *
 * 输出：test/evp/ 下 e<N>.png（编辑器画布）与 p<N>.png（预览端），并打印逐页差异比例。
 */
import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const ROOT = '/Users/jiamao/project/github/pptx-parser';
const OUT = path.join(ROOT, 'test', 'evp');
const MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const PIXEL_TOLERANCE = 24;
const CHROME = '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell';

fs.mkdirSync(OUT, { recursive: true });

const arg = process.argv[2] || 'examples/fefeding金腾科技前端开发通道答辨.pptx';
const file = path.resolve(ROOT, arg);
if (!fs.existsSync(file)) { console.error('文件不存在:', file); process.exit(1); }

const server = http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/') p = '/examples/index.html';
  const fp = path.join(ROOT, p);
  if (!fp.startsWith(ROOT)) { res.writeHead(403); res.end(); return; }
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

const buf = fs.readFileSync(file);

/** 预览端：上传文件，逐页截图 .slide */
async function renderPreview() {
  const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
  const errs = [];
  page.on('pageerror', (e) => errs.push(e.message));
  await page.route('**/pptx.js.org/**', (r) => r.abort());
  await page.goto(`${BASE}/examples/index.html`);
  await page.waitForTimeout(600);
  await (await page.$('#uploadFileInput')).setInputFiles({ name: path.basename(file), mimeType: MIME, buffer: buf });
  await page.waitForFunction(() => document.querySelectorAll('#result .slide').length >= 1, null, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(6000);
  // 强制预览端按原生尺寸（100%，约 1280px，与编辑器同一 96px/in 比例）渲染，
  // 否则预览端 fitToWidth 会把 slide 缩到 ~1104px，文字分辨率低于编辑器，
  // 缩放比对时产生大量“模糊伪差”。
  await page.evaluate(() => {
    document.querySelectorAll('#result .slide-scaler').forEach((sc) => {
      const inner = sc.firstElementChild;
      if (!inner) return;
      sc.style.width = inner.offsetWidth + 'px';
      sc.style.height = inner.offsetHeight + 'px';
      inner.style.transform = 'scale(1)';
    });
  });
  // 等待 transform: scale(1) 的布局回流完成，否则截到的是旧的 0.863 缩放版本
  await page.waitForTimeout(500);
  const slides = await page.$$('#result .slide');
  const urls = [];
  for (let i = 0; i < slides.length; i++) {
    const b = await slides[i].screenshot({ path: path.join(OUT, `p${i}.png`) });
    urls.push('data:image/png;base64,' + b.toString('base64'));
  }
  console.log(`预览: ${slides.length} 页`, errs.length ? '错误: ' + errs.slice(0, 3).join(' | ') : '');
  await page.close();
  return urls;
}

/** 编辑器：用渲染 harness 渲染所有幻灯片，逐页截图 .slide-frame */
async function renderEditor() {
  const page = await browser.newPage({ viewport: { width: 2200, height: 1200 } });
  const errs = [];
  page.on('pageerror', (e) => errs.push(e.message));
  await page.route('**/pptx.js.org/**', (r) => r.abort());
  const fileUrl = '/' + arg.split('\\').join('/');
  await page.goto(`${BASE}/examples/editor/_harness.html?file=${encodeURIComponent(fileUrl)}`);
  await page.waitForFunction(() => document.querySelectorAll('#root .slide-frame').length >= 1, null, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(4000);
  const frames = await page.$$('#root .slide-frame');
  const urls = [];
  for (let i = 0; i < frames.length; i++) {
    const b = await frames[i].screenshot({ path: path.join(OUT, `e${i}.png`) });
    urls.push('data:image/png;base64,' + b.toString('base64'));
  }
  console.log(`编辑器: ${frames.length} 页`, errs.length ? '错误: ' + errs.slice(0, 3).join(' | ') : '');
  await page.close();
  return urls;
}

const p = await renderPreview();
const e = await renderEditor();

const diffPage = await browser.newPage();
await diffPage.goto('about:blank');
async function diffRatio(ua, ub) {
  return await diffPage.evaluate(async ({ ua, ub, tol }) => {
    const load = (u) => new Promise((res) => { const img = new Image(); img.onload = () => res(img); img.src = u; });
    const [ia, ib] = await Promise.all([load(ua), load(ub)]);
    // 归一化到预览端尺寸：编辑器按自身 px 比例渲染（1279×720），预览端按视口自适应（~1104×622），
    // 二者尺度不同，直接按 min 裁剪会产生海量伪差异。先把编辑器缩放对齐到预览尺寸再比对。
    const w = ib.width, h = ib.height;
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    const ctx = c.getContext('2d');
    ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(ia, 0, 0, w, h); const da = ctx.getImageData(0, 0, w, h).data;
    ctx.drawImage(ib, 0, 0); const db = ctx.getImageData(0, 0, w, h).data;
    let diff = 0;
    for (let i = 0; i < da.length; i += 4) {
      if (Math.abs(da[i] - db[i]) > tol || Math.abs(da[i + 1] - db[i + 1]) > tol || Math.abs(da[i + 2] - db[i + 2]) > tol) diff++;
    }
    return diff / (w * h);
  }, { ua, ub, tol: PIXEL_TOLERANCE });
}

console.log('\n=== 编辑器 vs 预览 逐页差异比例 (0=完全相同) ===');
const stats = [];
for (let i = 0; i < Math.min(e.length, p.length); i++) {
  const d = await diffRatio(e[i], p[i]);
  console.log(`slide ${i + 1}: ${(d * 100).toFixed(1)}%`);
  stats.push({ i: i + 1, d });
}
stats.sort((x, y) => y.d - x.d);
console.log('\n差异最大:', stats.slice(0, 8).map((w) => `slide${w.i}(${(w.d * 100).toFixed(1)}%)`).join(', '));

await browser.close();
server.close();
