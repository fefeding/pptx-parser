/**
 * 导出保真度对比工具。
 *
 * 用预览端（examples/index.html）渲染「原始 PPTX」与「导出 PPTX」，逐页截图并计算像素差异比例。
 *
 * 用法：
 *   node test/roundtrip-cmp.mjs                       # 程序化 round-trip：解析→JSON→序列化，再与原始对比
 *   node test/roundtrip-cmp.mjs a.pptx b.pptx         # 直接对比两个已有文件
 *
 * 截图输出到 test/rt/（a<N>.png = 原始，b<N>.png = 导出）。
 */
import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const ROOT = '/Users/jiamao/project/github/pptx-parser';
const OUT = path.join(ROOT, 'test', 'rt');
const MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
/** 像素差分阈值：任一通道差值超过该值即计为差异像素（0~255） */
const PIXEL_TOLERANCE = 24;

fs.mkdirSync(OUT, { recursive: true });

const argv = process.argv.slice(2);
const roundTrip = argv.length === 0;
// 中间产物放系统临时目录：样例动辄 20MB+，不应写进仓库
const tmpPptx = path.join(os.tmpdir(), 'pptx-parser-roundtrip.pptx');
let fileA;
let fileB;

if (roundTrip) {
  // 程序化 round-trip：pptxToStandard → jsonToPptx
  const { pptxToStandard, jsonToPptx } = await import(path.join(ROOT, 'dist/ppt-parser.esm.js'));
  fileA = path.join(ROOT, 'examples/Sample_12.pptx');
  const doc = await pptxToStandard(fs.readFileSync(fileA));
  const out = await jsonToPptx(doc, { outputType: 'uint8array' });
  fs.writeFileSync(tmpPptx, Buffer.from(out));
  fileB = tmpPptx;
  console.log('round-trip 生成:', tmpPptx, out.length, 'bytes');
} else {
  [fileA, fileB] = argv.map((f) => path.resolve(ROOT, f));
}

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
// 监听临时端口：避免上次残留进程占用固定端口导致 EADDRINUSE
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const PORT = server.address().port;
const BASE = `http://127.0.0.1:${PORT}`;

const browser = await pw.chromium.launch({
  executablePath: '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell'
});

/** 渲染一个 PPTX，逐页截图落盘并返回 data URL 列表 */
async function renderAll(file, tag) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
  const errs = [];
  page.on('pageerror', (e) => errs.push(e.message));
  await page.route('**/pptx.js.org/**', (r) => r.abort());
  await page.goto(`${BASE}/examples/index.html`);
  await page.waitForTimeout(800);
  await (await page.$('#uploadFileInput')).setInputFiles({
    name: path.basename(file), mimeType: MIME, buffer: fs.readFileSync(file)
  });
  await page.waitForFunction(() => document.querySelectorAll('#result .slide').length >= 1, null, { timeout: 20000 }).catch(() => {});
  await page.waitForTimeout(6000);
  const slides = await page.$$('.slide');
  const urls = [];
  for (let i = 0; i < slides.length; i++) {
    const buf = await slides[i].screenshot({ path: path.join(OUT, `${tag}${i}.png`) });
    urls.push('data:image/png;base64,' + buf.toString('base64'));
  }
  console.log(`${tag}: ${slides.length} 页`, errs.length ? '错误: ' + errs.slice(0, 3).join(' | ') : '');
  await page.close();
  return urls;
}

const a = await renderAll(fileA, 'a');
const b = await renderAll(fileB, 'b');
console.log('原始', a.length, '导出', b.length);

/** 在浏览器内对两张图做逐像素差分，返回差异像素占比（0~1） */
const diffPage = await browser.newPage();
await diffPage.goto('about:blank');
async function diffRatio(urlA, urlB) {
  return await diffPage.evaluate(async ({ ua, ub, tol }) => {
    const load = (u) => new Promise((res) => {
      const img = new Image();
      img.onload = () => res(img);
      img.src = u;
    });
    const [ia, ib] = await Promise.all([load(ua), load(ub)]);
    const w = Math.min(ia.width, ib.width);
    const h = Math.min(ia.height, ib.height);
    const c = document.createElement('canvas');
    c.width = w; c.height = h;
    const ctx = c.getContext('2d');
    ctx.drawImage(ia, 0, 0);
    const da = ctx.getImageData(0, 0, w, h).data;
    ctx.drawImage(ib, 0, 0);
    const db = ctx.getImageData(0, 0, w, h).data;
    let diff = 0;
    for (let i = 0; i < da.length; i += 4) {
      if (Math.abs(da[i] - db[i]) > tol || Math.abs(da[i + 1] - db[i + 1]) > tol || Math.abs(da[i + 2] - db[i + 2]) > tol) diff++;
    }
    return diff / (w * h);
  }, { ua: urlA, ub: urlB, tol: PIXEL_TOLERANCE });
}

console.log('\n=== 逐页差异比例 (0=完全相同) ===');
const stats = [];
for (let i = 0; i < Math.min(a.length, b.length); i++) {
  const d = await diffRatio(a[i], b[i]);
  console.log(`slide ${i + 1}: ${(d * 100).toFixed(1)}%`);
  stats.push({ i: i + 1, d });
}
stats.sort((x, y) => y.d - x.d);
console.log('\n差异最大:', stats.slice(0, 4).map((w) => `slide${w.i}(${(w.d * 100).toFixed(1)}%)`).join(', '));

await browser.close();
server.close();
