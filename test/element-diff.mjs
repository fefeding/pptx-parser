/**
 * 逐元素差异定位：把每一页的差异按元素包围盒聚合，找出贡献最大的元素。
 * 用法：node test/element-diff.mjs [pptx]
 */
import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = '/Users/jiamao/project/github/pptx-parser';
const OUT = path.join(ROOT, 'test', 'evp');
const arg = process.argv[2] || 'examples/fefeding金腾科技前端开发通道答辨.pptx';
const file = path.resolve(ROOT, arg);
const CHROME = '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell';

const { pptxToStandard } = await import(path.join(ROOT, 'dist/ppt-parser.cjs'));
const doc = await pptxToStandard(fs.readFileSync(file));
const SW = doc.slideSize.width, SH = doc.slideSize.height;

const browser = await pw.chromium.launch({ executablePath: CHROME });
const page = await browser.newPage();
await page.goto('about:blank');

function b64(p) { return 'data:image/png;base64,' + fs.readFileSync(p).toString('base64'); }

async function analyze(idx) {
  const eUrl = b64(path.join(OUT, `e${idx}.png`));
  const pUrl = b64(path.join(OUT, `p${idx}.png`));
  const els = (doc.slides[idx].elements || []).map((e) => ({
    type: e.type, shapeType: e.shapeType || '', name: e.name || '',
    x: e.x || 0, y: e.y || 0, w: e.width || 0, h: e.height || 0,
    text: (e.paragraphs && e.paragraphs[0] && e.paragraphs[0].runs && e.paragraphs[0].runs[0] && e.paragraphs[0].runs[0].text || (typeof e.text === 'string' ? e.text : '')).slice(0, 16)
  }));
  return await page.evaluate(async ({ eUrl, pUrl, els, SW, SH }) => {
    const load = (u) => new Promise((r) => { const i = new Image(); i.onload = () => r(i); i.src = u; });
    const [ie, ip] = await Promise.all([load(eUrl), load(pUrl)]);
    const pw = ip.width, ph = ip.height;
    const sx = pw / SW, sy = ph / SH; // 预览相对标准坐标的缩放
    const ctxE = document.createElement('canvas').getContext('2d');
    const ctxP = document.createElement('canvas').getContext('2d');
    const tol = 24;
    const res = [];
    for (const el of els) {
      const ex = Math.max(0, el.x), ey = Math.max(0, el.y);
      const ew = el.w, eh = el.h;
      const px = Math.max(0, ex * sx), py = Math.max(0, ey * sy);
      const pw2 = ew * sx, ph2 = eh * sy;
      const cw = Math.min(Math.floor(ew), ie.width - Math.floor(ex));
      const ch = Math.min(Math.floor(eh), ie.height - Math.floor(ey));
      const cw2 = Math.min(Math.floor(pw2), pw - Math.floor(px));
      const ch2 = Math.min(Math.floor(ph2), ph - Math.floor(py));
      if (cw < 2 || ch < 2 || cw2 < 2 || ch2 < 2) { res.push({ ...el, d: 0, area: 0 }); continue; }
      ctxE.canvas.width = cw; ctxE.canvas.height = ch;
      ctxE.drawImage(ie, ex, ey, cw, ch, 0, 0, cw, ch);
      const de = ctxE.getImageData(0, 0, cw, ch).data;
      // 预览子区缩放到与编辑器子区同尺寸再比
      const tmp = document.createElement('canvas'); tmp.width = cw; tmp.height = ch;
      const tctx = tmp.getContext('2d'); tctx.imageSmoothingEnabled = true; tctx.imageSmoothingQuality = 'high';
      tctx.drawImage(ip, px, py, cw2, ch2, 0, 0, cw, ch);
      const dp = tctx.getImageData(0, 0, cw, ch).data;
      let diff = 0;
      for (let i = 0; i < de.length; i += 4) {
        if (Math.abs(de[i] - dp[i]) > tol || Math.abs(de[i + 1] - dp[i + 1]) > tol || Math.abs(de[i + 2] - dp[i + 2]) > tol) diff++;
      }
      res.push({ ...el, d: diff, area: cw * ch });
    }
    return res;
  }, { eUrl, pUrl, els, SW, SH });
}

for (let i = 0; i < doc.slides.length; i++) {
  const r = await analyze(i);
  r.sort((a, b) => b.d - a.d);
  const top = r.filter((x) => x.d > 0).slice(0, 6);
  const total = r.reduce((s, x) => s + x.d, 0);
  console.log(`\nslide ${i + 1} 总差异像素 ${total}`);
  top.forEach((x, k) => console.log(`  ${k + 1}. ${x.type}/${x.shapeType} "${x.text}" @(${Math.round(x.x)},${Math.round(x.y)},${Math.round(x.w)}x${Math.round(x.h)}) diff=${x.d}`));
}
await browser.close();
