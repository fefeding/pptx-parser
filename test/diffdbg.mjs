import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import fs from 'node:fs';
const { chromium } = pw;
const browser = await chromium.launch({
  executablePath: '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell'
});
const page = await browser.newPage();
await page.setContent('<canvas id="a"></canvas>');
const W = 1000, H = 562;
const i = process.argv[2] || '4';
const ea = fs.readFileSync(`/Users/jiamao/project/github/pptx-parser/test/cmp/e${i}.png`).toString('base64');
const pa = fs.readFileSync(`/Users/jiamao/project/github/pptx-parser/test/cmp/p${i}.png`).toString('base64');
const res = await page.evaluate(async ({ ea, pa, W, H }) => {
  const draw = (b64) => new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const c = document.createElement('canvas');
      c.width = W; c.height = H;
      const ctx = c.getContext('2d');
      ctx.drawImage(img, 0, 0, W, H);
      resolve(ctx.getImageData(0, 0, W, H).data);
    };
    img.src = 'data:image/png;base64,' + b64;
  });
  const d1 = await draw(ea), d2 = await draw(pa);
  const lum = (r, g, b) => 0.299 * r + 0.587 * g + 0.114 * b;
  // 8x8 网格分区 MAE
  const grid = [];
  const gw = 8, gh = 8;
  for (let gy = 0; gy < gh; gy++) {
    const row = [];
    for (let gx = 0; gx < gw; gx++) {
      let diff = 0, n = 0;
      const x0 = Math.floor(gx * W / gw), x1 = Math.floor((gx + 1) * W / gw);
      const y0 = Math.floor(gy * H / gh), y1 = Math.floor((gy + 1) * H / gh);
      for (let y = y0; y < y1; y += 2) for (let x = x0; x < x1; x += 2) {
        const p = (y * W + x) * 4;
        diff += Math.abs(lum(d1[p], d1[p+1], d1[p+2]) - lum(d2[p], d2[p+1], d2[p+2]));
        n++;
      }
      row.push(Math.round(diff / n));
    }
    grid.push(row.join('\t'));
  }
  let total = 0, cnt = 0;
  for (let p = 0; p < d1.length; p += 4) { total += Math.abs(lum(d1[p],d1[p+1],d1[p+2]) - lum(d2[p],d2[p+1],d2[p+2])); cnt++; }
  return { grid: grid.join('\n'), mae: Math.round(total / cnt * 100) / 100 };
}, { ea, pa, W, H });
console.log('MAE', res.mae);
console.log(res.grid);
await browser.close();
