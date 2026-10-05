import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import fs from 'node:fs';
import path from 'node:path';
const { chromium } = pw;
const ROOT = '/Users/jiamao/project/github/pptx-parser';
const OUT = path.join(ROOT, 'test', 'cmp');
const W = 1000, H = 562;

const browser = await chromium.launch({
  executablePath: '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell'
});
const page = await browser.newPage();
await page.setContent('<canvas id="a"></canvas><canvas id="b"></canvas>');

async function loadAndDiff(i) {
  const ep = path.join(OUT, `e${i}.png`);
  const pp = path.join(OUT, `p${i}.png`);
  if (!fs.existsSync(ep) || !fs.existsSync(pp)) return null;
  const ea = fs.readFileSync(ep).toString('base64');
  const pa = fs.readFileSync(pp).toString('base64');
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
    const d1 = await draw(ea);
    const d2 = await draw(pa);
    let diff = 0, n = W * H;
    const lum = (r, g, b) => 0.299 * r + 0.587 * g + 0.114 * b;
    for (let p = 0; p < d1.length; p += 4) {
      const l1 = lum(d1[p], d1[p+1], d1[p+2]);
      const l2 = lum(d2[p], d2[p+1], d2[p+2]);
      diff += Math.abs(l1 - l2);
    }
    const mae = diff / n; // 0..255
    return { mae: Math.round(mae * 100) / 100 };
  }, { ea, pa, W, H });
  return res.mae;
}

const results = [];
for (let i = 0; i < 12; i++) {
  const mae = await loadAndDiff(i);
  results.push(`slide ${String(i).padStart(2)}: ${mae === null ? 'N/A' : mae}`);
}
console.log(results.join('\n'));
await browser.close();
