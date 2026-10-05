import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
const { chromium } = pw;
const browser = await chromium.launch({ executablePath: '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell' });
const page = await browser.newPage();
await page.setContent(`<html lang="zh-CN"><body><span id=a style="font-family:Calibri;font-size:40px">Need more info?</span><span id=c lang="en" style="font-family:Calibri;font-size:40px">Need more info?</span></body></html>`);
const r = await page.evaluate(() => {
  const w = (id) => document.getElementById(id).getBoundingClientRect().width;
  const ff = (id) => { const c = document.getElementById(id); const cv = document.createElement('canvas').getContext('2d'); cv.font = '40px Calibri'; return w(id); };
  return { zh: w('a'), en: w('c') };
});
console.log(r);
await browser.close();
