import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
import fs from 'node:fs';
import path from 'path';
const ROOT = '/Users/jiamao/project/github/pptx-parser';
const CHROME = '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell';
const arg = 'examples/企业微信应用介绍.pptx';
const server = http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/') p = '/examples/editor/index.html';
  const fp = path.join(ROOT, p);
  if (!fp.startsWith(ROOT)) { res.writeHead(403); res.end(); return; }
  fs.readFile(fp, (err, data) => {
    if (err) { res.writeHead(404); res.end(); return; }
    res.writeHead(200, { 'Content-Type': fp.endsWith('.js') ? 'application/javascript' : fp.endsWith('.css') ? 'text/css' : 'text/html' });
    res.end(data);
  });
});
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const BASE = 'http://127.0.0.1:' + server.address().port;
const browser = await pw.chromium.launch({ executablePath: CHROME });
const page = await browser.newPage({ viewport: { width: 1700, height: 1100 } });
page.on('console', (m) => console.log('PAGE>', m.type(), m.text()));
page.on('pageerror', (e) => console.log('PAGEERR>', e.message));
await page.goto(BASE + '/examples/editor/index.html');
await page.waitForTimeout(2500);
await page.evaluate(async (url) => {
  const resp = await fetch(url);
  const buf = await resp.arrayBuffer();
  const { pptxToStandard } = await import('/dist/ppt-parser.browser.js');
  const { docFromPptx } = await import('/examples/editor/src/model.js');
  const { store } = await import('/examples/editor/src/store.js');
  store.setDoc(docFromPptx(await pptxToStandard(buf)), { noHistory: true });
  store.fitted = true;
  store.setSlide(0);
}, '/' + arg);
await page.waitForTimeout(5000);
const d = await page.evaluate(() => {
  const fr = document.querySelector('#frame.slide-frame');
  if (!fr) return { err: 'no frame' };
  const cs = getComputedStyle(fr);
  const els = [...fr.querySelectorAll('.el')];
  return {
    bg: cs.backgroundColor, bgImg: cs.backgroundImage.length,
    nEls: els.length,
    texts: els.map((e) => {
      const t = e.querySelector('.tb-body');
      const firstSpan = e.querySelector('span');
      const rects = firstSpan ? firstSpan.getClientRects().length : 0;
      const cs2 = firstSpan ? getComputedStyle(firstSpan) : null;
      // 行数：用 tb-body 的 scrollHeight / lineHeight 估算，或 span 的 client rects
      const lines = (() => {
        const sp = firstSpan;
        if (!sp) return 1;
        const lh = parseFloat(getComputedStyle(sp).lineHeight) || 0;
        return lh > 0 ? Math.round(sp.getBoundingClientRect().height / lh) : sp.getClientRects().length;
      })();
      return {
        text: (t ? t.textContent : e.textContent || '').slice(0, 24),
        color: cs2 ? cs2.color : null,
        fontSize: cs2 ? cs2.fontSize : null,
        lines,
        align: t ? getComputedStyle(t.firstElementChild || t).textAlign : null
      };
    })
  };
});
console.log(JSON.stringify(d, null, 1));
await (await page.$$('#frame.slide-frame'))[0].screenshot({ path: '/tmp/wx_editor.png' });
await browser.close(); server.close();
