import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
import fs from 'node:fs';
import path from 'path';
const ROOT = '/Users/jiamao/project/github/pptx-parser';
const fileBuffer = fs.readFileSync(path.join(ROOT, 'examples/Sample_12.pptx'));
const MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const server = http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/') p = '/examples/index.html';
  fs.readFile(path.join(ROOT, p), (e, d) => { if (e){res.writeHead(404);res.end();return;} res.writeHead(200,{'Content-Type':p.endsWith('.js')?'application/javascript':'text/html','Access-Control-Allow-Origin':'*'}); res.end(d); });
}).listen(8771,'127.0.0.1');
const { chromium } = pw;
const browser = await chromium.launch({ executablePath: '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell' });

async function getPreviewTables() {
  const page = await browser.newPage();
  await page.goto('http://127.0.0.1:8771/examples/index.html');
  await page.waitForTimeout(800);
  const input = await page.$('#uploadFileInput');
  await input.setInputFiles({ name: 'Sample_12.pptx', mimeType: MIME, buffer: fileBuffer });
  await input.evaluate((el) => el.dispatchEvent(new Event('change', { bubbles: true })));
  await page.waitForTimeout(6000);
  const out = await page.evaluate(() => {
    const res = [];
    document.querySelectorAll('.slide').forEach((s, i) => {
      const tbl = s.querySelector('table');
      if (!tbl) return;
      const m = tbl.outerHTML.match(/background(?:-color)?:\s*(#[0-9A-Fa-f]{3,8})/g) || [];
      res.push([i, [...new Set(m.map(x=>x.split('#')[1].toUpperCase()))].slice(0,6)]);
    });
    return res;
  });
  await page.close();
  return out;
}
async function getEditorTables() {
  const page = await browser.newPage();
  await page.goto('http://127.0.0.1:8771/examples/editor/index.html');
  await page.waitForTimeout(500);
  await page.evaluate(({bufArr,mime})=>new Promise((resolve,reject)=>{const i=document.getElementById('hiddenFile');i.onchange=async()=>{try{const {importPptxFile}=await import('./src/io.js');await importPptxFile(i.files[0]);resolve('ok');}catch(e){reject(e);}};const b=new Blob([new Uint8Array(bufArr)],{type:mime});const dt=new DataTransfer();dt.items.add(new File([b],'Sample_12.pptx'));i.files=dt.files;i.dispatchEvent(new Event('change',{bubbles:true}));}),{bufArr:Array.from(fileBuffer),mime:MIME});
  await page.waitForTimeout(2000);
  const out = await page.evaluate(async () => {
    const { store } = await import('./src/store.js');
    const res = [];
    store.doc.slides.forEach((sl, i) => {
      const tbl = sl.elements.find(e=>e.type==='table');
      if (!tbl) return;
      res.push([i, (tbl.rows||[]).flatMap(r=>(r.cells||[]).map(c=>c.fill||null)).filter(Boolean).slice(0,6)]);
    });
    return res;
  });
  await page.close();
  return out;
}
const pv = await getPreviewTables();
const ed = await getEditorTables();
console.log('PREVIEW table slide fills:');
pv.forEach(r=>console.log('  slide', r[0], r[1].join(' ')));
console.log('EDITOR table slide fills:');
ed.forEach(r=>console.log('  slide', r[0], r[1].join(' ')));
await browser.close(); server.close();
