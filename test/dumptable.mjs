import fs from 'node:fs';
import pw from '/Users/jiamao/.npm/_npx/e41f203b7505f1fb/node_modules/playwright/index.js';
import http from 'node:http';
const ROOT = '/Users/jiamao/project/github/pptx-parser';
const fileBuffer = fs.readFileSync(ROOT + '/examples/Sample_12.pptx');
const MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const srv = http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/') p = '/examples/editor/index.html';
  fs.readFile(ROOT + p, (e, d) => { if (e){res.writeHead(404);res.end();return;} res.writeHead(200,{'Content-Type':p.endsWith('.js')?'application/javascript':p.endsWith('.css')?'text/css':'text/html'}); res.end(d); });
}).listen(8782,'127.0.0.1');
const { chromium } = pw;
const browser = await chromium.launch({ executablePath: '/Users/jiamao/Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell' });
const page = await browser.newPage();
await page.goto('http://127.0.0.1:8782/examples/editor/index.html');
await page.waitForTimeout(400);
await page.evaluate(({bufArr,mime})=>new Promise((resolve,reject)=>{const i=document.getElementById('hiddenFile');i.onchange=async()=>{try{const {importPptxFile}=await import('./src/io.js');await importPptxFile(i.files[0]);resolve('ok');}catch(e){reject(e);}};const b=new Blob([new Uint8Array(bufArr)],{type:mime});const dt=new DataTransfer();dt.items.add(new File([b],'Sample_12.pptx'));i.files=dt.files;i.dispatchEvent(new Event('change',{bubbles:true}));}),{bufArr:Array.from(fileBuffer),mime:MIME});
await page.waitForTimeout(1500);
const dump = await page.evaluate(() => {
  const { store } = window.__editorStore || {};
  return null;
});
// 改用 dom 查询表格单元格颜色
const cells = await page.evaluate(async () => {
  const { store } = await import('./src/store.js');
  const out = {};
  [5,6].forEach((idx)=>{
    const tbl = store.doc.slides[idx].elements.find(e=>e.type==='table');
    if (!tbl) { out[idx] = 'no table'; return; }
    out[idx] = (tbl.rows||[]).map(r=>(r.cells||[]).map(c=>c.fill||null));
  });
  return out;
});
console.log(JSON.stringify(cells));
await browser.close(); srv.close();
