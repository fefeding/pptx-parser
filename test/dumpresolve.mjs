import fs from 'node:fs';
import { pptxToStandard } from '/Users/jiamao/project/github/pptx-parser/dist/ppt-parser.esm.js';
const buf = fs.readFileSync('/Users/jiamao/project/github/pptx-parser/examples/Sample_12.pptx');
const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
const doc = await pptxToStandard(ab, { themeProcess: true });
const sl = doc.slides[6];
const tbl = sl.elements.find(e => e.type === 'table');
const ts = tbl && tbl.__raw;
// 找到 tableStyles._themeContent 不容易，改为直接看标准元素首格 fill 与 doc.theme
console.log('doc.theme.accent1 =', doc.theme && doc.theme.colors && doc.theme.colors.accent1);
console.log('first cell fill =', tbl.rows[0].cells[0].fill);
console.log('second row cell0 fill =', tbl.rows[1].cells[0].fill);
console.log('third row cell0 fill =', tbl.rows[2].cells[0].fill);
// 反推 accent1：header=wholeTbl tint20000
function revTint(mixed, t) {
  // mixed = (1-t)*base + t*255 ; here t is fraction of white = (100 - tintPct)/100
  const f = (m) => ((m - 255*t)/(1-t));
  return [f(mixed[0]), f(mixed[1]), f(mixed[2])].map(Math.round);
}
const h = tbl.rows[0].cells[0].fill.replace('#','');
const hb = [parseInt(h.slice(0,2),16),parseInt(h.slice(2,4),16),parseInt(h.slice(4,6),16)];
// try tint20000 -> white fraction 0.8
console.log('implied accent1 (tint20%):', revTint(hb, 0.8));
console.log('implied accent1 (tint80%):', revTint(hb, 0.2));
