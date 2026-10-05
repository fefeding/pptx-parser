import fs from 'node:fs';
import { pptxToStandard } from '/Users/jiamao/project/github/pptx-parser/dist/ppt-parser.esm.js';
const buf = fs.readFileSync('/Users/jiamao/project/github/pptx-parser/examples/Sample_12.pptx');
const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
const doc = await pptxToStandard(ab, { themeProcess: true });
const sl = doc.slides[5];
console.log('slide theme?', Object.keys(sl.themeContent || {}));
const tc = sl.themeContent && sl.themeContent['a:theme'];
if (tc) {
  const cs = tc['a:themeElements'] && tc['a:themeElements']['a:clrScheme'];
  console.log('accent1', JSON.stringify(cs && cs['a:accent1']));
}
const tbl = sl.elements.find(e => e.type === 'table');
console.log('styleId', tbl.tableStyleId);
console.log('flags', JSON.stringify(tbl.tableStyleFlags));
console.log('rows', tbl.rows.length);
(tbl.rows || []).forEach((r, ri) => {
  const c0 = r.cells[0];
  const c1 = r.cells[1];
  const c2 = r.cells[2];
  console.log(`row${ri} c0=${c0.fill || 'none'} c1=${c1.fill || 'none'} c2=${c2.fill || 'none'} span=${c0.colSpan || 1}/${c0.rowSpan || 1}`);
});
