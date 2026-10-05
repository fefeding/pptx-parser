import fs from 'node:fs';
import { JSDOM } from 'jsdom';
const dom = new JSDOM('<!DOCTYPE html><html><body></body></html>');
globalThis.document = dom.window.document;
globalThis.window = dom.window;
const { pptxToHtml } = await import('/Users/jiamao/project/github/pptx-parser/dist/ppt-parser.esm.js');
const buf = fs.readFileSync('/Users/jiamao/project/github/pptx-parser/examples/Sample_12.pptx');
const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
const res = await pptxToHtml(ab, {});
const s7 = res.slides[6];
const tc = s7 && s7.data && s7.data.themeContent;
if (tc) {
  const cs = tc['a:theme'] && tc['a:theme']['a:themeElements'] && tc['a:theme']['a:themeElements']['a:clrScheme'];
  console.log('theme accent1 =', JSON.stringify(cs && cs['a:accent1']));
  console.log('theme accent6 =', JSON.stringify(cs && cs['a:accent6']));
}
console.log('themeContent keys', tc ? Object.keys(tc) : 'none');
