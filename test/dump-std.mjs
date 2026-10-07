/**
 * 诊断：dump pptxToStandard 的单页元素清单（类型/几何/文本/来源标记）。
 * 用法：node test/dump-std.mjs <pptx相对路径> [slideIndex 0-based]
 */
import fs from 'node:fs/promises';
import path from 'node:path';

const ROOT = '/Users/jiamao/project/github/pptx-parser';
const { pptxToStandard } = await import(path.join(ROOT, 'dist/ppt-parser.esm.js'));

const file = path.resolve(ROOT, process.argv[2] || 'examples/企业微信应用介绍.pptx');
const slideIdx = parseInt(process.argv[3] || '0', 10);

const buf = await fs.readFile(file);
const doc = await pptxToStandard(buf, { mediaProcess: true, themeProcess: true });
const slide = doc.slides[slideIdx];

console.log(`\n${path.basename(file)} — slide ${slideIdx + 1}`);
console.log(`slideSize: ${doc.slideSize.width}x${doc.slideSize.height}`);
console.log(`background: ${JSON.stringify(slide.background)?.slice(0, 200)}`);
console.log(`elements: ${slide.elements.length}\n`);

const pad = (s, n) => String(s).padEnd(n);
let i = 0;
function dump(el, depth) {
  const ind = '  '.repeat(depth);
  const geo = `${pad(Math.round(el.x), 7)},${pad(Math.round(el.y), 7)} ${pad(Math.round(el.width), 6)}x${pad(Math.round(el.height), 6)}`;
  const bits = [pad(geo, 26), pad(el.type, 8)];
  if (el.shapeType) bits.push(pad(`prst=${el.shapeType}`, 26));
  if (el.inherited) bits.push('[inherited]');
  if (el.adjust && Object.keys(el.adjust).length) bits.push(`adj=${JSON.stringify(el.adjust)}`);
  const txt = (el.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text).join('')).join('');
  if (txt) bits.push(`text="${txt.slice(0, 28)}"`);
  const rf = el.paragraphs && el.paragraphs[0] && el.paragraphs[0].runs && el.paragraphs[0].runs[0];
  if (rf && (rf.fontFace || rf.fontSize)) {
    bits.push(`run=${pad(rf.fontFace || '(继承)', 16)}${rf.fontSize ? rf.fontSize + 'pt' : ''}${rf.color ? ' ' + rf.color : ''}`);
  }
  console.log(`${ind}${pad('#' + (i++), 4)}${bits.join(' ')}`);
  for (const c of el.children || []) dump(c, depth + 1);
}

for (const el of slide.elements) dump(el, 0);
console.log();
