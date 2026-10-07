/**
 * 诊断：检查单页 p:spTree 的节点标签/prstGeom/占位符，以及版式/母版 spTree 的装饰形状。
 * 用法：node test/dump-xml.mjs <pptx相对路径> [slideIndex 0-based]
 */
import fs from 'node:fs/promises';
import path from 'node:path';

const ROOT = '/Users/jiamao/project/github/pptx-parser';
const { pptxToJson } = await import(path.join(ROOT, 'dist/ppt-parser.esm.js'));

const file = path.resolve(ROOT, process.argv[2] || 'examples/企业微信应用介绍.pptx');
const slideIdx = parseInt(process.argv[3] || '0', 10);
const asArray = (v) => (v == null ? [] : Array.isArray(v) ? v : [v]);
const pad = (s, n) => String(s == null ? '-' : s).padEnd(n);

const buf = await fs.readFile(file);
const json = await pptxToJson(buf, { mediaProcess: true, themeProcess: true });
const d = json.slides[slideIdx].data;

function walk(tree, title) {
  console.log('\n--- ' + title + ' ---');
  if (!tree) { console.log('(无)'); return; }
  for (const key of Object.keys(tree)) {
    const val = tree[key];
    if (val == null || typeof val !== 'object') continue;
    for (const node of asArray(val)) {
      const nvKey = ['p:nvSpPr','p:nvPicPr','p:nvGraphicFramePr','p:nvCxnSpPr','p:nvGrpSpPr'].find((k) => node[k]);
      const nv = nvKey ? node[nvKey] : null;
      const ph = nv && nv['p:nvPr'] && nv['p:nvPr']['p:ph'];
      const phs = ph ? ' ph(' + pad(ph.attrs && ph.attrs.type, 10) + 'idx=' + pad(ph.attrs && ph.attrs.idx, 4) + ')' : ' (无ph)';
      const name = nv && nv['p:cNvPr'] && nv['p:cNvPr'].attrs && nv['p:cNvPr'].attrs.name;
      const spPr = node['p:spPr'] || node['p:grpSpPr'] || node['p:xfrm'];
      let prst = '无';
      if (spPr) {
        if (spPr['a:prstGeom']) prst = spPr['a:prstGeom'].attrs.prst;
        else if (spPr['a:custGeom']) prst = 'custGeom';
      }
      const xf = spPr && spPr['a:xfrm'];
      const off = xf && xf['a:off'] && xf['a:off'].attrs;
      const ext = xf && xf['a:ext'] && xf['a:ext'].attrs;
      const geo = off && ext ? ' @(' + off.x + ',' + off.y + ') ' + ext.cx + 'x' + ext.cy : ' (无xfrm)';
      console.log('  ' + pad(key, 18) + 'name=' + pad(name, 24) + phs + ' prst=' + pad(prst, 16) + geo);
      if (key === 'p:grpSp') walk(node, title + ' > ' + (name || 'grpSp'));
    }
  }
}

const treeOf = (c, tag) => { const r = c && c[tag]; return r && r['p:cSld'] && r['p:cSld']['p:spTree']; };

console.log('\n' + path.basename(file) + ' slide ' + (slideIdx + 1));
const fsch = d.themeContent && d.themeContent['a:theme'] && d.themeContent['a:theme']['a:themeElements'] && d.themeContent['a:theme']['a:themeElements']['a:fontScheme'];
console.log('theme fontScheme:');
for (const t of ['a:majorFont', 'a:minorFont']) {
  const f = fsch && fsch[t];
  if (!f) continue;
  const g = (k) => (f[k] && f[k].attrs ? f[k].attrs.typeface : '-');
  console.log('  ' + t + ': latin=' + g('a:latin') + '  ea=' + g('a:ea') + '  cs=' + g('a:cs'));
}
walk(treeOf(d.slideContent, 'p:sld'), 'slide spTree');
walk(treeOf(d.slideLayoutContent, 'p:sldLayout'), 'layout spTree');
walk(treeOf(d.slideMasterContent, 'p:sldMaster'), 'master spTree');
console.log();
