#!/usr/bin/env node
/**
 * 由标准 JSON（PptxDocument）生成 PPTX，并可选做一次往返自检。
 *
 * 用法：
 *   node skills/pptx-parser/scripts/json-to-pptx.mjs <doc.json> [--out out.pptx] [--check]
 *     --out    输出 PPTX 路径（默认 <doc>.pptx）
 *     --check  生成后重新解析，对比页数/元素类型/文本保真
 *
 * JSON 结构见 references/json-schema.md；也可以先用 pptx-info.mjs 了解一份现成 PPTX，
 * 或用 round-trip 方式获得：pptxToStandard(file) → JSON.stringify → 本脚本。
 */

import { readFile, writeFile } from 'node:fs/promises';
import { resolve, basename, extname } from 'node:path';

const lib = await loadLib();
const { jsonToPptx, pptxToJson } = lib;

const args = process.argv.slice(2);
const file = args.find((a) => !a.startsWith('--'));
const outIdx = args.indexOf('--out');
const out = outIdx >= 0 ? args[outIdx + 1] : undefined;
const check = args.includes('--check');

if (!file) {
  console.error('用法: node json-to-pptx.mjs <doc.json> [--out out.pptx] [--check]');
  process.exit(1);
}

const doc = JSON.parse(await readFile(resolve(file), 'utf8'));
const data = await jsonToPptx(doc, { outputType: 'uint8array' });
const outPath = resolve(out || `${basename(file, extname(file))}.pptx`);
await writeFile(outPath, Buffer.from(data));
console.log(`已生成: ${outPath} (${(data.byteLength / 1024).toFixed(1)} KB)`);

if (!check) process.exit(0);

// ===== 往返自检：重新解析产物，与原 JSON 对比 =====
const back = await pptxToJson(toArrayBuffer(Buffer.from(data)), { mode: 'semantic' });
const before = doc.slides || [];
const after = back.document?.slides || [];

const typeStats = (elements = []) => {
  const m = {};
  for (const el of elements) m[el?.type ?? 'unknown'] = (m[el?.type ?? 'unknown'] || 0) + 1;
  return m;
};
const collect = (slide, out = []) => {
  for (const el of slide.elements || []) {
    walk(el, out);
  }
  return out;
};
const walk = (el, out = []) => {
  if (!el || typeof el !== 'object') return out;
  if (typeof el.text === 'string' && el.text) out.push(el.text);
  for (const p of el.paragraphs || []) {
    if (p.text) out.push(p.text);
    for (const r of p.runs || []) if (r?.text) out.push(r.text);
  }
  for (const r of el.runs || []) if (r?.text) out.push(r.text);
  for (const row of el.rows || []) for (const c of row.cells || []) walk(c, out);
  for (const ch of el.children || []) walk(ch, out);
  for (const n of el.nodes || []) walk(n, out);
  if (el.title) out.push(el.title);
  return out;
};
const counter = (list) => list.reduce((m, t) => {
  const k = String(t).trim();
  if (k) m.set(k, (m.get(k) || 0) + 1);
  return m;
}, new Map());

console.log('\n=== 往返自检 ===');
console.log(`页数: ${before.length} → ${after.length}`);
let diffPages = 0;
for (let i = 0; i < Math.max(before.length, after.length); i++) {
  const a = JSON.stringify(typeStats(before[i]?.elements));
  const b = JSON.stringify(typeStats(after[i]?.elements));
  const mark = a === b ? '=' : '≠';
  if (a !== b) diffPages++;
  console.log(`  第 ${i + 1} 页 ${mark}  ${a}  →  ${b}`);
}

const cb = counter(before.flatMap((s) => collect(s)));
const ca = counter(after.flatMap((s) => collect(s)));
const missing = [...cb].filter(([k, v]) => (ca.get(k) || 0) < v).map(([k]) => k);
console.log(`元素类型不一致页数: ${diffPages}`);
console.log(`文本保真: ${missing.length === 0 ? '是' : `否（缺失 ${missing.length} 条）`}`);
for (const m of missing.slice(0, 10)) console.log(`  - ${m.slice(0, 80)}`);

process.exit(diffPages === 0 && missing.length === 0 ? 0 : 2);

/** Buffer → 独立 ArrayBuffer */
function toArrayBuffer(b) {
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
}

/** 优先用已安装的包，否则回退到仓库构建产物 */
async function loadLib() {
  try {
    return await import('@fefeding/ppt-parser');
  } catch {
    return await import('../../../dist/ppt-parser.esm.js');
  }
}
