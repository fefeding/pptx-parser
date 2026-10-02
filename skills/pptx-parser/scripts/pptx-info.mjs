#!/usr/bin/env node
/**
 * 快速了解一份 PPTX：页数、画布、元数据、自定义属性、每页元素类型与文本、备注、批注、图表。
 *
 * 用法：
 *   node skills/pptx-parser/scripts/pptx-info.mjs <file.pptx> [--text] [--json]
 *     --text  打印每页文本（默认开启，量大时可关闭）
 *     --json  以 JSON 输出统计结果（便于程序消费）
 *
 * 依赖：优先用已安装的 @fefeding/ppt-parser，未安装时回退到本仓库构建产物 dist/。
 */

import { readFile } from 'node:fs/promises';
import { resolve, basename } from 'node:path';

const lib = await loadLib();
const { pptxToJson } = lib;

const args = process.argv.slice(2);
const file = args.find((a) => !a.startsWith('--'));
const showText = !args.includes('--no-text');
const asJson = args.includes('--json');

if (!file) {
  console.error('用法: node pptx-info.mjs <file.pptx> [--no-text] [--json]');
  process.exit(1);
}

const buf = await readFile(resolve(file));
const result = await pptxToJson(toArrayBuffer(buf), {
  mediaProcess: true,
  themeProcess: true,
  mode: 'semantic'
});

const doc = result.document;
const slides = doc?.slides ?? result.slides;

/** 统计一页内各类型元素数量 */
function typeStats(elements = []) {
  const stats = {};
  for (const el of elements) {
    const t = el?.type ?? 'unknown';
    stats[t] = (stats[t] || 0) + 1;
  }
  return stats;
}

/** 递归收集一个元素内的全部文本 */
function collectText(el, out = []) {
  if (!el || typeof el !== 'object') return out;
  if (typeof el.text === 'string' && el.text) out.push(el.text);
  for (const p of el.paragraphs || []) {
    if (typeof p.text === 'string' && p.text) out.push(p.text);
    for (const r of p.runs || []) if (r?.text) out.push(r.text);
  }
  for (const r of el.runs || []) if (r?.text) out.push(r.text);
  for (const row of el.rows || []) for (const c of row.cells || []) collectText(c, out);
  for (const child of el.children || []) collectText(child, out);
  for (const n of el.nodes || []) collectText(n, out);
  if (el.title) out.push(el.title);
  for (const s of el.series || []) if (s?.name) out.push(s.name);
  for (const c of el.categories || []) out.push(String(c));
  for (const t of el.texts || []) out.push(String(t));
  return out;
}

const report = {
  file: basename(file),
  bytes: buf.length,
  slideSize: `${result.slideSize?.width} x ${result.slideSize?.height} (px)`,
  pages: slides.length,
  metadata: result.metadata,
  customProps: result.customProps,
  charts: (result.charts || []).map((c) => ({ id: c.chartId, type: c.type, series: c.data?.length ?? 0 })),
  slides: slides.map((s, i) => {
    const elements = s.elements || [];
    const texts = collectText({ children: elements });
    return {
      page: i + 1,
      stats: typeStats(elements),
      notes: s.notes || '',
      comments: (s.comments || []).map((c) => `${c.author || 'Author'}: ${c.text}`),
      hidden: !!s.hidden,
      texts
    };
  })
};

if (asJson) {
  console.log(JSON.stringify(report, null, 2));
} else {
  console.log(`\n文件: ${report.file}  (${(report.bytes / 1024).toFixed(1)} KB)`);
  console.log(`画布: ${report.slideSize}`);
  console.log(`页数: ${report.pages}`);
  if (report.metadata && Object.keys(report.metadata).length) {
    console.log(`元数据: ${JSON.stringify(report.metadata)}`);
  }
  if (report.customProps && Object.keys(report.customProps).length) {
    console.log(`自定义属性: ${JSON.stringify(report.customProps)}`);
  }
  if (report.charts.length) {
    console.log(`图表: ${report.charts.length} 个 ${report.charts.map((c) => `[${c.type}]`).join(' ')}`);
  }
  for (const s of report.slides) {
    const fmt = Object.entries(s.stats).map(([k, v]) => `${k}:${v}`).join(' ') || '（无元素）';
    console.log(`\n[第 ${s.page} 页]${s.hidden ? ' (隐藏)' : ''} ${fmt}`);
    if (s.notes) console.log(`  备注: ${s.notes.slice(0, 80)}`);
    for (const c of s.comments) console.log(`  批注: ${c}`);
    if (showText) {
      const list = s.texts.filter((t) => String(t).trim());
      for (const t of list.slice(0, 20)) console.log(`  · ${String(t).slice(0, 100)}`);
      if (list.length > 20) console.log(`  … 共 ${list.length} 条`);
    }
  }
  console.log('');
}

/** Buffer → 独立 ArrayBuffer（避免共享内存池读到脏数据） */
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
