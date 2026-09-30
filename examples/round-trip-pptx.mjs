#!/usr/bin/env node
/**
 * PPTX → 标准 JSON → PPTX 往返示例
 *
 * 演示统一契约 PptxDocument 的双向能力：
 *   1) pptxToStandard(file)  解析 PPTX 为标准 JSON
 *   2) jsonToPptx(json)      由该 JSON 重新生成 PPTX
 *   3) 再次解析产物，与原解析结果做逐页对比，观察保真度
 *
 * 运行前请先构建产物（示例从 dist 导入）：
 *   npm run build
 *
 * 用法：
 *   node examples/round-trip-pptx.mjs                        # 默认 examples/test-sample.pptx
 *   node examples/round-trip-pptx.mjs 路径/xxx.pptx          # 指定文件
 *
 * 产物输出到 examples/out/：
 *   <name>.standard.json         完整标准 JSON（含 __raw 兜底载荷）
 *   <name>.standard.no-raw.json  去掉 __raw 的精简版（便于阅读）
 *   <name>.roundtrip.pptx        由 JSON 重新生成的 PPTX
 */

import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, dirname, basename, extname } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));

let lib;
try {
    lib = await import('@fefeding/ppt-parser');
} catch {
    lib = await import('../dist/ppt-parser.esm.js');
}
const { pptxToStandard, jsonToPptx } = lib;

/** 字节数格式化 */
function fmtSize(n) {
    if (n < 1024) return `${n} B`;
    if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
    return `${(n / 1024 / 1024).toFixed(2)} MB`;
}

/** 深度删除 __raw / rawFallback，得到可阅读的精简 JSON */
function stripRaw(value) {
    if (Array.isArray(value)) return value.map(stripRaw);
    if (value && typeof value === 'object') {
        const out = {};
        for (const [k, v] of Object.entries(value)) {
            if (k === '__raw' || k === 'rawFallback') continue;
            out[k] = stripRaw(v);
        }
        return out;
    }
    return value;
}

/** 统计一页内各类型元素数量 */
function typeStats(elements = []) {
    const stats = {};
    for (const el of elements) {
        const t = el && el.type ? el.type : 'unknown';
        stats[t] = (stats[t] || 0) + 1;
    }
    return stats;
}

/** 收集一个元素内的全部文本（用于文本量对比） */
function collectElementText(el, out = []) {
    if (!el || typeof el !== 'object') return out;
    if (typeof el.text === 'string' && el.text) out.push(el.text);
    for (const p of el.paragraphs || []) {
        if (typeof p.text === 'string' && p.text) out.push(p.text);
        for (const r of p.runs || []) if (r && r.text) out.push(r.text);
    }
    for (const r of el.runs || []) if (r && r.text) out.push(r.text);
    // 表格单元格
    for (const row of el.rows || []) {
        for (const cell of row.cells || []) collectElementText(cell, out);
    }
    if (el.title) out.push(el.title);
    for (const s of el.series || []) if (s && s.name) out.push(s.name);
    for (const c of el.categories || []) out.push(String(c));
    for (const t of el.texts || []) out.push(t);
    return out;
}

/** 收集整份文档的文本（含备注） */
function collectDocText(doc) {
    const out = [];
    for (const slide of doc.slides || []) {
        if (slide.notes) out.push(slide.notes);
        for (const el of slide.elements || []) collectElementText(el, out);
    }
    return out.filter((t) => String(t).trim().length > 0);
}

async function main() {
    const inputPath = resolve(__dirname, process.argv[2] || 'test-sample.pptx');
    const outDir = resolve(__dirname, 'out');
    await mkdir(outDir, { recursive: true });
    const name = basename(inputPath, extname(inputPath));

    const input = await readFile(inputPath);
    console.log(`\n输入文件: ${inputPath}`);
    console.log(`输入大小: ${fmtSize(input.length)}`);

    // ===== 1. PPTX → 标准 JSON =====
    const t0 = Date.now();
    const doc = await pptxToStandard(input, { mediaProcess: true, themeProcess: true });
    const t1 = Date.now();

    const fullJson = JSON.stringify(doc, null, 2);
    const slimJson = JSON.stringify(stripRaw(doc), null, 2);
    await writeFile(resolve(outDir, `${name}.standard.json`), fullJson);
    await writeFile(resolve(outDir, `${name}.standard.no-raw.json`), slimJson);

    // ===== 2. 标准 JSON → PPTX =====
    const output = await jsonToPptx(doc, { outputType: 'uint8array' });
    const outPath = resolve(outDir, `${name}.roundtrip.pptx`);
    await writeFile(outPath, output);
    const t2 = Date.now();

    // ===== 3. 再次解析产物做对比 =====
    const back = await pptxToStandard(output, { mediaProcess: true, themeProcess: true });

    console.log('\n=== 基础信息 ===');
    console.log(`解析耗时: ${t1 - t0} ms   生成耗时: ${t2 - t1} ms`);
    console.log(`画布尺寸: ${doc.slideSize?.width} x ${doc.slideSize?.height} (px)`);
    console.log(`元数据: ${doc.metadata ? Object.keys(doc.metadata).length + ' 项' : '无'}`);
    console.log(`JSON 大小: ${fmtSize(Buffer.byteLength(fullJson))}（精简版 ${fmtSize(Buffer.byteLength(slimJson))}）`);
    console.log(`输出文件: ${outPath}`);
    console.log(`输出大小: ${fmtSize(output.length)}（原文件 ${fmtSize(input.length)}，${(output.length / input.length * 100).toFixed(0)}%）`);

    console.log('\n=== 逐页元素对比（原始解析 → 重新生成后再解析）===');
    const beforeCount = (doc.slides || []).reduce((n, s) => n + (s.elements?.length || 0), 0);
    const afterCount = (back.slides || []).reduce((n, s) => n + (s.elements?.length || 0), 0);
    console.log(`页数: ${doc.slides?.length} → ${back.slides?.length}`);
    for (let i = 0; i < Math.max(doc.slides?.length || 0, back.slides?.length || 0); i++) {
        const a = doc.slides[i];
        const b = back.slides[i];
        const sa = typeStats(a?.elements);
        const sb = typeStats(b?.elements);
        const fmt = (s) => Object.entries(s).map(([k, v]) => `${k}:${v}`).join(' ') || '（无元素）';
        const mark = JSON.stringify(sa) === JSON.stringify(sb) ? '=' : '≠';
        console.log(`  第 ${i + 1} 页 ${mark}  [${fmt(sa)}]  →  [${fmt(sb)}]`);
        if (a?.notes) console.log(`        备注: ${String(a.notes).slice(0, 40)}`);
    }
    console.log(`元素总数: ${beforeCount} → ${afterCount}`);

    // 文本比对：忽略空片段与顺序，按多重集合统计缺失/新增
    const beforeTexts = collectDocText(doc);
    const afterTexts = collectDocText(back);
    const counter = (list) => list.reduce((m, t) => {
        const k = String(t).trim();
        if (k) m.set(k, (m.get(k) || 0) + 1);
        return m;
    }, new Map());
    const cb = counter(beforeTexts);
    const ca = counter(afterTexts);
    const missing = [...cb].filter(([k, v]) => (ca.get(k) || 0) < v).map(([k]) => k);
    const extra = [...ca].filter(([k, v]) => (cb.get(k) || 0) < v).map(([k]) => k);

    console.log('\n=== 文本保真 ===');
    console.log(`非空文本片段数: ${cb.size} → ${ca.size}`);
    console.log(`文本一致: ${missing.length === 0 && extra.length === 0 ? '是' : '否'}`);
    if (missing.length) {
        console.log(`缺失片段（共 ${missing.length} 条，最多显示 10 条）:`);
        for (const m of missing.slice(0, 10)) console.log(`    - ${m.slice(0, 60)}`);
    }
    if (extra.length) {
        console.log(`新增片段（共 ${extra.length} 条，最多显示 10 条）:`);
        for (const m of extra.slice(0, 10)) console.log(`    + ${m.slice(0, 60)}`);
    }

    console.log('\n产出目录: examples/out/');
    console.log('提示: .no-raw.json 去掉了 __raw 兜底载荷，适合直接阅读/编辑后再用 jsonToPptx 生成。\n');
}

main().catch((err) => {
    console.error('往返失败:', err);
    console.error('\n提示：若报错找不到模块，请先执行 `npm run build` 生成 dist 目录。');
    process.exit(1);
});
