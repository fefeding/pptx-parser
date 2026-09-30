/**
 * PPTX → 标准 JSON（pptxToJson 语义模式的反向提取）
 *
 * 将 pptxToJson 解析出的底层 OOXML 树（SlideDataRecord.data）反向提取为
 * 语义化、与 jsonToPptx 输入同源的 PptxDocument（见 types/pptx-document.ts），
 * 从而实现 JSON 级 round-trip：pptxToJson(..., {mode:'semantic'}) 或 pptxToStandard()
 * 产出 PptxDocument，再交给 jsonToPptx(PptxDocument) 可还原。
 *
 * 提取策略：
 * - 遍历 spTree 中的 p:sp / p:pic / p:graphicFrame（groups 递归展开）。
 * - 文本优先：含文本内容的 p:sp 视为 text 元素；仅含几何的视为 shape。
 * - 坐标 EMU → px（SLIDE_FACTOR）；线宽 EMU → pt（/12700）；旋转 rot → 度（/60000）。
 * - 媒体内联为 dataURL（保证 round-trip 自包含）；图表反向解析 c:chartSpace。
 * - 每个元素附 __raw（原始 OOXML 子树）作为无损兜底；单元素解析失败不影响整体。
 *
 * @module serializer/json-from-pptx
 */

import JSZip from 'jszip';
import { SLIDE_FACTOR } from '../core/constants';
import { PPTXXmlUtils } from '../utils/xml';
import type {
    PptxDocument, PptxSlide, PptxElement, PptxTextElement, PptxShapeElement,
    PptxImageElement, PptxChartElement, PptxParagraph, PptxTextRun, PptxChartSeries,
    PptxTransition, PptxBackground
} from '../types/pptx-document';

/** 1pt = 12700 EMU */
const EMU_PER_PT = 12700;
/** 线宽 EMU 默认值（缺省按 1pt 处理时参考） */
const DEFAULT_LN_PT = 1;

/** 将任意值规整为数组（tXml 单节点即对象，多节点为数组） */
function asArray<T = any>(v: T | T[] | undefined): T[] {
    if (v === undefined || v === null) return [];
    return Array.isArray(v) ? v : [v];
}

/** EMU → px（保留两位小数） */
function emuToPx(emu: unknown): number {
    const n = Number(emu) || 0;
    return Math.round(n * SLIDE_FACTOR * 100) / 100;
}

/** EMU → pt（保留两位小数） */
function emuToPt(emu: unknown): number {
    const n = Number(emu) || 0;
    return Math.round((n / EMU_PER_PT) * 100) / 100;
}

/** rot（60000 分之一度）→ 度 */
function rotToDeg(rot: unknown): number {
    const n = Number(rot) || 0;
    return n === 0 ? 0 : Math.round(n / 60000);
}

/** 从相对路径（如 ../media/image1.png）解析为 zip 内绝对部件路径 */
function resolvePart(target: string | undefined): string | undefined {
    if (!target) return undefined;
    return target.replace(/\.\.\//g, 'ppt/').replace(/^\/+/, '');
}

/** 从节点读取 a:srgbClr 的颜色值 */
function readSrgbClr(node: any): string | undefined {
    const c = node && (node['a:srgbClr'] || (node['a:solidFill'] && node['a:solidFill']['a:srgbClr']));
    return c && c.attrs && c.attrs.val ? String(c.attrs.val) : undefined;
}

/** 读取文本运行中的文本（a:t 在 simplify 形态下为字符串） */
function readRunText(runNode: any): string {
    const t = runNode && runNode['a:t'];
    return typeof t === 'string' ? t : (t ? String(t) : '');
}

/** 读取运行级样式（a:rPr） */
function readRunStyle(rPr: any): Partial<PptxTextRun> {
    const style: Partial<PptxTextRun> = {};
    if (!rPr) return style;
    const attrs = rPr.attrs || {};
    if (attrs.sz) style.fontSize = Math.round(Number(attrs.sz) / 100 * 100) / 100; // 百分之一 pt → pt
    if (attrs.b === '1' || attrs.b === 1) style.bold = true;
    if (attrs.i === '1' || attrs.i === 1) style.italic = true;
    if (attrs.u && attrs.u !== 'none') style.underline = true;
    const color = readSrgbClr(rPr);
    if (color) style.color = color;
    const latin = rPr['a:latin'];
    if (latin && latin.attrs && latin.attrs.typeface) style.fontFace = String(latin.attrs.typeface);
    const hlink = rPr['a:hlinkClick'];
    if (hlink && hlink.attrs && hlink.attrs['r:id']) style.href = String(hlink.attrs['r:id']);
    return style;
}

/** 提取一个 p:sp 的文本为正文段落 */
function extractTextBody(spNode: any): { paragraphs: PptxParagraph[]; hasText: boolean } {
    const paragraphs: PptxParagraph[] = [];
    let hasText = false;
    const txBody = spNode && spNode['p:txBody'];
    if (!txBody) return { paragraphs, hasText };

    const bodyPr = txBody['a:bodyPr'];
    const valignMap: Record<string, 'top' | 'middle' | 'bottom'> = { ctr: 'middle', b: 'bottom' };
    const defaultValign = bodyPr && bodyPr.attrs && bodyPr.attrs.anchor
        ? valignMap[bodyPr.attrs.anchor] : undefined;

    for (const pNode of asArray(txBody['a:p'])) {
        const pPr = pNode['a:pPr'];
        const pAttrs = (pPr && pPr.attrs) || {};
        const alignMap: Record<string, 'left' | 'center' | 'right' | 'justify'> = { l: 'left', ctr: 'center', r: 'right', just: 'justify' };
        const align = pAttrs.algn ? alignMap[pAttrs.algn] : undefined;
        // 项目符号：存在 buChar/buAutoNum 且无 buNone
        const bullet = !!(pNode['a:pPr'] && (pNode['a:pPr']['a:buChar'] || pNode['a:pPr']['a:buAutoNum']))
            && !(pNode['a:pPr']['a:buNone']);

        const runs: PptxTextRun[] = [];
        for (const runNode of asArray(pNode['a:r'])) {
            const text = readRunText(runNode);
            if (text) hasText = true;
            const style = readRunStyle(runNode['a:rPr']);
            runs.push({ text, ...style });
        }
        const para: PptxParagraph = { runs };
        if (align) para.align = align;
        if (bullet) para.bullet = true;
        if (defaultValign) (para as any).valign = defaultValign;
        paragraphs.push(para);
    }
    return { paragraphs, hasText };
}

/** 从 p:spPr 读取几何/填充/边框 */
function readSpPr(spPr: any): Pick<PptxShapeElement, 'shapeType' | 'fill' | 'line'> {
    const out: Pick<PptxShapeElement, 'shapeType' | 'fill' | 'line'> = { shapeType: 'rect' };
    if (!spPr) return out;
    const prst = spPr['a:prstGeom'];
    if (prst && prst.attrs && prst.attrs.prst) out.shapeType = String(prst.attrs.prst);

    // 填充
    if (spPr['a:noFill']) {
        out.fill = 'none';
    } else {
        const color = readSrgbClr(spPr);
        if (color) out.fill = color;
    }

    // 边框
    const ln = spPr['a:ln'];
    if (ln) {
        if (ln['a:noFill']) {
            out.line = 'none';
        } else {
            const color = readSrgbClr(ln);
            const w = ln.attrs && ln.attrs.w ? emuToPt(ln.attrs.w) : DEFAULT_LN_PT;
            out.line = color ? { color, width: w } : { width: w };
        }
    }
    return out;
}

/** 读取 xfrm 位置（graphicFrame 用 p:xfrm，其余用 p:spPr/a:xfrm） */
function readXfrm(node: any, isGraphicFrame: boolean) {
    const xf = isGraphicFrame ? node['p:xfrm'] : (node['p:spPr'] && node['p:spPr']['a:xfrm']);
    if (!xf || !xf['a:off']) return null;
    const off = (xf['a:off'] && xf['a:off'].attrs) || {};
    const ext = (xf['a:ext'] && xf['a:ext'].attrs) || {};
    return {
        x: emuToPx(off.x),
        y: emuToPx(off.y),
        width: emuToPx(ext.cx),
        height: emuToPx(ext.cy),
        rotation: rotToDeg(xf.attrs && xf.attrs.rot)
    };
}

/** 提取备注文本（notesSlide） */
function extractNotes(notesContent: any): string | undefined {
    if (!notesContent) return undefined;
    const spTree = notesContent['p:notesSlide']
        && notesContent['p:notesSlide']['p:cSld']
        && notesContent['p:notesSlide']['p:cSld']['p:spTree'];
    if (!spTree) return undefined;
    const texts: string[] = [];
    for (const sp of asArray(spTree['p:sp'])) {
        const txBody = sp && sp['p:txBody'];
        if (!txBody) continue;
        for (const p of asArray(txBody['a:p'])) {
            for (const r of asArray(p['a:r'])) {
                const t = readRunText(r);
                if (t) texts.push(t);
            }
        }
    }
    const joined = texts.join('\n').trim();
    return joined.length ? joined : undefined;
}

/** 提取过渡效果 */
function extractTransition(slideContent: any): PptxTransition | undefined {
    const sld = slideContent && slideContent['p:sld'];
    if (!sld) return undefined;
    const t = sld['p:transition'];
    if (!t) return undefined;
    const types = ['p:blinds', 'p:checker', 'p:circle', 'p:comb', 'p:cover', 'p:dissolve', 'p:fade', 'p:push', 'p:random', 'p:split', 'p:strips', 'p:wipe'];
    let type = 'fade';
    for (const ty of types) {
        if (t[ty]) { type = ty.replace('p:', ''); break; }
    }
    let duration = 1000;
    if (t.attrs && t.attrs['spd']) {
        const m: Record<string, number> = { '1': 500, '2': 1000, '3': 2000 };
        duration = m[String(t.attrs['spd'])] || 1000;
    }
    return { type, duration };
}

/** 提取背景 */
function extractBackground(slideContent: any): PptxBackground | undefined {
    const sld = slideContent && slideContent['p:sld'];
    if (!sld) return undefined;
    const bg = sld['p:cSld'] && sld['p:cSld']['p:bg'];
    if (!bg) return undefined;
    const bgPr = bg['p:bgPr'];
    if (bgPr) {
        const color = readSrgbClr(bgPr);
        if (color) return color;
        if (bgPr['a:gradFill']) {
            // 简化：记录为 gradient，stops 尽力提取
            const gsLst = bgPr['a:gradFill']['a:gsLst'];
            const stops: { color: string; position: number }[] = [];
            for (const gs of asArray(gsLst && gsLst['a:gs'])) {
                const pos = gs && gs.attrs && gs.attrs.pos ? Number(gs.attrs.pos) / 100000 : 0;
                const c = readSrgbClr(gs);
                if (c) stops.push({ color: c, position: pos });
            }
            return { type: 'gradient', direction: 'horizontal', stops };
        }
    }
    // 主题引用 bgRef：无法解析为具体色，标记继承（不写 background）
    return undefined;
}

/** 从 c:chartSpace 反向提取图表语义 */
function extractChart(chartXml: any): Partial<PptxChartElement> | undefined {
    try {
        const chart = chartXml && chartXml['c:chartSpace'] && chartXml['c:chartSpace']['c:chart'];
        if (!chart) return undefined;
        const plotArea = chart['c:plotArea'];
        if (!plotArea) return undefined;

        // 图表类型
        const chartTypes = ['c:barChart', 'c:lineChart', 'c:areaChart', 'c:pieChart', 'c:pie3DChart', 'c:scatterChart'];
        let chartNode: any = null;
        let chartType = 'barChart';
        for (const ct of chartTypes) {
            if (plotArea[ct]) { chartNode = plotArea[ct]; chartType = ct.replace('c:', ''); break; }
        }
        if (!chartNode) return undefined;

        const isScatter = chartType === 'scatterChart';

        // 类别（位于首个 series 的 c:cat 下）
        const firstSer = asArray(chartNode['c:ser'])[0];
        const catNode = firstSer && firstSer['c:cat'];
        let categories: string[] = [];
        if (catNode) {
            const strCache = catNode['c:strRef'] && catNode['c:strRef']['c:strCache'];
            const pts = strCache ? asArray(strCache['c:pt']) : [];
            categories = pts.map((p: any) => (p && p['c:v'] !== undefined ? String(p['c:v']) : '')).filter(Boolean);
        }

        // 系列
        const series: PptxChartSeries[] = [];
        for (const ser of asArray(chartNode['c:ser'])) {
            const nameNode = ser['c:tx'] && ser['c:tx']['c:strRef'] && ser['c:tx']['c:strRef']['c:strCache'];
            const namePts = nameNode ? asArray(nameNode['c:pt']) : [];
            const name = namePts.length && namePts[0]['c:v'] ? String(namePts[0]['c:v']) : undefined;

            const s: PptxChartSeries = {};
            if (name) s.name = name;
            if (isScatter) {
                s.x = numCacheValues(ser['c:xVal']);
                s.y = numCacheValues(ser['c:yVal']);
            } else {
                s.values = numCacheValues(ser['c:val']);
            }
            const serColor = readSrgbClr(ser);
            if (serColor) s.color = serColor;
            series.push(s);
        }

        // 标题
        let title: string | undefined;
        const titleRich = chart['c:title'] && chart['c:title']['c:tx'] && chart['c:title']['c:tx']['c:rich'];
        if (titleRich) {
            for (const p of asArray(titleRich['a:p'])) {
                for (const r of asArray(p['a:r'])) {
                    const t = readRunText(r);
                    if (t) title = (title || '') + t;
                }
            }
        }

        const legend = !!chart['c:legend'];
        const out: Partial<PptxChartElement> = { chartType, series };
        if (categories.length) out.categories = categories;
        if (title) out.title = title.trim();
        out.legend = legend;
        return out;
    } catch {
        return undefined;
    }
}

/** 从 c:numRef/c:numCache 读取数值数组 */
function numCacheValues(refNode: any): number[] {
    if (!refNode) return [];
    const cache = refNode['c:numRef'] && refNode['c:numRef']['c:numCache'];
    if (!cache) return [];
    return asArray(cache['c:pt'])
        .map((p: any) => (p && p['c:v'] !== undefined ? Number(p['c:v']) : NaN))
        .filter((v: number) => !isNaN(v));
}

/** 递归收集 spTree 下的图形节点（展开 group） */
function collectShapeNodes(spTree: any, acc: any[]) {
    if (!spTree || typeof spTree !== 'object') return;
    for (const key of Object.keys(spTree)) {
        const val = spTree[key];
        if (val === undefined || val === null) continue;
        const nodes = asArray(val);
        for (const node of nodes) {
            if (key === 'p:grpSp') {
                const inner = node && node['p:spTree'];
                if (inner) collectShapeNodes(inner, acc);
            } else if (['p:sp', 'p:pic', 'p:graphicFrame', 'p:cxnSp'].includes(key)) {
                acc.push({ key, node });
            }
        }
    }
}

/**
 * 将单页解析数据提取为标准 PptxSlide
 * @param slideData - pptxToJson 产出的每页 data（SlideDataRecord）
 * @param zip - JSZip 实例（用于读取媒体/图表部件）
 */
export async function extractSlideToStandard(slideData: any, zip: JSZip): Promise<PptxSlide> {
    const slide: PptxSlide = { elements: [] };

    try {
        const slideContent = slideData && slideData.slideContent;
        const spTree = slideContent
            && slideContent['p:sld']
            && slideContent['p:sld']['p:cSld']
            && slideContent['p:sld']['p:cSld']['p:spTree'];

        const resObj: Record<string, { type?: string; target?: string }> = slideData.slideResObj || {};

        // 背景 / 过渡 / 备注
        const bg = extractBackground(slideContent);
        if (bg !== undefined) slide.background = bg;
        const transition = extractTransition(slideContent);
        if (transition) slide.transition = transition;
        const notes = extractNotes(slideData.notesContent);
        if (notes) slide.notes = notes;

        if (spTree) {
            const shapeNodes: { key: string; node: any }[] = [];
            collectShapeNodes(spTree, shapeNodes);

            for (const { key, node } of shapeNodes) {
                try {
                    const el = await nodeToElement(key, node, resObj, zip);
                    if (el) slide.elements.push(el);
                } catch {
                    // 单元素失败：保留原始节点，不影响其它元素
                    slide.elements.push({ type: 'text', x: 0, y: 0, width: 0, height: 0, __raw: node } as any);
                }
            }
        }
    } catch {
        // 整页失败：返回空元素列表（保持结构合法）
    }

    return slide;
}

/** 单个 OOXML 节点 → PptxElement */
async function nodeToElement(
    key: string,
    node: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip
): Promise<PptxElement | null> {
    if (key === 'p:graphicFrame') {
        return await graphicFrameToChart(node, resObj, zip);
    }
    if (key === 'p:pic') {
        return await picToImage(node, resObj, zip);
    }
    // p:sp / p:cxnSp → text 或 shape
    const spPr = node['p:spPr'];
    const { paragraphs, hasText } = extractTextBody(node);
    const geom = spPr && spPr['a:prstGeom'];

    if (hasText || (node['p:nvSpPr'] && node['p:nvSpPr']['p:cNvSpPr'] && node['p:nvSpPr']['p:cNvSpPr'].attrs && node['p:nvSpPr']['p:cNvSpPr'].attrs.txBox === '1')) {
        // 文本元素
        const xf = readXfrm(node, false);
        const name = node['p:nvSpPr'] && node['p:nvSpPr']['p:cNvPr'] && node['p:nvSpPr']['p:cNvPr'].attrs && node['p:nvSpPr']['p:cNvPr'].attrs.name;
        const textEl: PptxTextElement = {
            type: 'text',
            x: xf ? xf.x : 0,
            y: xf ? xf.y : 0,
            width: xf ? xf.width : 300,
            height: xf ? xf.height : 60,
            paragraphs,
            __raw: node
        };
        if (xf && xf.rotation) textEl.rotation = xf.rotation;
        if (name) textEl.name = String(name);
        // 元素级默认对齐（取首段）
        if (paragraphs[0]) {
            if (paragraphs[0].align) textEl.align = paragraphs[0].align;
            if ((paragraphs[0] as any).valign) textEl.valign = (paragraphs[0] as any).valign;
        }
        // 将首个 run 的统一样式镜像到元素级（jsonToPptx 以元素级为默认，提升 round-trip 保真）
        const firstRun = paragraphs[0] && paragraphs[0].runs && paragraphs[0].runs[0];
        if (firstRun) {
            if (firstRun.fontSize !== undefined) textEl.fontSize = firstRun.fontSize;
            if (firstRun.color !== undefined) textEl.color = firstRun.color;
            if (firstRun.bold !== undefined) textEl.bold = firstRun.bold;
            if (firstRun.italic !== undefined) textEl.italic = firstRun.italic;
            if (firstRun.underline !== undefined) textEl.underline = firstRun.underline;
            if (firstRun.fontFace !== undefined) textEl.fontFace = firstRun.fontFace;
        }
        return textEl;
    }

    if (geom) {
        // 形状元素
        const xf = readXfrm(node, false);
        const sp = readSpPr(spPr);
        const name = node['p:nvSpPr'] && node['p:nvSpPr']['p:cNvPr'] && node['p:nvSpPr']['p:cNvPr'].attrs && node['p:nvSpPr']['p:cNvPr'].attrs.name;
        const shapeEl: PptxShapeElement = {
            type: 'shape',
            shapeType: sp.shapeType,
            x: xf ? xf.x : 0,
            y: xf ? xf.y : 0,
            width: xf ? xf.width : 200,
            height: xf ? xf.height : 120,
            __raw: node
        };
        if (sp.fill !== undefined) shapeEl.fill = sp.fill;
        if (sp.line !== undefined) shapeEl.line = sp.line;
        if (xf && xf.rotation) shapeEl.rotation = xf.rotation;
        if (name) shapeEl.name = String(name);
        return shapeEl;
    }

    // 其余（无文本的占位/连接符等）：跳过，避免噪音
    return null;
}

/** p:graphicFrame → chart 元素 */
async function graphicFrameToChart(
    node: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip
): Promise<PptxElement | null> {
    const chartRef = node['a:graphic']
        && node['a:graphic']['a:graphicData']
        && node['a:graphic']['a:graphicData']['c:chart'];
    if (!chartRef || !chartRef.attrs || !chartRef.attrs['r:id']) return null;

    const rid = String(chartRef.attrs['r:id']);
    const target = resObj[rid] && resObj[rid].target;
    const part = resolvePart(target);
    let chartSemantic: Partial<PptxChartElement> | undefined;
    if (part) {
        const chartXml = await PPTXXmlUtils.readXmlFile(zip, part);
        chartSemantic = extractChart(chartXml);
    }

    const xf = readXfrm(node, true);
    const chartEl: PptxChartElement = {
        type: 'chart',
        chartType: (chartSemantic && chartSemantic.chartType) || 'barChart',
        x: xf ? xf.x : 0,
        y: xf ? xf.y : 0,
        width: xf ? xf.width : 600,
        height: xf ? xf.height : 400,
        series: (chartSemantic && chartSemantic.series) || [],
        __raw: node
    };
    if (chartSemantic && chartSemantic.categories) chartEl.categories = chartSemantic.categories;
    if (chartSemantic && chartSemantic.title) chartEl.title = chartSemantic.title;
    if (chartSemantic && chartSemantic.legend !== undefined) chartEl.legend = chartSemantic.legend;
    return chartEl;
}

/** p:pic → image 元素（内联 dataURL） */
async function picToImage(
    node: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip
): Promise<PptxElement | null> {
    const blip = node['p:blipFill'] && node['p:blipFill']['a:blip'];
    if (!blip || !blip.attrs || !blip.attrs['r:embed']) return null;
    const rid = String(blip.attrs['r:embed']);
    const target = resObj[rid] && resObj[rid].target;
    const part = resolvePart(target);
    if (!part) return null;

    const xf = readXfrm(node, false);
    const name = node['p:nvPicPr'] && node['p:nvPicPr']['p:cNvPr'] && node['p:nvPicPr']['p:cNvPr'].attrs && node['p:nvPicPr']['p:cNvPr'].attrs.name;

    const ext = (part.split('.').pop() || 'png').toLowerCase();
    const mimeMap: Record<string, string> = {
        png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml'
    };
    const mime = mimeMap[ext] || 'application/octet-stream';

    let data: string | undefined;
    try {
        const file = zip.file(part);
        if (file) {
            const b64 = await file.async('base64');
            data = `data:${mime};base64,${b64}`;
        }
    } catch { /* 忽略媒体读取失败 */ }

    const imgEl: PptxImageElement = {
        type: 'image',
        x: xf ? xf.x : 0,
        y: xf ? xf.y : 0,
        width: xf ? xf.width : 300,
        height: xf ? xf.height : 200,
        extension: ext,
        __raw: node
    };
    if (data) imgEl.data = data;
    else if (target) (imgEl as any).src = target;
    if (xf && xf.rotation) imgEl.rotation = xf.rotation;
    if (name) imgEl.name = String(name);
    return imgEl;
}

/**
 * 将 pptxToJson 的解析结果（parsedData）构建为标准 PptxDocument
 * @param parsedData - processToJson 返回的 parsedData（含 slides/data、slideSize、metadata）
 * @param zip - JSZip 实例
 */
export async function buildStandardDocument(parsedData: any, zip: JSZip): Promise<PptxDocument> {
    const slides = await Promise.all(
        (parsedData.slides || []).map(async (s: any) => extractSlideToStandard(s.data, zip))
    );

    const doc: PptxDocument = {
        version: '1.0',
        slideSize: parsedData.slideSize,
        slides
    };
    if (parsedData.metadata) doc.metadata = parsedData.metadata;
    return doc;
}
