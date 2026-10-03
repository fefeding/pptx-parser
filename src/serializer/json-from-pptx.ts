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
    PptxTransition, PptxBackground, PptxTableElement, PptxTableRow, PptxTableCell,
    PptxDiagramElement, PptxRawElement, PptxAnimation, PptxGroupElement, TextAlign, VAlign, ChartGrouping
} from '../types/pptx-document';

/** plotArea 下可能出现的全部图表节点（ECMA-376 全集），用于反向提取时判定图表类型 */
const CHART_PLOT_TYPES = [
    'c:barChart', 'c:bar3DChart',
    'c:lineChart', 'c:line3DChart',
    'c:areaChart', 'c:area3DChart',
    'c:pieChart', 'c:pie3DChart', 'c:doughnutChart', 'c:ofPieChart',
    'c:scatterChart', 'c:bubbleChart',
    'c:radarChart', 'c:stockChart',
    'c:surfaceChart', 'c:surface3DChart'
];

/** 1pt = 12700 EMU */
const EMU_PER_PT = 12700;
/** 线宽 EMU 默认值（缺省按 1pt 处理时参考） */
const DEFAULT_LN_PT = 1;
/** 段落对齐映射（OOXML algn → 标准 align） */
const ALIGN_MAP: Record<string, TextAlign> = { l: 'left', ctr: 'center', r: 'right', just: 'justify' };
/** 垂直对齐映射（OOXML anchor → 标准 valign） */
const VALIGN_MAP: Record<string, VAlign> = { t: 'top', ctr: 'middle', b: 'bottom' };
/** graphicData uri 中的表格/图示标识 */
const URI_TABLE = 'http://schemas.openxmlformats.org/drawingml/2006/table';
const URI_DIAGRAM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

/** 将任意值规整为数组（tXml 单节点即对象，多节点为数组） */
function asArray<T = any>(v: T | T[] | undefined): T[] {
    if (v === undefined || v === null) return [];
    return Array.isArray(v) ? v : [v];
}

/** 取节点的首个同名子节点（tXml 中重复节点为数组） */
function firstChild(parent: any, tag: string): any {
    if (!parent || typeof parent !== 'object') return null;
    const v = parent[tag];
    if (v === undefined || v === null) return null;
    return Array.isArray(v) ? v[0] : v;
}

/** 读取节点属性中的数值（缺失/空/非法返回 undefined） */
function readNumAttr(node: any, attr: string): number | undefined {
    const raw = node && node.attrs ? node.attrs[attr] : undefined;
    if (raw === undefined || raw === null || raw === '') return undefined;
    const n = Number(raw);
    return isFinite(n) ? n : undefined;
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

/**
 * 关系目标 → zip 内绝对部件路径（ppt/...）
 * 兼容 ../media/x.png、media/x.png、/ppt/media/x.png 等写法；外部链接返回 undefined。
 * @param {string} target - 关系目标
 * @returns {string|undefined} 绝对部件路径
 */
function resolvePart(target: string | undefined): string | undefined {
    if (!target) return undefined;
    if (/^https?:/i.test(target)) return undefined;
    if (target.startsWith('ppt/')) return target;
    return `ppt/${target.replace(/^(\.\.\/)+/, '').replace(/^\/+/, '')}`;
}

/** 从节点读取 a:srgbClr 的颜色值 */
function readSrgbClr(node: any): string | undefined {
    const c = node && (node['a:srgbClr'] || (node['a:solidFill'] && node['a:solidFill']['a:srgbClr']));
    return c && c.attrs && c.attrs.val ? String(c.attrs.val) : undefined;
}

/**
 * 读取文本运行中的文本（a:t 在 simplify 形态下为字符串）
 * 注：<a:t/> 空元素或仅带属性时解析结果非字符串，此时按空文本处理，
 * 避免 String(对象) 产生字面量 "[object Object]"。
 */
function readRunText(runNode: any): string {
    const t = runNode && runNode['a:t'];
    return typeof t === 'string' ? t : '';
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

/**
 * 提取 txBody（p:txBody 或表格单元格 a:txBody）为正文段落
 * @returns paragraphs 段落列表；hasText 是否含文本；valign 文本体垂直对齐；text 纯文本拼接
 */
function extractTxBody(txBody: any): { paragraphs: PptxParagraph[]; hasText: boolean; valign?: VAlign; text: string; textDirection?: string } {
    const paragraphs: PptxParagraph[] = [];
    let hasText = false;
    let text = '';
    if (!txBody) return { paragraphs, hasText, text };

    const bodyPr = txBody['a:bodyPr'];
    const valign = bodyPr && bodyPr.attrs && bodyPr.attrs.anchor
        ? VALIGN_MAP[bodyPr.attrs.anchor] : undefined;
    const textDirection = bodyPr && bodyPr.attrs && bodyPr.attrs.vert
        ? String(bodyPr.attrs.vert) : undefined;

    for (const pNode of asArray(txBody['a:p'])) {
        const pPr = pNode['a:pPr'];
        const pAttrs = (pPr && pPr.attrs) || {};
        const align = pAttrs.algn ? ALIGN_MAP[pAttrs.algn] : undefined;

        // 列表样式：自动编号 / 项目符号
        let bullet: any;
        if (pPr && pPr['a:buAutoNum'] && !pPr['a:buNone']) {
            const auto = pPr['a:buAutoNum'];
            bullet = { type: 'number', fmt: (auto.attrs && auto.attrs.type) || 'arabic', start: (auto.attrs && auto.attrs.startAt != null) ? Number(auto.attrs.startAt) : 1 };
        } else if (pPr && pPr['a:buChar'] && !pPr['a:buNone']) {
            const bc = pPr['a:buChar'];
            bullet = { type: 'bullet', char: (bc.attrs && bc.attrs.char) || '•' };
        }

        // 行距 / 段间距 / 缩进
        let lineSpacing: any;
        const lnSpc = pPr && pPr['a:lnSpc'];
        if (lnSpc) {
            if (lnSpc['a:spcPct']) lineSpacing = { type: 'percent', value: Number(lnSpc['a:spcPct'].attrs.val) / 1000 };
            else if (lnSpc['a:spcPts']) lineSpacing = { type: 'pt', value: Number(lnSpc['a:spcPts'].attrs.val) / 100 };
        }
        const spcBef = pPr && pPr['a:spcBef'] && pPr['a:spcBef']['a:spcPts'];
        const spcAft = pPr && pPr['a:spcAft'] && pPr['a:spcAft']['a:spcPts'];
        const spaceBefore = spcBef ? Number(spcBef.attrs.val) / 100 : undefined;
        const spaceAfter = spcAft ? Number(spcAft.attrs.val) / 100 : undefined;
        const indentLeft = pAttrs.marL != null ? emuToPt(pAttrs.marL) : undefined;
        const indentRight = pAttrs.marR != null ? emuToPt(pAttrs.marR) : undefined;
        const indent = pAttrs.indent != null ? emuToPt(pAttrs.indent) : undefined;

        const runs: PptxTextRun[] = [];
        let paraText = '';
        for (const runNode of asArray(pNode['a:r'])) {
            const t = readRunText(runNode);
            if (t) hasText = true;
            paraText += t;
            runs.push({ text: t, ...readRunStyle(runNode['a:rPr']) });
        }
        if (paraText) text += (text ? '\n' : '') + paraText;

        const para: PptxParagraph = { runs };
        if (align) para.align = align;
        if (bullet) para.bullet = bullet;
        if (lineSpacing) para.lineSpacing = lineSpacing;
        if (spaceBefore != null) para.spaceBefore = spaceBefore;
        if (spaceAfter != null) para.spaceAfter = spaceAfter;
        if (indentLeft != null) para.indentLeft = indentLeft;
        if (indentRight != null) para.indentRight = indentRight;
        if (indent != null) para.indent = indent;
        if (valign) (para as any).valign = valign;
        paragraphs.push(para);
    }
    return { paragraphs, hasText, valign, text, textDirection };
}

/** 提取一个 p:sp 的文本为正文段落 */
function extractTextBody(spNode: any): { paragraphs: PptxParagraph[]; hasText: boolean; textDirection?: string } {
    const { paragraphs, hasText, textDirection } = extractTxBody(spNode && spNode['p:txBody']);
    return { paragraphs, hasText, textDirection };
}

/** 从 p:spPr 读取几何/填充/边框/特效 */
function readSpPr(spPr: any): Pick<PptxShapeElement, 'shapeType' | 'fill' | 'line' | 'effects'> {
    const out: any = { shapeType: 'rect' };
    if (!spPr) return out;
    const prst = spPr['a:prstGeom'];
    if (prst && prst.attrs && prst.attrs.prst) out.shapeType = String(prst.attrs.prst);

    // 填充
    if (spPr['a:noFill']) {
        out.fill = 'none';
    } else if (spPr['a:gradFill']) {
        const gf = spPr['a:gradFill'];
        const gsLst = gf['a:gsLst'];
        const stops = (gsLst ? asArray(gsLst['a:gs']) : []).map((gs: any) => {
            const pos = gs && gs.attrs ? Number(gs.attrs.pos) / 100000 : 0;
            return { color: readSrgbClr(gs) || '#000000', position: pos };
        });
        const lin = gf['a:lin'];
        let direction: 'horizontal' | 'vertical' | 'diagonal' = 'horizontal';
        if (lin && lin.attrs) {
            const ang = Number(lin.attrs.ang) / 60000; // 度
            direction = ang >= 45 && ang < 135 ? 'vertical' : (ang >= 22.5 && ang < 67.5 ? 'diagonal' : 'horizontal');
        }
        out.fill = { type: 'gradient', direction, stops };
    } else {
        const color = readSrgbClr(spPr);
        if (color) {
            const solid = spPr['a:solidFill'];
            const srgb = solid && solid['a:srgbClr'];
            const alphaNode = (solid && solid['a:alpha']) || (srgb && srgb['a:alpha']);
            const transparency = alphaNode && alphaNode.attrs ? Math.round(100 - Number(alphaNode.attrs.val) / 1000) : undefined;
            out.fill = transparency != null ? { type: 'solid', color, transparency } : color;
        }
    }

    // 边框
    const ln = spPr['a:ln'];
    if (ln) {
        if (ln['a:noFill']) {
            out.line = 'none';
        } else {
            const color = readSrgbClr(ln);
            const w = ln.attrs && ln.attrs.w ? emuToPt(ln.attrs.w) : DEFAULT_LN_PT;
            const srgb = ln['a:solidFill'] && ln['a:solidFill']['a:srgbClr'];
            const alphaNode = srgb && srgb['a:alpha'];
            const transparency = alphaNode && alphaNode.attrs ? Math.round(100 - Number(alphaNode.attrs.val) / 1000) : undefined;
            const dash = ln['a:prstDash'] && ln['a:prstDash'].attrs && ln['a:prstDash'].attrs.val;
            const lineObj: any = { width: w };
            if (color) lineObj.color = color;
            if (transparency != null) lineObj.transparency = transparency;
            if (dash) lineObj.dashType = String(dash);
            out.line = lineObj;
        }
    }

    // 特效（a:effectLst：阴影 / 发光）
    const effLst = spPr['a:effectLst'];
    if (effLst) {
        const effects: any = {};
        const inner = effLst['a:innerShdw'];
        const outer = effLst['a:outerShdw'] || inner;
        if (outer) {
            const shadow: any = { type: inner ? 'inner' : 'outer' };
            if (outer.attrs) {
                // OOXML 属性名为 blurRad（旧文件可能写作 blur，做兼容兜底）
                const blurAttr = outer.attrs.blurRad != null ? outer.attrs.blurRad : outer.attrs.blur;
                if (blurAttr != null) shadow.blur = emuToPt(blurAttr);
                if (outer.attrs.dist != null) shadow.distance = emuToPt(outer.attrs.dist);
                if (outer.attrs.dir != null) shadow.angle = Math.round(Number(outer.attrs.dir) / 60000);
            }
            const srgb = outer['a:srgbClr'];
            if (srgb && srgb.attrs && srgb.attrs.val) shadow.color = String(srgb.attrs.val);
            const alphaNode = srgb && srgb['a:alpha'];
            if (alphaNode && alphaNode.attrs) shadow.transparency = Math.round(100 - Number(alphaNode.attrs.val) / 1000);
            effects.shadow = shadow;
        }
        const glow = effLst['a:glow'];
        if (glow) {
            const g: any = {};
            // OOXML 属性名为 rad（旧文件可能写作 blur，做兼容兜底）
            const glowRad = glow.attrs ? (glow.attrs.rad != null ? glow.attrs.rad : glow.attrs.blur) : undefined;
            if (glowRad != null) g.blur = emuToPt(glowRad);
            const srgb = glow['a:srgbClr'];
            if (srgb && srgb.attrs && srgb.attrs.val) g.color = String(srgb.attrs.val);
            effects.glow = g;
        }
        if (effects.shadow || effects.glow) out.effects = effects;
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

/** 取形状节点的 cNvPr id（= OOXML 形状 spid） */
function getShapeId(key: string, node: any): string | undefined {
    const nvKey: Record<string, string> = {
        'p:sp': 'p:nvSpPr', 'p:pic': 'p:nvPicPr', 'p:graphicFrame': 'p:nvGraphicFramePr',
        'p:cxnSp': 'p:nvCxnSpPr', 'p:grpSp': 'p:nvGrpSpPr'
    };
    const nv = node && nvKey[key] && node[nvKey[key]];
    const cNvPr = nv && nv['p:cNvPr'];
    return cNvPr && cNvPr.attrs && cNvPr.attrs.id != null ? String(cNvPr.attrs.id) : undefined;
}

/** 解析 p:timing：自动播放 afterTime + 元素进入动画 p:spTgt */
function extractTiming(slideContent: any, spidToIndex: Map<string, number>): { advanceTime?: number; animations?: PptxAnimation[] } {
    const sld = slideContent && slideContent['p:sld'];
    const timing = sld && sld['p:timing'];
    if (!timing) return {};
    const result: { advanceTime?: number; animations?: PptxAnimation[] } = {};
    const animations: PptxAnimation[] = [];

    // 递归收集所有 p:cTn
    const allCtn: any[] = [];
    (function walk(n: any) {
        if (!n || typeof n !== 'object') return;
        for (const k of Object.keys(n)) {
            if (k === 'p:cTn') {
                for (const c of asArray(n[k])) { allCtn.push(c); walk(c); }
            } else {
                const v = n[k];
                if (Array.isArray(v)) v.forEach(walk);
                else if (v && typeof v === 'object') walk(v);
            }
        }
    })(timing);

    for (const c of allCtn) {
        // 自动播放：stCondLst 下存在 afterTime cond
        const stCondLst = c['p:stCondLst'];
        if (stCondLst) {
            for (const cond of asArray(stCondLst['p:cond'])) {
                if (cond && cond.attrs && cond.attrs.type === 'afterTime' && cond.attrs.val != null) {
                    result.advanceTime = Number(cond.attrs.val);
                }
            }
        }
        // 元素动画：tgtEl > spTgt spid + childTnLst > cTn preset/dur
        const tgtEl = c['p:tgtEl'];
        const spTgt = tgtEl && tgtEl['p:spTgt'];
        const spid = spTgt && spTgt.attrs && spTgt.attrs.spid;
        if (spid != null && spidToIndex.has(String(spid))) {
            let preset = 'fade';
            let dur: number | undefined;
            let presetClass: string | undefined;
            let presetId: number | undefined;
            let delay: number | undefined;
            let repeat: number | 'indefinite' | undefined;
            const childTnLst = c['p:childTnLst'];
            if (childTnLst) {
                for (const inner of asArray(childTnLst['p:cTn'])) {
                    if (inner && inner.attrs) {
                        // preset 透传真实名称（不再收敛为 4 种，避免 swivel/bounce 等退化为 fade）
                        if (inner.attrs.preset) preset = String(inner.attrs.preset);
                        if (inner.attrs.dur != null) dur = Number(inner.attrs.dur) / 1000;
                        if (inner.attrs.presetClass) presetClass = String(inner.attrs.presetClass);
                        if (inner.attrs.presetId != null) presetId = Number(inner.attrs.presetId);
                        if (inner.attrs.delay != null && inner.attrs.delay !== 'indefinite') delay = Number(inner.attrs.delay) / 1000;
                        if (inner.attrs.repeatCount != null) {
                            repeat = inner.attrs.repeatCount === 'indefinite' ? 'indefinite' : Number(inner.attrs.repeatCount) / 1000;
                        }
                    }
                }
            }
            const anim: PptxAnimation = {
                target: spidToIndex.get(String(spid))!,
                type: preset,
                duration: dur ?? 1
            };
            if (presetClass) anim.presetClass = presetClass as PptxAnimation['presetClass'];
            if (presetId != null) anim.presetId = presetId;
            if (delay != null) anim.delay = delay;
            if (repeat != null) anim.repeat = repeat;
            animations.push(anim);
        }
    }
    if (animations.length) result.animations = animations;
    return result;
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
            const lin = bgPr['a:gradFill']['a:lin'];
            const ang = lin && lin.attrs && lin.attrs.ang !== undefined ? Number(lin.attrs.ang) : 0;
            const direction = ang === 90 ? 'vertical' : ang === 45 ? 'diagonal' : 'horizontal';
            return { type: 'gradient', direction, stops };
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

        // 图表类型：取 plotArea 下首个图表节点（组合图取第一个 plot）
        let chartNode: any = null;
        let chartType = 'barChart';
        for (const ct of CHART_PLOT_TYPES) {
            const node = asArray(plotArea[ct])[0];
            if (node) { chartNode = node; chartType = ct.replace('c:', ''); break; }
        }
        if (!chartNode) return undefined;

        const isScatter = chartType === 'scatterChart';
        const isBubble = chartType === 'bubbleChart';
        const isStock = chartType === 'stockChart';

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
            } else if (isBubble) {
                // 气泡图：x/y 为坐标，values 复用为气泡大小（与生成端一致）
                s.x = numCacheValues(ser['c:xVal']);
                s.y = numCacheValues(ser['c:yVal']);
                s.values = numCacheValues(ser['c:bubbleSize']);
            } else if (isStock) {
                s.open = numCacheValues(ser['c:openVal']);
                s.high = numCacheValues(ser['c:highVal']);
                s.low = numCacheValues(ser['c:lowVal']);
                s.close = numCacheValues(ser['c:closeVal']);
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

        // 读取子节点属性值的快捷方式
        const attrOf = (parent: any, tag: string): string | undefined => {
            const n = parent && parent[tag];
            return n && n.attrs ? n.attrs.val : undefined;
        };

        // 分组/堆叠与方向
        const grouping = attrOf(chartNode, 'c:grouping');
        if (grouping) out.grouping = grouping as ChartGrouping;
        const barDir = attrOf(chartNode, 'c:barDir');
        if (barDir === 'bar' || barDir === 'col') out.barDir = barDir;
        const varyColors = attrOf(chartNode, 'c:varyColors');
        if (varyColors !== undefined) out.varyColors = varyColors === '1';

        // 甜甜圈内径 / 子母饼图类型
        const holeSize = attrOf(chartNode, 'c:holeSize');
        if (holeSize !== undefined && holeSize !== '') out.holeSize = Number(holeSize);
        const ofPieType = attrOf(chartNode, 'c:ofPieType');
        if (ofPieType === 'pie' || ofPieType === 'bar') out.ofPieType = ofPieType;

        // 平滑线 / 数据标记（取自首个系列）
        const smoothVal = attrOf(firstSer, 'c:smooth');
        if (smoothVal !== undefined) out.smooth = smoothVal !== '0';
        const markerNode = firstSer && firstSer['c:marker'];
        if (markerNode) {
            const symbol = attrOf(markerNode, 'c:symbol');
            out.marker = symbol !== 'none';
        }

        // 数字格式：数值轴优先，其次数据标签
        const valAx = asArray(plotArea['c:valAx'])[0];
        const numFmt = (valAx && valAx['c:numFmt'] && valAx['c:numFmt'].attrs && valAx['c:numFmt'].attrs.formatCode)
            || (chartNode['c:dLbls'] && chartNode['c:dLbls']['c:numFmt']
                && chartNode['c:dLbls']['c:numFmt'].attrs && chartNode['c:dLbls']['c:numFmt'].attrs.formatCode);
        if (numFmt) out.numberFormat = String(numFmt);

        // 气泡图属性
        if (isBubble) {
            const b3d = attrOf(chartNode, 'c:bubble3D');
            const negB = attrOf(chartNode, 'c:showNegBubbles');
            const scale = attrOf(chartNode, 'c:bubbleScale');
            if (b3d !== undefined) out.bubble3D = b3d === '1';
            if (negB !== undefined) out.showNegBubbles = negB === '1';
            if (scale !== undefined) out.bubbleScale = Number(scale);
        }

        // 曲面图线框
        const wireframe = attrOf(chartNode, 'c:wireframe');
        if (wireframe !== undefined) out.wireframe = wireframe === '1';

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

/** 语义层已支持的类型（其余类型需靠 __raw 回退，故默认要附带关系与部件依赖） */
const SEMANTIC_TYPES = new Set(['text', 'shape', 'image', 'chart', 'table']);

/**
 * 标准 JSON 解析选项
 * rawDeps 控制 __raw 依赖（rels/parts）的携带范围：
 * - 'auto'（默认）：仅为语义层不支持的类型附带，避免 JSON 体积膨胀；
 * - 'all'：为所有元素附带，使语义类型也能用 rawFallback 无损回写。
 */
export interface StandardExtractOptions {
    rawDeps?: 'auto' | 'all';
}

/**
 * OOXML 关系类型前缀
 * 注：slideResObj 中的 type 为去掉该前缀的短名（如 diagramData/image/chart），
 * 写入 slide rels 时需还原为完整 URI。
 */
const REL_PREFIX = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

/** 关系类型短名 → 完整 URI（已是完整 URI 时原样返回） */
function normalizeRelType(type: string): string {
    if (!type) return '';
    return type.includes('/') ? type : REL_PREFIX + type;
}

/** 关系类型 URI → 部件 Content-Type */
const REL_CONTENT_TYPE: Record<string, string> = {
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData':
        'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout':
        'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle':
        'application/vnd.openxmlformats-officedocument.drawingml.diagramQuickStyle+xml',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors':
        'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart':
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
};

/** 递归收集节点内所有 r:* 关系引用 id（r:dm/r:embed/r:id/r:link...） */
function collectRelIds(node: any, acc: Set<string>) {
    if (!node || typeof node !== 'object') return;
    for (const key of Object.keys(node)) {
        if (key === 'attrs') {
            const attrs = node.attrs || {};
            for (const a of Object.keys(attrs)) {
                if (!a.startsWith('r:')) continue;
                const v = String(attrs[a]);
                if (/^rId\d+$/.test(v)) acc.add(v);
            }
            continue;
        }
        for (const child of asArray(node[key])) {
            if (child && typeof child === 'object') collectRelIds(child, acc);
        }
    }
}

/**
 * 为 __raw 载荷补充关系与部件依赖，使其可在无源 PPTX 的情况下独立回写
 * @param {Object} el - 已提取的元素（其 __raw 需已就位）
 * @param {Object} node - 原始 OOXML 节点
 * @param {Object} resObj - 幻灯片关系表（rId → { type, target }）
 * @param {Object} zip - JSZip 实例
 */
async function attachRawDeps(
    el: any,
    node: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip
) {
    const relIds = new Set<string>();
    collectRelIds(node, relIds);
    if (!relIds.size) return;

    const rels: Record<string, { type: string; target: string; external?: boolean }> = {};
    const parts: { path: string; content?: string; base64?: string; contentType: string; media?: boolean }[] = [];

    for (const rid of relIds) {
        const rel = resObj[rid];
        if (!rel || !rel.target) continue;
        // resObj.type 为短名（diagramData/image/...），统一还原为完整 URI 便于回写
        const type = normalizeRelType(rel.type || '');
        const external = /^https?:/i.test(rel.target) || /(^|\/)hyperlink$/.test(type);
        rels[rid] = { type, target: rel.target };
        if (external) {
            rels[rid].external = true;
            continue;
        }
        // resObj.target 已是 zip 内绝对路径（ppt/...），resolvePart 对已绝对路径幂等
        const partPath = resolvePart(rel.target);
        if (!partPath) continue;
        const file = zip.file(partPath);
        if (!file) continue;

        const isMedia = /(image|video|audio|media)$/.test(type);
        const contentType = REL_CONTENT_TYPE[type]
            || `application/vnd.openxmlformats-officedocument.${(partPath.split('.').pop() || 'xml')}`;
        try {
            if (isMedia) {
                parts.push({ path: partPath, base64: await file.async('base64'), contentType, media: true });
            } else {
                parts.push({ path: partPath, content: await file.async('string'), contentType });
            }
        } catch { /* 部件读取失败则跳过，回退时该引用会悬空 */ }
    }

    if (Object.keys(rels).length) el.__raw.rels = rels;
    if (parts.length) el.__raw.parts = parts;
}

/**
 * 递归收集 spTree 下的图形节点。
 * @param keepGroups - true 时把 p:grpSp 也作为条目保留（供语义层产出 group 元素）；
 *                    false 时展开 group（兼容旧调用方 / HTML 渲染链路）。
 */
function collectShapeNodes(spTree: any, acc: any[], keepGroups = false) {
    if (!spTree || typeof spTree !== 'object') return;
    for (const key of Object.keys(spTree)) {
        const val = spTree[key];
        if (val === undefined || val === null) continue;
        const nodes = asArray(val);
        for (const node of nodes) {
            if (key === 'p:grpSp') {
                if (keepGroups) {
                    acc.push({ key, node });
                } else {
                    const inner = node && node['p:spTree'];
                    if (inner) collectShapeNodes(inner, acc, keepGroups);
                }
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
export async function extractSlideToStandard(
    slideData: any,
    zip: JSZip,
    options: StandardExtractOptions = {}
): Promise<PptxSlide> {
    const slide: PptxSlide = { elements: [] };
    const allDeps = options.rawDeps === 'all';

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
        // 批注（解析端已从 ppt/comments/commentsN.xml 解析并挂在 slideData.comments）
        const comments = slideData && slideData.comments;
        if (Array.isArray(comments) && comments.length) {
            slide.comments = comments.map((c: any) => ({
                author: c.author,
                text: c.text || '',
                dt: c.dt,
                pos: (c.x != null || c.y != null) ? { x: c.x, y: c.y } : undefined
            }));
        }

        if (spTree) {
            // spid → 元素扁平序号映射（用于动画 p:spTgt 回指）。
            // 序号按 DFS 顺序递增（含 group 自身及其子孙），与生成端元素编号一致。
            const spidToIndex = new Map<string, number>();
            let elemIndex = 0;

            const registerSpid = (spid: string | undefined, el: PptxElement) => {
                if (spid != null) spidToIndex.set(spid, elemIndex);
                elemIndex++;
            };

            const processNode = async (key: string, node: any): Promise<PptxElement | null> => {
                const spid = getShapeId(key, node);
                try {
                    const el = await nodeToElement(key, node, resObj, zip);
                    if (!el) return null;
                    // 统一挂载 __raw 载荷（含标签名，供生成端无损回写）
                    (el as any).__raw = { tag: key, node };
                    // 依赖携带：语义层未覆盖的类型必须附带，否则 __raw 无法独立回写；
                    // 语义类型仅在 rawDeps:'all' 时附带（供 rawFallback 使用）
                    if (allDeps || !SEMANTIC_TYPES.has(el.type)) {
                        await attachRawDeps(el, node, resObj, zip);
                    }
                    registerSpid(spid, el);
                    return el;
                } catch {
                    // 单元素失败：仅保留原始节点，由 __raw 回退承载
                    const rawEl: PptxRawElement = {
                        type: 'raw', x: 0, y: 0, width: 0, height: 0,
                        __raw: { tag: key, node }, rawFallback: true
                    };
                    registerSpid(spid, rawEl);
                    if (allDeps) await attachRawDeps(rawEl, node, resObj, zip);
                    return rawEl;
                }
            };

            // 处理一个 spTree → 元素数组（递归保留 group 层级）
            const processTree = async (tree: any): Promise<PptxElement[]> => {
                const acc: { key: string; node: any }[] = [];
                collectShapeNodes(tree, acc, true);
                const out: PptxElement[] = [];
                for (const { key, node } of acc) {
                    if (key === 'p:grpSp') {
                        const g = await processGroup(node);
                        if (g) out.push(g);
                    } else {
                        const el = await processNode(key, node);
                        if (el) out.push(el);
                    }
                }
                return out;
            };

            // 组合：将 group 内部子元素坐标由「局部（chOff/chExt 空间）」转换为
            // 「相对 group 左上角的偏移」（childrenCoordinates='relative'），生成端据此重建。
            const processGroup = async (node: any): Promise<PptxGroupElement | null> => {
                const gid = getShapeId('p:grpSp', node);
                const inner = node['p:spTree'];
                const children = inner ? await processTree(inner) : [];

                const gxf = node['p:grpSpPr'] && node['p:grpSpPr']['a:xfrm'];
                const gOff = gxf && gxf['a:off'] && gxf['a:off'].attrs;
                const gExt = gxf && gxf['a:ext'] && gxf['a:ext'].attrs;
                const chOff = gxf && gxf['a:chOff'] && gxf['a:chOff'].attrs;
                const chExt = gxf && gxf['a:chExt'] && gxf['a:chExt'].attrs;
                const gx = emuToPx(gOff?.x), gy = emuToPx(gOff?.y);
                const gw = emuToPx(gExt?.cx), gh = emuToPx(gExt?.cy);
                const chx = emuToPx(chOff?.x), chy = emuToPx(chOff?.y);
                const chw = emuToPx(chExt?.cx) || 1, chh = emuToPx(chExt?.cy) || 1;
                const sx = gw / chw, sy = gh / chh;

                for (const c of children) {
                    const lx = c.x || 0, ly = c.y || 0;
                    c.x = (lx - chx) * sx;
                    c.y = (ly - chy) * sy;
                    if (c.width != null) c.width = c.width * sx;
                    if (c.height != null) c.height = c.height * sy;
                }
                const g: PptxGroupElement = {
                    type: 'group',
                    x: gx, y: gy, width: gw, height: gh,
                    children,
                    childrenCoordinates: 'relative'
                };
                registerSpid(gid, g as PptxElement);
                return g;
            };

            slide.elements = await processTree(spTree);

            // 自动播放 / 元素动画（p:timing）
            const timing = extractTiming(slideContent, spidToIndex);
            if (timing.advanceTime != null) slide.advanceTime = timing.advanceTime;
            if (timing.animations) slide.animations = timing.animations;
        }

        // 隐藏幻灯片（p:sld show="0"）
        const sldAttrs = slideContent && slideContent['p:sld'] && slideContent['p:sld'].attrs;
        if (sldAttrs && String(sldAttrs.show) === '0') slide.hidden = true;
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
        return await graphicFrameToElement(node, resObj, zip);
    }
    if (key === 'p:pic') {
        return await picToImage(node, resObj, zip);
    }
    // p:sp / p:cxnSp → text 或 shape
    const spPr = node['p:spPr'];
    const { paragraphs, hasText, textDirection } = extractTextBody(node);
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
            paragraphs
        };
        if (xf && xf.rotation) textEl.rotation = xf.rotation;
        if (name) textEl.name = String(name);
        if (textDirection) textEl.textDirection = textDirection;
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
            height: xf ? xf.height : 120
        };
        if (sp.fill !== undefined) shapeEl.fill = sp.fill;
        if (sp.line !== undefined) shapeEl.line = sp.line;
        if (sp.effects) shapeEl.effects = sp.effects;
        if (xf && xf.rotation) shapeEl.rotation = xf.rotation;
        if (name) shapeEl.name = String(name);
        return shapeEl;
    }

    // 其余（无文本的占位/连接符等）：跳过，避免噪音
    return null;
}

/** 递归收集节点下所有 a:t 文本（用于 SmartArt 数据部件） */
function collectTexts(node: any, acc: string[] = []): string[] {
    if (!node || typeof node !== 'object') return acc;
    for (const key of Object.keys(node)) {
        const val = node[key];
        if (val === undefined || val === null) continue;
        if (key === 'a:t') {
            for (const t of asArray(val)) {
                if (typeof t !== 'string') continue;
                if (t.trim()) acc.push(t);
            }
            continue;
        }
        for (const child of asArray(val)) {
            if (child && typeof child === 'object') collectTexts(child, acc);
        }
    }
    return acc;
}

/** 读取 graphicFrame 的名称（p:nvGraphicFramePr/p:cNvPr） */
function readGraphicFrameName(node: any): string | undefined {
    const cNvPr = node && node['p:nvGraphicFramePr'] && node['p:nvGraphicFramePr']['p:cNvPr'];
    const name = cNvPr && cNvPr.attrs && cNvPr.attrs.name;
    return name ? String(name) : undefined;
}

/** a:tbl（表格）→ PptxTableElement */
function tableToElement(tbl: any, node: any): PptxTableElement {
    const xf = readXfrm(node, true);
    const el: PptxTableElement = {
        type: 'table',
        x: xf ? xf.x : 0,
        y: xf ? xf.y : 0,
        width: xf ? xf.width : 400,
        height: xf ? xf.height : 200,
        rows: []
    };
    const name = readGraphicFrameName(node);
    if (name) el.name = name;

    // 列宽（a:tblGrid/a:gridCol）
    const grid = tbl && tbl['a:tblGrid'];
    const colWidths = asArray(grid && grid['a:gridCol'])
        .map((c: any) => (c && c.attrs && c.attrs.w ? emuToPx(c.attrs.w) : 0));
    if (colWidths.length) el.colWidths = colWidths;

    const rowHeights: number[] = [];
    for (const tr of asArray(tbl && tbl['a:tr'])) {
        const row: PptxTableRow = { cells: [] };
        if (tr && tr.attrs && tr.attrs.h) row.height = emuToPx(tr.attrs.h);

        for (const tc of asArray(tr && tr['a:tc'])) {
            const cell: PptxTableCell = {};
            const attrs = (tc && tc.attrs) || {};
            if (attrs.gridSpan && Number(attrs.gridSpan) > 1) cell.colSpan = Number(attrs.gridSpan);
            if (attrs.rowSpan && Number(attrs.rowSpan) > 1) cell.rowSpan = Number(attrs.rowSpan);

            const { paragraphs, text } = extractTxBody(tc && tc['a:txBody']);
            if (paragraphs.length > 1) {
                // 多段保留段落结构，避免丢段
                cell.paragraphs = paragraphs;
            } else {
                // 单段用 text 简写，并把 run 样式镜像到单元格级
                cell.text = text;
                const r0 = paragraphs[0] && paragraphs[0].runs && paragraphs[0].runs[0];
                if (r0) {
                    if (r0.fontSize !== undefined) cell.fontSize = r0.fontSize;
                    if (r0.color !== undefined) cell.color = r0.color;
                    if (r0.bold) cell.bold = true;
                    if (r0.italic) cell.italic = true;
                    if (r0.underline) cell.underline = true;
                    if (r0.fontFace !== undefined) cell.fontFace = r0.fontFace;
                }
            }
            if (paragraphs[0] && paragraphs[0].align) cell.align = paragraphs[0].align;

            // 单元格属性：底色、垂直对齐与边框
            const tcPr = tc && tc['a:tcPr'];
            if (tcPr) {
                const fill = readSrgbClr(tcPr);
                if (fill) cell.fill = fill;
                if (tcPr.attrs && tcPr.attrs.anchor) cell.valign = VALIGN_MAP[tcPr.attrs.anchor];

                // 边框回读（a:lnL / a:lnR / a:lnT / a:lnB）
                const edgeKey: Record<string, 'left' | 'right' | 'top' | 'bottom'> = { L: 'left', R: 'right', T: 'top', B: 'bottom' };
                const borders: Record<string, { color: string; width: number } | 'none'> = {};
                let hasBorder = false;
                for (const e of Object.keys(edgeKey)) {
                    const ln = tcPr['a:ln' + e];
                    if (!ln) continue;
                    // 显式无边框（<a:lnX><a:noFill/></a:lnX>）→ 回读为 'none'，避免与「未指定（继承表格样式）」混淆
                    if (ln['a:noFill']) {
                        borders[edgeKey[e]] = 'none';
                        hasBorder = true;
                        continue;
                    }
                    if (ln.attrs) {
                        const w = ln.attrs.w !== undefined ? Math.round(Number(ln.attrs.w) / EMU_PER_PT) : DEFAULT_LN_PT;
                        const color = readSrgbClr(ln) || '#000000';
                        borders[edgeKey[e]] = { color, width: w };
                        hasBorder = true;
                    }
                }
                if (hasBorder) cell.borders = borders;
            }
            row.cells.push(cell);
        }

        rowHeights.push(row.height !== undefined ? row.height : 0);
        el.rows.push(row);
    }
    if (rowHeights.length && rowHeights.every((h) => h > 0)) el.rowHeights = rowHeights;
    return el;
}

/** p:graphicFrame（SmartArt 图示）→ PptxDiagramElement：提取数据部件文本 */
async function diagramToElement(
    node: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip
): Promise<PptxDiagramElement> {
    const graphicData = node['a:graphic'] && node['a:graphic']['a:graphicData'];
    // OOXML 标准写法是 dgm:relIds（含 r:dm/r:cs/r:lo/r:qs），历史代码只认单数 dgm:rel
    const rel = graphicData && (graphicData['dgm:relIds'] || graphicData['dgm:rel']);
    const rid = rel && rel.attrs && (rel.attrs['r:dm'] || rel.attrs['r:id']);
    const target = rid ? resObj[String(rid)] && resObj[String(rid)].target : undefined;
    const part = resolvePart(target);

    let texts: string[] = [];
    if (part) {
        try {
            const dataXml = await PPTXXmlUtils.readXmlFile(zip, part);
            texts = collectTexts(dataXml);
        } catch { /* 数据部件缺失/解析失败则仅保留 __raw */ }
    }

    const xf = readXfrm(node, true);
    const el: PptxDiagramElement = {
        type: 'diagram',
        x: xf ? xf.x : 0,
        y: xf ? xf.y : 0,
        width: xf ? xf.width : 400,
        height: xf ? xf.height : 300
    };
    const name = readGraphicFrameName(node);
    if (name) el.name = name;
    if (texts.length) el.texts = texts;
    if (part) el.dataPath = part;
    return el;
}

/** p:graphicFrame → 按 graphicData 类型分派到 table / diagram / chart */
async function graphicFrameToElement(
    node: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip
): Promise<PptxElement | null> {
    const graphicData = node['a:graphic'] && node['a:graphic']['a:graphicData'];
    if (!graphicData) return null;
    const uri = graphicData.attrs && graphicData.attrs.uri ? String(graphicData.attrs.uri) : '';

    // 表格：以 a:tbl 实际存在为准（仅凭 uri 声明无法还原内容）
    const tbl = graphicData['a:tbl'];
    if (tbl) return tableToElement(tbl, node);
    // 图示 / SmartArt：dgm:rel 或其命名空间 uri
    if (graphicData['dgm:rel'] || uri === URI_DIAGRAM || /diagram/.test(uri)) {
        return await diagramToElement(node, resObj, zip);
    }
    return await graphicFrameToChart(node, resObj, zip);
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
        series: (chartSemantic && chartSemantic.series) || []
    };
    if (!chartSemantic) return chartEl;

    // 透传语义层已提取的图表属性（缺失字段保持不写，交由生成端取默认）
    const passKeys: (keyof PptxChartElement)[] = [
        'categories', 'title', 'legend', 'grouping', 'varyColors', 'barDir',
        'holeSize', 'smooth', 'marker', 'ofPieType', 'numberFormat',
        'bubble3D', 'showNegBubbles', 'bubbleScale', 'wireframe'
    ];
    const src = chartSemantic as unknown as Record<string, unknown>;
    const dst = chartEl as unknown as Record<string, unknown>;
    for (const k of passKeys) {
        const v = src[k as string];
        if (v !== undefined) dst[k as string] = v;
    }
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
    // 外链图片无包内部件，保留为 src 交由生成端下载；其余无法定位的则跳过
    if (!part && !/^https?:/i.test(String(target))) return null;

    const xf = readXfrm(node, false);
    const name = node['p:nvPicPr'] && node['p:nvPicPr']['p:cNvPr'] && node['p:nvPicPr']['p:cNvPr'].attrs && node['p:nvPicPr']['p:cNvPr'].attrs.name;

    // ===== 媒体（视频 / 音频）=====
    // 媒体在 OOXML 中仍是 p:pic，区别是 p:nvPr 下挂 a:videoFile / a:audioFile（r:link 指向媒体部件）。
    // 此前一律按图片解析，导致视频/音频在语义层退化为 image、往返丢失。
    const nvPr = node['p:nvPicPr'] && node['p:nvPicPr']['p:nvPr'];
    const mediaNode = nvPr && (nvPr['a:videoFile'] || nvPr['a:audioFile']);
    if (mediaNode) {
        const kind: 'video' | 'audio' = nvPr['a:videoFile'] ? 'video' : 'audio';
        const linkRid = mediaNode.attrs && mediaNode.attrs['r:link'] ? String(mediaNode.attrs['r:link']) : '';
        const mTarget = resObj[linkRid] && resObj[linkRid].target;
        const mPart = resolvePart(mTarget);
        const mExt = ((mPart || mTarget || '').split('.').pop() || (kind === 'video' ? 'mp4' : 'mp3')).toLowerCase();

        let mData: string | undefined;
        if (mPart) {
            try {
                const f = zip.file(mPart);
                if (f) mData = await f.async('base64');
            } catch { /* 忽略媒体读取失败 */ }
        }

        const mediaEl: any = {
            type: kind,
            x: xf ? xf.x : 0,
            y: xf ? xf.y : 0,
            width: xf ? xf.width : 300,
            height: xf ? xf.height : 200,
            extension: mExt
        };
        if (mData) mediaEl.data = mData;
        else if (/^https?:/i.test(String(mTarget))) mediaEl.src = mTarget;
        else return null;
        if (xf && xf.rotation) mediaEl.rotation = xf.rotation;
        if (name) mediaEl.name = String(name);

        // 预览图（poster）：媒体区域的显示帧，走 a:blip@r:embed
        const pRid = blip && blip.attrs && blip.attrs['r:embed'];
        if (pRid) {
            const pTarget = resObj[String(pRid)] && resObj[String(pRid)].target;
            const pPart = resolvePart(pTarget);
            if (pPart) {
                try {
                    const f = zip.file(pPart);
                    if (f) {
                        mediaEl.poster = {
                            data: await f.async('base64'),
                            extension: (pPart.split('.').pop() || 'png').toLowerCase()
                        };
                    }
                } catch { /* 忽略预览图读取失败 */ }
            }
        }
        return mediaEl as PptxElement;
    }

    const ext = ((part || target || '').split('.').pop() || 'png').toLowerCase();
    const mimeMap: Record<string, string> = {
        png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml'
    };
    const mime = mimeMap[ext] || 'application/octet-stream';

    let data: string | undefined;
    if (part) {
        try {
            const file = zip.file(part);
            if (file) {
                const b64 = await file.async('base64');
                data = `data:${mime};base64,${b64}`;
            }
        } catch { /* 忽略媒体读取失败 */ }
    }

    const imgEl: PptxImageElement = {
        type: 'image',
        x: xf ? xf.x : 0,
        y: xf ? xf.y : 0,
        width: xf ? xf.width : 300,
        height: xf ? xf.height : 200,
        extension: ext
    };
    if (data) {
        imgEl.data = data;
    } else if (/^https?:/i.test(String(target))) {
        // 外链图片：保留 URL 交由生成端下载
        (imgEl as any).src = target;
    } else {
        // 部件缺失且非外链（源文件不完整）：跳过该图片，避免生成端得到不可用的 src
        return null;
    }
    if (xf && xf.rotation) imgEl.rotation = xf.rotation;
    if (name) imgEl.name = String(name);

    // 裁剪：a:srcRect 是 p:blipFill 的子节点（与 a:blip 同级），千分比 → 0~1 比例
    const srcRect = firstChild(node['p:blipFill'], 'a:srcRect');
    if (srcRect) {
        const crop: { l?: number; r?: number; t?: number; b?: number } = {};
        for (const side of ['l', 'r', 't', 'b'] as const) {
            const v = readNumAttr(srcRect, side);
            if (v) crop[side] = v / 100000;
        }
        if (Object.keys(crop).length > 0) imgEl.crop = crop;
    }
    // 调整：a:lum（bright/contrast 属性）与 a:alphaModFix（amt 属性）是 a:blip 的子节点
    const lum = firstChild(blip, 'a:lum');
    const bright = readNumAttr(lum, 'bright');
    const contrast = readNumAttr(lum, 'contrast');
    const alphaFix = readNumAttr(firstChild(blip, 'a:alphaModFix'), 'amt');
    if (bright !== undefined || contrast !== undefined || alphaFix !== undefined) {
        const adj: { brightness?: number; contrast?: number; transparency?: number } = {};
        if (bright !== undefined) adj.brightness = bright / 1000;
        if (contrast !== undefined) adj.contrast = contrast / 1000;
        if (alphaFix !== undefined) adj.transparency = Math.round((100 - alphaFix / 1000) * 100) / 100;
        imgEl.imageAdjust = adj;
    }
    return imgEl;
}

/**
 * 将 pptxToJson 的解析结果（parsedData）构建为标准 PptxDocument
 * @param parsedData - processToJson 返回的 parsedData（含 slides/data、slideSize、metadata）
 * @param zip - JSZip 实例
 */
export async function buildStandardDocument(
    parsedData: any,
    zip: JSZip,
    options: StandardExtractOptions = {}
): Promise<PptxDocument> {
    const slides = await Promise.all(
        (parsedData.slides || []).map(async (s: any) => extractSlideToStandard(s.data, zip, options))
    );

    const doc: PptxDocument = {
        version: '1.0',
        slideSize: parsedData.slideSize,
        slides
    };
    if (parsedData.metadata) doc.metadata = parsedData.metadata;
    // 文档自定义属性（解析端已自 docProps/custom.xml 解析）
    if (parsedData.customProps && Object.keys(parsedData.customProps).length) {
        doc.customProps = parsedData.customProps;
    }
    return doc;
}
