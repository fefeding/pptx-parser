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
import TinyColor from 'tinycolor2';
import { SLIDE_FACTOR } from '../core/constants';
import { PPTXXmlUtils } from '../utils/xml';

/** tinycolor 工厂函数 */
const tinycolor = (color: any, opts?: any) => new TinyColor(color, opts);
import type {
    PptxDocument, PptxSlide, PptxElement, PptxTextElement, PptxShapeElement,
    PptxImageElement, PptxChartElement, PptxParagraph, PptxTextRun, PptxChartSeries,
    PptxTransition, PptxBackground, PptxTableElement, PptxTableRow, PptxTableCell,
    PptxDiagramElement, PptxDiagramShape, PptxRawElement, PptxAnimation, PptxGroupElement, TextAlign, VAlign, ChartGrouping,
    PptxFillImage, PptxTheme, PptxThemeColorScheme, PptxCustomGeometry, PptxGeometryPath, PptxGeometryCommand, PptxGradientFill
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

/** 从节点读取颜色：a:srgbClr → hex；a:schemeClr → 'scheme:<name>'（主题色引用，生成端 colorNode 支持） */
function readSrgbClr(node: any): string | undefined {
    const srgb = node && (node['a:srgbClr'] || (node['a:solidFill'] && node['a:solidFill']['a:srgbClr']));
    if (srgb && srgb.attrs && srgb.attrs.val) return String(srgb.attrs.val);
    const sch = node && (node['a:schemeClr'] || (node['a:solidFill'] && node['a:solidFill']['a:schemeClr']));
    if (sch && sch.attrs && sch.attrs.val) return 'scheme:' + String(sch.attrs.val);
    return undefined;
}

/** 读取填充节点（a:solidFill）的颜色（srgb/scheme 均可） */
function readSolidColor(solid: any): string | undefined {
    if (!solid) return undefined;
    const srgb = solid['a:srgbClr'];
    if (srgb && srgb.attrs && srgb.attrs.val) return String(srgb.attrs.val);
    const sch = solid['a:schemeClr'];
    if (sch && sch.attrs && sch.attrs.val) return 'scheme:' + String(sch.attrs.val);
    return undefined;
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

/** 读取运行级样式（a:rPr）：含文字描边 a:ln 与超链接 a:hlinkClick */
function readRunStyle(rPr: any, themeMap: Record<string, string> = {}, resolveHref?: (rid: string) => string | undefined): Partial<PptxTextRun> {
    const style: Partial<PptxTextRun> = {};
    if (!rPr) return style;
    const attrs = rPr.attrs || {};
    if (attrs.sz) style.fontSize = Math.round(Number(attrs.sz) / 100 * 100) / 100; // 百分之一 pt → pt
    if (attrs.b === '1' || attrs.b === 1) style.bold = true;
    if (attrs.i === '1' || attrs.i === 1) style.italic = true;
    if (attrs.u && attrs.u !== 'none') style.underline = true;
    // 某些生成器会在 a:rPr 下直接写 a:srgbClr/a:schemeClr（与 a:solidFill 并存），
    // 按实际渲染优先级优先取直接子节点颜色，再回退到 a:solidFill
    const directColorNode = rPr['a:srgbClr'] || rPr['a:schemeClr'];
    const color = (directColorNode ? spColor(directColorNode, themeMap) : undefined)
        || spColor(rPr['a:solidFill'], themeMap);
    if (color) style.color = color;
    const latin = rPr['a:latin'];
    if (latin && latin.attrs && latin.attrs.typeface) style.fontFace = String(latin.attrs.typeface);
    const hlink = rPr['a:hlinkClick'] || rPr['a:hlinkHover'];
    if (hlink && hlink.attrs && hlink.attrs['r:id']) {
        // r:id 是关系 id，需解析成实际 URL（外部 http(s) 或内部 '#N' 跳转）。
        // 存裸 rId 的话，渲染端无从下手、回写端还会把它再当成一个新关系的目标。
        const resolved = resolveHref ? resolveHref(String(hlink.attrs['r:id'])) : undefined;
        if (resolved) {
            style.href = resolved;
            if (hlink.attrs.tooltip) style.hrefTooltip = String(hlink.attrs.tooltip);
        }
    }
    // 文字描边（a:ln）
    const ln = rPr['a:ln'];
    if (ln) {
        if (ln['a:noFill']) {
            style.outline = 'none';
        } else {
            const lnColor = spColor(ln['a:solidFill'], themeMap);
            const w = ln.attrs && ln.attrs.w != null ? emuToPt(ln.attrs.w) : 1;
            style.outline = { color: lnColor, width: w };
        }
    }
    // 文字外阴影（a:effectLst/a:outerShdw）
    const outerShdw = rPr['a:effectLst'] && rPr['a:effectLst']['a:outerShdw'];
    if (outerShdw && outerShdw.attrs) {
        const attrs = outerShdw.attrs;
        const dist = Number(attrs.dist) || 0;
        const dir = Number(attrs.dir) || 0; // 1/60000 度
        const blur = Number(attrs.blurRad) || 0;
        const rad = (dir / 60000) * Math.PI / 180;
        const shColor = spColor(outerShdw, themeMap);
        const alphaNode = (outerShdw['a:srgbClr'] || outerShdw['a:schemeClr']) && (outerShdw['a:srgbClr'] || outerShdw['a:schemeClr'])['a:alpha'];
        const alpha = alphaNode && alphaNode.attrs ? Math.round(Number(alphaNode.attrs.val) / 1000) / 100 : undefined;
        style.shadow = {
            color: shColor,
            blur: Math.round(emuToPt(blur) * 100) / 100,
            x: Math.round(Math.cos(rad) * emuToPt(dist) * 100) / 100,
            y: Math.round(Math.sin(rad) * emuToPt(dist) * 100) / 100,
            alpha
        };
    }
    return style;
}

/**
 * 提取 txBody（p:txBody 或表格单元格 a:txBody）为正文段落
 * @returns paragraphs 段落列表；hasText 是否含文本；valign 文本体垂直对齐；text 纯文本拼接
 */
function extractTxBody(txBody: any, themeMap: Record<string, string> = {}, fallbackColor?: string, resolveHref?: (rid: string) => string | undefined): { paragraphs: PptxParagraph[]; hasText: boolean; valign?: VAlign; text: string; textDirection?: string; inset?: { l?: number; r?: number; t?: number; b?: number }; noWrap?: boolean; rtlCol?: boolean } {
    const paragraphs: PptxParagraph[] = [];
    let hasText = false;
    let text = '';
    if (!txBody) return { paragraphs, hasText, text };

    const bodyPr = txBody['a:bodyPr'];
    const valign = bodyPr && bodyPr.attrs && bodyPr.attrs.anchor
        ? VALIGN_MAP[bodyPr.attrs.anchor] : undefined;
    const textDirection = bodyPr && bodyPr.attrs && bodyPr.attrs.vert
        ? String(bodyPr.attrs.vert) : undefined;
    // 不换行（a:bodyPr@wrap="none"）：与预览端一致，文本不自动折行
    const noWrap = bodyPr && bodyPr.attrs && bodyPr.attrs.wrap === 'none' ? true : undefined;
    // 从右向左列排布（a:bodyPr@rtlCol），影响 RTL 文本折行
    const rtlCol = bodyPr && bodyPr.attrs && (bodyPr.attrs.rtlCol === '1' || bodyPr.attrs.rtlCol === 1) ? true : undefined;
    // 内边距（a:bodyPr/@lIns/rIns/tIns/bIns，EMU → px）
    let inset: { l?: number; r?: number; t?: number; b?: number } | undefined;
    if (bodyPr && bodyPr.attrs) {
        const a = bodyPr.attrs;
        const l = a.lIns != null ? emuToPx(a.lIns) : undefined;
        const r = a.rIns != null ? emuToPx(a.rIns) : undefined;
        const t = a.tIns != null ? emuToPx(a.tIns) : undefined;
        const b = a.bIns != null ? emuToPx(a.bIns) : undefined;
        if (l != null || r != null || t != null || b != null) inset = { l, r, t, b };
    }

    for (const pNode of asArray(txBody['a:p'])) {
        const pPr = pNode['a:pPr'];
        const pAttrs = (pPr && pPr.attrs) || {};
        const isRtlPara = pAttrs.rtl === '1' || pAttrs.rtl === 1;
        // RTL 段落（a:pPr@rtl="1"）未显式给 algn 时按右对齐处理：
        // 预览端依 rtl 渲染为右对齐，生成端需落成 algn="r" 才能保持一致
        const align = pAttrs.algn ? ALIGN_MAP[pAttrs.algn] : (isRtlPara ? 'right' : undefined);

        // 列表样式：自动编号 / 字符项目符号 / 图片项目符号
        let bullet: any;
        if (pPr && pPr['a:buNone']) {
            bullet = undefined;
        } else if (pPr && pPr['a:buBlip'] && pPr['a:buBlip']['a:blip']) {
            // 图片项目符号（a:buBlip）：data URL 由 resolveBulletBlips 预先注入
            const blip = pPr['a:buBlip']['a:blip'];
            const src = (blip.attrs && (blip.attrs.__data || blip.attrs['r:embed'])) || undefined;
            if (src) bullet = { type: 'picture', data: String(src).startsWith('data:') ? String(src) : undefined, rid: String(src) };
            const szPct = pPr['a:buSzPct'] && pPr['a:buSzPct'].attrs && Number(pPr['a:buSzPct'].attrs.val) / 1000;
            if (bullet && szPct) bullet.sizePct = szPct;
        } else if (pPr && pPr['a:buAutoNum']) {
            const auto = pPr['a:buAutoNum'];
            bullet = { type: 'number', fmt: (auto.attrs && auto.attrs.type) || 'arabic', start: (auto.attrs && auto.attrs.startAt != null) ? Number(auto.attrs.startAt) : 1 };
        } else if (pPr && pPr['a:buChar']) {
            const bc = pPr['a:buChar'];
            bullet = { type: 'bullet', char: (bc.attrs && bc.attrs.char) || '•' };
            // 符号字体（a:buFont，如 Wingdings / Wingdings 3）与字号比例：
            // 缺了字体名，符号字体字符会退化成普通字母，编辑器画不出图标
            const bf = pPr['a:buFont'];
            if (bf && bf.attrs && bf.attrs.typeface) bullet.font = String(bf.attrs.typeface);
            const szPct = pPr['a:buSzPct'] && pPr['a:buSzPct'].attrs && Number(pPr['a:buSzPct'].attrs.val) / 1000;
            if (szPct) bullet.sizePct = szPct;
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
        // 按文档顺序交错处理 a:r 与 a:br（simplify 为每个节点写入全局递增 attrs.order）
        const ordered: Array<{ n: any; br: boolean }> = [
            ...asArray(pNode['a:r']).map((n: any) => ({ n, br: false })),
            ...asArray(pNode['a:br']).map((n: any) => ({ n, br: true }))
        ].sort((x, y) => ((x.n && x.n.attrs && x.n.attrs.order) || 0) - ((y.n && y.n.attrs && y.n.attrs.order) || 0));
        for (const { n: runNode, br } of ordered) {
            if (br) {
                // <a:br/>：段内软换行
                runs.push({ text: '', break: true });
                continue;
            }
            const t = readRunText(runNode);
            if (t) hasText = true;
            paraText += t;
            const st = readRunStyle(runNode['a:rPr'], themeMap, resolveHref);
            // run 无显式字色时回退到形状 p:style/a:fontRef 的默认色
            if (!st.color && fallbackColor) st.color = fallbackColor;
            runs.push({ text: t, ...st });
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
        // 从右到左段落（a:pPr@rtl="1"）：RTL 段落缺省右对齐
        if (isRtlPara) para.rtl = true;
        if (valign) (para as any).valign = valign;
        paragraphs.push(para);
    }
    return { paragraphs, hasText, valign, text, textDirection, inset, noWrap, rtlCol };
}

/**
 * 预解析 txBody 内所有图片项目符号（a:pPr/a:buBlip/a:blip/@r:embed），
 * 就地写入 __data（data URL）。extractTxBody 是同步函数，拿不到 zip，
 * 因此在有 resObj/zip 的调用点先把图读出来注入节点。
 * @returns 是否注入了至少一个图片项目符号
 */
async function resolveBulletBlips(
    txBody: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip
): Promise<boolean> {
    if (!txBody) return false;
    let touched = false;
    for (const pNode of asArray(txBody['a:p'])) {
        const buBlip = pNode && pNode['a:pPr'] && pNode['a:pPr']['a:buBlip'];
        const blip = buBlip && buBlip['a:blip'];
        if (!blip || !blip.attrs || !blip.attrs['r:embed'] || blip.attrs.__data) continue;
        const target = resObj[String(blip.attrs['r:embed'])] && resObj[String(blip.attrs['r:embed'])].target;
        const part = resolvePart(target);
        if (!part) continue;
        try {
            const file = zip.file(part);
            if (!file) continue;
            const ext = ((part || String(target || '')).split('.').pop() || 'png').toLowerCase();
            blip.attrs.__data = `data:${IMAGE_MIME[ext] || 'application/octet-stream'};base64,${await file.async('base64')}`;
            touched = true;
        } catch { /* 忽略媒体读取失败 */ }
    }
    return touched;
}

/** 提取一个 p:sp 的文本为正文段落 */
function extractTextBody(spNode: any, themeMap: Record<string, string> = {}, resolveHref?: (rid: string) => string | undefined): { paragraphs: PptxParagraph[]; hasText: boolean; textDirection?: string; inset?: { l?: number; r?: number; t?: number; b?: number }; noWrap?: boolean; rtlCol?: boolean } {
    // p:style/a:fontRef 的颜色是形状文本的默认字色（run 无显式色时生效，与预览端一致）
    const fontRefColor = spColor(spNode && spNode['p:style'] && spNode['p:style']['a:fontRef'], themeMap);
    return extractTxBody(spNode && spNode['p:txBody'], themeMap, fontRefColor, resolveHref);
}

/**
 * 关系 id → 超链接 URL：外部 http(s) 原样返回；
 * 内部幻灯片跳转（关系 target 指向 slideN.xml）→ '#N'，与生成端 buildHyperlink 约定一致。
 */
function resolveHyperlink(resObj: Record<string, { type?: string; target?: string }>): (rid: string) => string | undefined {
    return (rid: string) => {
        const rel = resObj[String(rid)];
        const target = rel && rel.target ? String(rel.target) : '';
        if (!target) return undefined;
        const m = /slide(\d+)\.xml$/i.exec(target);
        return m ? `#${m[1]}` : target;
    };
}

/** 解析颜色并去掉前导 # / alpha，保持与历史 readSrgbClr 返回格式一致（RRGGBB） */
function spColor(node: any, themeMap: Record<string, string>): string | undefined {
    const c = resolveColorNode(node, themeMap);
    return c ? tinycolor(c).toHexString().toUpperCase().replace(/^#/, '') : undefined;
}

/** 从 p:spPr 读取几何/填充/边框/特效（themeMap 用于把 schemeClr 解析为实际 RRGGBB） */
function readSpPr(spPr: any, themeMap: Record<string, string> = {}): Pick<PptxShapeElement, 'shapeType' | 'fill' | 'line' | 'effects'> {
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
            return { color: spColor(gs, themeMap) || '000000', position: pos };
        });
        const lin = gf['a:lin'];
        let direction: 'horizontal' | 'vertical' | 'diagonal' = 'horizontal';
        if (lin && lin.attrs) {
            const ang = Number(lin.attrs.ang) / 60000; // 度
            direction = ang >= 45 && ang < 135 ? 'vertical' : (ang >= 22.5 && ang < 67.5 ? 'diagonal' : 'horizontal');
        }
        // a:path（径向渐变）：没有 a:lin。此前只按线性处理，导出后径向会退化成水平线性
        const path = gf['a:path'];
        out.fill = path
            ? { type: 'gradient', direction, stops, gradientType: 'radial', gradientPath: (path.attrs && path.attrs.path) || 'circle' }
            : { type: 'gradient', direction, stops };
    } else if (spPr['a:pattFill']) {
        // 图案填充（a:pattFill）：prst + 前景/背景色，往返必须保留，否则形状会退化为空白
        const pf = Array.isArray(spPr['a:pattFill']) ? spPr['a:pattFill'][0] : spPr['a:pattFill'];
        const prst = pf && pf.attrs && pf.attrs.prst ? String(pf.attrs.prst) : undefined;
        if (prst) {
            const patt: any = { type: 'pattern', prst };
            const fg = spColor(firstChild(pf, 'a:fgClr'), themeMap);
            const bg = spColor(firstChild(pf, 'a:bgClr'), themeMap);
            if (fg) patt.fg = fg;
            if (bg) patt.bg = bg;
            out.fill = patt;
        }
    } else if (spPr['a:blipFill']) {
        // 形状图片填充（a:blipFill）需要异步读取媒体，占位标记，由调用方用 readImageFill 覆盖
        out.fill = undefined;
    } else {
        const solid = spPr['a:solidFill'];
        const color = spColor(solid, themeMap);
        if (color) {
            const colorNodeAny = solid && (solid['a:srgbClr'] || solid['a:schemeClr']);
            const alphaNode = colorNodeAny && colorNodeAny['a:alpha'];
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
            const color = spColor(ln['a:solidFill'], themeMap);
            const w = ln.attrs && ln.attrs.w ? emuToPt(ln.attrs.w) : DEFAULT_LN_PT;
            const solidLn = ln['a:solidFill'];
            const colorNodeAny = solidLn && (solidLn['a:srgbClr'] || solidLn['a:schemeClr']);
            const alphaNode = colorNodeAny && colorNodeAny['a:alpha'];
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

/** 读取 xfrm 位置（graphicFrame 用 p:xfrm，其余用 p:spPr/a:xfrm；graphicFrame 兼容部分生成器写出的非标准 a:xfrm） */
function readXfrm(node: any, isGraphicFrame: boolean) {
    const xf = isGraphicFrame
        ? (node['p:xfrm'] || node['a:xfrm'])
        : (node['p:spPr'] && node['p:spPr']['a:xfrm']);
    if (!xf || !xf['a:off']) return null;
    const off = (xf['a:off'] && xf['a:off'].attrs) || {};
    const ext = (xf['a:ext'] && xf['a:ext'].attrs) || {};
    const attrs = xf.attrs || {};
    const out: any = {
        x: emuToPx(off.x),
        y: emuToPx(off.y),
        width: emuToPx(ext.cx),
        height: emuToPx(ext.cy),
        rotation: rotToDeg(attrs.rot)
    };
    // 水平/垂直翻转（a:xfrm/@flipH/@flipV，值 1/0 或 undefined）
    if (attrs.flipH === 1 || attrs.flipH === '1' || attrs.flipH === true) out.flipH = true;
    if (attrs.flipV === 1 || attrs.flipV === '1' || attrs.flipV === true) out.flipV = true;
    return out;
}

/** 读取形状预设几何的调整值（a:prstGeom/a:avLst/a:gd → { adj1: 50000 }，单位为 OOXML 原生千分比，与生成端一致） */
function readAdjust(prstGeom: any): Record<string, number> | undefined {
    const avLst = prstGeom && prstGeom['a:avLst'];
    if (!avLst) return undefined;
    const gds = asArray(avLst['a:gd']);
    if (!gds.length) return undefined;
    const adj: Record<string, number> = {};
    for (const g of gds) {
        const name = g && g.attrs && g.attrs.name;
        const fmla = g && g.attrs && g.attrs.fmla; // 形如 "val 50000"
        const m = fmla && /val\s+(\d+)/.exec(String(fmla));
        if (name && m) adj[name] = Number(m[1]);
    }
    return Object.keys(adj).length ? adj : undefined;
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
function extractTiming(slideContent: any): { advanceTime?: number; animations?: PptxAnimation[] } {
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
        if (spid != null) {
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
                target: String(spid),
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

/** 提取背景：slide → layout → master 逐级继承（与预览端 getSlideBackgroundFill 同语义） */
async function extractBackground(
    slideContent: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip,
    fallbacks: Array<{ content: any; res?: Record<string, { type?: string; target?: string }> }> = [],
    themeMap: Record<string, string> = {}
): Promise<PptxBackground | undefined> {
    const parts: Array<{ content: any; res?: Record<string, { type?: string; target?: string }> }> = [
        { content: slideContent, res: resObj }, ...fallbacks
    ];
    for (const part of parts) {
        const sld = part.content && (part.content['p:sld'] || part.content['p:sldLayout'] || part.content['p:sldMaster']);
        const bg = sld && sld['p:cSld'] && sld['p:cSld']['p:bg'];
        if (!bg) continue;
        const bgPr = bg['p:bgPr'];
        if (bgPr) {
            const color = spColor(bgPr, themeMap) || readSrgbClr(bgPr);
            if (color) return color;
            // 图片背景（p:bgPr/a:blipFill）：与形状图片填充一致地内联 base64
            if (bgPr['a:blipFill'] && part.res) {
                const img = await readImageFill(bgPr['a:blipFill'], part.res, zip);
                if (img) return img;
            }
            if (bgPr['a:gradFill']) {
                const grad = readGradientFill(bgPr['a:gradFill'], themeMap);
                if (grad) return grad;
            }
        }
        // 主题引用 bgRef：无法解析为具体色，标记继承（不写 background）
        return undefined;
    }
    return undefined;
}

/**
 * a:gradFill → 渐变填充描述。
 * stop 颜色经 resolveColorNode 解析（scheme 引用 + tint/lumMod 等修饰）；
 * a:path（circle/shape/ray）为径向渐变，a:lin 按角度换算方向。
 * 无任何可解析 stop 时返回 undefined（调用方退回纯色近似）。
 */
function readGradientFill(gradFill: any, themeMap: Record<string, string> = {}): PptxGradientFill | undefined {
    if (!gradFill) return undefined;
    const gsLst = gradFill['a:gsLst'];
    const stops: { color: string; position: number }[] = [];
    for (const gs of asArray(gsLst && gsLst['a:gs'])) {
        const pos = gs && gs.attrs && gs.attrs.pos ? Number(gs.attrs.pos) / 100000 : 0;
        const c = spColor(gs, themeMap) || readSrgbClr(gs);
        if (c) stops.push({ color: c, position: pos });
    }
    if (!stops.length) return undefined;
    const lin = gradFill['a:lin'];
    // ang 为 1/60000 度，需换算：90°→vertical、45°→diagonal、其余 horizontal
    const ang = lin && lin.attrs && lin.attrs.ang !== undefined ? Number(lin.attrs.ang) / 60000 : 0;
    const direction = ang === 90 ? 'vertical' : ang === 45 ? 'diagonal' : 'horizontal';
    const path = gradFill['a:path'];
    return path
        ? { type: 'gradient', direction, stops, gradientType: 'radial', gradientPath: (path.attrs && path.attrs.path) || 'circle' }
        : { type: 'gradient', direction, stops };
}

/** 从 c:chartSpace 反向提取图表语义 */
function extractChart(chartXml: any, themeMap: Record<string, string> = {}): Partial<PptxChartElement> | undefined {
    try {
        const chartSpace = chartXml && chartXml['c:chartSpace'];
        const chart = chartSpace && chartSpace['c:chart'];
        if (!chart) return undefined;
        const plotArea = chart['c:plotArea'];
        if (!plotArea) return undefined;

        // 图表区填充（c:chartSpace/c:spPr）：solidFill 直接取色；gradFill 取首个渐变 stop 近似
        let spaceFill: string | undefined;
        const spaceSpPr = chartSpace['c:spPr'];
        if (spaceSpPr) {
            const solid = spColor(spaceSpPr['a:solidFill'], themeMap);
            if (solid) {
                spaceFill = solid;
            } else if (spaceSpPr['a:noFill']) {
                spaceFill = 'none';
            } else {
                const gsLst = spaceSpPr['a:gradFill'] && spaceSpPr['a:gradFill']['a:gsLst'];
                const gs0 = asArray(gsLst && gsLst['a:gs'])[0];
                const gradCol = gs0 ? spColor(gs0, themeMap) : undefined;
                if (gradCol) spaceFill = gradCol;
            }
        }

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
            if (isScatter || isBubble) {
                // 散点/气泡：优先 c:xVal；个别导出器把 X 放在 c:cat（类别）下，做回退
                let xv = numCacheValues(ser['c:xVal']);
                if (!xv.length) xv = numCacheValues(ser['c:cat']);
                s.x = xv;
                s.y = numCacheValues(ser['c:yVal']);
                if (isBubble) {
                    // 气泡图：values 复用为气泡大小（与生成端一致）
                    s.values = numCacheValues(ser['c:bubbleSize']);
                }
            } else if (isStock) {
                s.open = numCacheValues(ser['c:openVal']);
                s.high = numCacheValues(ser['c:highVal']);
                s.low = numCacheValues(ser['c:lowVal']);
                s.close = numCacheValues(ser['c:closeVal']);
            } else {
                s.values = numCacheValues(ser['c:val']);
            }
            // 系列填充色：c:ser/c:spPr/a:solidFill（srgb 或 schemeClr，schemeClr 用主题色映射解析）
            const serSpPr = ser['c:spPr'];
            const serColor = serSpPr ? spColor(serSpPr['a:solidFill'], themeMap) : undefined;
            if (serColor) s.color = serColor;
            // 逐点填充（c:dPt）：饼/环等 varyColors 场景每点可带独立 spPr；
            // solidFill 取色，gradFill 保留完整渐变（3D 饼/环多为 accent1 径向渐变，
            // 压成纯色会让渲染端退回默认调色板导致整图配色错位）
            const dPts = asArray(ser['c:dPt']);
            if (dPts.length) {
                const nPts = Math.max(s.values ? s.values.length : 0, ...(dPts.map((d: any) => {
                    const idx = d && d['c:idx'] && d['c:idx'].attrs ? Number(d['c:idx'].attrs.val) : -1;
                    return idx + 1;
                })));
                if (nPts > 0) {
                    const pc: (string | PptxGradientFill | undefined)[] = new Array(nPts).fill(undefined);
                    for (const d of dPts) {
                        const idx = d && d['c:idx'] && d['c:idx'].attrs ? Number(d['c:idx'].attrs.val) : -1;
                        if (idx < 0) continue;
                        const dSpPr = d['c:spPr'];
                        if (!dSpPr) continue;
                        const grad = readGradientFill(dSpPr['a:gradFill'], themeMap);
                        if (grad) { pc[idx] = grad; continue; }
                        const solid = spColor(dSpPr['a:solidFill'], themeMap);
                        if (solid) pc[idx] = solid;
                    }
                    s.pointColors = pc;
                }
            }
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

        const legendNode = chart['c:legend'];
        const legend = !!legendNode;
        const legendPosition = legendNode && legendNode['c:legendPos'] && legendNode['c:legendPos'].attrs
            ? String(legendNode['c:legendPos'].attrs.val)
            : undefined;
        const out: Partial<PptxChartElement> = { chartType, series };
        if (spaceFill !== undefined) out.spaceFill = spaceFill;
        if (categories.length) out.categories = categories;
        if (title) out.title = title.trim();
        out.legend = legend;
        if (legendPosition) out.legendPosition = legendPosition;

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

        // 三维视角（c:view3D）：旋转/厚度/直角轴
        const v3dNode = chartNode['c:view3D'];
        if (v3dNode && v3dNode.attrs) {
            const va = v3dNode.attrs;
            const v3d: any = {};
            if (va.rotX !== undefined) v3d.rotX = parseFloat(va.rotX);
            if (va.rotY !== undefined) v3d.rotY = parseFloat(va.rotY);
            if (va.depthPercent !== undefined) v3d.depthPercent = parseFloat(va.depthPercent);
            if (va.rAngAx !== undefined) v3d.rAngAx = va.rAngAx === '1';
            if (Object.keys(v3d).length) out.view3D = v3d;
        }

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
    /** 当前 PPTX 主题配色，用于把表格样式里的 schemeClr 解析为具体颜色 */
    theme?: { colors?: Record<string, string> };
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
    // Microsoft 缓存绘图（diagramDrawing）—— 常量定义在文件后部，这里内联 URI
    'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing':
        'application/vnd.ms-office.drawingml.diagramDrawing+xml',
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
    // SmartArt 缓存绘图（diagramDrawing 关系）不经节点上的 r: 属性引用
    // （dgm:relIds 只带 dm/cs/lo/qs），按关系类型补入，否则导出后丢失 drawingN.xml、
    // SmartArt 的全部图形在解析器/预览端消失
    for (const [rid, rel] of Object.entries(resObj || {})) {
        if (rel && (rel.type === MS_DIAGRAM_DRAWING_REL || /diagrams\/drawing\d+\.xml$/i.test(String(rel.target || '')))) {
            relIds.add(rid);
        }
    }
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
                    // grpSp 的子形状是其直接子节点（p:sp/p:pic/...），不存在 p:spTree 包装；
                    // collectShapeNodes 只挑形状键、忽略 nvGrpSpPr/grpSpPr/attrs，可直接传入 grpSp 节点
                    collectShapeNodes(node, acc, keepGroups);
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

        // 主题色映射（scheme 名 → #RRGGBB）：优先用该页实际主题（slideData.themeContent），
        // 回退到全局默认主题（options.theme）。避免多主题文件（如 Sample_12）用错主题。
        const themeMap: Record<string, string> = {};
        if (slideData.themeContent) {
            Object.assign(themeMap, themeColorsFromContent(slideData.themeContent));
        }
        if (!Object.keys(themeMap).length && options.theme) {
            Object.assign(themeMap, themeColorsFromTheme(options.theme));
        }

        // 背景 / 过渡 / 备注（slide → layout → master 继承）
        const bg = await extractBackground(slideContent, resObj, zip, [
            { content: slideData.slideLayoutContent, res: slideData.layoutResObj },
            { content: slideData.slideMasterContent, res: slideData.masterResObj }
        ], themeMap);
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
            const processNode = async (key: string, node: any): Promise<PptxElement | null> => {
                try {
                    const el = await nodeToElement(key, node, resObj, zip, themeMap, slideData.themeContent);
                    if (!el) return null;
                    // 统一挂载 __raw 载荷（含标签名，供生成端无损回写）
                    (el as any).__raw = { tag: key, node };
                    // 依赖携带：语义层未覆盖的类型必须附带，否则 __raw 无法独立回写；
                    // 语义类型仅在 rawDeps:'all' 时附带（供 rawFallback 使用）
                    if (allDeps || !SEMANTIC_TYPES.has(el.type)) {
                        await attachRawDeps(el, node, resObj, zip);
                    }
                    return el;
                } catch {
                    // 单元素失败：仅保留原始节点，由 __raw 回退承载
                    const rawEl: PptxRawElement = {
                        type: 'raw', x: 0, y: 0, width: 0, height: 0,
                        __raw: { tag: key, node }, rawFallback: true
                    };
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
                // 标准 OOXML（CT_GroupShape）：grpSp 的子形状是其直接子节点（p:sp/p:pic/...）；
                // 兼容回退：旧版本生成端曾错误地在 grpSp 内包一层 p:spTree，优先按其存在与否选择
                const legacyInner = node['p:spTree'];
                const children = legacyInner ? await processTree(legacyInner) : await processTree(node);

                const gxf = node['p:grpSpPr'] && node['p:grpSpPr']['a:xfrm'];
                const gOff = gxf && gxf['a:off'] && gxf['a:off'].attrs;
                const gExt = gxf && gxf['a:ext'] && gxf['a:ext'].attrs;
                const chOff = gxf && gxf['a:chOff'] && gxf['a:chOff'].attrs;
                const chExt = gxf && gxf['a:chExt'] && gxf['a:chExt'].attrs;
                const gx = emuToPx(gOff?.x), gy = emuToPx(gOff?.y);
                const gw = emuToPx(gExt?.cx), gh = emuToPx(gExt?.cy);
                const chx = emuToPx(chOff?.x), chy = emuToPx(chOff?.y);
                const chw = emuToPx(chExt?.cx), chh = emuToPx(chExt?.cy);
                // chExt 缺失或为零时跳过缩放（避免除零导致坐标放大数百倍）
                const sx = chw > 0 ? gw / chw : 1;
                const sy = chh > 0 ? gh / chh : 1;

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
                return g;
            };

            slide.elements = await processTree(spTree);

            // 对表格元素应用 tableStyles.xml 中定义的样式（主题色、填充、文字色、边框）
            if (slideData.tableStyles) {
                (slideData.tableStyles as any)._themeContent = slideData.themeContent;
                for (const el of slide.elements) {
                    if (el.type === 'table') {
                        applyTableStyle(el as PptxTableElement, slideData.tableStyles, options.theme);
                    }
                }
            }

            // 元素内残留的 'scheme:<name>' 引用按该页真实主题解析为绝对色。
            // PPTX 标准：主题色由该页所属母版绑定的主题决定；多主题文件（如 Sample_12）
            // 中不同页可绑定不同主题，若留给渲染端按全局主题解析会取错颜色。
            if (Object.keys(themeMap).length) {
                for (const el of slide.elements) resolveSchemeRefs(el, themeMap);
            }

            // 自动播放 / 元素动画（p:timing）
            const timing = extractTiming(slideContent);
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

/** 图片扩展名 → MIME（形状图片填充内联为 dataURL 时使用） */
const IMAGE_MIME: Record<string, string> = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
    bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tiff: 'image/tiff',
    emf: 'image/emf', wmf: 'image/wmf'
};

/**
 * 图片填充（a:blipFill，用于形状 p:spPr 或背景 p:bgPr）→ PptxFillImage。
 *
 * 与 p:pic 的处理保持一致：包内部件内联为 dataURL（保证 round-trip 自包含），
 * 外链图片保留 src 交由生成端下载；同时读取 a:srcRect（裁剪）与 a:tile（平铺），
 * 这两者在 OOXML 中均为千分比，统一转成 0~1 比例。
 *
 * 部件缺失且非外链时返回 undefined（避免生成端拿到不可用的填充）。
 */
async function readImageFill(
    blipFill: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip
): Promise<PptxFillImage | undefined> {
    const blip = blipFill && blipFill['a:blip'];
    if (!blip || !blip.attrs || !blip.attrs['r:embed']) return undefined;

    const rid = String(blip.attrs['r:embed']);
    const target = resObj[rid] && resObj[rid].target;
    const part = resolvePart(target);
    const ext = ((part || target || '').split('.').pop() || 'png').toLowerCase();
    const out: PptxFillImage = { type: 'image', extension: ext };

    if (part) {
        try {
            const file = zip.file(part);
            if (file) {
                out.data = `data:${IMAGE_MIME[ext] || 'application/octet-stream'};base64,${await file.async('base64')}`;
            }
        } catch { /* 忽略媒体读取失败 */ }
    }
    if (!out.data) {
        if (/^https?:/i.test(String(target))) out.src = String(target);
        else return undefined;
    }

    // 裁剪：a:srcRect（WPS 也可能写 a:stretch/a:fillRect），千分比 → 0~1 比例
    const srcRectNode = firstChild(blipFill, 'a:srcRect')
        || firstChild(firstChild(blipFill, 'a:stretch'), 'a:fillRect');
    if (srcRectNode) {
        const rect: { l?: number; t?: number; r?: number; b?: number } = {};
        for (const side of ['l', 't', 'r', 'b'] as const) {
            const v = readNumAttr(srcRectNode, side);
            if (v) rect[side] = v / 100000;
        }
        if (Object.keys(rect).length) out.srcRect = rect;
    }
    // 平铺：a:tile，sx/sy（每格占原图比例）与 tx/ty（偏移）均为千分比 → 0~1 比例
    const tileNode = firstChild(blipFill, 'a:tile');
    if (tileNode) {
        const tile: { sx?: number; sy?: number; tx?: number; ty?: number } = {};
        for (const k of ['sx', 'sy', 'tx', 'ty'] as const) {
            const v = readNumAttr(tileNode, k);
            if (v !== undefined) tile[k] = v / 100000;
        }
        if (Object.keys(tile).length) out.tile = tile;
    }
    return out;
}

/** 提取 a:custGeom（自定义自由曲线/任意多边形）为归一化路径结构。
 *  OOXML 路径命令按 attrs.order 保序；坐标保持原始 EMU 空间（path 的 w/h），
 *  编辑器端据此生成 SVG。 */
function readCustGeom(custGeom: any): PptxCustomGeometry | undefined {
    const pathLst = custGeom && custGeom['a:pathLst'];
    const pathArr = asArray(pathLst && pathLst['a:path']);
    const paths: PptxGeometryPath[] = [];
    for (const path of pathArr) {
        if (!path || typeof path !== 'object') continue;
        const w = Number(path.attrs && path.attrs.w) || undefined;
        const h = Number(path.attrs && path.attrs.h) || undefined;
        const cmds: PptxGeometryCommand[] = [];
        const collected: Array<{ order: number; key: string; node: any }> = [];
        for (const k of Object.keys(path)) {
            if (k === 'attrs') continue;
            for (const child of asArray(path[k])) {
                if (child && typeof child === 'object') {
                    collected.push({ order: Number(child.attrs && child.attrs.order) || 0, key: k, node: child });
                }
            }
        }
        collected.sort((a, b) => a.order - b.order);
        const pt = (node: any, i: number) => {
            const p = asArray(node && node['a:pt'])[i];
            return p && p.attrs ? { x: Number(p.attrs.x) || 0, y: Number(p.attrs.y) || 0 } : { x: 0, y: 0 };
        };
        let hasClose = false;
        for (const { key, node } of collected) {
            switch (key) {
                case 'a:moveTo': { const p = pt(node, 0); cmds.push({ type: 'moveTo', x: p.x, y: p.y }); break; }
                case 'a:lnTo': { const p = pt(node, 0); cmds.push({ type: 'lnTo', x: p.x, y: p.y }); break; }
                case 'a:cubicBezTo': {
                    const p1 = pt(node, 0), p2 = pt(node, 1), p3 = pt(node, 2);
                    cmds.push({ type: 'cubicBezTo', x1: p1.x, y1: p1.y, x2: p2.x, y2: p2.y, x: p3.x, y: p3.y });
                    break;
                }
                case 'a:quadBezTo': {
                    const p1 = pt(node, 0), p2 = pt(node, 1);
                    cmds.push({ type: 'quadBezTo', x1: p1.x, y1: p1.y, x: p2.x, y: p2.y });
                    break;
                }
                case 'a:arcTo': {
                    const a = node.attrs || {};
                    cmds.push({ type: 'arcTo', wR: Number(a.wR) || 0, hR: Number(a.hR) || 0, stAng: Number(a.stAng) || 0, swAng: Number(a.swAng) || 0 });
                    break;
                }
                case 'a:close': cmds.push({ type: 'close' }); hasClose = true; break;
                default: break;
            }
        }
        if (cmds.length) paths.push({ w, h, commands: cmds, closed: hasClose || undefined });
    }
    return paths.length ? { paths } : undefined;
}

/** 单个 OOXML 节点 → PptxElement */
async function nodeToElement(
    key: string,
    node: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip,
    themeMap: Record<string, string> = {},
    themeContent?: any
): Promise<PptxElement | null> {
    if (key === 'p:graphicFrame') {
        return await graphicFrameToElement(node, resObj, zip, themeMap, themeContent);
    }
    if (key === 'p:pic') {
        return await picToImage(node, resObj, zip);
    }
    // p:sp / p:cxnSp → text 或 shape
    const spPr = node['p:spPr'];
    // 图片项目符号需要读 zip，先把 buBlip 的 r:embed 就地解析成 data URL
    if (node && node['p:txBody'] && node['p:txBody']['a:p']) {
        try { await resolveBulletBlips(node['p:txBody'], resObj, zip); } catch { /* 忽略 */ }
    }
    const { paragraphs, hasText, textDirection, inset, noWrap, rtlCol } = extractTextBody(node, themeMap, resolveHyperlink(resObj));
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
        if (xf && xf.flipH) textEl.flipH = true;
        if (xf && xf.flipV) textEl.flipV = true;
        if (name) textEl.name = String(name);
        if (textDirection) textEl.textDirection = textDirection;
        if (inset) textEl.inset = inset;
        if (noWrap) textEl.noWrap = noWrap;
        if (rtlCol) textEl.rtlCol = true;
        // 提取形状外观（填充/边框/特效）。纯文本框（txBox=1）也可能带背景填充，
        // 因此统一读取 spPr；只有「非矩形 + 非纯文本框」才保留 shapeType/adjust，
        // 避免把普通矩形文本框渲染成异形。
        const isPureTextBox = node['p:nvSpPr'] && node['p:nvSpPr']['p:cNvSpPr'] && node['p:nvSpPr']['p:cNvSpPr'].attrs && node['p:nvSpPr']['p:cNvSpPr'].attrs.txBox === '1';
        const hasCustGeom = !!(spPr && spPr['a:custGeom']);
        const hasNonRectShape = (geom && geom.attrs && geom.attrs.prst && String(geom.attrs.prst) !== 'rect' && !isPureTextBox) || hasCustGeom;
        try {
            const spVis = readSpPr(spPr, themeMap);
            if (spPr && spPr['a:blipFill']) {
                const imgFill = await readImageFill(spPr['a:blipFill'], resObj, zip);
                if (imgFill) spVis.fill = imgFill;
            }
            const pStyle = node['p:style'];
            if (pStyle && typeof pStyle === 'object') {
                const tables = await getThemeStyleTables(zip, themeContent);
                if (spVis.fill === undefined && pStyle['a:fillRef']) {
                    const ref = resolveThemeStyleRef(pStyle['a:fillRef'], tables.fills, themeMap);
                    if (ref) spVis.fill = ref.gradient || ref.color;
                }
                if (spVis.line === undefined && pStyle['a:lnRef']) {
                    const ref = resolveThemeStyleRef(pStyle['a:lnRef'], tables.lines, themeMap);
                    if (ref) spVis.line = { width: ref.widthPt || DEFAULT_LN_PT, color: ref.color };
                }
                if (pStyle['a:effectRef']) {
                    const ref = resolveThemeEffectRef(pStyle['a:effectRef'], tables.effects, themeMap);
                    if (ref && ref.shadow) {
                        spVis.effects = spVis.effects || {};
                        spVis.effects.shadow = ref.shadow;
                    }
                }
            }
            if (hasNonRectShape && spVis.shapeType) textEl.shapeType = spVis.shapeType;
            if (hasNonRectShape) {
                const visAdjust = readAdjust(geom);
                if (visAdjust) textEl.adjust = visAdjust;
            }
            if (hasCustGeom) {
                const cg = readCustGeom(spPr['a:custGeom']);
                if (cg) textEl.custGeom = cg;
            }
            if (spVis.fill !== undefined) textEl.fill = spVis.fill;
            if (spVis.line !== undefined) textEl.line = spVis.line;
            if (spVis.effects) textEl.effects = spVis.effects;
        } catch (e) {
            console.warn('[json-from-pptx] 文本形状外观提取失败（跳过外观）:', e instanceof Error ? (e.message + '\n' + String(e.stack).split('\n')[1]) : e);
        }
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

    if (geom || (spPr && spPr['a:custGeom'])) {
        // 形状元素（预设几何或自定义自由曲线）
        const xf = readXfrm(node, false);
        const sp = readSpPr(spPr, themeMap);
        // 形状图片填充（a:blipFill）需异步读取媒体部件，readSpPr 无法处理，这里覆盖
        if (spPr && spPr['a:blipFill']) {
            const imgFill = await readImageFill(spPr['a:blipFill'], resObj, zip);
            if (imgFill) sp.fill = imgFill;
        }
        // p:style 主题样式引用（fillRef/lnRef）：spPr 无显式填充/边框时按主题样式表解析，
        // 否则大量 PowerPoint 原生形状（仅带样式引用）会丢失填充与描边
        const pStyle = node['p:style'];
        if (pStyle && typeof pStyle === 'object') {
            const tables = await getThemeStyleTables(zip, themeContent);
            if (sp.fill === undefined && pStyle['a:fillRef']) {
                const ref = resolveThemeStyleRef(pStyle['a:fillRef'], tables.fills, themeMap);
                if (ref) sp.fill = ref.gradient || ref.color;
            }
            if (pStyle['a:lnRef']) {
                const ref = resolveThemeStyleRef(pStyle['a:lnRef'], tables.lines, themeMap);
                if (ref) {
                    if (sp.line === undefined) {
                        sp.line = { width: ref.widthPt || DEFAULT_LN_PT, color: ref.color };
                    } else if (sp.line !== 'none' && sp.line && typeof sp.line === 'object' && !sp.line.color) {
                        sp.line.color = ref.color;
                    }
                }
            }
            if (pStyle['a:effectRef']) {
                const ref = resolveThemeEffectRef(pStyle['a:effectRef'], tables.effects, themeMap);
                if (ref && ref.shadow) {
                    sp.effects = sp.effects || {};
                    sp.effects.shadow = ref.shadow;
                }
            }
        }
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
        if (xf && xf.flipH) shapeEl.flipH = true;
        if (xf && xf.flipV) shapeEl.flipV = true;
        const adjust = geom ? readAdjust(geom) : undefined;
        if (adjust) shapeEl.adjust = adjust;
        const custGeom = spPr && spPr['a:custGeom'];
        if (custGeom) {
            const cg = readCustGeom(custGeom);
            if (cg) shapeEl.custGeom = cg;
        }
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

/** 读取单元格内边距属性（marL/marR/marT/marB），EMU → px */
function readInsetsFromAttrs(attrs: any): { l?: number; r?: number; t?: number; b?: number } | undefined {
    if (!attrs) return undefined;
    const out: any = {};
    if (attrs.marL != null) out.l = emuToPx(attrs.marL);
    if (attrs.marR != null) out.r = emuToPx(attrs.marR);
    if (attrs.marT != null) out.t = emuToPx(attrs.marT);
    if (attrs.marB != null) out.b = emuToPx(attrs.marB);
    return Object.keys(out).length ? out : undefined;
}

/** 读取表格级默认内边距（a:tableCellMar） */
function readTableInsets(tblPr: any): { l?: number; r?: number; t?: number; b?: number } | undefined {
    if (!tblPr) return undefined;
    const mar = tblPr['a:tableCellMar'];
    if (!mar) return undefined;
    const side = (k: string) => {
        const n = mar['a:' + k];
        return n && n.attrs && n.attrs.w != null ? emuToPx(n.attrs.w) : undefined;
    };
    const out: any = {};
    const l = side('left'), r = side('right'), t = side('top'), b = side('bottom');
    if (l != null) out.l = l;
    if (r != null) out.r = r;
    if (t != null) out.t = t;
    if (b != null) out.b = b;
    return Object.keys(out).length ? out : undefined;
}

/** a:tbl（表格）→ PptxTableElement */
function tableToElement(tbl: any, node: any, themeMap: Record<string, string> = {}, resObj?: Record<string, { type?: string; target?: string }>): PptxTableElement {
    const xf = readXfrm(node, true);
    const el: any = {
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

    // 表格属性：样式 ID、应用标志、默认内边距
    const tblPr = tbl && tbl['a:tblPr'];
    if (tblPr) {
        const styleIdNode = tblPr['a:tableStyleId'];
        if (styleIdNode) {
            const sid = typeof styleIdNode === 'string' ? styleIdNode : (styleIdNode.text != null ? String(styleIdNode.text) : undefined);
            if (sid) el.tableStyleId = sid;
        }
        const attrs = tblPr.attrs || {};
        const flags: any = {};
        if (attrs.firstRow != null) flags.firstRow = String(attrs.firstRow) === '1';
        if (attrs.bandRow != null) flags.bandRow = String(attrs.bandRow) === '1';
        if (attrs.lastRow != null) flags.lastRow = String(attrs.lastRow) === '1';
        if (attrs.firstCol != null) flags.firstCol = String(attrs.firstCol) === '1';
        if (attrs.lastCol != null) flags.lastCol = String(attrs.lastCol) === '1';
        if (attrs.bandCol != null) flags.bandCol = String(attrs.bandCol) === '1';
        if (Object.keys(flags).length) el.tableStyleFlags = flags;
        const inset = readTableInsets(tblPr);
        if (inset) el.inset = inset;
    }

    const rowHeights: number[] = [];
    for (const tr of asArray(tbl && tbl['a:tr'])) {
        const row: PptxTableRow = { cells: [] };
        if (tr && tr.attrs && tr.attrs.h) row.height = emuToPx(tr.attrs.h);

        for (const tc of asArray(tr && tr['a:tc'])) {
            const cell: any = {};
            const attrs = (tc && tc.attrs) || {};
            if (attrs.gridSpan && Number(attrs.gridSpan) > 1) cell.colSpan = Number(attrs.gridSpan);
            if (attrs.rowSpan && Number(attrs.rowSpan) > 1) cell.rowSpan = Number(attrs.rowSpan);

            const { paragraphs, text } = extractTxBody(tc && tc['a:txBody'], themeMap, undefined, resObj ? resolveHyperlink(resObj) : undefined);
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
            else if (paragraphs[0] && (paragraphs[0] as any).rtl) cell.align = 'right'; // RTL 段落缺省起点对齐=右
            // RTL 标记在单元格级冗余：单段单元格走 text 简写时段落级 rtl 会丢失
            if (paragraphs[0] && (paragraphs[0] as any).rtl) cell.rtl = true;

            // 单元格属性：底色、垂直对齐、边框、内边距、对角线
            const tcPr = tc && tc['a:tcPr'];
            if (tcPr) {
                // 底色走 spColor（resolveColorNode）：既解析 schemeClr 主题引用，
                // 也应用 lumMod/lumOff/tint/shade/alpha 修饰符。
                // readSrgbClr 只取引用名（scheme:accent4）会把亮度修饰丢掉，颜色偏深；
                // 且该字符串交给生成端 colorToHex 会被判为非法色而回写成 000000。
                // 注意 resolveColorNode 要求色节点是容器（其下直接挂 a:srgbClr/a:schemeClr），
                // 单元格底色在 tcPr/a:solidFill 下，两种层级都试一遍。
                const fill = spColor(tcPr['a:solidFill'], themeMap)
                    || spColor(tcPr, themeMap)
                    || readSrgbClr(tcPr);
                if (fill) cell.fill = fill;
                if (tcPr.attrs && tcPr.attrs.anchor) cell.valign = VALIGN_MAP[tcPr.attrs.anchor];
                const cellInset = readInsetsFromAttrs(tcPr.attrs);
                if (cellInset) cell.inset = cellInset;

                // 边框回读（a:lnL / a:lnR / a:lnT / a:lnB）
                const edgeKey: Record<string, 'left' | 'right' | 'top' | 'bottom'> = { L: 'left', R: 'right', T: 'top', B: 'bottom' };
                const borders: any = {};
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
                // 对角线边框（a:lnTlToBr / a:lnBlToTr）
                const diagTl = tcPr['a:lnTlToBr'];
                const diagBl = tcPr['a:lnBlToTr'];
                if (diagTl && diagBl) {
                    borders.diagonal = 'both';
                    hasBorder = true;
                } else if (diagTl) {
                    borders.diagonal = 'tlBr';
                    hasBorder = true;
                } else if (diagBl) {
                    borders.diagonal = 'blTr';
                    hasBorder = true;
                }
                if (hasBorder) cell.borders = borders;
            }
            row.cells.push(cell);
        }

        rowHeights.push(row.height !== undefined ? row.height : 0);
        el.rows.push(row);
    }
    if (rowHeights.length && rowHeights.every((h) => h > 0)) el.rowHeights = rowHeights;
    return el as PptxTableElement;
}

/* ======================= 表格样式解析（把 tableStyles.xml + 主题色解析到单元格） ======================= */

/** 从主题对象建立 scheme 名 → #RRGGBB 映射 */
function themeColorsFromTheme(theme?: { colors?: Record<string, string> }): Record<string, string> {
    const map: Record<string, string> = {};
    if (!theme || !theme.colors) return map;
    for (const [k, v] of Object.entries(theme.colors)) {
        const hex = String(v).replace(/^#/, '').toUpperCase();
        if (/^[0-9A-F]{6}$/.test(hex)) map[k.toLowerCase()] = '#' + hex;
    }
    return map;
}

/**
 * 递归把元素树中残留的 'scheme:<name>' 颜色引用解析为绝对色（themeMap 命中时替换）。
 * 未命中的引用保持原样（由渲染端按全局主题兜底）。
 */
function resolveSchemeRefs(value: any, themeMap: Record<string, string>): void {
    if (Array.isArray(value)) {
        for (const v of value) resolveSchemeRefs(v, themeMap);
        return;
    }
    if (!value || typeof value !== 'object') return;
    for (const [k, v] of Object.entries(value)) {
        if (typeof v === 'string' && v.startsWith('scheme:')) {
            const hex = themeMap[v.slice(7).toLowerCase()];
            if (hex) (value as any)[k] = hex;
        } else if (v && typeof v === 'object') {
            resolveSchemeRefs(v, themeMap);
        }
    }
}

/** 从 theme1.xml 原始内容建立 scheme 色映射 */
function themeColorsFromContent(themeContent: any): Record<string, string> {
    const map: Record<string, string> = {};
    if (!themeContent) return map;
    const clrScheme = themeContent['a:theme']
        && themeContent['a:theme']['a:themeElements']
        && themeContent['a:theme']['a:themeElements']['a:clrScheme'];
    if (!clrScheme) return map;
    const slots = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    for (const slot of slots) {
        const node = clrScheme['a:' + slot];
        if (!node) continue;
        const srgb = node['a:srgbClr'];
        const sys = node['a:sysClr'];
        const val = (srgb && srgb.attrs && srgb.attrs.val)
            || (sys && sys.attrs && (sys.attrs.lastClr || sys.attrs.val));
        if (val) map[slot] = '#' + String(val).replace(/^#/, '').toUpperCase();
    }
    // clrMap 别名：tx1/bg1/tx2/bg2 是引用方常用写法（母版 clrMap 默认把 bg2→lt2 等）
    const aliases: Record<string, string> = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
    for (const [alias, slot] of Object.entries(aliases)) {
        if (map[slot] && !map[alias]) map[alias] = map[slot];
    }
    return map;
}

/** 主题样式表（fmtScheme 的 fillStyleLst / lnStyleLst / effectStyleLst），供 p:style 的 fillRef/lnRef/effectRef 解析 */
interface ThemeStyleTables { fills: any[]; lines: any[]; effects: any[]; }
const themeStylesCache = new WeakMap<object, ThemeStyleTables>();

/** 读取并缓存主题样式表（同一 zip / 主题只解析一次） */
async function getThemeStyleTables(zip: JSZip, themeContent?: any): Promise<ThemeStyleTables> {
    const cacheKey = themeContent || zip;
    const cached = themeStylesCache.get(cacheKey);
    if (cached) return cached;
    const tables: ThemeStyleTables = { fills: [], lines: [], effects: [] };
    try {
        const xml = themeContent || await PPTXXmlUtils.readXmlFile(zip, 'ppt/theme/theme1.xml');
        const fmt = xml && xml['a:theme']
            && xml['a:theme']['a:themeElements']
            && xml['a:theme']['a:themeElements']['a:fmtScheme'];
        if (fmt) {
            const listOf = (lstNode: any) => {
                if (!lstNode || typeof lstNode !== 'object') return [];
                const out: any[] = [];
                for (const key of Object.keys(lstNode)) {
                    if (key === 'attrs') continue;
                    for (const child of asArray(lstNode[key])) {
                        if (child && typeof child === 'object') out.push(child);
                    }
                }
                return out;
            };
            tables.fills = listOf(fmt['a:fillStyleLst']);
            tables.lines = listOf(fmt['a:lnStyleLst']);
            tables.effects = listOf(fmt['a:effectStyleLst']);
        }
    } catch { /* 主题缺失时跳过样式引用解析 */ }
    themeStylesCache.set(zip, tables);
    return tables;
}

/**
 * 解析 p:style 的 fillRef/lnRef 主题样式引用。
 * idx 为 1 基（PowerPoint 中 >3 时循环使用并叠加透明度变化，这里按 1~3 取模近似）。
 */
function resolveThemeStyleRef(
    refNode: any,
    styleList: any[],
    themeMap: Record<string, string>
): { color: string; gradient?: { type: 'gradient'; direction?: 'horizontal' | 'vertical' | 'diagonal'; stops: { color: string; position: number }[]; gradientType?: 'linear' | 'radial'; gradientPath?: string }; widthPt?: number } | undefined {
    if (!refNode || !refNode.attrs || !styleList.length) return undefined;
    // idx=0 在 OOXML 中表示 noFill/noLine（不引用任何样式），直接返回 undefined
    const rawIdx = Number(refNode.attrs.idx);
    if (rawIdx === 0) return undefined;
    let idx = rawIdx || 1;
    if (idx < 1) idx = 1;
    const raw = styleList[(idx - 1) % styleList.length];
    if (!raw || typeof raw !== 'object') return undefined;
    // lnStyleLst 的子项包在 a:ln 里；fillStyleLst 的子项即填充节点本身
    const lnInner = asArray(raw['a:ln'])[0];
    const entry = (lnInner && typeof lnInner === 'object') ? lnInner : raw;
    // 样式表内颜色以 phClr 占位，实际色取引用自带的 schemeClr/srgbClr（如 accent1）
    const refColor = resolveColorNode(refNode, themeMap);
    const colorMap = refColor ? { ...themeMap, phclr: refColor } : themeMap;
    // entry 可能是 a:solidFill/a:gradFill 包装层，也可能本身就是填充节点
    const fillNode = entry['a:solidFill'] || entry['a:gradFill'] || entry;
    let color = resolveColorNode(fillNode['a:solidFill'], colorMap)
        || ((fillNode['a:schemeClr'] || fillNode['a:srgbClr']) ? resolveColorNode(fillNode, colorMap) : undefined);
    // 渐变样式（fillStyleLst[1] 等）：完整提取所有 stop（按 phClr 逐个解析）与方向
    let gradient: { type: 'gradient'; direction?: 'horizontal' | 'vertical' | 'diagonal'; stops: { color: string; position: number }[]; gradientType?: 'linear' | 'radial'; gradientPath?: string } | undefined;
    if (!color) {
        const grad = fillNode['a:gradFill'] || (fillNode['a:gsLst'] ? fillNode : undefined);
        const gsLst = grad && grad['a:gsLst'];
        const gsNodes = asArray(gsLst && gsLst['a:gs']);
        const stops = gsNodes
            .map((gs: any) => {
                const c = resolveColorNode(gs, colorMap);
                const pos = gs && gs.attrs && gs.attrs.pos != null ? Number(gs.attrs.pos) / 100000 : 0;
                return c ? { color: c, position: pos } : undefined;
            })
            .filter((s: any): s is { color: string; position: number } => !!s);
        if (stops.length) {
            color = stops[0].color;
            gradient = { type: 'gradient', stops };
            const lin = grad && grad['a:lin'];
            const ang = lin && lin.attrs && lin.attrs.ang !== undefined ? Number(lin.attrs.ang) / 60000 : 0;
            gradient.direction = ang === 90 ? 'vertical' : ang === 45 ? 'diagonal' : 'horizontal';
            // a:path：径向渐变（主题 fillStyleLst 常用），此前会被当成水平线性
            const path = grad && grad['a:path'];
            if (path) {
                gradient.gradientType = 'radial';
                gradient.gradientPath = (path.attrs && path.attrs.path) || 'circle';
            }
        }
    }
    if (!color) return undefined;
    const wAttr = (entry.attrs && entry.attrs.w != null) ? entry.attrs.w
        : (raw.attrs && raw.attrs.w != null) ? raw.attrs.w : undefined;
    const w = wAttr != null ? emuToPt(wAttr) : undefined;
    return { color, gradient, widthPt: w };
}

/** 把 resolveColorNode 的结果拆成 #RRGGBB + 不透明度百分比（兼容 rgb()/rgba() 字符串） */
function splitColorAlpha(color?: string): { hex?: string; alphaPct?: number } {
    if (!color) return {};
    const m = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+)\s*)?\)$/i.exec(String(color).trim());
    if (m) {
        const hex = '#' + [m[1], m[2], m[3]]
            .map((n) => Number(n).toString(16).padStart(2, '0'))
            .join('').toUpperCase();
        return { hex, alphaPct: m[4] != null ? Math.round(Number(m[4]) * 100) : undefined };
    }
    return { hex: (color.startsWith('#') ? color : '#' + color).toUpperCase() };
}

/**
 * 解析 p:style 的 effectRef 主题样式引用，得到形状/文本的外阴影。
 * OOXML 中 effectRef@idx 为样式表索引（0 = 无主题效果）。PowerPoint/预览端对
 * effectStyleLst 实际以 0 基访问（idx=2 → 第三项「强烈阴影」），本解析与之对齐，
 * 让带 effectRef 的形状（如 arc2）显示与预览一致的外阴影。
 */
function resolveThemeEffectRef(
    refNode: any,
    styleList: any[],
    themeMap: Record<string, string>
): { shadow?: { type: 'outer'; blur?: number; distance?: number; angle?: number; color?: string; transparency?: number } } | undefined {
    if (!refNode || !refNode.attrs || !styleList.length) return undefined;
    const idx = Number(refNode.attrs.idx) || 0;
    if (idx <= 0) return undefined;
    const raw = styleList[idx < styleList.length ? idx : idx - 1];
    if (!raw || typeof raw !== 'object') return undefined;
    const effLst = raw['a:effectLst'];
    if (!effLst) return undefined;
    const outer = effLst['a:outerShdw'];
    if (!outer || !outer.attrs) return undefined;
    const a = outer.attrs;
    const refColor = resolveColorNode(refNode, themeMap);
    const colorMap = refColor ? { ...themeMap, phclr: refColor } : themeMap;
    let color = resolveColorNode(outer, colorMap)
        || resolveColorNode(outer['a:srgbClr'], colorMap)
        || resolveColorNode(outer['a:schemeClr'], colorMap);
    // resolveColorNode 在带 a:alpha 时返回 rgb()/rgba() 字符串，这里拆成 hex + 不透明度
    const parsed = splitColorAlpha(color);
    color = parsed.hex;
    const alphaNode = (outer['a:srgbClr'] || outer['a:schemeClr'] || {})['a:alpha'];
    // a:alpha val=60000 → 60% 不透明度；PptxShapeEffects 内部统一用 transparency = 100 - alpha%
    const alphaPct = parsed.alphaPct != null
        ? parsed.alphaPct
        : (alphaNode && alphaNode.attrs ? Math.round(Number(alphaNode.attrs.val) / 1000) : undefined);
    const shadow: any = { type: 'outer' };
    if (a.blurRad != null) shadow.blur = Math.round(emuToPt(a.blurRad) * 100) / 100;
    if (a.dist != null) shadow.distance = Math.round(emuToPt(a.dist) * 100) / 100;
    if (a.dir != null) shadow.angle = Math.round(Number(a.dir) / 60000);
    if (color) shadow.color = color.replace(/^#/, '').toUpperCase();
    if (alphaPct != null) shadow.transparency = 100 - alphaPct;
    if (shadow.color || shadow.blur != null || shadow.distance != null) return { shadow };
    return undefined;
}

/** 解析 a:solidFill / a:srgbClr / a:schemeClr 为 #RRGGBB（含主题色变换） */
function resolveColorNode(node: any, themeMap: Record<string, string>): string | undefined {
    if (!node) return undefined;
    let base: string | undefined;
    let modNode: any;
    const srgb = node['a:srgbClr'];
    if (srgb && srgb.attrs && srgb.attrs.val) {
        base = '#' + String(srgb.attrs.val).replace(/^#/, '').toUpperCase();
        modNode = srgb;
    }
    const sch = node['a:schemeClr'];
    if (sch && sch.attrs && sch.attrs.val) {
        const key = String(sch.attrs.val).toLowerCase();
        const mapped = themeMap[key];
        base = mapped || '#' + String(sch.attrs.val).replace(/^#/, '').toUpperCase();
        modNode = sch;
    }
    if (!base) return undefined;
    let tc = tinycolor(base);
    if (modNode) {
        const readMod = (tag: string) => {
            const n = modNode[tag];
            return n && n.attrs && n.attrs.val != null ? Number(n.attrs.val) : undefined;
        };
        const tint = readMod('a:tint');
        const shade = readMod('a:shade');
        const lumMod = readMod('a:lumMod');
        const lumOff = readMod('a:lumOff');
        const alpha = readMod('a:alpha');
        // 与预览端 getSolidFill/applyTint 保持一致：采用 OOXML 标准 HSL 公式
        // tint：l' = l * t + (1 - t)（变亮）；shade：l' = l * s（变暗）。t/s 为 0~1 比例。
        if (tint != null) {
            const t = Math.max(0, Math.min(1, tint / 100000));
            const hsl = tc.toHsl();
            hsl.l = Math.max(0, Math.min(1, hsl.l * t + (1 - t)));
            tc = tinycolor(hsl);
        }
        if (shade != null) {
            const s = Math.max(0, Math.min(1, shade / 100000));
            const hsl = tc.toHsl();
            hsl.l = Math.max(0, Math.min(1, hsl.l * s));
            tc = tinycolor(hsl);
        }
        if (lumMod != null) {
            const hsl = tc.toHsl();
            hsl.l = Math.max(0, Math.min(1, hsl.l * (lumMod / 100000)));
            tc = tinycolor(hsl);
        }
        if (lumOff != null) {
            const hsl = tc.toHsl();
            hsl.l = Math.max(0, Math.min(1, hsl.l + lumOff / 100000));
            tc = tinycolor(hsl);
        }
        if (alpha != null) {
            const a = Math.max(0, Math.min(1, tc.getAlpha() * (alpha / 100000)));
            tc.setAlpha(a);
            return tc.toRgbString();
        }
    }
    return tc.toHexString().toUpperCase();
}

/** 读取一条线属性（颜色/宽度/无填充）。兼容直接传 a:ln 或包了一层的边节点（a:left 等） */
function readLineStyle(ln: any, themeMap: Record<string, string>): { color?: string; width?: number } | 'none' | undefined {
    if (!ln) return undefined;
    const node = ln['a:ln'] || ln;
    if (node['a:noFill']) return 'none';
    const attrs = node.attrs || {};
    const width = attrs.w != null ? Math.round(Number(attrs.w) / EMU_PER_PT) : DEFAULT_LN_PT;
    const color = resolveColorNode(node['a:solidFill'], themeMap)
        || resolveColorNode(node, themeMap) || '#000000';
    return { color, width };
}

/** 读取表格样式某部分（wholeTbl/firstRow/...）的填充、文字色、边框 */
function readTableStylePart(part: any, themeMap: Record<string, string>) {
    if (!part) return undefined;
    const out: any = {};
    const tcStyle = part['a:tcStyle'];
    if (tcStyle) {
        // a:fill 下是 a:solidFill / a:gradFill 等，solidFill 内才是具体颜色节点
        const fillNode = tcStyle['a:fill'] && (tcStyle['a:fill']['a:solidFill'] || tcStyle['a:fill']);
        if (fillNode) out.fill = resolveColorNode(fillNode, themeMap);
        const bdr = tcStyle['a:tcBdr'];
        if (bdr) {
            const sides: any = {};
            const sideMap: Record<string, string> = { left: 'left', right: 'right', top: 'top', bottom: 'bottom' };
            let has = false;
            for (const [xml, key] of Object.entries(sideMap)) {
                const v = readLineStyle(bdr['a:' + xml], themeMap);
                if (v !== undefined) { sides[key] = v; has = true; }
            }
            const insideH = bdr['a:insideH'];
            const insideV = bdr['a:insideV'];
            if (insideH !== undefined) { sides.insideH = readLineStyle(insideH, themeMap); has = true; }
            if (insideV !== undefined) { sides.insideV = readLineStyle(insideV, themeMap); has = true; }
            if (has) out.borders = sides;
        }
    }
    const txStyle = part['a:tcTxStyle'];
    if (txStyle) {
        const fontRef = txStyle['a:fontRef'];
        const defRPr = txStyle['a:defRPr'];
        const colorNode = txStyle['a:schemeClr'] || txStyle['a:srgbClr'] || (fontRef && (fontRef['a:schemeClr'] || fontRef['a:srgbClr']));
        if (colorNode) out.color = resolveColorNode(colorNode, themeMap) || resolveColorNode(txStyle, themeMap);
        if (txStyle.attrs) {
            if (txStyle.attrs.b != null) out.bold = String(txStyle.attrs.b) === '1' || String(txStyle.attrs.b) === 'on';
            if (txStyle.attrs.i != null) out.italic = String(txStyle.attrs.i) === '1' || String(txStyle.attrs.i) === 'on';
        }
        if (defRPr && defRPr.attrs) {
            if (defRPr.attrs.b != null) out.bold = String(defRPr.attrs.b) === '1' || String(defRPr.attrs.b) === 'on';
            if (defRPr.attrs.sz != null) out.fontSize = Math.round(Number(defRPr.attrs.sz) / 100);
        }
    }
    return Object.keys(out).length ? out : undefined;
}

/** 把 tableStyles.xml 中的样式应用到表格各单元格（仅处理 solidFill 主题色） */
function applyTableStyle(el: PptxTableElement, tableStyles: any, theme?: { colors?: Record<string, string> }) {
    const styleId = el.tableStyleId;
    if (!styleId || !tableStyles) return;
    // 优先用该页真实主题（slideData.themeContent，可能为 theme2 等多主题文件），
    // 回退到全局 options.theme，确保多主题 PPTX 的表格配色与 PowerPoint 一致。
    const themeMap = tableStyles._themeContent
        ? themeColorsFromContent(tableStyles._themeContent)
        : (theme ? themeColorsFromTheme(theme) : {});
    const styleLst = tableStyles['a:tblStyleLst'] || tableStyles;
    const styles = asArray(styleLst['a:tblStyle']);
    const style = styles.find((s: any) => s && s.attrs && s.attrs.styleId === styleId);
    if (!style) return;

    const flags = (el as any).tableStyleFlags || {};
    const parts: Record<string, any> = {};
    const names = ['wholeTbl', 'firstRow', 'lastRow', 'band1H', 'band2H', 'firstCol', 'lastCol', 'band1V', 'band2V', 'nwCell', 'neCell', 'swCell', 'seCell'];
    for (const n of names) {
        const part = readTableStylePart(style['a:' + n], themeMap);
        if (part) parts[n] = part;
    }
    if (Object.keys(parts).length === 0) return;

    const rowCount = el.rows.length;
    const colCount = Math.max(1, ...el.rows.map((r) => (r.cells || []).length));

    function mergeInto(cell: any, part: any) {
        if (!part) return;
        // 样式部件按优先级顺序依次覆盖（wholeTbl → band → 首末行列），
        // 单元格自身的显式样式（OOXML tcPr 直填）始终优先于任何样式部件
        if (part.fill != null) cell.fill = cell.fillExplicit ?? part.fill;
        if (part.color != null) cell.color = cell.colorExplicit ?? part.color;
        if (part.bold != null && !cell.boldExplicit) cell.bold = part.bold;
        if (part.italic != null && !cell.italicExplicit) cell.italic = part.italic;
        if (part.fontSize != null && cell.fontSizeExplicit == null) cell.fontSize = part.fontSize;
        if (part.borders) {
            if (!cell.borders) cell.borders = {};
            for (const [k, v] of Object.entries(part.borders)) {
                // 单元格显式边框优先于样式部件
                if (cell.bordersExplicit && cell.bordersExplicit[k] !== undefined) continue;
                cell.borders[k] = v;
            }
        }
    }

    for (let ri = 0; ri < rowCount; ri++) {
        const row = el.rows[ri];
        for (let ci = 0; ci < (row.cells || []).length; ci++) {
            const cell = row.cells[ci] as any;
            // 先保存单元格显式样式，样式部件覆盖后再恢复
            cell.fillExplicit = cell.fill;
            cell.colorExplicit = cell.color;
            cell.boldExplicit = cell.bold;
            cell.italicExplicit = cell.italic;
            cell.fontSizeExplicit = cell.fontSize;
            cell.bordersExplicit = cell.borders ? { ...cell.borders } : undefined;

            mergeInto(cell, parts.wholeTbl);
            if (flags.bandRow) {
                if ((ri % 2) === 0 && parts.band2H) mergeInto(cell, parts.band2H);
                if ((ri % 2) === 1 && parts.band1H) mergeInto(cell, parts.band1H);
            }
            if (flags.bandCol) {
                if ((ci % 2) === 0 && parts.band2V) mergeInto(cell, parts.band2V);
                if ((ci % 2) === 1 && parts.band1V) mergeInto(cell, parts.band1V);
            }
            if (flags.firstRow && ri === 0 && parts.firstRow) mergeInto(cell, parts.firstRow);
            if (flags.lastRow && ri === rowCount - 1 && parts.lastRow) mergeInto(cell, parts.lastRow);
            if (flags.firstCol && ci === 0 && parts.firstCol) mergeInto(cell, parts.firstCol);
            if (flags.lastCol && ci === colCount - 1 && parts.lastCol) mergeInto(cell, parts.lastCol);

            // 清理临时字段
            delete cell.fillExplicit;
            delete cell.colorExplicit;
            delete cell.boldExplicit;
            delete cell.italicExplicit;
            delete cell.fontSizeExplicit;
            delete cell.bordersExplicit;
        }
    }
}

/** Microsoft 缓存绘图部件（diagramDrawing）关系类型 */
const MS_DIAGRAM_DRAWING_REL = 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';

/** 从 resObj 定位图示缓存绘图部件（ppt/diagrams/drawingN.xml） */
function findDiagramDrawingPath(resObj: Record<string, { type?: string; target?: string }>): string | undefined {
    for (const rel of Object.values(resObj || {})) {
        if (!rel || !rel.target) continue;
        if (rel.type === MS_DIAGRAM_DRAWING_REL || /diagrams\/drawing\d+\.xml$/i.test(String(rel.target))) {
            return resolvePart(rel.target);
        }
    }
    return undefined;
}

/** 从 dsp:txBody 提取文字与基础文字样式 */
function readDiagramShapeText(txBody: any, themeMap: Record<string, string>, defaultColor?: string) {
    if (!txBody) return undefined;
    const anchor = txBody['a:bodyPr'] && txBody['a:bodyPr'].attrs ? String(txBody['a:bodyPr'].attrs.anchor || '') : '';
    const lines: string[] = [];
    let fontSize: number | undefined, color: string | undefined, bold: boolean | undefined, align: string | undefined;
    // 形状级默认字号：a:lstStyle/a:lvl1pPr/a:defRPr@sz（run 上常省略 sz）
    const lstLvl1 = txBody['a:lstStyle'] && txBody['a:lstStyle']['a:lvl1pPr'];
    const lstDefRPr = lstLvl1 && lstLvl1['a:defRPr'];
    // 形状级默认对齐：a:lstStyle/a:lvl1pPr@algn（pPr 上常省略 algn）
    if (!align && lstLvl1 && lstLvl1.attrs && lstLvl1.attrs.algn) align = String(lstLvl1.attrs.algn);
    if (lstDefRPr && lstDefRPr.attrs && lstDefRPr.attrs.sz != null) {
        fontSize = Math.round(Number(lstDefRPr.attrs.sz) / 100) || undefined;
    }
    if (lstDefRPr && lstDefRPr.attrs && lstDefRPr.attrs.b != null) {
        bold = String(lstDefRPr.attrs.b) === '1' || String(lstDefRPr.attrs.b) === 'on';
    }
    for (const p of asArray(txBody['a:p'])) {
        if (!p || typeof p !== 'object') continue;
        if (!align && p['a:pPr'] && p['a:pPr'].attrs && p['a:pPr'].attrs.algn) align = String(p['a:pPr'].attrs.algn);
        let line = '';
        for (const r of asArray(p['a:r'])) {
            if (!r || typeof r !== 'object') continue;
            const t = r['a:t'];
            if (typeof t !== 'string') continue;
            line += t;
            const rPr = r['a:rPr'];
            if (rPr && rPr.attrs) {
                if (fontSize === undefined && rPr.attrs.sz != null) fontSize = Math.round(Number(rPr.attrs.sz) / 100) || undefined;
                if (bold === undefined && rPr.attrs.b != null) bold = String(rPr.attrs.b) === '1' || String(rPr.attrs.b) === 'on';
            }
            if (color === undefined && rPr) color = resolveColorNode(rPr['a:solidFill'], themeMap) || readSrgbClr(rPr);
        }
        lines.push(line);
    }
    const text = lines.join('\n');
    // run 无显式色时用形状的 fontRef 默认字色
    const finalColor = color || defaultColor;
    if (!text.trim() && fontSize === undefined && finalColor === undefined) return { text: text || undefined, anchor: anchor || undefined };
    return { text: text || undefined, fontSize, color: finalColor, bold, align, anchor: anchor || undefined };
}

/**
 * 从缓存绘图部件（drawingN.xml）提取图示形状树（坐标 px，相对图示框）。
 * 该部件由 PowerPoint 写入，含按布局算好的形状位置/填充/文字与连接线，
 * 是预览端（pptxToHtml）树形展示效果的同一数据源。
 */
async function extractDiagramShapes(
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip,
    themeMap: Record<string, string>,
    frameW: number,
    frameH: number,
    themeContent?: any
): Promise<PptxDiagramShape[] | undefined> {
    const drawingPath = findDiagramDrawingPath(resObj);
    if (!drawingPath) return undefined;
    let drawingXml: any;
    try {
        drawingXml = await PPTXXmlUtils.readXmlFile(zip, drawingPath);
    } catch {
        return undefined;
    }
    const spTree = drawingXml && drawingXml['dsp:drawing'] && drawingXml['dsp:drawing']['dsp:spTree'];
    if (!spTree) return undefined;

    // 绘图坐标空间归一化（chOff/chExt → 图示框 px）；chExt 缺失时按 EMU 直转
    const gxf = spTree['dsp:grpSpPr'] && spTree['dsp:grpSpPr']['a:xfrm'];
    const chOffAttrs = gxf && gxf['a:chOff'] && gxf['a:chOff'].attrs || {};
    const chExtAttrs = gxf && gxf['a:chExt'] && gxf['a:chExt'].attrs || {};
    const chX = Number(chOffAttrs.x) || 0, chY = Number(chOffAttrs.y) || 0;
    const chW = Number(chExtAttrs.cx) || 0, chH = Number(chExtAttrs.cy) || 0;
    const mapX = (v: unknown) => chW > 0 ? ((Number(v) || 0) - chX) / chW * frameW : emuToPx(v);
    const mapY = (v: unknown) => chH > 0 ? ((Number(v) || 0) - chY) / chH * frameH : emuToPx(v);
    const mapW = (v: unknown) => chW > 0 ? (Number(v) || 0) / chW * frameW : emuToPx(v);
    const mapH = (v: unknown) => chH > 0 ? (Number(v) || 0) / chH * frameH : emuToPx(v);

    // dsp:spPr 里通常没有显式 a:solidFill/a:ln，填充与描边来自 dsp:style 的
    // fillRef/lnRef 主题样式引用（SmartArt 专属机制，与 p:style 同构）。
    // 预览端把 dsp: 前缀重写为 p: 后复用形状渲染路径的 p:style 解析，
    // 这里显式按主题样式表解析，保持两端一致。
    const tables = await getThemeStyleTables(zip, themeContent);

    const shapes: PptxDiagramShape[] = [];
    const pushShape = (sp: any, isConnector: boolean) => {
        if (!sp || typeof sp !== 'object') return;
        const spPr = sp['dsp:spPr'];
        const dspStyle = sp['dsp:style'];
        const xf = spPr && spPr['a:xfrm'];
        const off = (xf && xf['a:off'] && xf['a:off'].attrs) || {};
        const ext = (xf && xf['a:ext'] && xf['a:ext'].attrs) || {};
        const width = mapW(ext.cx), height = mapH(ext.cy);
        if (!isFinite(width) || !isFinite(height) || (width <= 0 && height <= 0)) return;
        const geom = spPr && spPr['a:prstGeom'];
        const shape: PptxDiagramShape = {
            x: Math.round(mapX(off.x) * 100) / 100,
            y: Math.round(mapY(off.y) * 100) / 100,
            width: Math.round(width * 100) / 100,
            height: Math.round(height * 100) / 100,
            adjust: {}
        };
        if (geom && geom.attrs && geom.attrs.prst) shape.prst = String(geom.attrs.prst);
        // 预设几何的调整值（a:prstGeom/a:avLst/a:gd，如 arc 的起止角度 adj1/adj2）
        const avLst = geom && geom['a:avLst'];
        if (avLst && typeof avLst === 'object') {
            for (const gd of asArray(avLst['a:gd'])) {
                if (!gd || !gd.attrs) continue;
                const nm = gd.attrs.name ? String(gd.attrs.name) : '';
                const raw = gd.attrs.fmla != null ? String(gd.attrs.fmla) : (gd.attrs.val != null ? String(gd.attrs.val) : '');
                const v = /(-?\d+(?:\.\d+)?)\s*$/.exec(raw);
                if (nm && v) (shape.adjust as Record<string, number>)[nm] = Math.round(Number(v[1]) * 100) / 100;
            }
        }
        const xattrs = (xf && xf.attrs) || {};
        if (String(xattrs.flipH) === '1') shape.flipH = true;
        if (String(xattrs.flipV) === '1') shape.flipV = true;

        // 样式引用回退：spPr 无显式填充/描边时按 dsp:style 的 fillRef/lnRef 解析
        const fillRef = dspStyle && dspStyle['a:fillRef'];
        const lnRef = dspStyle && dspStyle['a:lnRef'];

        if (isConnector) {
            shape.connector = true;
            const ln = spPr && spPr['a:ln'];
            if (ln) {
                const lc = resolveColorNode(ln['a:solidFill'], themeMap);
                if (lc) shape.lineColor = lc;
                if (ln.attrs && ln.attrs.w != null) shape.lineWidth = emuToPt(ln.attrs.w);
            }
            if (!shape.lineColor && lnRef) {
                const ref = resolveThemeStyleRef(lnRef, tables.lines, themeMap);
                if (ref) {
                    shape.lineColor = ref.color;
                    if (shape.lineWidth == null) shape.lineWidth = ref.widthPt || DEFAULT_LN_PT;
                }
            }
            shapes.push(shape);
            return;
        }

        const fill = spPr && resolveColorNode(spPr['a:solidFill'], themeMap);
        if (fill) shape.fill = fill;
        else if (spPr && spPr['a:noFill']) shape.fill = 'none';
        else if (fillRef) {
            // fillRef@idx=0（或 1000）表示无填充（连接线/弧形等）
            const refIdx = Number(fillRef.attrs && fillRef.attrs.idx) || 0;
            if (refIdx === 0 || refIdx === 1000) shape.fill = 'none';
            else {
                const ref = resolveThemeStyleRef(fillRef, tables.fills, themeMap);
                if (ref) shape.fill = ref.color;
            }
        }
        const ln = spPr && spPr['a:ln'];
        if (ln && !ln['a:noFill']) {
            const lc = resolveColorNode(ln['a:solidFill'], themeMap);
            if (lc) {
                shape.lineColor = lc;
                shape.lineWidth = ln.attrs && ln.attrs.w != null ? emuToPt(ln.attrs.w) : 1;
            }
        }
        if (!shape.lineColor && lnRef) {
            const ref = resolveThemeStyleRef(lnRef, tables.lines, themeMap);
            if (ref) {
                shape.lineColor = ref.color;
                if (shape.lineWidth == null) shape.lineWidth = ref.widthPt || DEFAULT_LN_PT;
            }
        }
        // dsp:style/a:fontRef 是形状内文字的默认字色（run 无显式色时生效）
        const fontRefColor = dspStyle ? spColor(dspStyle['a:fontRef'], themeMap) : undefined;
        const txt = readDiagramShapeText(sp['dsp:txBody'], themeMap, fontRefColor);
        if (txt) {
            if (txt.text) shape.text = txt.text;
            if (txt.fontSize) shape.fontSize = txt.fontSize;
            if (txt.color) shape.color = txt.color;
            if (txt.bold) shape.bold = true;
            if (txt.align) shape.align = txt.align;
            if (txt.anchor) shape.anchor = txt.anchor;
        }
        shapes.push(shape);
    };

    for (const sp of asArray(spTree['dsp:sp'])) pushShape(sp, false);
    for (const cxn of asArray(spTree['dsp:cxnSp'])) pushShape(cxn, true);
    return shapes.length ? shapes : undefined;
}

/** p:graphicFrame（SmartArt 图示）→ PptxDiagramElement：提取数据部件文本 + 缓存绘图形状 */
async function diagramToElement(
    node: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip,
    themeMap: Record<string, string> = {},
    themeContent?: any
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
    // 缓存绘图形状：与预览端同源的树形布局（节点框 + 连接线）
    try {
        const shapes = await extractDiagramShapes(resObj, zip, themeMap, el.width, el.height, themeContent);
        if (shapes) el.shapes = shapes;
    } catch { /* 形状提取失败不影响文本兜底 */ }
    return el;
}

/** p:graphicFrame → 按 graphicData 类型分派到 table / diagram / chart */
async function graphicFrameToElement(
    node: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip,
    themeMap: Record<string, string> = {},
    themeContent?: any
): Promise<PptxElement | null> {
    const graphicData = node['a:graphic'] && node['a:graphic']['a:graphicData'];
    if (!graphicData) return null;
    const uri = graphicData.attrs && graphicData.attrs.uri ? String(graphicData.attrs.uri) : '';

    // 表格：以 a:tbl 实际存在为准（仅凭 uri 声明无法还原内容）
    const tbl = graphicData['a:tbl'];
    if (tbl) return tableToElement(tbl, node, themeMap, resObj);
    // 图示 / SmartArt：dgm:rel 或其命名空间 uri
    if (graphicData['dgm:rel'] || uri === URI_DIAGRAM || /diagram/.test(uri)) {
        return await diagramToElement(node, resObj, zip, themeMap, themeContent);
    }
    return await graphicFrameToChart(node, resObj, zip, themeMap);
}

/** p:graphicFrame → chart 元素 */
async function graphicFrameToChart(
    node: any,
    resObj: Record<string, { type?: string; target?: string }>,
    zip: JSZip,
    themeMap: Record<string, string> = {}
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
        chartSemantic = extractChart(chartXml, themeMap);
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
        'categories', 'title', 'legend', 'legendPosition', 'grouping', 'varyColors', 'barDir',
        'holeSize', 'smooth', 'marker', 'ofPieType', 'numberFormat',
        'bubble3D', 'showNegBubbles', 'bubbleScale', 'wireframe', 'spaceFill'
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
        const mediaMimeMap: Record<string, string> = {
            mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', webm: 'video/webm', avi: 'video/x-msvideo',
            mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', aac: 'audio/aac', ogg: 'audio/ogg', wma: 'audio/x-ms-wma'
        };

        let mData: string | undefined;
        if (mPart) {
            try {
                const f = zip.file(mPart);
                if (f) {
                    const b64 = await f.async('base64');
                    mData = `data:${mediaMimeMap[mExt] || 'application/octet-stream'};base64,${b64}`;
                }
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
                        const pExt = (pPart.split('.').pop() || 'png').toLowerCase();
                        const pb64 = await f.async('base64');
                        mediaEl.poster = {
                            data: `data:${IMAGE_MIME[pExt] || 'image/png'};base64,${pb64}`,
                            extension: pExt
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
    // 先提取主题配色，后续表格样式解析需要使用
    const theme = await extractTheme(zip);
    const slides = await Promise.all(
        (parsedData.slides || []).map(async (s: any) => extractSlideToStandard(s.data, zip, { ...options, theme: theme as any }))
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
    // 主题配色方案（a:theme/a:clrScheme）：供 scheme 主题色引用解析为具体色，
    // 避免退化为默认蓝主题（如本样例自定义主题覆盖了 accent4/accent5）。
    if (theme) doc.theme = theme;
    // 原始 theme1.xml 整串：用于生成端无损回退，保留 fmtScheme/fontScheme 等细节（SmartArt/样式引用保真）
    try {
        const themeFile = zip.file('ppt/theme/theme1.xml');
        if (themeFile) doc.themeXml = await themeFile.async('string');
    } catch { /* ignore */ }

    // 原始 tableStyles.xml 整串：表格样式的网格线颜色/底纹/条带由该文件的 GUID 定义决定，
    // 重新生成的等价定义只有通用黑网格，会让白网格表格变黑
    try {
        const tsFile = zip.file('ppt/tableStyles.xml');
        if (tsFile) doc.tableStylesXml = await tsFile.async('string');
    } catch { /* ignore */ }

    // 多主题/母版/版式无损回退：解析端回读原始母版、版式、主题整串 XML，
    // 并维护 幻灯片→版式→母版→主题 的映射，使生成端按原结构写回（保留各页绑定的真实主题）。
    try {
        const slideFileNames = (parsedData.slides || []).map((s: any) => String(s.fileName || ''));
        const pkg = await extractPackageParts(zip, slideFileNames);
        if (pkg) {
            doc.themeXmls = pkg.themeXmls;
            doc.masters = pkg.masters as any;
            pkg.slideLayoutIndex.forEach((li, si) => { if (slides[si]) slides[si].layout = li; });
        }
    } catch { /* ignore */ }

    return doc;
}

/**
 * 关系部件 → 关系 id 映射（{ Id: { type: 短名, target } }）
 * type 去掉 relationships 前缀（slideLayout / slideMaster / theme / image ...）
 */
async function readRelMap(zip: JSZip, relsPath: string): Promise<Record<string, { type: string; target: string }>> {
    const out: Record<string, { type: string; target: string }> = {};
    let xml: any;
    try { xml = await PPTXXmlUtils.readXmlFile(zip, relsPath); } catch { return out; }
    const relsRoot = xml && (xml.Relationships || (xml['Relationships:Relationships'] as any));
    if (!relsRoot) return out;
    const raw = relsRoot.Relationship || (relsRoot as any)['Relationship:Relationship'];
    const list: any[] = Array.isArray(raw) ? raw : (raw ? [raw] : []);
    for (const rel of list) {
        const id = rel && rel.attrs && rel.attrs.Id;
        if (!id) continue;
        const fullType = String((rel.attrs && rel.attrs.Type) || '');
        out[String(id)] = {
            type: fullType.replace(REL_PREFIX, '').replace(REL_PREFIX_MS, 'ms:'),
            target: String((rel.attrs && rel.attrs.Target) || '')
        };
    }
    return out;
}

/** Microsoft Office 2007 关系前缀 */
const REL_PREFIX_MS = 'http://schemas.microsoft.com/office/2007/relationships/';

/** 母版/版式/主题原始部件与 幻灯片→版式 映射（多主题无损回退用） */
interface PackageParts {
    /** 主题部件整串，索引 0 对应 ppt/theme/theme1.xml */
    themeXmls: string[];
    /** 母版（__rawXml 原始整串 + 绑定的 themeXml + 其版式原始整串） */
    masters: { __rawXml: string; themeXml: string; layouts: { __rawXml: string }[] }[];
    /** 每页（按显示顺序）所用版式在展平版式序列中的下标 */
    slideLayoutIndex: number[];
}

/**
 * 回读母版/版式/主题原始部件整串 XML，并解析 幻灯片→版式→母版→主题 关系链。
 * PPTX 允许不同母版绑定不同主题部件（多主题文件）；生成端若统一写单一 theme1.xml，
 * 各页 schemeClr 引用会解析到错误配色（SmartArt dsp:style、图表系列色尤其明显）。
 */
async function extractPackageParts(zip: JSZip, slideFileNames: string[]): Promise<PackageParts | null> {
    /** 收集形如 prefix + N + suffix 的部件编号，按 N 升序 */
    const partNos = (re: RegExp): number[] => Object.keys(zip.files)
        .map((p) => { const m = re.exec(p); return m ? Number(m[1]) : NaN; })
        .filter((n) => Number.isFinite(n))
        .sort((a, b) => a - b);

    const themeNos = partNos(/^ppt\/theme\/theme(\d+)\.xml$/);
    if (!themeNos.length) return null;
    const themeXmls: string[] = [];
    for (const n of themeNos) {
        const f = zip.file(`ppt/theme/theme${n}.xml`);
        themeXmls[n - 1] = f ? await f.async('string') : '';
    }

    // 版式 → 所属母版号，并回读版式原始整串
    const layoutNos = partNos(/^ppt\/slideLayouts\/slideLayout(\d+)\.xml$/);
    const layoutMaster = new Map<number, number>();
    const layoutRaw = new Map<number, string>();
    for (const n of layoutNos) {
        const f = zip.file(`ppt/slideLayouts/slideLayout${n}.xml`);
        if (f) { try { layoutRaw.set(n, await f.async('string')); } catch { /* 忽略 */ } }
        const rels = await readRelMap(zip, `ppt/slideLayouts/_rels/slideLayout${n}.xml.rels`);
        for (const r of Object.values(rels)) {
            if (r.type !== 'slideMaster') continue;
            const m = /slideMaster(\d+)\.xml/i.exec(r.target);
            if (m) layoutMaster.set(n, Number(m[1]));
        }
    }

    // 母版 → 绑定主题号 + 其版式（按 rels 声明顺序，保持与原包一致的版式次序）
    const masterNos = partNos(/^ppt\/slideMasters\/slideMaster(\d+)\.xml$/);
    if (!masterNos.length) return null;
    const masters: PackageParts['masters'] = [];
    const layoutNoToFlat = new Map<number, number>();
    let flat = 0;
    for (const n of masterNos) {
        const f = zip.file(`ppt/slideMasters/slideMaster${n}.xml`);
        let raw = '';
        if (f) { try { raw = await f.async('string'); } catch { /* 忽略 */ } }
        const rels = await readRelMap(zip, `ppt/slideMasters/_rels/slideMaster${n}.xml.rels`);
        let themeNo = 1;
        const relLayoutNos: number[] = [];
        for (const r of Object.values(rels)) {
            if (r.type === 'theme') {
                const m = /theme(\d+)\.xml/i.exec(r.target);
                if (m) themeNo = Number(m[1]);
            } else if (r.type === 'slideLayout') {
                const m = /slideLayout(\d+)\.xml/i.exec(r.target);
                if (m) relLayoutNos.push(Number(m[1]));
            }
        }
        // 母版 rels 偶有缺失，补齐归属该母版的其余版式
        for (const ln of layoutNos) {
            if (layoutMaster.get(ln) === n && relLayoutNos.indexOf(ln) < 0) relLayoutNos.push(ln);
        }
        const layouts = relLayoutNos.map((ln) => {
            layoutNoToFlat.set(ln, flat++);
            return { __rawXml: layoutRaw.get(ln) || '' };
        });
        masters.push({ __rawXml: raw, themeXml: themeXmls[themeNo - 1] || '', layouts });
    }

    // 幻灯片（按显示顺序）→ 版式展平下标
    const slideLayoutIndex: number[] = [];
    for (const name of slideFileNames) {
        let idx = 0;
        if (name) {
            // fileName 为不带扩展名的基名（如 slide1），rels 位于 ppt/slides/_rels/slideN.xml.rels
            const base = name.indexOf('/') >= 0 ? name.slice(name.lastIndexOf('/') + 1) : name;
            const rels = await readRelMap(zip, `ppt/slides/_rels/${base}.xml.rels`);
            for (const r of Object.values(rels)) {
                if (r.type !== 'slideLayout') continue;
                const m = /slideLayout(\d+)\.xml/i.exec(r.target);
                if (m) {
                    const fi = layoutNoToFlat.get(Number(m[1]));
                    if (fi !== undefined) idx = fi;
                }
                break;
            }
        }
        slideLayoutIndex.push(idx);
    }

    return { themeXmls, masters, slideLayoutIndex };
}

/**
 * 提取文档主题配色（ppt/theme/theme1.xml 的 a:clrScheme）。
 * @returns PptxTheme（含 colors：12 个色槽，hex 大写带 #）；无主题时 undefined
 */
async function extractTheme(zip: JSZip): Promise<PptxTheme | undefined> {
    try {
        const file = zip.file('ppt/theme/theme1.xml');
        if (!file) return undefined;
        const xml = await PPTXXmlUtils.readXmlFile(zip, 'ppt/theme/theme1.xml');
        const clr = xml && xml['a:theme'] && xml['a:theme']['a:themeElements'] && xml['a:theme']['a:themeElements']['a:clrScheme'];
        if (!clr) return undefined;
        const colors: Record<string, string> = {};
        const slots = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
        for (const slot of slots) {
            const node = clr['a:' + slot];
            if (!node) continue;
            const srgb = node['a:srgbClr'];
            const sys = node['a:sysClr'];
            const val = (srgb && srgb.attrs && srgb.attrs.val)
                || (sys && sys.attrs && sys.attrs.lastClr);
            if (val) colors[slot] = '#' + String(val).toUpperCase();
        }
        const name = xml['a:theme'] && xml['a:theme'].attrs && xml['a:theme'].attrs.name;
        return { name: name ? String(name) : undefined, colors: colors as PptxThemeColorScheme };
    } catch {
        return undefined;
    }
}
