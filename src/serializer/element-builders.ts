/**
 * 元素构建器模块
 *
 * 将幻灯片元素（文本/形状/图片）的 JSON 描述转换为 OOXML 节点。
 * 构建过程中通过 ctx 收集图片媒体文件与超链接关系，
 * 供上层（jsonToPptx / editPptx.addSlide）打包 zip 时使用。
 *
 * 支持的元素 JSON 格式：
 * - 文本：{ type:'text', x,y,width,height, text | runs | paragraphs,
 *          align, valign, fontSize(pt), color, bold, italic, underline, fontFace, href }
 * - 形状：{ type:'shape', shapeType:'rect'|'roundRect'|'ellipse'|..., x,y,width,height,
 *          fill:{color}|'none', line:{color,width(pt)}|'none', rotation(deg) }
 * - 图片：{ type:'image', x,y,width,height, data(dataURL/base64)|src(URL), extension, href }
 *
 * @module serializer/element-builders
 */

import { xmlNode, pxToEmu, ptToSz, ptToEmu, degToRot, colorToHex, NS, escapeXml, type BuilderNode } from './xml-builder';
import { REL_TYPES } from './templates';
import type { PptxBackground, PptxTransition } from '../types/pptx-document';

/** 关系记录（写入 slide rels） */
export interface SerializerRel {
    relId: string;
    type: string;
    target: string;
    external: boolean;
}
/** 媒体文件记录 */
export interface SerializerMedia {
    name: string;
    base64: string;
}
/** 图表部件记录 */
export interface SerializerChart {
    name: string;
    xml: string;
}
/** SmartArt 图示部件记录（data/layout/colors/quickStyle 四件套，编号共享） */
export interface SerializerDiagram {
    /** 部件编号（与 dataN/layoutN/colorsN/quickStyleN 的 N 一致） */
    index: number;
    dataXml: string;
    layoutXml: string;
    colorsXml: string;
    quickStyleXml: string;
}
/** SmartArt 图示节点（层级结构，叶子含 text） */
export interface DiagramNode {
    /** 节点文本 */
    text: string;
    /** 子节点（层级） */
    children?: DiagramNode[];
}
/**
 * __raw 回退所需的附属部件（如 SmartArt 的 diagrams/*.xml）
 * media=true 表示二进制资源（以 base64 落盘，走 Default 扩展名声明）
 */
export interface SerializerPart {
    path: string;
    content?: string;
    base64?: string;
    contentType: string;
    media?: boolean;
}
/** __raw 载荷：原始 OOXML 子树 + 其依赖的关系与部件 */
export interface SerializerRawPayload {
    /** 原始节点标签名（如 p:graphicFrame） */
    tag: string;
    /** tXml simplify 形态的节点内容 */
    node: unknown;
    /** 节点引用的关系：旧 rId → { type, target, external } */
    rels?: Record<string, { type: string; target: string; external?: boolean }>;
    /** 关系指向的部件内容（自包含，保证 round-trip 不丢件） */
    parts?: SerializerPart[];
}
/** 元素构建上下文：构建过程中收集 rels / media / charts / parts */
export interface SerializerContext {
    rels: SerializerRel[];
    media: SerializerMedia[];
    charts: SerializerChart[];
    diagrams: SerializerDiagram[];
    parts: SerializerPart[];
    nextElementId: number;
    nextRelId: number;
    mediaIndex: number;
    chartIndex: number;
    diagramIndex: number;
}
/** createElementContext 的选项 */
export interface ElementContextOptions {
    /** 媒体文件起始编号（避免与已有文件冲突） */
    startMediaIndex?: number;
    /** 图表部件起始编号（避免与已有部件冲突） */
    startChartIndex?: number;
    /** 图示部件起始编号（避免与已有部件冲突） */
    startDiagramIndex?: number;
}
/** 运行级样式（文本默认样式 / 单段 run 样式均使用） */
export interface RunStyle {
    align?: string;
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
    href?: string;
    lang?: string;
}
/** 文本运行：{ text, options } 规范格式，或扁平简写格式 */
export interface TextRunSpec extends RunStyle {
    text?: string;
    options?: RunStyle;
}
/** 段落 */
export interface ParagraphSpec {
    text?: string;
    runs?: TextRunSpec[];
    align?: string;
    bullet?: boolean | string | { char?: string; type?: 'number' | 'bullet'; fmt?: string; start?: number };
    lineSpacing?: number | { type: 'pt' | 'percent'; value: number };
    spaceBefore?: number;
    spaceAfter?: number;
    indentLeft?: number;
    indentRight?: number;
    indent?: number;
}
/** 单元格/表格边框：颜色 + 线宽(pt) */
export interface CellBorder {
    color?: string;
    width?: number;
}
/** 分边边框：每边为 CellBorder 或 'none'（显式无该边） */
export interface TableBorders {
    left?: CellBorder | 'none';
    right?: CellBorder | 'none';
    top?: CellBorder | 'none';
    bottom?: CellBorder | 'none';
    /** 对角线边框：'tlBr'(左上-右下) / 'blTr'(左下-右上) / 'both' */
    diagonal?: 'tlBr' | 'blTr' | 'both';
}

/** 形状填充/边框/特效（供 SerializerElement 使用） */
export interface ShapeFillSolid { type?: 'solid'; color?: string; transparency?: number; }
export interface ShapeFillGradient { type: 'gradient'; direction?: 'horizontal' | 'vertical' | 'diagonal'; stops: { color: string; position: number }[]; }
export interface ShapeLineSpec { color?: string; width?: number; transparency?: number; dashType?: string; }
export interface ShapeShadowSpec { type?: 'outer' | 'inner'; color?: string; blur?: number; distance?: number; angle?: number; transparency?: number; }
export interface ShapeGlowSpec { color?: string; blur?: number; }
export interface ShapeEffectsSpec { shadow?: ShapeShadowSpec | boolean; glow?: ShapeGlowSpec | boolean; }

/** 表格单元格 */
export interface SerializerTableCell {
    text?: string;
    paragraphs?: ParagraphSpec[];
    colSpan?: number;
    rowSpan?: number;
    fill?: string;
    /** 四边统一边框（优先于表格级默认） */
    border?: CellBorder;
    /** 分边边框（最高优先级，覆盖统一边框） */
    borders?: TableBorders;
    align?: string;
    valign?: string;
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
    /** 单元格内边距（px）：{ l, r, t, b } */
    inset?: { l?: number; r?: number; t?: number; b?: number };
}
/** 表格行 */
export interface SerializerTableRow {
    height?: number;
    cells: SerializerTableCell[];
}
/** 图表系列 */
export interface ChartSeriesSpec {
    name?: string;
    /** 非散点图数值 */
    values?: number[];
    /** 散点图 x / y 值 */
    x?: number[];
    y?: number[];
    /** 股票图：开盘/最高/最低/收盘值（open/high/low/close） */
    open?: number[];
    high?: number[];
    low?: number[];
    close?: number[];
    color?: string;
}
/** 幻灯片元素 JSON（text / shape / image / chart） */
export interface SerializerElement {
    type?: string;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    rotation?: number;
    /** 文本：直接文本 */
    text?: string;
    runs?: TextRunSpec[];
    paragraphs?: ParagraphSpec[];
    align?: string;
    valign?: string;
    /** 段落级默认样式（纯 text 模式透传给每个段落）：列表/行距/段间距/缩进 */
    bullet?: boolean | string | { char?: string; type?: 'number' | 'bullet'; fmt?: string; start?: number };
    lineSpacing?: number | { type: 'pt' | 'percent'; value: number };
    spaceBefore?: number;
    spaceAfter?: number;
    indentLeft?: number;
    indentRight?: number;
    indent?: number;
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
    lang?: string;
    href?: string;
    /** 形状填充：颜色串 / {type:'solid',color,transparency} / {type:'gradient',...} / 'none' / null */
    fill?: string | ShapeFillSolid | ShapeFillGradient | null;
    /** 形状边框：{ color, width, transparency, dashType } / 'none' / null */
    line?: ShapeLineSpec | 'none' | null;
    /** 形状特效（阴影 / 发光） */
    effects?: ShapeEffectsSpec | null;
    shapeType?: string;
    /** 图片：dataURL / base64 / 远程 URL */
    data?: string;
    src?: string;
    extension?: string;
    chartType?: string;
    categories?: string[];
    series?: ChartSeriesSpec[];
    varyColors?: boolean;
    barDir?: string;
    title?: string;
    legend?: boolean;
    /** 分组（type:'group'）：子元素坐标体系。'local'=相对组左上角的局部坐标（OOXML 标准，默认）；'page'=页绝对坐标（构建时减 group 偏移做相对化） */
    childrenCoordinates?: 'local' | 'page';
    /** 表格：行数据 */
    rows?: SerializerTableRow[];
    /** 表格：列宽（px，缺省均分） */
    colWidths?: number[];
    /** 表格：行高（px，缺省均分） */
    rowHeights?: number[];
    /** 表格级默认边框（四边统一） */
    border?: CellBorder;
    /** 表格级分边默认边框 */
    borders?: TableBorders;
    /** 分组：子元素列表（type:'group' 时） */
    children?: SerializerElement[];
    /** 水平翻转 */
    flipH?: boolean;
    /** 垂直翻转 */
    flipV?: boolean;
    /** 几何调整值（圆角半径 / 箭头尺寸 / 星形尖角等），如 { adj: 25000 } */
    adjust?: Record<string, number>;
    /** 文本框内边距（px）：{ l, r, t, b } */
    inset?: { l?: number; r?: number; t?: number; b?: number };
    /** 文字方向：'wordArtVertical' / 'eaVertical' / 'vert' / 'horz'（默认横排） */
    textDirection?: string;
    /** 图片裁剪（百分比 0-100）：{ l, r, t, b } */
    crop?: { l?: number; r?: number; t?: number; b?: number };
    /** 图片调整：{ brightness(-100..100), contrast(-100..100), transparency(0..100) } */
    imageAdjust?: { brightness?: number; contrast?: number; transparency?: number };
    /** 表格级单元格内边距（px）：{ l, r, t, b }（单元格未显式时继承） */
    tableInset?: { l?: number; r?: number; t?: number; b?: number };
    /** 表格级对角线边框：'tlBr' / 'blTr' / 'both' */
    tableDiagonal?: 'tlBr' | 'blTr' | 'both';
    /** 表格样式 ID（引用内置 tableStyles.xml） */
    tableStyleId?: string;
    /** SmartArt 图示类型：'list' | 'hierarchy' | 'process' | 'cycle' | 'pyramid' */
    diagramType?: string;
    /** SmartArt 图示节点（层级结构，叶子含 text） */
    nodes?: DiagramNode[];
    /** 原始 OOXML 载荷（解析端产出，语义层未覆盖时用于无损回写） */
    __raw?: unknown;
    /** 强制以 __raw 回写（即使 type 已受语义层支持） */
    rawFallback?: boolean;
}
/** 幻灯片 JSON */
export interface SerializerSlide {
    /** 背景：纯色串 / 渐变 / 图片（与 PptxBackground 同源） */
    background?: string | PptxBackground | null;
    /** 演讲者备注 */
    notes?: string;
    /** 过渡效果 */
    transition?: PptxTransition;
    /** 自动播放：停留毫秒后切换（设置后生成 p:timing 计时） */
    advanceTime?: number;
    /** 是否允许点击切换（默认 true；与 advanceTime 配合） */
    advanceOnClick?: boolean;
    /** 元素动画：[{ target?, type:'fade'|'flyIn'|'zoom'|'wipe', duration? }] */
    animations?: Array<{ target?: number; type?: string; duration?: number }>;
    /** 隐藏幻灯片（show="0"） */
    hidden?: boolean;
    /** 幻灯片批注（生成 ppt/comments/commentsN.xml） */
    comments?: SerializerComment[];
    elements?: SerializerElement[];
}

/** 幻灯片批注（commentsN.xml 的 p:cm） */
export interface SerializerComment {
    /** 作者（用于 commentAuthors；缺省 'Author'） */
    author?: string;
    /** 批注正文 */
    text: string;
    /** 批注时间（ISO 8601）；缺省取当前时间 */
    dt?: string;
    /** 批注锚点位置（EMU）；缺省 1 英寸处 */
    pos?: { x?: number; y?: number };
}

/**
 * 创建元素构建上下文
 * @param {Object} [options]
 * @param {number} [options.startMediaIndex=0] - 媒体文件起始编号（避免与已有文件冲突）
 * @returns {Object} 构建上下文
 */
export function createElementContext(options: ElementContextOptions = {}): SerializerContext {
    return {
        /** 关系列表 {relId, type, target, external} */
        rels: ([] as SerializerRel[]),
        /** 媒体文件列表 {name, base64} */
        media: ([] as SerializerMedia[]),
        /** __raw 回退附属部件列表 */
        parts: ([] as SerializerPart[]),
        /** 幻灯片内元素自增 id（1 被 spTree 根占用） */
        nextElementId: 2,
        /** 关系自增 id（rId1 固定为版式引用） */
        nextRelId: 2,
        /** 媒体文件自增编号 */
        mediaIndex: options.startMediaIndex || 0,
        /** 图表部件列表 {name, xml}（生成后由上层写入 ppt/charts/） */
        charts: ([] as SerializerChart[]),
        /** 图表自增编号 */
        chartIndex: options.startChartIndex || 0,
        /** 图示部件列表（生成后由上层写入 ppt/diagrams/） */
        diagrams: ([] as SerializerDiagram[]),
        /** 图示自增编号 */
        diagramIndex: options.startDiagramIndex || 0
    };
}

/**
 * 添加关系并返回关系 id
 * @param {Object} ctx - 构建上下文
 * @param {string} type - 关系类型（REL_TYPES 值）
 * @param {string} target - 目标路径
 * @param {boolean} [external=false] - 是否为外部链接
 * @returns {string} 关系 id（如 rId3）
 */
function addRelationship(ctx: SerializerContext, type: string, target: string, external?: boolean) {
    const relId = `rId${ctx.nextRelId++}`;
    ctx.rels.push({ relId, type, target, external: !!external });
    return relId;
}

/**
 * 解析图片来源数据
 * @param {Object} el - 图片元素
 * @returns {Promise<{base64: string, ext: string}>} 图片数据
 */
async function resolveImageData(el: SerializerElement) {
    if (el.data) {
        const str = String(el.data);
        // 兼容任意 mime 的 dataURL（如非常规扩展名产生的 data:application/octet-stream;base64,...）
        const dataUrlMatch = str.match(/^data:([^;,]*);base64,(.+)$/is);
        if (dataUrlMatch) {
            return { base64: dataUrlMatch[2], ext: el.extension || mimeToExt(dataUrlMatch[1]) };
        }
        // 裸 base64
        return { base64: str, ext: el.extension || 'png' };
    }

    if (el.src) {
        if (!/^https?:\/\//i.test(String(el.src))) {
            // 非 URL（如包内相对路径）无法下载，给出明确报错而不是交给 fetch 失败
            throw new Error(`图片元素 src 不是可下载的 URL: ${el.src}。请改用 data（dataURL/base64）或绝对 http(s) 地址`);
        }
        if (typeof fetch !== 'function') {
            throw new Error(`无法获取远程图片 ${el.src}：当前环境不支持 fetch。请将图片下载后以 data(base64/dataURL) 方式提供`);
        }
        const resp = await fetch(el.src);
        if (!resp.ok) {
            throw new Error(`下载远程图片失败: ${el.src} (HTTP ${resp.status})`);
        }
        const buffer = await resp.arrayBuffer();
        const base64 = arrayBufferToBase64(buffer);
        const mimeFromHeader = (resp.headers.get('content-type') || '').match(/^image\/([a-z0-9.+-]+)/i);
        let ext = mimeFromHeader ? mimeFromHeader[1].toLowerCase() : null;
        if (ext === 'jpg') ext = 'jpeg';
        if (ext === 'svg+xml') ext = 'svg';
        return { base64, ext: el.extension || ext || guessExtFromUrl(el.src) || 'png' };
    }

    throw new Error("图片元素需要提供 data（dataURL/base64）或 src（远程 URL）字段");
}

/**
 * mime 类型 → 文件扩展名（未知 mime 回退 png）
 * @param {string} mime - mime 类型，如 image/png
 * @returns {string} 扩展名（不含点）
 */
function mimeToExt(mime: string): string {
    const map: Record<string, string> = {
        'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg', 'image/gif': 'gif',
        'image/bmp': 'bmp', 'image/svg+xml': 'svg', 'image/tiff': 'tiff', 'image/webp': 'webp',
        'image/x-icon': 'ico', 'image/vnd.microsoft.icon': 'ico',
        'image/x-emf': 'emf', 'image/x-wmf': 'wmf'
    };
    const m = String(mime || '').toLowerCase().trim();
    if (map[m]) return map[m];
    const sub = m.split('/')[1];
    if (!sub) return 'png';
    return sub.replace(/\+.*$/, '').replace(/^x-/, '');
}

/**
 * 从 URL 猜测图片扩展名
 * @param {string} url - 图片 URL
 * @returns {string|null} 扩展名
 */
function guessExtFromUrl(url: string) {
    const match = String(url).split('?')[0].match(/\.([a-z0-9]+)$/i);
    return match ? match[1].toLowerCase() : null;
}

/**
 * ArrayBuffer 转 base64（兼容浏览器与 Node）
 * @param {ArrayBuffer} buffer - 二进制数据
 * @returns {string} base64 字符串
 */
function arrayBufferToBase64(buffer: ArrayBuffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)));
    }
    if (typeof btoa === 'function') {
        return btoa(binary);
    }
    return Buffer.from(bytes).toString('base64');
}

/**
 * 构建位置节点（a:xfrm）
 * @param {Object} el - 含 x/y/width/height/rotation 的元素
 * @returns {Object} a:xfrm 节点
 */
function buildXfrm(el: SerializerElement) {
    const xfrmAttrs: Record<string, number | string | null> = { rot: el.rotation ? degToRot(el.rotation) : null };
    if (el.flipH) xfrmAttrs.flipH = 1;
    if (el.flipV) xfrmAttrs.flipV = 1;
    return xmlNode('a:xfrm',
        xfrmAttrs,
        xmlNode('a:off', { x: pxToEmu(el.x || 0), y: pxToEmu(el.y || 0) }),
        xmlNode('a:ext', { cx: pxToEmu(el.width || 0), cy: pxToEmu(el.height || 0) })
    );
}

/**
 * 构建超链接 rPr 子节点
 * @param {Object} ctx - 构建上下文
 * @param {string} href - 链接（http(s):// 外部；'#N' 内部跳转到第 N 页）
 * @returns {Object|null} a:hlinkClick 节点
 */
function buildHyperlink(ctx: SerializerContext, href: string | undefined) {
    if (!href) return null;
    const internalMatch = String(href).match(/^#(\d+)$/);
    if (internalMatch) {
        // 内部跳转：指向对应 slide 部件（slide rels 位于 ppt/slides/_rels/，同目录相对路径）
        const relId = addRelationship(ctx, REL_TYPES.slide, `slide${internalMatch[1]}.xml`);
        return xmlNode('a:hlinkClick', { 'r:id': relId, action: 'ppaction://hlinksldjump' });
    }
    const relId = addRelationship(ctx, REL_TYPES.hyperlink, String(href), true);
    return xmlNode('a:hlinkClick', { 'r:id': relId });
}

/**
 * 构建文本运行（a:r）
 * @param {Object} ctx - 构建上下文
 * @param {string} text - 运行文本
 * @param {Object} opts - 运行选项（fontSize/color/bold/italic/underline/fontFace/href）
 * @returns {Object} a:r 节点
 */
function buildTextRun(ctx: SerializerContext, text: string | undefined, opts: RunStyle = {}) {
    const rPrChildren = [];

    if (opts.color) {
        rPrChildren.push(xmlNode('a:solidFill', colorNode(opts.color)));
    }
    if (opts.fontFace) {
        rPrChildren.push(xmlNode('a:latin', { typeface: opts.fontFace }));
    }
    const hlink = buildHyperlink(ctx, opts.href);
    if (hlink) rPrChildren.push(hlink);

    return xmlNode('a:r',
        null,
        xmlNode('a:rPr',
            {
                lang: opts.lang || 'zh-CN',
                sz: opts.fontSize !== undefined ? ptToSz(opts.fontSize) : null,
                b: opts.bold ? 1 : null,
                i: opts.italic ? 1 : null,
                u: opts.underline ? 'sng' : null,
                dirty: 0
            },
            ...rPrChildren
        ),
        xmlNode('a:t', null, String(text))
    );
}

/**
 * 构建段落（a:p）
 * @param {Object} ctx - 构建上下文
 * @param {Object} paragraph - 段落 { text, runs, align, bullet }
 * @param {Object} defaults - 元素级默认运行选项
 * @returns {Object} a:p 节点
 */
function buildParagraph(ctx: SerializerContext, paragraph: ParagraphSpec, defaults: RunStyle) {
    const alignMap: Record<string, string | null> = { left: 'l', center: 'ctr', right: 'r', justify: 'just' };
    const p = paragraph || {};

    const pPrAttrs: Record<string, number | string | null> = { algn: alignMap[p.align || defaults.align || 'left'] || null };
    if (p.indentLeft != null) pPrAttrs.marL = ptToEmu(p.indentLeft);
    if (p.indentRight != null) pPrAttrs.marR = ptToEmu(p.indentRight);
    if (p.indent != null) pPrAttrs.indent = ptToEmu(p.indent);

    const pPrChildren: BuilderNode[] = [];

    // 行距 / 段间距
    if (p.lineSpacing != null) {
        const ls = p.lineSpacing;
        if (typeof ls === 'number') {
            pPrChildren.push(xmlNode('a:lnSpc', null, xmlNode('a:spcPct', { val: Math.round(ls * 1000) })));
        } else if (ls.type === 'pt') {
            pPrChildren.push(xmlNode('a:lnSpc', null, xmlNode('a:spcPts', { val: ptToSz(ls.value) })));
        } else {
            pPrChildren.push(xmlNode('a:lnSpc', null, xmlNode('a:spcPct', { val: Math.round(ls.value * 1000) })));
        }
    }
    if (p.spaceBefore != null) pPrChildren.push(xmlNode('a:spcBef', null, xmlNode('a:spcPts', { val: ptToSz(p.spaceBefore) })));
    if (p.spaceAfter != null) pPrChildren.push(xmlNode('a:spcAft', null, xmlNode('a:spcPts', { val: ptToSz(p.spaceAfter) })));

    // 列表符号：自动编号 / 项目符号 / 无
    const b = p.bullet;
    if (b === 'number' || (b && typeof b === 'object' && b.type === 'number')) {
        const fmt = (b && typeof b === 'object' && b.fmt) || 'arabic';
        const start = (b && typeof b === 'object' && b.start != null) ? b.start : 1;
        pPrChildren.push(xmlNode('a:buAutoNum', { type: fmt, startAt: start }));
    } else if (b === true || b === 'bullet' || (b && typeof b === 'object' && (b.type === 'bullet' || b.char))) {
        const char = (b && typeof b === 'object' && b.char) ? b.char : '•';
        pPrChildren.push(xmlNode('a:buFont', { typeface: 'Arial' }));
        pPrChildren.push(xmlNode('a:buChar', { char }));
    } else {
        pPrChildren.push(xmlNode('a:buNone'));
    }

    const pPr = xmlNode('a:pPr', pPrAttrs, ...pPrChildren);

    // 运行列表：显式 runs 优先，否则用 text + 元素级默认样式
    let runs;
    if (Array.isArray(p.runs) && p.runs.length > 0) {
        runs = p.runs.map((r: TextRunSpec) => {
            // 兼容两种格式：{ text, options: {...} }（规范）与 { text, color, ... }（扁平简写）
            const { text, options, ...flat } = r;
            return buildTextRun(ctx, text, { ...defaults, ...(options || {}), ...flat });
        });
    } else {
        runs = [buildTextRun(ctx, p.text !== undefined ? p.text : '', defaults)];
    }

    return xmlNode('a:p', null, pPr, ...runs);
}

/**
 * 规范化段落列表
 * @param {Object} el - 文本元素
 * @returns {Array<Object>} 段落列表
 */
function normalizeParagraphs(el: SerializerElement): ParagraphSpec[] {
    let paras: ParagraphSpec[];
    if (Array.isArray(el.paragraphs) && el.paragraphs.length > 0) {
        paras = el.paragraphs;
    } else if (Array.isArray(el.runs) && el.runs.length > 0) {
        paras = [{ runs: el.runs }];
    } else if (el.text !== undefined) {
        paras = String(el.text).split('\n').map(t => ({ text: t }));
    } else {
        paras = [{ text: '' }];
    }
    // 元素级段落默认样式（纯 text 模式）：段落自身未显式设置时继承
    const pDefaults: ParagraphSpec = {
        bullet: el.bullet,
        lineSpacing: el.lineSpacing,
        spaceBefore: el.spaceBefore,
        spaceAfter: el.spaceAfter,
        indentLeft: el.indentLeft,
        indentRight: el.indentRight,
        indent: el.indent
    };
    return paras.map(p => ({ ...pDefaults, ...p }));
}

/**
 * 构建文本框元素（p:sp）
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 文本元素 JSON
 * @returns {Object} p:sp 节点
 */
function buildTextElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;
    const anchorMap: Record<string, string | null> = { top: null, middle: 'ctr', bottom: 'b' };
    const bodyPrAttrs: Record<string, number | string | null> = { wrap: 'square', rtlCol: 0 };
    const anchor = anchorMap[el.valign ?? 'top'];
    if (anchor) bodyPrAttrs.anchor = anchor;
    if (el.textDirection) bodyPrAttrs.vert = el.textDirection;
    if (el.inset) {
        if (el.inset.l != null) bodyPrAttrs.lIns = pxToEmu(el.inset.l);
        if (el.inset.r != null) bodyPrAttrs.rIns = pxToEmu(el.inset.r);
        if (el.inset.t != null) bodyPrAttrs.tIns = pxToEmu(el.inset.t);
        if (el.inset.b != null) bodyPrAttrs.bIns = pxToEmu(el.inset.b);
    }
    const defaults = {
        align: el.align,
        fontSize: el.fontSize,
        color: el.color,
        bold: el.bold,
        italic: el.italic,
        underline: el.underline,
        fontFace: el.fontFace,
        href: el.href,
        lang: el.lang
    };

    return xmlNode('p:sp',
        null,
        xmlNode('p:nvSpPr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `TextBox ${id - 1}` }),
            xmlNode('p:cNvSpPr', { txBox: 1 }),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:spPr',
            null,
            buildXfrm(el),
            xmlNode('a:prstGeom', { prst: 'rect' }, xmlNode('a:avLst'))
        ),
        xmlNode('p:txBody',
            null,
            xmlNode('a:bodyPr',
                bodyPrAttrs
            ),
            xmlNode('a:lstStyle'),
            ...normalizeParagraphs(el).map((p: ParagraphSpec) => buildParagraph(ctx, p, defaults))
        )
    );
}

/**
 * 构建形状元素（p:sp）
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 形状元素 JSON
 * @returns {Object} p:sp 节点
 */
/** 构造颜色节点：支持 scheme:<name>（a:schemeClr）与绝对色（a:srgbClr） */
function colorNode(color: string): BuilderNode {
    if (typeof color === 'string' && color.startsWith('scheme:')) {
        return xmlNode('a:schemeClr', { val: color.slice(7) });
    }
    return xmlNode('a:srgbClr', { val: colorToHex(color) });
}

/** 构造填充节点（solidFill / gradFill / blipFill / pattFill / noFill）；transparency 写入 a:alpha */
async function buildFillNode(ctx: SerializerContext, fill: SerializerElement['fill']): Promise<BuilderNode | null> {
    if (fill === undefined) return null; // 未指定填充则继承主题
    if (fill === 'none' || fill === null) return xmlNode('a:noFill');
    if (typeof fill === 'string') return xmlNode('a:solidFill', colorNode(fill));
    if ((fill as ShapeFillGradient).type === 'gradient') {
        const g = fill as ShapeFillGradient;
        const stops = (g.stops || []).map((s) =>
            xmlNode('a:gs', { pos: Math.round((s.position || 0) * 100000) }, colorNode(s.color))
        );
        const ang = (g.direction === 'vertical' ? 90 : g.direction === 'diagonal' ? 45 : 0) * 60000;
        return xmlNode('a:gradFill', null, xmlNode('a:gsLst', null, ...stops), xmlNode('a:lin', { ang, scaled: 1 }));
    }
    if ((fill as any).type === 'image') {
        const img = fill as { type: 'image'; data?: string; src?: string; extension?: string };
        const { base64, ext } = await resolveImageData({ type: 'image', data: img.data, src: img.src, extension: img.extension } as SerializerElement);
        ctx.mediaIndex++;
        const mediaName = `image${ctx.mediaIndex}.${ext}`;
        ctx.media.push({ name: mediaName, base64 });
        const embedRelId = addRelationship(ctx, REL_TYPES.image, `../media/${mediaName}`);
        return xmlNode('a:blipFill', null, xmlNode('a:blip', { 'r:embed': embedRelId }), xmlNode('a:stretch', null, xmlNode('a:fillRect')));
    }
    if ((fill as any).type === 'pattern') {
        const p = fill as { type: 'pattern'; prst: string; fg?: string; bg?: string };
        const fg = p.fg ? colorNode(p.fg) : xmlNode('a:schemeClr', { val: 'tx1' });
        const bg = p.bg ? colorNode(p.bg) : xmlNode('a:schemeClr', { val: 'bg1' });
        return xmlNode('a:pattFill', { prst: p.prst }, xmlNode('a:fgClr', null, fg), xmlNode('a:bgClr', null, bg));
    }
    // solid（{color} 或 {type:'solid'}）
    const color = (fill as ShapeFillSolid).color;
    if (!color) return null;
    const alpha = (fill as ShapeFillSolid).transparency;
    // 注意：a:alpha 必须嵌套在颜色节点（srgbClr/schemeClr）内部，不能放在 a:solidFill 下，否则不符合 OOXML schema
    const srgb = colorNode(color);
    if (alpha != null) (srgb.children as BuilderNode[]).push(xmlNode('a:alpha', { val: Math.round((100 - alpha) * 1000) }));
    return xmlNode('a:solidFill', srgb);
}

async function buildShapeElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;

    const fillNode = await buildFillNode(ctx, el.fill);

    // 边框
    let lineNode;
    if (el.line === 'none' || el.line === null) {
        lineNode = xmlNode('a:ln', null, xmlNode('a:noFill'));
    } else if (el.line) {
        const w = el.line.width !== undefined ? el.line.width : 1;
        const srgb = colorNode(el.line.color);
        if (el.line.transparency != null) (srgb.children as BuilderNode[]).push(xmlNode('a:alpha', { val: Math.round((100 - el.line.transparency) * 1000) }));
        const lnChildren: BuilderNode[] = [xmlNode('a:solidFill', srgb)];
        if (el.line.dashType && el.line.dashType !== 'solid') lnChildren.push(xmlNode('a:prstDash', { val: el.line.dashType }));
        lineNode = xmlNode('a:ln', { w: ptToEmu(w) }, ...lnChildren);
    }

    // 特效（a:effectLst：阴影 / 发光）
    let effectNode: BuilderNode | null = null;
    if (el.effects) {
        const effChildren: BuilderNode[] = [];
        const sh = el.effects.shadow;
        if (sh && sh !== true) {
            const shType = sh.type || 'outer';
            const blur = sh.blur !== undefined ? sh.blur : 4;
            const dist = sh.distance !== undefined ? sh.distance : 3;
            const ang = sh.angle !== undefined ? sh.angle : 90;
            const color = sh.color || '#000000';
            const alpha = sh.transparency != null ? Math.round((100 - sh.transparency) * 1000) : 60000;
            effChildren.push(xmlNode(shType === 'inner' ? 'a:innerShdw' : 'a:outerShdw',
                { blur: ptToEmu(blur), dist: ptToEmu(dist), dir: Math.round(ang * 60000) },
                xmlNode('a:srgbClr', { val: colorToHex(color) }, xmlNode('a:alpha', { val: alpha }))
            ));
        } else if (sh === true) {
            effChildren.push(xmlNode('a:outerShdw', { blur: ptToEmu(4), dist: ptToEmu(3), dir: 5400000 },
                xmlNode('a:srgbClr', { val: '000000' }, xmlNode('a:alpha', { val: 60000 }))
            ));
        }
        const gw = el.effects.glow;
        if (gw && gw !== true) {
            effChildren.push(xmlNode('a:glow', { blur: ptToEmu(gw.blur !== undefined ? gw.blur : 5) },
                xmlNode('a:srgbClr', { val: colorToHex(gw.color || '#FFFF00') })));
        } else if (gw === true) {
            effChildren.push(xmlNode('a:glow', { blur: ptToEmu(5) },
                xmlNode('a:srgbClr', { val: colorToHex('#FFFF00') })));
        }
        if (effChildren.length) effectNode = xmlNode('a:effectLst', null, ...effChildren);
    }

    return xmlNode('p:sp',
        null,
        xmlNode('p:nvSpPr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `Shape ${id - 1}` }),
            xmlNode('p:cNvSpPr'),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:spPr',
            null,
            buildXfrm(el),
            xmlNode('a:prstGeom', { prst: el.shapeType || 'rect' },
                el.adjust && Object.keys(el.adjust).length
                    ? xmlNode('a:avLst', null, ...Object.entries(el.adjust).map(([name, val]) => xmlNode('a:gd', { name, fmla: `val ${val}` })))
                    : xmlNode('a:avLst')),
            fillNode,
            lineNode,
            ...(effectNode ? [effectNode] : [])
        )
    );
}

/**
 * 构建图片元素（p:pic）
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 图片元素 JSON
 * @returns {Promise<Object>} p:pic 节点
 */
async function buildImageElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;
    const { base64, ext } = await resolveImageData(el);

    ctx.mediaIndex++;
    const mediaName = `image${ctx.mediaIndex}.${ext}`;
    ctx.media.push({ name: mediaName, base64 });

    const embedRelId = addRelationship(ctx, REL_TYPES.image, `../media/${mediaName}`);

    // 图片级超链接挂在 cNvPr 上
    let cNvPrChildren = null;
    if (el.href && !/^#\d+$/.test(String(el.href))) {
        const hlinkRelId = addRelationship(ctx, REL_TYPES.hyperlink, String(el.href), true);
        cNvPrChildren = xmlNode('a:hlinkClick', { 'r:id': hlinkRelId });
    }

    // 裁剪 / 调整 → blip 子节点（a:srcRect / a:lum / a:contrast / a:alphaModFix）
    const blipChildren: BuilderNode[] = [];
    if (el.crop) {
        const c = el.crop;
        const srect: Record<string, number> = {};
        if (c.l != null) srect.l = Math.round(c.l * 1000);
        if (c.r != null) srect.r = Math.round(c.r * 1000);
        if (c.t != null) srect.t = Math.round(c.t * 1000);
        if (c.b != null) srect.b = Math.round(c.b * 1000);
        blipChildren.push(xmlNode('a:srcRect', srect));
    }
    if (el.imageAdjust) {
        const adj = el.imageAdjust;
        if (adj.brightness != null) blipChildren.push(xmlNode('a:lum', { val: Math.round((100 + adj.brightness) * 1000) }));
        if (adj.contrast != null) blipChildren.push(xmlNode('a:contrast', { val: Math.round((100 + adj.contrast) * 1000) }));
        if (adj.transparency != null) blipChildren.push(xmlNode('a:alphaModFix', { val: Math.round((100 - adj.transparency) * 1000) }));
    }

    return xmlNode('p:pic',
        null,
        xmlNode('p:nvPicPr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `Image ${id - 1}` }, cNvPrChildren),
            xmlNode('p:cNvPicPr', null, xmlNode('a:picLocks', { noChangeAspect: 1 })),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:blipFill',
            null,
            xmlNode('a:blip', { 'r:embed': embedRelId }, ...blipChildren),
            xmlNode('a:stretch', null, xmlNode('a:fillRect'))
        ),
        xmlNode('p:spPr',
            null,
            buildXfrm(el),
            xmlNode('a:prstGeom', { prst: 'rect' }, xmlNode('a:avLst'))
        )
    );
}

/**
 * 构建图表元素（p:graphicFrame + 原生 c:chartSpace 部件）
 *
 * 输入 el 格式：
 * { type:'chart', x,y,width,height, name,
 *   chartType: 'barChart'|'lineChart'|'areaChart'|'pieChart'|'pie3DChart'|'scatterChart',
 *   title, legend(true), varyColors(true),
 *   categories: ['A','B','C'],
 *   series: [ { name, values:[..], color? }, ... ]            // 非散点
 *   series: [ { name, x:[..], y:[..] } ]                      // 散点
 * }
 *
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 图表元素 JSON
 * @returns {Object} p:graphicFrame 节点
 */
function buildChartElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;
    ctx.chartIndex++;
    const chartNum = ctx.chartIndex;
    const chartName = `chart${chartNum}.xml`;
    const relId = addRelationship(ctx, REL_TYPES.chart, `../charts/${chartName}`);
    const xml = buildChartXml(el);
    ctx.charts.push({ name: chartName, xml });

    return xmlNode('p:graphicFrame',
        null,
        xmlNode('p:nvGraphicFramePr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `Chart ${chartNum}` }),
            xmlNode('p:cNvGraphicFramePr'),
            xmlNode('p:nvPr')
        ),
        buildXfrm(el),
        xmlNode('a:graphic',
            null,
            xmlNode('a:graphicData',
                { uri: NS.c },
                xmlNode('c:chart', { 'xmlns:c': NS.c, 'xmlns:r': NS.r, 'r:id': relId })
            )
        )
    );
}

/** 构造 c:strRef（类别标签缓存） */
function strRefXml(values: unknown[], col: string) {
    const n = values.length;
    let pts = '';
    for (let i = 0; i < n; i++) {
        pts += `<c:pt idx="${i}"><c:v>${escapeXml(String(values[i]))}</c:v></c:pt>`;
    }
    const last = n > 0 ? n - 1 : 0;
    return `<c:strRef><c:f>Sheet1!$${col}$2:$${col}$${2 + last}</c:f>` +
        `<c:strCache><c:ptCount val="${n}"/>${pts}</c:strCache></c:strRef>`;
}

/** 构造 c:numRef（数值缓存） */
function numRefXml(values: unknown[], col: string) {
    const n = values.length;
    let pts = '';
    for (let i = 0; i < n; i++) {
        pts += `<c:pt idx="${i}"><c:v>${Number(values[i])}</c:v></c:pt>`;
    }
    const last = n > 0 ? n - 1 : 0;
    return `<c:numRef><c:f>Sheet1!$${col}$2:$${col}$${2 + last}</c:f>` +
        `<c:numCache><c:fmtCode>General</c:fmtCode><c:ptCount val="${n}"/>${pts}</c:numCache></c:numRef>`;
}

/**
 * 生成 c:chartSpace 原生图表 XML（自包含，内联数据缓存，无需外部工作簿）
 * @param {Object} el - 图表元素 JSON
 * @returns {string} chart 部件 XML
 */
function buildChartXml(el: SerializerElement) {
    const type = el.chartType || 'barChart';
    const isPie = /pie/i.test(type);
    const isScatter = type === 'scatterChart';
    const isStock = type === 'stockChart';
    const isRadar = type === 'radarChart';
    const isSurface = type === 'surfaceChart';
    const is3D = /3D$/i.test(type);
    const isDoughnut = type === 'doughnutChart';
    const isBubble = type === 'bubbleChart';
    const cats = el.categories || [];
    const series = el.series || [];
    const varyColors = el.varyColors !== undefined ? (el.varyColors ? 1 : 0) : (isPie ? 1 : 0);

    const serXml = series.map((s: ChartSeriesSpec, i: number) => {
        const tx = `<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f>` +
            `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(s.name || `Series${i + 1}`)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
        let data;
        if (isScatter) {
            data = `<c:xVal>${numRefXml(s.x || [], 'B')}</c:xVal><c:yVal>${numRefXml(s.y || [], 'C')}</c:yVal>`;
        } else if (isStock) {
            data = `<c:openVal>${numRefXml(s.open || [], 'B')}</c:openVal>` +
                `<c:highVal>${numRefXml(s.high || [], 'C')}</c:highVal>` +
                `<c:lowVal>${numRefXml(s.low || [], 'D')}</c:lowVal>` +
                `<c:closeVal>${numRefXml(s.close || s.values || [], 'E')}</c:closeVal>`;
        } else {
            data = `<c:cat>${strRefXml(cats, 'A')}</c:cat><c:val>${numRefXml(s.values || [], 'B')}</c:val>`;
        }
        const spPr = s.color ? `<c:spPr><a:solidFill><a:srgbClr val="${colorToHex(s.color)}"/></a:solidFill></c:spPr>` : '';
        return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${spPr}${data}</c:ser>`;
    }).join('');

    // 图表类型特定根（含轴 id，散点/柱状/折线/面积需要）
    let plotChart;
    if (isPie || isDoughnut) {
        plotChart = `<c:${type}><c:varyColors val="${varyColors}"/>${serXml}</c:${type}>`;
    } else if (isBubble) {
        plotChart = `<c:${type}>${serXml}</c:${type}>`;
    } else if (isRadar) {
        plotChart = `<c:${type}><c:radarStyle val="standard"/>${serXml}</c:${type}>`;
    } else if (isStock) {
        plotChart = `<c:${type}><c:hiLowLines/><c:serLines/>${serXml}</c:${type}>`;
    } else if (isSurface) {
        plotChart = `<c:${type}><c:bandFmts/>${serXml}</c:${type}>`;
    } else {
        const dir = (type === 'barChart' || type === 'bar3DChart') ? `<c:barDir val="${el.barDir || 'col'}"/>` : '';
        const grouping = (type === 'lineChart' || type === 'areaChart') ? '<c:grouping val="standard"/>' : '';
        plotChart = `<c:${type}>${dir}${grouping}<c:varyColors val="${varyColors}"/>${serXml}` +
            `<c:axId val="111"/><c:axId val="112"/></c:${type}>`;
    }
    const view3D = is3D ? '<c:view3D><c:rotX val="30"/><c:rotY val="0"/></c:view3D>' : '';

    // 坐标轴（饼图除外）
    let axes = '';
    if (!isPie) {
        axes = `<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="112"/></c:catAx><c:valAx><c:axId val="112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/><c:majorGridlines/></c:valAx>`;
    }

    const titleXml = el.title
        ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr lang="zh-CN"/><a:t>${escapeXml(el.title)}</a:t></a:r></a:p>` +
          `</c:rich></c:tx><c:overlay val="0"/></c:title>`
        : '';

    const legendXml = el.legend !== false && el.legend !== undefined
        ? `<c:legend><c:legendPos val="${typeof el.legend === 'string' ? el.legend : 'r'}"/><c:overlay val="0"/></c:legend>`
        : '';
    const dLblsXml = el.dataLabels
        ? `<c:dLbls><c:showVal val="${(el.dataLabels === true || el.dataLabels.showValue) ? 1 : 0}"/>` +
          `<c:showPercent val="${el.dataLabels.showPercent ? 1 : 0}"/><c:showSer val="${el.dataLabels.showSeries ? 1 : 0}"/>` +
          `<c:showCatName val="${el.dataLabels.showCategory ? 1 : 0}"/></c:dLbls>`
        : '';

    const autoTitleDeleted = `<c:autoTitleDeleted val="${el.title ? 0 : 1}"/>`;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
        `<c:chart>${titleXml}${autoTitleDeleted}` +
        `<c:plotArea><c:layout/>${plotChart}${view3D}${axes}</c:plotArea>` +
        `${dLblsXml}${legendXml}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

/** 表格级边框默认（type:'table' 元素的 border/borders 透传） */
type TableBorderSpec = { border?: CellBorder; borders?: TableBorders };

/** 读取边框对象（'none' 或空 → undefined 表示不生成该边） */
function normalizeCellBorder(b: CellBorder | 'none' | undefined): CellBorder | undefined {
    if (!b || b === 'none') return undefined;
    return { color: b.color, width: b.width };
}

/** 解析某条边最终采用的边框：单元格分边 > 单元格统一 > 表格分边 > 表格统一 */
function resolveBorderSide(side: 'L' | 'R' | 'T' | 'B', cell: SerializerTableCell, tableBorder?: TableBorderSpec): CellBorder | undefined {
    const key = side === 'L' ? 'left' : side === 'R' ? 'right' : side === 'T' ? 'top' : 'bottom';
    const cSides = cell.borders || {};
    const tSides = (tableBorder && tableBorder.borders) || {};
    if (cSides[key] !== undefined) return normalizeCellBorder(cSides[key]);
    if (cell.border) return { color: cell.border.color, width: cell.border.width };
    if (tSides[key] !== undefined) return normalizeCellBorder(tSides[key]);
    if (tableBorder && tableBorder.border) return { color: tableBorder.border.color, width: tableBorder.border.width };
    return undefined;
}

/** 构造一条单元格边框线 a:ln{side}（含 solidFill 颜色） */
function edgeLineXml(side: 'L' | 'R' | 'T' | 'B', b: CellBorder): BuilderNode {
    const w = b.width !== undefined ? b.width : 1;
    const color = b.color || '#000000';
    return xmlNode(`a:ln${side}`, { w: ptToEmu(w) },
        xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(color) })));
}

/**
 * 构建表格单元格（a:tc）
 * @param {Object} ctx - 构建上下文
 * @param {Object} cell - 单元格 { text | paragraphs, colSpan, rowSpan, fill, border, borders, align, valign, ...样式 }
 * @param {Object} [tableBorder] - 表格级边框默认（type:'table' 的 border/borders）
 * @returns {Object} a:tc 节点
 */
function buildTableCell(ctx: SerializerContext, cell: SerializerTableCell, tableBorder?: TableBorderSpec): BuilderNode {
    const attrs: Record<string, unknown> = {};
    if (cell.colSpan && cell.colSpan > 1) attrs.gridSpan = cell.colSpan;
    if (cell.rowSpan && cell.rowSpan > 1) attrs.rowSpan = cell.rowSpan;

    const defaults: RunStyle = {
        align: cell.align,
        fontSize: cell.fontSize,
        color: cell.color,
        bold: cell.bold,
        italic: cell.italic,
        underline: cell.underline,
        fontFace: cell.fontFace
    };
    const paragraphs = Array.isArray(cell.paragraphs) && cell.paragraphs.length
        ? cell.paragraphs
        : [{ text: cell.text !== undefined ? cell.text : '' }];

    const tcPrChildren: BuilderNode[] = [];
    // 边框（OOXML 顺序位在填充之前）
    for (const side of ['L', 'R', 'T', 'B'] as const) {
        const b = resolveBorderSide(side, cell, tableBorder);
        if (b) tcPrChildren.push(edgeLineXml(side, b));
    }
    if (cell.fill) {
        tcPrChildren.push(xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(cell.fill) })));
    }
    // 单元格内边距
    if (cell.inset) {
        const ins = cell.inset;
        const insAttrs: Record<string, number> = {};
        if (ins.l != null) insAttrs.l = pxToEmu(ins.l);
        if (ins.r != null) insAttrs.r = pxToEmu(ins.r);
        if (ins.t != null) insAttrs.t = pxToEmu(ins.t);
        if (ins.b != null) insAttrs.b = pxToEmu(ins.b);
        tcPrChildren.push(xmlNode('a:tableCellInsets', insAttrs));
    }
    // 对角线边框
    const diag = cell.borders && cell.borders.diagonal;
    if (diag) {
        const dc = (cell.border && cell.border.color) || '#000000';
        const dw = (cell.border && cell.border.width != null) ? cell.border.width : 1;
        const dLine = xmlNode('a:ln', { w: ptToEmu(dw) }, xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(dc) })));
        if (diag === 'tlBr' || diag === 'both') tcPrChildren.push(xmlNode('a:lnTlToBr', null, dLine));
        if (diag === 'blTr' || diag === 'both') tcPrChildren.push(xmlNode('a:lnBlToTr', null, dLine));
    }
    const anchorMap: Record<string, string | null> = { top: 't', middle: 'ctr', bottom: 'b' };

    return xmlNode('a:tc', attrs,
        xmlNode('a:txBody', null,
            xmlNode('a:bodyPr', { wrap: 'square', rtlCol: 0 }),
            xmlNode('a:lstStyle'),
            ...paragraphs.map((p) => buildParagraph(ctx, p, defaults))
        ),
        xmlNode('a:tcPr', { anchor: anchorMap[cell.valign ?? 'top'] ?? null }, ...tcPrChildren)
    );
}

/**
 * 构建表格元素（p:graphicFrame + a:graphic/a:graphicData/a:tbl）
 *
 * 输入 el 格式：
 * { type:'table', x,y,width,height, colWidths?: number[], rowHeights?: number[],
 *   rows: [ { height?, cells: [ { text | paragraphs, colSpan, rowSpan, fill, align, valign, ... } ] } ] }
 * 缺省 colWidths/rowHeights 时按整体尺寸均分。
 *
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 表格元素 JSON
 * @returns {Object} p:graphicFrame 节点
 */
function buildTableElement(ctx: SerializerContext, el: SerializerElement): BuilderNode {
    const id = ctx.nextElementId++;
    const rows = el.rows || [];
    const width = el.width || 0;
    const height = el.height || 0;

    // 列宽：显式优先，否则均分
    const colCount = el.colWidths && el.colWidths.length
        ? el.colWidths.length
        : Math.max(0, ...rows.map((r) => (r.cells || []).length));
    const colWidths = el.colWidths && el.colWidths.length
        ? el.colWidths
        : new Array(colCount).fill(colCount ? width / colCount : 0);

    // 行高：行级 height > rowHeights > 均分
    const rowHeights = rows.map((r, i) =>
        r.height !== undefined
            ? r.height
            : (el.rowHeights && el.rowHeights[i] !== undefined
                ? el.rowHeights[i]
                : (rows.length ? height / rows.length : 0)));

    const gridCols = colWidths.map((w) => xmlNode('a:gridCol', { w: pxToEmu(w) }));
    const trNodes = rows.map((row, ri) =>
        xmlNode('a:tr', { h: pxToEmu(rowHeights[ri]) },
            ...(row.cells || []).map((cell) => {
                const merged: SerializerTableCell = { ...cell };
                if (merged.inset == null && el.tableInset != null) merged.inset = el.tableInset;
                const cdiag = merged.borders && merged.borders.diagonal;
                if (el.tableDiagonal && !cdiag) merged.borders = { ...(merged.borders || {}), diagonal: el.tableDiagonal } as TableBorders;
                return buildTableCell(ctx, merged, { border: el.border, borders: el.borders });
            })
        )
    );

    return xmlNode('p:graphicFrame', null,
        xmlNode('p:nvGraphicFramePr', null,
            xmlNode('p:cNvPr', { id, name: el.name || `Table ${id - 1}` }),
            xmlNode('p:cNvGraphicFramePr', null, xmlNode('a:graphicFrameLocks', { noGrp: 1 })),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:xfrm', null,
            xmlNode('a:off', { x: pxToEmu(el.x || 0), y: pxToEmu(el.y || 0) }),
            xmlNode('a:ext', { cx: pxToEmu(width), cy: pxToEmu(height) })
        ),
        xmlNode('a:graphic', null,
            xmlNode('a:graphicData', { uri: NS.table },
                xmlNode('a:tbl', null,
                    xmlNode('a:tblPr', { firstRow: 1, bandRow: 1, tableStyleId: el.tableStyleId || '2D1D2E6E-4B9A-4B3C-9B6E-7B5C8D9E0F1A' }),
                    xmlNode('a:tblGrid', null, ...gridCols),
                    ...trNodes
                )
            )
        )
    );
}

/** 规整为数组（tXml 单节点为对象、多节点为数组） */
function toArray(v: unknown): unknown[] {
    if (v === undefined || v === null) return [];
    return Array.isArray(v) ? v : [v];
}

/**
 * 将 tXml simplify 形态的原始节点还原为构建器节点（含 rId 重映射）
 * @param {string} tag - 标签名
 * @param {Object|string} node - 原始节点内容
 * @param {Object} remap - 旧 rId → 新 rId 映射
 * @returns {Object} BuilderNode
 */
function rawNodeToBuilder(tag: string, node: unknown, remap: Record<string, string>): BuilderNode {
    if (node === null || node === undefined) return xmlNode(tag, null);
    if (typeof node !== 'object') return xmlNode(tag, null, String(node));

    const src = node as Record<string, unknown>;
    const attrs: Record<string, unknown> = {};
    const srcAttrs = (src.attrs || {}) as Record<string, unknown>;
    for (const k of Object.keys(srcAttrs)) {
        const raw = String(srcAttrs[k]);
        attrs[k] = remap[raw] || raw;
    }

    const children: (BuilderNode | string)[] = [];
    for (const k of Object.keys(src)) {
        if (k === 'attrs') continue;
        for (const child of toArray(src[k])) {
            if (child === null || child === undefined) continue;
            if (typeof child === 'object') children.push(rawNodeToBuilder(k, child, remap));
            else children.push(xmlNode(k, null, String(child)));
        }
    }
    return xmlNode(tag, attrs, ...children);
}

/**
 * 用 __raw 原样回写元素（语义层未覆盖时的无损兜底）
 *
 * 载荷形如 { tag, node, rels, parts }（解析端产出）：
 * - rels 中的旧 rId 会在当前 ctx 中重新登记，节点内的 r:* 属性同步重写为新 rId；
 * - parts 登记到 ctx.parts，由上层（jsonToPptx / addSlide）写入 zip 并声明 Content-Types。
 *
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 元素 JSON（需含 __raw）
 * @returns {Object|null} 构建器节点；__raw 缺失或格式不符时返回 null
 */
export function buildRawElement(ctx: SerializerContext, el: SerializerElement): BuilderNode | null {
    const payload = el.__raw as SerializerRawPayload | undefined;
    if (!payload || typeof payload !== 'object' || !payload.tag) return null;

    // 关系重映射：旧 rId → 新 rId
    // 载荷中的 target 为 zip 内绝对路径（ppt/...），写入 slide rels 需转换为 ../ 相对路径
    const remap: Record<string, string> = {};
    for (const [oldRid, rel] of Object.entries(payload.rels || {})) {
        if (!rel || !rel.target) continue;
        const target = rel.target.startsWith('ppt/') ? `../${rel.target.slice(4)}` : rel.target;
        remap[oldRid] = addRelationship(ctx, rel.type, target, rel.external);
    }

    // 附属部件登记
    for (const part of payload.parts || []) {
        if (part && part.path) ctx.parts.push(part);
    }

    return rawNodeToBuilder(payload.tag, payload.node, remap);
}

/**
 * 构建单个幻灯片元素节点
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 元素 JSON
 * @returns {Promise<Object|null>} 元素节点；语义层不支持且无 __raw 时返回 null
 */
/** 构建分组元素（p:grpSp），递归构建子元素 */
async function buildGroupElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;
    const children = el.children || [];
    const childNodes: BuilderNode[] = [];
    // 'page' 模式：子元素坐标是页绝对坐标，构建时减 group 偏移转为组内局部坐标（OOXML 标准）
    const pageCoords = el.childrenCoordinates === 'page';
    const gx = el.x || 0;
    const gy = el.y || 0;
    for (const c of children) {
        const cc = pageCoords ? { ...c, x: (c.x || 0) - gx, y: (c.y || 0) - gy } : c;
        const n = await buildElement(ctx, cc);
        if (n) childNodes.push(n);
    }
    const w = pxToEmu(el.width || 0);
    const h = pxToEmu(el.height || 0);
    return xmlNode('p:grpSp',
        null,
        xmlNode('p:nvGrpSpPr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `Group ${id - 1}` }),
            xmlNode('p:cNvGrpSpPr'),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:grpSpPr',
            null,
            xmlNode('a:xfrm',
                null,
                xmlNode('a:off', { x: pxToEmu(el.x || 0), y: pxToEmu(el.y || 0) }),
                xmlNode('a:ext', { cx: w, cy: h }),
                xmlNode('a:chOff', { x: 0, y: 0 }),
                xmlNode('a:chExt', { cx: w, cy: h })
            )
        ),
        ...childNodes
    );
}

/** 构建视频/音频媒体元素（p:pic + 媒体关系） */
async function buildMediaElement(ctx: SerializerContext, el: SerializerElement, kind: 'video' | 'audio') {
    const id = ctx.nextElementId++;
    ctx.mediaIndex++;
    const ext = el.extension || (kind === 'video' ? 'mp4' : 'm4a');
    const mediaName = `${kind}${ctx.mediaIndex}.${ext}`;
    const { base64 } = await resolveImageData({ type: 'image', data: el.data, src: el.src, extension: ext } as SerializerElement);
    ctx.media.push({ name: mediaName, base64 });
    const embedRelId = addRelationship(ctx, kind === 'video' ? REL_TYPES.video : REL_TYPES.audio, `../media/${mediaName}`);
    let posterRelId = embedRelId;
    if (el.poster) {
        ctx.mediaIndex++;
        const pExt = el.poster.extension || 'png';
        const pName = `image${ctx.mediaIndex}.${pExt}`;
        const pData = await resolveImageData({ type: 'image', data: el.poster.data, src: el.poster.src, extension: pExt } as SerializerElement);
        ctx.media.push({ name: pName, base64: pData.base64 });
        posterRelId = addRelationship(ctx, REL_TYPES.image, `../media/${pName}`);
    }
    const mediaFileNode = kind === 'video'
        ? xmlNode('a:videoFile', { 'r:link': embedRelId })
        : xmlNode('a:audioFile', { 'r:link': embedRelId });
    return xmlNode('p:pic',
        null,
        xmlNode('p:nvPicPr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `${kind} ${id}` }),
            xmlNode('p:cNvPicPr'),
            xmlNode('p:nvPr', null, mediaFileNode)
        ),
        xmlNode('p:blipFill', null, xmlNode('a:blip', { 'r:embed': posterRelId }), xmlNode('a:stretch', null, xmlNode('a:fillRect'))),
        xmlNode('p:spPr', null, xmlNode('a:prstGeom', { prst: 'rect' }, xmlNode('a:avLst')))
    );
}

/** SmartArt 图示命名空间 */
const DGML_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

/** 生成确定性的 uniqueId（GUID 形态，便于被 PowerPoint 接受） */
function diagramUniqueId(seed: number): string {
    const hex = (seed & 0xfffffff).toString(16).padStart(8, '0');
    return `{A1B2C3D4-0000-0000-0000-${('00000000' + hex).slice(-12)}}`;
}

/** diagramType → dgm:cat type 缩写 */
function diagramCat(type: string): string {
    const map: Record<string, string> = {
        list: 'list', hierarchy: 'hier', orgChart: 'hier',
        process: 'process', cycle: 'cycle', pyramid: 'pyra', matrix: 'matrix'
    };
    return map[type] || 'list';
}

/** 构建图示数据模型（dataN.xml）：point + connector 树 */
function buildDiagramDataModel(nodes: DiagramNode[]): string {
    let modelId = 0;
    const points = [`<dsdgm:pt modelId="0" type="doc"><dsdgm:extLst/></dsdgm:pt>`];
    const connectors: string[] = [];
    const walk = (ns: DiagramNode[], parentId: number) => {
        for (const node of ns) {
            modelId++;
            const mid = modelId;
            const text = escapeXml(node.text || '');
            points.push(
                `<dsdgm:pt modelId="${mid}" type="node">` +
                `<dsdgm:prSet loColr="none" hiColr="none"/>` +
                `<dsdgm:spPr/>` +
                `<dsdgm:t><dsdgm:str val="${text}"/></dsdgm:t>` +
                `<dsdgm:extLst/></dsdgm:pt>`
            );
            const cxnId = 1000 + mid;
            connectors.push(`<dsdgm:cxn modelId="${cxnId}" type="parOf" srcId="${parentId}" destId="${mid}"/>`);
            if (node.children && node.children.length) walk(node.children, mid);
        }
    };
    walk(nodes, 0);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<dsdgm:dataModel xmlns:dsdgm="${DGML_NS}">` +
        `<dsdgm:ptLst>${points.join('')}</dsdgm:ptLst>` +
        `<dsdgm:cxnLst>${connectors.join('')}</dsdgm:cxnLst>` +
        `</dsdgm:dataModel>`;
}

/** 构建图示布局（layoutN.xml）：最小但结构完整的 layoutDef */
function buildDiagramLayout(type: string, seed: number): string {
    const cat = diagramCat(type);
    const uid = diagramUniqueId(seed * 4 + 1);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<dgm:layoutDef xmlns:dgm="${DGML_NS}" uniqueId="${uid}" minVer="2.0.0.0" ` +
        `defStyle="urn:microsoft.com/office/officeart/2005/8/diagramStyle" name="${escapeXml(type)}">` +
        `<dgm:title val="${escapeXml(type)}"/><dgm:desc val="pptx-parser generated diagram"/>` +
        `<dgm:catLst><dgm:cat type="${cat}" pri="1000"/></dgm:catLst>` +
        `<dgm:styleData/><dgm:adjLst/><dgm:nodeLst/><dgm:connLst/><dgm:foreachLst/>` +
        `<dgm:layoutNode name="ROOT" styleLbl="node"><dgm:alg type="hierChild"/>` +
        `<dgm:forEach val="__NODE__" refType="begin"><dgm:layoutNode name="NODE" styleLbl="node">` +
        `<dgm:alg type="tx"/><dgm:shp txBox="1" type="rect"/>` +
        `<dgm:forEach val="__CHILD__" refType="child">` +
        `<dgm:layoutNode name="CHILD" styleLbl="node"><dgm:alg type="tx"/><dgm:shp txBox="1" type="rect"/></dgm:layoutNode>` +
        `</dgm:forEach></dgm:layoutNode></dgm:forEach></dgm:layoutNode>` +
        `<dgm:qryLst/><dgm:clrData/><dgm:ruleLst/><dgm:layoutVarLst/><dgm:algLst/>` +
        `</dgm:layoutDef>`;
}

/** 构建图示配色（colorsN.xml）：最小 colorsDef */
function buildDiagramColors(type: string, seed: number): string {
    const uid = diagramUniqueId(seed * 4 + 2);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<dgm:colorsDef xmlns:dgm="${DGML_NS}" uniqueId="${uid}" minVer="2.0.0.0" ` +
        `defStyle="urn:microsoft.com/office/officeart/2005/8/diagramColors" name="${escapeXml(type)} colors">` +
        `<dgm:title val="${escapeXml(type)} colors"/><dgm:desc val="generated"/>` +
        `<dgm:catLst/><dgm:varLst/><dgm:styleLblLst/><dgm:sampRagLst/></dgm:colorsDef>`;
}

/** 构建图示快速样式（quickStyleN.xml）：最小 quickStyleDef */
function buildDiagramQuickStyle(type: string, seed: number): string {
    const uid = diagramUniqueId(seed * 4 + 3);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<dgm:quickStyleDef xmlns:dgm="${DGML_NS}" uniqueId="${uid}" minVer="2.0.0.0" ` +
        `defStyle="urn:microsoft.com/office/officeart/2005/8/diagramQuickStyle" name="${escapeXml(type)} quick style">` +
        `<dgm:title val="${escapeXml(type)} quick style"/><dgm:desc val="generated"/>` +
        `<dgm:catLst/><dgm:varLst/><dgm:styleLblLst/></dgm:quickStyleDef>`;
}

/**
 * 构建 SmartArt 图示元素（p:graphicFrame + 原生 diagrams/* 部件）
 * 生成 data/layout/colors/quickStyle 四件套，并登记幻灯片 → dataN.xml 关系；
 * 四部件由上层（jsonToPptx / editPptx.addSlide）写入 ppt/diagrams/ 并补 dataN.xml.rels。
 *
 * 限制：生成的 layoutDef/colorsDef/quickStyleDef 为结构级最小骨架，可被本解析器 round-trip、
 * 可被 PowerPoint 打开，但 PowerPoint 对 SmartArt 布局引擎校验严格，自定义布局可能不渲染为
 * 完整图示图形。如需保真渲染，应提供完整的标准 layoutDef 模板替换 buildDiagramLayout 等。
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 图示元素 JSON { type:'diagram', diagramType, nodes:[{text,children?}] }
 * @returns {Promise<Object>} p:graphicFrame 节点
 */
async function buildDiagramElement(ctx: SerializerContext, el: SerializerElement) {
    ctx.diagramIndex++;
    const n = ctx.diagramIndex;
    const dgmType = el.diagramType || 'list';
    const nodes: DiagramNode[] = el.nodes || [];
    const dataXml = buildDiagramDataModel(nodes);
    const layoutXml = buildDiagramLayout(dgmType, n);
    const colorsXml = buildDiagramColors(dgmType, n);
    const quickStyleXml = buildDiagramQuickStyle(dgmType, n);
    const dataRelId = addRelationship(ctx, REL_TYPES.diagramData, `../diagrams/data${n}.xml`);
    ctx.diagrams.push({ index: n, dataXml, layoutXml, colorsXml, quickStyleXml });

    const id = ctx.nextElementId++;
    return xmlNode('p:graphicFrame',
        null,
        xmlNode('p:nvGraphicFramePr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `Diagram ${id}` }),
            xmlNode('p:cNvGraphicFramePr'),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:xfrm',
            null,
            xmlNode('a:off', { x: pxToEmu(el.x || 0), y: pxToEmu(el.y || 0) }),
            xmlNode('a:ext', { cx: pxToEmu(el.width || 400), cy: pxToEmu(el.height || 300) })
        ),
        xmlNode('a:graphic',
            null,
            xmlNode('a:graphicData', { uri: DGML_NS },
                xmlNode('dgm:rel', { 'xmlns:dgm': DGML_NS, 'xmlns:r': NS.r, 'r:id': dataRelId }))
        )
    );
}

export async function buildElement(ctx: SerializerContext, el: SerializerElement) {
    if (!el || typeof el !== 'object') return null;
    // 显式回退：已支持的语义类型也可用 __raw 原样回写（语义层可能丢失主题色/动画等细节）
    if (el.rawFallback && el.__raw) return buildRawElement(ctx, el);
    switch (el.type) {
        case 'text':
            return buildTextElement(ctx, el);
        case 'shape':
            return buildShapeElement(ctx, el);
        case 'image':
            return buildImageElement(ctx, el);
        case 'video':
        case 'audio':
            return buildMediaElement(ctx, el, el.type as 'video' | 'audio');
        case 'chart':
            return buildChartElement(ctx, el);
        case 'table':
            return buildTableElement(ctx, el);
        case 'group':
            return buildGroupElement(ctx, el);
        case 'diagram':
            // 有 __raw 载荷（解析端回读）时回落无损回退；否则按语义层生成原生 diagrams 部件
            return el.__raw ? buildRawElement(ctx, el) : buildDiagramElement(ctx, el);
        default:
            // 语义层未覆盖（SmartArt / 组合 / 连接符 / OLE 等）→ 回退 __raw
            return buildRawElement(ctx, el);
    }
}

/**
 * 构建背景节点（p:bg）
 * @param {Object} bg - 背景描述（字符串 / {type:'solid'} / {type:'gradient'} / {type:'image'}）
 * @param {Object} ctx - 构建上下文（图片型背景需登记媒体与关系）
 * @returns {Promise<Object|null>} p:bg 节点，无背景返回 null
 */
export async function buildBackground(bg: string | PptxBackground | null | undefined, ctx: SerializerContext): Promise<BuilderNode | null> {
    if (!bg || bg === 'none') return null;

    let bgPrChildren: (BuilderNode | string)[];
    if (typeof bg === 'string') {
        bgPrChildren = [
            xmlNode('a:solidFill', colorNode(bg)),
            xmlNode('a:effectLst')
        ];
    } else if (bg.type === 'solid') {
        bgPrChildren = [
            xmlNode('a:solidFill', colorNode(bg.color)),
            xmlNode('a:effectLst')
        ];
    } else if (bg.type === 'gradient') {
        const stops = (bg.stops || []).map(s =>
            xmlNode('a:gs', { pos: Math.round((s.position || 0) * 100000) },
                xmlNode('a:srgbClr', { val: colorToHex(s.color) }))
        );
        const ang = bg.direction === 'vertical' ? 90 : bg.direction === 'diagonal' ? 45 : 0;
        bgPrChildren = [
            xmlNode('a:gradFill', null,
                xmlNode('a:gsLst', null, ...stops),
                xmlNode('a:lin', { ang, scaled: 1 })
            )
        ];
    } else {
        // 图片型背景：内联媒体并登记关系
        const imgEl: SerializerElement = { type: 'image', data: bg.data, src: bg.src, extension: bg.extension };
        const { base64, ext } = await resolveImageData(imgEl);
        ctx.mediaIndex++;
        const mediaName = `image${ctx.mediaIndex}.${ext}`;
        ctx.media.push({ name: mediaName, base64 });
        const embedRelId = addRelationship(ctx, REL_TYPES.image, `../media/${mediaName}`);
        bgPrChildren = [
            xmlNode('a:blipFill', null,
                xmlNode('a:blip', { 'r:embed': embedRelId }),
                xmlNode('a:stretch', null, xmlNode('a:fillRect'))
            )
        ];
    }

    return xmlNode('p:bg', null, xmlNode('p:bgPr', null, ...bgPrChildren));
}

/** 过渡类型 → OOXML p:* 子元素 */
const TRANSITION_TAG: Record<string, string> = {
    fade: 'p:fade', wipe: 'p:wipe', push: 'p:push', cover: 'p:cover',
    blinds: 'p:blinds', checker: 'p:checker', circle: 'p:circle', comb: 'p:comb',
    dissolve: 'p:dissolve', random: 'p:random', split: 'p:split', strips: 'p:strips'
};

/**
 * 构建过渡节点（p:transition）
 * @param {Object} t - 过渡描述 { type, duration(ms) }
 * @returns {Object|null} p:transition 节点
 */
export function buildTransition(t: PptxTransition | undefined): BuilderNode | null {
    if (!t) return null;
    const tag = TRANSITION_TAG[t.type] || 'p:fade';
    const spd = t.duration <= 750 ? '1' : t.duration >= 1500 ? '3' : '2';
    return xmlNode('p:transition', { spd }, xmlNode(tag));
}

/**
 * 构建备注幻灯片节点（p:notesSlide）
 * @param {string} notes - 备注文本
 * @returns {Object} p:notesSlide 根节点
 */
export function buildNotesSlide(notes: string): BuilderNode {
    return xmlNode('p:notesSlide',
        { 'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main' },
        xmlNode('p:cSld', null,
            xmlNode('p:spTree', null,
                xmlNode('p:nvGrpSpPr', null,
                    xmlNode('p:cNvPr', { id: 1, name: '' }),
                    xmlNode('p:cNvGrpSpPr'),
                    xmlNode('p:nvPr')
                ),
                xmlNode('p:grpSpPr', null,
                    xmlNode('a:xfrm', null,
                        xmlNode('a:off', { x: 0, y: 0 }),
                        xmlNode('a:ext', { cx: 0, cy: 0 }),
                        xmlNode('a:chOff', { x: 0, y: 0 }),
                        xmlNode('a:chExt', { cx: 0, cy: 0 })
                    )
                ),
                xmlNode('p:sp', null,
                    xmlNode('p:nvSpPr', null,
                        xmlNode('p:cNvPr', { id: 2, name: 'Notes Placeholder 1' }),
                        xmlNode('p:cNvSpPr', null, xmlNode('p:ph', { type: 'body', idx: 1 })),
                        xmlNode('p:nvPr')
                    ),
                    xmlNode('p:spPr'),
                    xmlNode('p:txBody', null,
                        xmlNode('a:bodyPr', { rtlCol: 0 }),
                        xmlNode('a:lstStyle'),
                        xmlNode('a:p', null,
                            xmlNode('a:r', null,
                                xmlNode('a:rPr', { lang: 'zh-CN' }),
                                xmlNode('a:t', null, notes)
                            )
                        )
                    )
                )
            )
        ),
        xmlNode('p:clrMapOvr', null, xmlNode('a:masterClrMapping'))
    );
}

/**
 * 构建完整幻灯片 XML 根节点（p:sld）
 * @param {Object} ctx - 构建上下文
 * @param {Object} slide - 幻灯片 JSON { background, notes, transition, elements }
 * @returns {Promise<Object>} p:sld 根节点
 */
/** 构建幻灯片计时（自动播放 + 元素动画），生成 p:timing */
function buildTimingNode(slide: SerializerSlide): BuilderNode | null {
    if (!slide) return null;
    const adv = slide.advanceTime;
    const anims = slide.animations || [];
    if (adv == null && anims.length === 0) return null;

    const childNodes: BuilderNode[] = [];
    if (adv != null) {
        childNodes.push(xmlNode('p:cTn',
            { id: 2, fill: 'hold' },
            xmlNode('p:stCondLst', null, xmlNode('p:cond', { type: 'afterTime', val: Math.round(adv) })),
            xmlNode('p:childTnLst', null, xmlNode('p:cTn', { id: 3, pres: 'sld', fill: 'hold' }))
        ));
    }
    let nid = 10;
    for (const a of anims) {
        const spid = a.target != null ? a.target + 1 : 2; // 元素 id 从 2 起
        const preset = a.type === 'flyIn' ? 'flyIn' : a.type === 'zoom' ? 'zoom' : a.type === 'wipe' ? 'wipe' : 'fade';
        const effectChildren: BuilderNode[] = [];
        if (a.duration != null) effectChildren.push(xmlNode('p:cTn', { id: nid++, dur: Math.round(a.duration * 1000), fill: 'hold' }));
        childNodes.push(xmlNode('p:cTn',
            { id: nid++, fill: 'hold' },
            xmlNode('p:tgtEl', null, xmlNode('p:spTgt', { spid })),
            xmlNode('p:childTnLst', null,
                xmlNode('p:cTn', { id: nid++, presetClass: 'entr', presetId: 1, type: 'withEffect', preset }, ...effectChildren)
            )
        ));
    }
    return xmlNode('p:timing',
        null,
        xmlNode('p:tnLst',
            null,
            xmlNode('p:par',
                null,
                xmlNode('p:cTn', { id: 1, fill: 'hold' },
                    xmlNode('p:stCondLst', null, xmlNode('p:cond', { type: 'afterPrev' })),
                    xmlNode('p:childTnLst', null, ...childNodes)
                )
            )
        )
    );
}

export async function buildSlideRoot(ctx: SerializerContext, slide: SerializerSlide) {
    const elementNodes = [];
    for (const el of (slide && slide.elements) || []) {
        const node = await buildElement(ctx, el);
        if (node) elementNodes.push(node);
    }

    // 背景（纯色 / 渐变 / 图片）
    const bgNode = await buildBackground(slide && slide.background, ctx);

    // 过渡效果
    const transitionNode = buildTransition(slide && slide.transition);
    const timingNode = buildTimingNode(slide);

    return xmlNode('p:sld',
        { 'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
          show: (slide && slide.hidden) ? 0 : null },
        xmlNode('p:cSld',
            null,
            bgNode,
            xmlNode('p:spTree',
                null,
                xmlNode('p:nvGrpSpPr',
                    null,
                    xmlNode('p:cNvPr', { id: 1, name: '' }),
                    xmlNode('p:cNvGrpSpPr'),
                    xmlNode('p:nvPr')
                ),
                xmlNode('p:grpSpPr',
                    null,
                    xmlNode('a:xfrm',
                        null,
                        xmlNode('a:off', { x: 0, y: 0 }),
                        xmlNode('a:ext', { cx: 0, cy: 0 }),
                        xmlNode('a:chOff', { x: 0, y: 0 }),
                        xmlNode('a:chExt', { cx: 0, cy: 0 })
                    )
                ),
                ...elementNodes
            )
        ),
        transitionNode,
        timingNode,
        xmlNode('p:clrMapOvr', null, xmlNode('a:masterClrMapping'))
    );
}
