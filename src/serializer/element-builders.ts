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
    parts: SerializerPart[];
    nextElementId: number;
    nextRelId: number;
    mediaIndex: number;
    chartIndex: number;
}
/** createElementContext 的选项 */
export interface ElementContextOptions {
    /** 媒体文件起始编号（避免与已有文件冲突） */
    startMediaIndex?: number;
    /** 图表部件起始编号（避免与已有部件冲突） */
    startChartIndex?: number;
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
    bullet?: boolean;
}
/** 表格单元格 */
export interface SerializerTableCell {
    text?: string;
    paragraphs?: ParagraphSpec[];
    colSpan?: number;
    rowSpan?: number;
    fill?: string;
    align?: string;
    valign?: string;
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
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
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
    lang?: string;
    href?: string;
    /** 形状填充：颜色字符串 / { color } / 'none' / null */
    fill?: string | { color?: string } | null;
    /** 形状边框：{ color, width } / 'none' / null */
    line?: { color?: string; width?: number } | 'none' | null;
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
    /** 表格：行数据 */
    rows?: SerializerTableRow[];
    /** 表格：列宽（px，缺省均分） */
    colWidths?: number[];
    /** 表格：行高（px，缺省均分） */
    rowHeights?: number[];
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
    elements?: SerializerElement[];
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
        chartIndex: options.startChartIndex || 0
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
    return xmlNode('a:xfrm',
        { rot: el.rotation ? degToRot(el.rotation) : null },
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
        rPrChildren.push(xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(opts.color) })));
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

    // 段落属性
    const pPrChildren = [];
    if (p.bullet) {
        pPrChildren.push(xmlNode('a:buFont', { typeface: 'Arial' }));
        pPrChildren.push(xmlNode('a:buChar', { char: '•' }));
    } else {
        pPrChildren.push(xmlNode('a:buNone'));
    }
    const pPr = xmlNode('a:pPr',
        { algn: alignMap[p.align || defaults.align || 'left'] || null },
        ...pPrChildren
    );

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
    if (Array.isArray(el.paragraphs) && el.paragraphs.length > 0) {
        return el.paragraphs;
    }
    if (Array.isArray(el.runs) && el.runs.length > 0) {
        return [{ runs: el.runs }];
    }
    if (el.text !== undefined) {
        return String(el.text).split('\n').map(t => ({ text: t }));
    }
    return [{ text: '' }];
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
                { wrap: 'square', rtlCol: 0, anchor: anchorMap[el.valign ?? 'top'] ?? null }
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
function buildShapeElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;

    // 填充
    let fillNode;
    if (el.fill === 'none' || el.fill === null) {
        fillNode = xmlNode('a:noFill');
    } else {
        const fillColor = typeof el.fill === 'string' ? el.fill : (el.fill && el.fill.color);
        if (fillColor) {
            fillNode = xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(fillColor) }));
        } else {
            fillNode = null; // 未指定填充则继承主题
        }
    }

    // 边框
    let lineNode;
    if (el.line === 'none' || el.line === null) {
        lineNode = xmlNode('a:ln', null, xmlNode('a:noFill'));
    } else if (el.line) {
        const w = el.line.width !== undefined ? el.line.width : 1;
        lineNode = xmlNode('a:ln',
            { w: ptToEmu(w) },
            xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(el.line.color) }))
        );
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
            xmlNode('a:prstGeom', { prst: el.shapeType || 'rect' }, xmlNode('a:avLst')),
            fillNode,
            lineNode
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
            xmlNode('a:blip', { 'r:embed': embedRelId }),
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
    const cats = el.categories || [];
    const series = el.series || [];
    const varyColors = el.varyColors !== undefined ? (el.varyColors ? 1 : 0) : (isPie ? 1 : 0);

    const serXml = series.map((s: ChartSeriesSpec, i: number) => {
        const tx = `<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f>` +
            `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(s.name || `Series${i + 1}`)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
        let data;
        if (isScatter) {
            data = `<c:xVal>${numRefXml(s.x || [], 'B')}</c:xVal><c:yVal>${numRefXml(s.y || [], 'C')}</c:yVal>`;
        } else {
            data = `<c:cat>${strRefXml(cats, 'A')}</c:cat><c:val>${numRefXml(s.values || [], 'B')}</c:val>`;
        }
        return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${data}</c:ser>`;
    }).join('');

    // 图表类型特定根（含轴 id，散点/柱状/折线/面积需要）
    let plotChart;
    if (isPie) {
        plotChart = `<c:${type}><c:varyColors val="${varyColors}"/>${serXml}</c:${type}>`;
    } else {
        const dir = type === 'barChart' ? `<c:barDir val="${el.barDir || 'col'}"/>` : '';
        const grouping = type === 'lineChart' ? '<c:grouping val="standard"/>' : '';
        plotChart = `<c:${type}>${dir}${grouping}<c:varyColors val="${varyColors}"/>${serXml}` +
            `<c:axId val="111"/><c:axId val="112"/></c:${type}>`;
    }

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

    const legendXml = el.legend !== false
        ? '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>'
        : '';

    const autoTitleDeleted = `<c:autoTitleDeleted val="${el.title ? 0 : 1}"/>`;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
        `<c:chart>${titleXml}${autoTitleDeleted}` +
        `<c:plotArea><c:layout/>${plotChart}${axes}</c:plotArea>` +
        `${legendXml}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

/**
 * 构建表格单元格（a:tc）
 * @param {Object} ctx - 构建上下文
 * @param {Object} cell - 单元格 { text | paragraphs, colSpan, rowSpan, fill, align, valign, ...样式 }
 * @returns {Object} a:tc 节点
 */
function buildTableCell(ctx: SerializerContext, cell: SerializerTableCell): BuilderNode {
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
    if (cell.fill) {
        tcPrChildren.push(xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(cell.fill) })));
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
            ...(row.cells || []).map((cell) => buildTableCell(ctx, cell))
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
                    xmlNode('a:tblPr', { firstRow: 1, bandRow: 1 }),
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
        case 'chart':
            return buildChartElement(ctx, el);
        case 'table':
            return buildTableElement(ctx, el);
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
            xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(bg) })),
            xmlNode('a:effectLst')
        ];
    } else if (bg.type === 'solid') {
        bgPrChildren = [
            xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(bg.color) })),
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

    return xmlNode('p:sld',
        { 'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main' },
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
        xmlNode('p:clrMapOvr', null, xmlNode('a:masterClrMapping'))
    );
}
