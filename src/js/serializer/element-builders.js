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

import { xmlNode, pxToEmu, ptToSz, ptToEmu, degToRot, colorToHex, NS, escapeXml } from './xml-builder.js';
import { REL_TYPES } from './templates.js';

/**
 * 创建元素构建上下文
 * @param {Object} [options]
 * @param {number} [options.startMediaIndex=0] - 媒体文件起始编号（避免与已有文件冲突）
 * @returns {Object} 构建上下文
 */
export function createElementContext(options = {}) {
    return {
        /** 关系列表 {relId, type, target, external} */
        rels: [],
        /** 媒体文件列表 {name, base64} */
        media: [],
        /** 幻灯片内元素自增 id（1 被 spTree 根占用） */
        nextElementId: 2,
        /** 关系自增 id（rId1 固定为版式引用） */
        nextRelId: 2,
        /** 媒体文件自增编号 */
        mediaIndex: options.startMediaIndex || 0,
        /** 图表部件列表 {name, xml}（生成后由上层写入 ppt/charts/） */
        charts: [],
        /** 图表自增编号 */
        chartIndex: 0
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
function addRelationship(ctx, type, target, external) {
    const relId = `rId${ctx.nextRelId++}`;
    ctx.rels.push({ relId, type, target, external: !!external });
    return relId;
}

/**
 * 解析图片来源数据
 * @param {Object} el - 图片元素
 * @returns {Promise<{base64: string, ext: string}>} 图片数据
 */
async function resolveImageData(el) {
    if (el.data) {
        const str = String(el.data);
        const dataUrlMatch = str.match(/^data:image\/([a-z0-9.+-]+);base64,(.+)$/i);
        if (dataUrlMatch) {
            let ext = dataUrlMatch[1].toLowerCase();
            if (ext === 'jpg') ext = 'jpeg';
            if (ext === 'svg+xml') ext = 'svg';
            return { base64: dataUrlMatch[2], ext: el.extension || ext };
        }
        // 裸 base64
        return { base64: str, ext: el.extension || 'png' };
    }

    if (el.src) {
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
 * 从 URL 猜测图片扩展名
 * @param {string} url - 图片 URL
 * @returns {string|null} 扩展名
 */
function guessExtFromUrl(url) {
    const match = String(url).split('?')[0].match(/\.([a-z0-9]+)$/i);
    return match ? match[1].toLowerCase() : null;
}

/**
 * ArrayBuffer 转 base64（兼容浏览器与 Node）
 * @param {ArrayBuffer} buffer - 二进制数据
 * @returns {string} base64 字符串
 */
function arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
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
function buildXfrm(el) {
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
function buildHyperlink(ctx, href) {
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
function buildTextRun(ctx, text, opts = {}) {
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
function buildParagraph(ctx, paragraph, defaults) {
    const alignMap = { left: 'l', center: 'ctr', right: 'r', justify: 'just' };
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
        { algn: alignMap[p.align || defaults.align] || null },
        ...pPrChildren
    );

    // 运行列表：显式 runs 优先，否则用 text + 元素级默认样式
    let runs;
    if (Array.isArray(p.runs) && p.runs.length > 0) {
        runs = p.runs.map(r => {
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
function normalizeParagraphs(el) {
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
function buildTextElement(ctx, el) {
    const id = ctx.nextElementId++;
    const anchorMap = { top: null, middle: 'ctr', bottom: 'b' };
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
                { wrap: 'square', rtlCol: 0, anchor: anchorMap[el.valign] || null }
            ),
            xmlNode('a:lstStyle'),
            ...normalizeParagraphs(el).map(p => buildParagraph(ctx, p, defaults))
        )
    );
}

/**
 * 构建形状元素（p:sp）
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 形状元素 JSON
 * @returns {Object} p:sp 节点
 */
function buildShapeElement(ctx, el) {
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
async function buildImageElement(ctx, el) {
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
function buildChartElement(ctx, el) {
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
function strRefXml(values, col) {
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
function numRefXml(values, col) {
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
function buildChartXml(el) {
    const type = el.chartType || 'barChart';
    const isPie = /pie/i.test(type);
    const isScatter = type === 'scatterChart';
    const cats = el.categories || [];
    const series = el.series || [];
    const varyColors = el.varyColors !== undefined ? (el.varyColors ? 1 : 0) : (isPie ? 1 : 0);

    const serXml = series.map((s, i) => {
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
        axes = '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
            '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="112"/></c:catAx>' +
            '<c:valAx><c:axId val="112"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
            '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/><c:majorGridlines/></c:valAx>';
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
 * 构建单个幻灯片元素节点
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 元素 JSON
 * @returns {Promise<Object|null>} 元素节点，不支持的类型返回 null
 */
export async function buildElement(ctx, el) {
    if (!el || typeof el !== 'object') return null;
    switch (el.type) {
        case 'text':
            return buildTextElement(ctx, el);
        case 'shape':
            return buildShapeElement(ctx, el);
        case 'image':
            return buildImageElement(ctx, el);
        case 'chart':
            return buildChartElement(ctx, el);
        default:
            return null;
    }
}

/**
 * 构建完整幻灯片 XML 根节点（p:sld）
 * @param {Object} ctx - 构建上下文
 * @param {Object} slide - 幻灯片 JSON { background, elements }
 * @returns {Promise<Object>} p:sld 根节点
 */
export async function buildSlideRoot(ctx, slide) {
    const elementNodes = [];
    for (const el of (slide && slide.elements) || []) {
        const node = await buildElement(ctx, el);
        if (node) elementNodes.push(node);
    }

    // 背景色
    let bgNode = null;
    const bgColor = slide && slide.background;
    if (bgColor && bgColor !== 'none') {
        bgNode = xmlNode('p:bg',
            null,
            xmlNode('p:bgPr',
                null,
                xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(bgColor) })),
                xmlNode('a:effectLst')
            )
        );
    }

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
        xmlNode('p:clrMapOvr', null, xmlNode('a:masterClrMapping'))
    );
}
