/**
 * @fefeding/ppt-parser v1.0.34
 * PPTX文件解析与序列化核心库，纯TS编写，支持解析PPTX为JSON结构、JSON序列化为标准PPTX文件，无框架依赖
 * MIT License
 */
'use strict';

var pptParser = require('@fefeding/ppt-parser');
var pdfLib = require('pdf-lib');
var marked = require('marked');
var parse5 = require('parse5');
var docx = require('docx');

class ConvertError extends Error {
    constructor(message, format, options = {}) {
        super(message);
        this.name = 'ConvertError';
        this.format = format;
        this.code = options.code;
        this.cause = options.cause;
        Object.setPrototypeOf(this, ConvertError.prototype);
    }
}
function describeError(error) {
    if (error instanceof Error)
        return error.message;
    if (typeof error === 'string')
        return error;
    try {
        return JSON.stringify(error);
    }
    catch {
        return String(error);
    }
}
function wrapError(format, action, error) {
    if (error instanceof ConvertError)
        return error;
    const code = error?.code;
    return new ConvertError(`${format} ${action}失败：${describeError(error)}`, format, {
        code,
        cause: error
    });
}

function stripDataUrlPrefix(value) {
    const match = /^data:[^,]*?(;base64)?,(.*)$/s.exec(value.trim());
    if (match) {
        if (!match[1])
            return '';
        return match[2];
    }
    return value.trim();
}
function isNodeBuffer(value) {
    return typeof Buffer !== 'undefined' && Buffer.isBuffer(value);
}
function base64ToBytes(base64) {
    const clean = base64.replace(/\s+/g, '');
    if (!clean)
        return new Uint8Array(0);
    const globalAtob = globalThis.atob;
    const binary = globalAtob
        ? globalAtob(clean)
        : Buffer.from(clean, 'base64').toString('binary');
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i) & 0xff;
    }
    return bytes;
}
function bytesToBase64(bytes) {
    const globalBtoa = globalThis.btoa;
    if (globalBtoa) {
        let binary = '';
        const chunkSize = 0x8000;
        for (let i = 0; i < bytes.length; i += chunkSize) {
            binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunkSize)));
        }
        return globalBtoa(binary);
    }
    return Buffer.from(bytes).toString('base64');
}
async function toUint8Array(input) {
    if (input == null) {
        throw new TypeError('convert: 缺少输入数据');
    }
    if (typeof input === 'string') {
        return base64ToBytes(stripDataUrlPrefix(input));
    }
    if (isNodeBuffer(input)) {
        const buf = input;
        return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
    }
    if (input instanceof Uint8Array) {
        return input;
    }
    if (ArrayBuffer.isView(input)) {
        return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
    }
    if (input instanceof ArrayBuffer) {
        return new Uint8Array(input);
    }
    const maybeBlob = input;
    if (typeof maybeBlob.arrayBuffer === 'function') {
        return new Uint8Array(await maybeBlob.arrayBuffer());
    }
    throw new TypeError('convert: 不支持的输入类型，请传入 ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串');
}
async function toArrayBuffer(input) {
    if (typeof input === 'string') {
        const bytes = base64ToBytes(stripDataUrlPrefix(input));
        return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    }
    if (input instanceof ArrayBuffer) {
        return input;
    }
    if (isNodeBuffer(input)) {
        const buf = input;
        return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    }
    const bytes = await toUint8Array(input);
    return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}
async function toBinaryPayload(input) {
    const bytes = await toUint8Array(input);
    if (typeof Buffer !== 'undefined') {
        return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    }
    return bytes;
}
function assertNonEmpty(bytes, format, minSize = 4) {
    const size = bytes instanceof ArrayBuffer ? bytes.byteLength : bytes.byteLength;
    if (size < minSize) {
        throw new TypeError(`convert: ${format} 数据无效（长度 ${size} 字节，至少需要 ${minSize} 字节）`);
    }
}
const MD_STRONG = [
    /^#{1,6}\s+\S/m,
    /^```/m,
    /\[[^\]\n]+\]\([^)\n]+\)/,
    /!\[[^\]\n]*\]\([^)\n]+\)/,
    /^\s*\|.*\|\s*$/m
];
const MD_WEAK = [
    /^\s*[-*+]\s+\S/m,
    /^\s*\d+\.\s+\S/m,
    /^>\s+\S/m,
    /^\s*(?:-{3,}|\*{3,}|_{3,})\s*$/m
];
function looksLikeMarkdown(text) {
    if (MD_STRONG.some((re) => re.test(text)))
        return true;
    return MD_WEAK.filter((re) => re.test(text)).length >= 2;
}
async function detectFormat(input) {
    const bytes = await toUint8Array(input);
    if (!bytes.length)
        return null;
    if (bytes.length >= 4) {
        if (bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46)
            return 'pdf';
        if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
            const isZip = (bytes[2] === 0x03 && bytes[3] === 0x04) ||
                (bytes[2] === 0x05 && bytes[3] === 0x06) ||
                (bytes[2] === 0x07 && bytes[3] === 0x08);
            if (!isZip)
                return null;
            const text = new TextDecoder('latin1').decode(bytes.slice(0, Math.min(bytes.length, 8192)));
            if (text.includes('word/'))
                return 'docx';
            if (text.includes('xl/'))
                return 'xlsx';
            if (text.includes('ppt/'))
                return 'pptx';
            const tail = new TextDecoder('latin1').decode(bytes.slice(Math.max(0, bytes.length - 65536)));
            if (tail.includes('word/'))
                return 'docx';
            if (tail.includes('xl/'))
                return 'xlsx';
            if (tail.includes('ppt/'))
                return 'pptx';
            return null;
        }
    }
    if (!bytes.includes(0)) {
        try {
            const head = new TextDecoder('utf-8', { fatal: true }).decode(bytes.slice(0, Math.min(bytes.length, 8192)));
            return looksLikeMarkdown(head) ? 'md' : 'txt';
        }
        catch {
        }
    }
    return null;
}

const MIME_BY_EXT = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    webp: 'image/webp',
    svg: 'image/svg+xml',
    tif: 'image/tiff',
    tiff: 'image/tiff',
    emf: 'image/x-emf',
    wmf: 'image/x-wmf',
    pdf: 'application/pdf',
    zip: 'application/zip',
    mp4: 'video/mp4',
    mp3: 'audio/mpeg',
    wav: 'audio/wav',
    ogg: 'audio/ogg',
    txt: 'text/plain',
    log: 'text/plain',
    csv: 'text/csv',
    md: 'text/markdown',
    json: 'application/json',
    xml: 'application/xml',
    html: 'text/html',
    htm: 'text/html',
    doc: 'application/msword',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ppt: 'application/vnd.ms-powerpoint',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    eml: 'message/rfc822',
    bin: 'application/octet-stream'
};
function mimeOf(nameOrPath, fallback = 'application/octet-stream') {
    const ext = (nameOrPath.split('.').pop() || '').toLowerCase();
    return MIME_BY_EXT[ext] || fallback;
}
const EXT_BY_MIME = {
    'image/png': '.png',
    'image/jpeg': '.jpg',
    'image/gif': '.gif',
    'image/bmp': '.bmp',
    'image/webp': '.webp',
    'image/svg+xml': '.svg',
    'image/tiff': '.tif',
    'image/x-emf': '.emf',
    'image/x-wmf': '.wmf',
    'application/pdf': '.pdf',
    'application/zip': '.zip',
    'video/mp4': '.mp4',
    'audio/mpeg': '.mp3',
    'audio/wav': '.wav',
    'audio/ogg': '.ogg'
};
function extOf(mime, fallback = '') {
    return EXT_BY_MIME[mime] ?? fallback;
}
function createAssetId(prefix, index) {
    return `${prefix}-${String(index).padStart(4, '0')}`;
}
function createAsset(params) {
    const mime = params.mime ||
        mimeOf(params.part || params.name || '', '') ||
        sniffMime(params.data);
    const ext = extOf(mime, params.part?.split('.').pop() ? `.${params.part.split('.').pop()}` : '');
    const asset = {
        id: params.id,
        name: params.name || params.id + ext,
        mime,
        data: params.data,
        anchor: params.anchor
    };
    if (params.part)
        asset.part = params.part;
    if (params.width != null)
        asset.width = params.width;
    if (params.height != null)
        asset.height = params.height;
    if (params.alt)
        asset.alt = params.alt;
    if (params.meta)
        asset.meta = params.meta;
    return asset;
}
function parseDataUrl(dataUrl) {
    const match = /^data:([^;,]*)(;base64)?,([\s\S]*)$/.exec(dataUrl.trim());
    if (!match)
        return null;
    const mime = match[1] || 'application/octet-stream';
    const payload = match[3] || '';
    if (match[2]) {
        return { bytes: base64ToBytes(payload), mime };
    }
    return { bytes: new TextEncoder().encode(decodeURIComponent(payload)), mime };
}
function assetToBase64(asset) {
    return bytesToBase64(toBytes(asset));
}
function assetToDataUrl(asset, mime) {
    const bytes = toBytes(asset);
    const type = mime || ('mime' in asset ? asset.mime : undefined) || mimeOf('', 'application/octet-stream');
    return `data:${type};base64,${bytesToBase64(bytes)}`;
}
function assetToBlob(asset, mime) {
    const bytes = toBytes(asset);
    const type = mime || ('mime' in asset ? asset.mime : undefined) || mimeOf('', 'application/octet-stream');
    return new Blob([bytes.slice().buffer], { type });
}
function assetToNodeStream(asset) {
    const bytes = toBytes(asset);
    const Ctor = globalThis.Readable;
    if (Ctor)
        return new Ctor(bytes);
    const RS = globalThis.ReadableStream;
    if (RS)
        return new RS(bytes.slice());
    throw new TypeError('convert: 当前环境既无 Node Readable 也无 ReadableStream，无法生成流');
}
function assetToWebStream(asset) {
    const bytes = toBytes(asset).slice();
    const RS = globalThis
        .ReadableStream;
    if (RS)
        return new RS(bytes);
    throw new TypeError('convert: 当前环境不支持 ReadableStream，请改用 assetToNodeStream 或 assetToArrayBuffer');
}
function assetToArrayBuffer(asset) {
    const bytes = toBytes(asset);
    return bytes.slice().buffer;
}
function toBytes(asset) {
    return asset instanceof Uint8Array ? asset : asset.data;
}
function toAssetTable(assets) {
    const table = {};
    for (const asset of assets || [])
        table[asset.id] = asset;
    return table;
}
function getAsset(source, id) {
    if (!source || !id)
        return undefined;
    if (Array.isArray(source))
        return source.find((asset) => asset.id === id);
    if (Array.isArray(source.assets)) {
        return source.assets.find((asset) => asset.id === id);
    }
    return source[id];
}
function summarizeAssets(assets) {
    let bytes = 0;
    for (const asset of assets || [])
        bytes += asset.data?.byteLength || 0;
    return { count: assets?.length || 0, bytes };
}
function isAssetRef(value) {
    if (!value || typeof value !== 'object')
        return false;
    const record = value;
    return typeof record.assetId === 'string' || (!!record.asset && typeof record.asset === 'object');
}
function isDocAsset(value) {
    if (!value || typeof value !== 'object')
        return false;
    const record = value;
    return typeof record.id === 'string' && record.data instanceof Uint8Array;
}
async function resolveAssetBytes(value, table) {
    if (value == null)
        return null;
    if (isDocAsset(value))
        return { bytes: value.data, mime: value.mime || sniffMime(value.data) };
    if (isAssetRef(value)) {
        const ref = value;
        const asset = ref.asset ?? getAsset(table, ref.assetId);
        if (!asset)
            return null;
        return { bytes: asset.data, mime: asset.mime };
    }
    if (value instanceof Uint8Array || value instanceof ArrayBuffer) {
        const bytes = await toUint8Array(value);
        return { bytes, mime: sniffMime(bytes) };
    }
    if (typeof value === 'string') {
        const bytes = await toUint8Array(value);
        const match = /^data:([^;,]+)/.exec(value.trim());
        return { bytes, mime: match ? match[1] : mimeOf(value, 'application/octet-stream') };
    }
    if (typeof value?.arrayBuffer === 'function') {
        const bytes = new Uint8Array(await value.arrayBuffer());
        return { bytes, mime: value.type || 'application/octet-stream' };
    }
    return null;
}
function startsWith(bytes, ...prefixes) {
    return prefixes.some((prefix) => prefix.every((b, i) => bytes[i] === b));
}
function sniffMime(bytes) {
    if (bytes.length >= 8 &&
        bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)
        return 'image/png';
    if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff)
        return 'image/jpeg';
    if (bytes.length >= 6 && startsWith(bytes, [0x47, 0x49, 0x46, 0x38, 0x37, 0x61], [0x47, 0x49, 0x46, 0x38, 0x39, 0x61])) {
        return 'image/gif';
    }
    if (bytes.length >= 12 && startsWith(bytes, [0x52, 0x49, 0x46, 0x46]) &&
        bytes[8] === 0x57 && bytes[9] === 0x45 && bytes[10] === 0x42 && bytes[11] === 0x50)
        return 'image/webp';
    if (bytes.length >= 2 && bytes[0] === 0x42 && bytes[1] === 0x4d)
        return 'image/bmp';
    if (bytes.length >= 5 &&
        bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46 && bytes[4] === 0x2d) {
        return 'application/pdf';
    }
    return 'application/octet-stream';
}

function isImageElement(element) {
    if (!element || typeof element !== 'object')
        return false;
    const record = element;
    return typeof record.data === 'string' && record.data.startsWith('data:');
}
function extractPptxAssets(document, mode = 'both') {
    if (mode === 'ignore')
        return [];
    const assets = [];
    const slides = document?.slides;
    if (!Array.isArray(slides))
        return assets;
    const walk = (node, slideIndex) => {
        if (!node || typeof node !== 'object')
            return;
        if (Array.isArray(node)) {
            for (const item of node)
                walk(item, slideIndex);
            return;
        }
        if (isImageElement(node)) {
            const parsed = parseDataUrl(node.data);
            if (parsed) {
                const id = createAssetId('pptx', assets.length + 1);
                const alt = node.alt;
                const asset = createAsset({
                    id,
                    data: parsed.bytes,
                    mime: parsed.mime,
                    part: `ppt/media/${id}${extOfMime(parsed.mime)}`,
                    anchor: {
                        scope: 'pptx',
                        kind: 'slide',
                        index: slideIndex + 1,
                        path: `slides[${slideIndex}].elements[]`,
                        label: `第 ${slideIndex + 1} 张幻灯片`
                    },
                    alt: typeof alt === 'string' ? alt : undefined
                });
                assets.push(asset);
                const element = node;
                element.assetId = id;
                if (mode === 'extract')
                    delete element.data;
            }
            return;
        }
        for (const value of Object.values(node)) {
            walk(value, slideIndex);
        }
    };
    slides.forEach((slide, index) => walk(slide, index));
    return assets;
}
function extOfMime(mime) {
    const map = {
        'image/png': '.png',
        'image/jpeg': '.jpg',
        'image/gif': '.gif',
        'image/webp': '.webp',
        'image/svg+xml': '.svg'
    };
    return map[mime] || '.bin';
}

function normalizeAssets$1(assets) {
    if (!assets)
        return {};
    if (Array.isArray(assets)) {
        const table = {};
        for (const asset of assets)
            table[asset.id] = asset;
        return table;
    }
    return assets;
}
function resolvePptxAssetRefs(document, options) {
    const messages = [];
    const table = normalizeAssets$1(options?.assets ?? document?.assets);
    if (!Object.keys(table).length)
        return { document, messages };
    const walk = (node) => {
        if (!node || typeof node !== 'object')
            return;
        if (Array.isArray(node)) {
            for (const item of node)
                walk(item);
            return;
        }
        const record = node;
        if (typeof record.assetId === 'string' && record.data === undefined) {
            const asset = getAsset(table, record.assetId);
            if (asset) {
                record.data = assetToDataUrl(asset);
            }
            else {
                messages.push(`资源引用无法解析：${record.assetId}`);
            }
        }
        for (const field of ['data', 'src']) {
            const value = record[field];
            if (isAssetRef(value)) {
                const ref = value;
                const asset = ref.asset ?? getAsset(table, ref.assetId);
                if (asset) {
                    record[field] = assetToDataUrl(asset);
                }
                else {
                    messages.push(`资源引用无法解析：${ref.assetId || '(缺少 assetId)'}`);
                    delete record[field];
                }
            }
        }
        if (typeof record.assetId === 'string' && typeof record.data === 'string') {
            delete record.assetId;
        }
        for (const value of Object.values(record))
            walk(value);
    };
    walk(document);
    return { document, messages };
}

const TYPE_MAP = {
    barChart: 'column',
    lineChart: 'line',
    areaChart: 'area',
    pieChart: 'pie',
    pie3DChart: 'pie',
    ofPieChart: 'pie',
    doughnutChart: 'doughnut',
    scatterChart: 'scatter',
    bubbleChart: 'scatter',
    radarChart: 'radar',
    line3DChart: 'line',
    area3DChart: 'area',
    bar3DChart: 'column',
    stockChart: 'line',
    surfaceChart: 'line',
    surface3DChart: 'line'
};
function seriesValues(values) {
    const out = [];
    const xs = [];
    let hasX = false;
    for (const v of values || []) {
        if (v === null || v === undefined) {
            out.push(null);
            xs.push(null);
            continue;
        }
        if (typeof v === 'number') {
            out.push(v);
            xs.push(null);
            continue;
        }
        const pt = v;
        if (typeof pt.x === 'number') {
            hasX = true;
            xs.push(pt.x);
        }
        else {
            xs.push(null);
        }
        out.push(typeof pt.y === 'number' ? pt.y : null);
    }
    return hasX ? { values: out, xValues: xs } : { values: out };
}
function pptxChartToSpec(el, theme) {
    const raw = String(el?.chartType || 'barChart');
    const [base, variant] = raw.split('|');
    let type = TYPE_MAP[base] || 'column';
    if (type === 'column' && variant === 'bar')
        type = 'bar';
    const series = (el?.series || []).map((s) => {
        const { values, xValues } = seriesValues(s.values);
        const out = { name: s.name, values };
        if (xValues)
            out.xValues = xValues;
        if (s.color)
            out.color = s.color;
        if (s.pointColors?.length)
            out.pointColors = s.pointColors;
        return out;
    });
    const spec = {
        type,
        categories: el?.categories || [],
        series
    };
    if (el?.title)
        spec.title = el.title;
    if (el?.legend) {
        const pos = String(el.legendPosition || 'r').toLowerCase();
        spec.legend = ['l', 'r', 't', 'b'].includes(pos)
            ? pos
            : 'r';
    }
    if (el?.dataLabels)
        spec.labels = { value: true };
    const grouping = String(el?.grouping || '').toLowerCase();
    if (grouping === 'stacked')
        spec.grouping = 'stacked';
    else if (grouping === 'percentstacked')
        spec.grouping = 'percentStacked';
    else if (grouping === 'clustered')
        spec.grouping = 'clustered';
    if (el?.holeSize !== undefined)
        spec.holeSize = el.holeSize;
    if (el?.smooth !== undefined)
        spec.smooth = el.smooth;
    if (el?.marker !== undefined)
        spec.marker = el.marker;
    if (el?.varyColors !== undefined)
        spec.varyColors = el.varyColors;
    if (theme?.palette?.length)
        spec.palette = theme.palette;
    if (theme?.textColor)
        spec.textColor = theme.textColor;
    if (theme?.titleColor)
        spec.titleColor = theme.titleColor;
    return spec;
}

function toUint8(result) {
    if (result instanceof Uint8Array)
        return result;
    if (typeof result === 'string') {
        const binary = globalThis.atob
            ? globalThis.atob(result)
            : Buffer.from(result, 'base64').toString('binary');
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++)
            bytes[i] = binary.charCodeAt(i) & 0xff;
        return bytes;
    }
    if (result instanceof ArrayBuffer)
        return new Uint8Array(result);
    if (Array.isArray(result))
        return new Uint8Array(result);
    throw new TypeError('PPTX 生成结果类型无法识别');
}
async function pptxToHtml(input, options = {}) {
    try {
        return await pptParser.pptxToHtml(await toArrayBuffer(input), options);
    }
    catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}
async function pptxToJson(input, options = {}) {
    try {
        return await pptParser.pptxToJson(await toArrayBuffer(input), options);
    }
    catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}
async function pptxToStandard(input, options = {}) {
    try {
        const { assets: assetOption, ...parserOptions } = options;
        const document = await pptParser.pptxToStandard(await toArrayBuffer(input), parserOptions);
        if (assetOption === 'ignore')
            return document;
        const assets = extractPptxAssets(document, assetOption ?? 'both');
        if (assets.length)
            document.assets = assets;
        return document;
    }
    catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}
async function pptxToFiles(input) {
    try {
        return await pptParser.pptxToFiles(await toArrayBuffer(input));
    }
    catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}
async function jsonToPptx(presentation, options = {}) {
    try {
        const { assets, ...rest } = options;
        const { document } = resolvePptxAssetRefs(presentation, {
            assets: assets ?? presentation?.assets
        });
        const result = await pptParser.jsonToPptx(document, { ...rest, outputType: 'uint8array' });
        return toUint8(result);
    }
    catch (error) {
        throw wrapError('pptx', '生成', error);
    }
}
const PPTXComposer = pptParser.PPTXComposer;
async function editPptx(input) {
    try {
        return await pptParser.editPptx(await toArrayBuffer(input));
    }
    catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}

let fontkitPromise;
async function registerPdfFontkit() {
    if (!fontkitPromise) {
        fontkitPromise = import('@pdf-lib/fontkit')
            .then((mod) => mod.default ?? mod)
            .catch((error) => {
            fontkitPromise = undefined;
            throw new ConvertError('嵌入自定义字体需要 @pdf-lib/fontkit，请先安装：npm i @pdf-lib/fontkit', 'pdf', { cause: error });
        });
    }
    return fontkitPromise;
}
const SYSTEM_FONT_CANDIDATES = {
    darwin: [
        '/System/Library/Fonts/Supplemental/Arial Unicode.ttf',
        '/Library/Fonts/Arial Unicode.ttf',
        '/System/Library/Fonts/Supplemental/Songti.ttc',
        '/System/Library/Fonts/STHeiti Light.ttc',
        '/System/Library/Fonts/Hiragino Sans GB.ttc',
        '/System/Library/Fonts/PingFang.ttc'
    ],
    win32: [
        'C:\\Windows\\Fonts\\simhei.ttf',
        'C:\\Windows\\Fonts\\Deng.ttf',
        'C:\\Windows\\Fonts\\msyh.ttc',
        'C:\\Windows\\Fonts\\simsun.ttc'
    ],
    linux: [
        '/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf',
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
        '/usr/share/fonts/truetype/arphic/uming.ttc'
    ]
};
function defaultCandidates(options) {
    const platform = globalThis.process?.platform || '';
    const list = [];
    if (platform === 'darwin')
        list.push(...(SYSTEM_FONT_CANDIDATES.darwin || []));
    else if (platform === 'win32') {
        list.push(...(SYSTEM_FONT_CANDIDATES.win32 || []));
        if (options.includeUserFonts !== false)
            list.push(`${process.env.LOCALAPPDATA || 'C:\\Users\\Default'}\\Microsoft\\Windows\\Fonts\\msyh.ttc`);
    }
    else
        list.push(...(SYSTEM_FONT_CANDIDATES.linux || []));
    return [...(options.candidates || []), ...list];
}
const detectCache = new Map();
async function detectSystemFonts(options = {}) {
    if (!globalThis.process?.versions?.node) {
        throw new ConvertError('系统字体探测仅支持 Node.js（浏览器端请自行传入 fonts）', 'pdf');
    }
    const key = JSON.stringify([options.candidates || [], options.includeUserFonts !== false]);
    if (detectCache.has(key))
        return detectCache.get(key);
    const fs = (await import('node:fs/promises'));
    const result = await findFirstUsableFont(defaultCandidates(options), fs);
    detectCache.set(key, result);
    return result;
}
async function findFirstUsableFont(candidates, fs) {
    const fontkit = await registerPdfFontkit();
    const { PDFDocument } = await import('pdf-lib');
    for (const path of candidates) {
        let bytes;
        try {
            const stat = await fs.stat(path);
            if (!stat.isFile() || stat.size === 0)
                continue;
            bytes = new Uint8Array(await fs.readFile(path));
        }
        catch {
            continue;
        }
        try {
            const probe = await PDFDocument.create();
            probe.registerFontkit(fontkit);
            await probe.embedFont(bytes, { subset: true });
            return { regular: bytes };
        }
        catch {
        }
    }
    return undefined;
}

const PAGE_SIZES$1 = {
    A3: [841.89, 1190.55],
    A4: [595.28, 841.89],
    A5: [419.53, 595.28],
    Letter: [612, 792],
    Legal: [612, 1008]
};
const NAMED_COLORS = {
    black: '#000000',
    white: '#FFFFFF',
    red: '#FF0000',
    green: '#008000',
    blue: '#0000FF',
    gray: '#808080',
    grey: '#808080',
    silver: '#C0C0C0',
    yellow: '#FFFF00',
    orange: '#FFA500',
    purple: '#800080',
    navy: '#000080',
    teal: '#008080',
    maroon: '#800000',
    olive: '#808000',
    lime: '#00FF00',
    aqua: '#00FFFF',
    cyan: '#00FFFF',
    fuchsia: '#FF00FF',
    magenta: '#FF00FF'
};
function parseColor(value, fallback = pdfLib.rgb(0, 0, 0)) {
    if (!value)
        return fallback;
    let hex = value.trim();
    const named = NAMED_COLORS[hex.toLowerCase()];
    if (named)
        hex = named;
    if (hex.startsWith('#'))
        hex = hex.slice(1);
    if (hex.length === 3)
        hex = hex.split('').map((ch) => ch + ch).join('');
    if (!/^[0-9a-fA-F]{6}$/.test(hex))
        return fallback;
    return pdfLib.rgb(parseInt(hex.slice(0, 2), 16) / 255, parseInt(hex.slice(2, 4), 16) / 255, parseInt(hex.slice(4, 6), 16) / 255);
}
function resolvePageSize(size, landscape) {
    let width;
    let height;
    if (Array.isArray(size)) {
        [width, height] = size;
    }
    else if (size && typeof size === 'object') {
        width = size.width;
        height = size.height;
    }
    else {
        const key = typeof size === 'string' ? size : 'A4';
        [width, height] = PAGE_SIZES$1[key] || PAGE_SIZES$1.A4;
    }
    return landscape ? [height, width] : [width, height];
}
function resolveMargin(margin) {
    const DEFAULT = 56;
    if (typeof margin === 'number') {
        return { top: margin, right: margin, bottom: margin, left: margin };
    }
    if (margin && typeof margin === 'object') {
        return {
            top: margin.top ?? DEFAULT,
            right: margin.right ?? DEFAULT,
            bottom: margin.bottom ?? DEFAULT,
            left: margin.left ?? DEFAULT
        };
    }
    return { top: DEFAULT, right: DEFAULT, bottom: DEFAULT, left: DEFAULT };
}
const STANDARD_FAMILIES = [
    { key: 'normal', constant: pdfLib.StandardFonts.Helvetica },
    { key: 'bold', constant: pdfLib.StandardFonts.HelveticaBold },
    { key: 'italic', constant: pdfLib.StandardFonts.HelveticaOblique },
    { key: 'boldItalic', constant: pdfLib.StandardFonts.HelveticaBoldOblique }
];
async function toFontBytes(source) {
    if (typeof source !== 'string')
        return toUint8Array(source);
    const match = /^data:[^;,]*(;base64)?,(.*)$/s.exec(source.trim());
    const payload = match ? match[2] : source;
    const clean = payload.replace(/\s+/g, '');
    const atobFn = globalThis.atob;
    const binary = atobFn ? atobFn(clean) : Buffer.from(clean, 'base64').toString('binary');
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++)
        bytes[i] = binary.charCodeAt(i) & 0xff;
    return bytes;
}
async function buildFonts(doc, fonts) {
    if (fonts && (fonts.regular || fonts.bold)) {
        doc.registerFontkit((await registerPdfFontkit()));
        const embed = async (source, fallback) => {
            const bytes = source ? await toFontBytes(source) : fallback ? await toFontBytes(fallback) : undefined;
            if (!bytes)
                return undefined;
            const magic = String.fromCharCode(bytes[0] ?? 0, bytes[1] ?? 0, bytes[2] ?? 0, bytes[3] ?? 0);
            if (magic === 'wOF2') {
                throw new Error('PDF 字体不支持 WOFF2 格式：请先解码为 TTF/OTF 再作为 fonts 传入（WOFF2 子集化后文字不可见）');
            }
            if (bytes[0] === 0x4c && bytes[1] === 0x0a) {
                throw new Error('PDF 字体不支持 EOT 格式：请使用 TTF/OTF 字体');
            }
            return doc.embedFont(bytes, { subset: true });
        };
        const regular = await embed(fonts.regular);
        const normal = regular || (await doc.embedFont(pdfLib.StandardFonts.Helvetica));
        const bold = (await embed(fonts.bold, typeof fonts.regular === 'string' ? fonts.regular : undefined)) || normal;
        const italic = (await embed(fonts.italic, typeof fonts.regular === 'string' ? fonts.regular : undefined)) || normal;
        const boldItalic = (await embed(fonts.boldItalic, typeof fonts.bold === 'string' ? fonts.bold : undefined)) || bold;
        return { normal, bold, italic, boldItalic, embedded: true };
    }
    const normal = await doc.embedFont(pdfLib.StandardFonts.Helvetica);
    const bold = await doc.embedFont(pdfLib.StandardFonts.HelveticaBold);
    const italic = await doc.embedFont(pdfLib.StandardFonts.HelveticaOblique);
    const boldItalic = await doc.embedFont(pdfLib.StandardFonts.HelveticaBoldOblique);
    return { normal, bold, italic, boldItalic, embedded: false };
}
function hasCjk$1(text) {
    return /[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF]/.test(text);
}
function wrapText(text, font, size, maxWidth) {
    if (!text)
        return [];
    if (!maxWidth || maxWidth <= 0)
        return text.split('\n');
    const lines = [];
    for (const paragraph of String(text).split('\n')) {
        if (!paragraph) {
            lines.push('');
            continue;
        }
        let line = '';
        for (const char of Array.from(paragraph)) {
            const candidate = line + char;
            let width;
            try {
                width = font.widthOfTextAtSize(candidate, size);
            }
            catch {
                width = candidate.length * size * 0.5;
            }
            if (width <= maxWidth || !line) {
                line = candidate;
                continue;
            }
            const lastSpace = line.lastIndexOf(' ');
            if (lastSpace > 0 && /[\x20-\x7E]/.test(char)) {
                lines.push(line.slice(0, lastSpace));
                line = line.slice(lastSpace + 1) + char;
            }
            else {
                lines.push(line);
                line = char;
            }
        }
        if (line)
            lines.push(line);
    }
    return lines;
}
function detectImageType(bytes) {
    if (bytes.length > 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)
        return 'png';
    if (bytes.length > 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff)
        return 'jpg';
    return undefined;
}
async function readImageBytes(src, assets) {
    if (isAssetRef(src)) {
        const ref = src;
        const asset = ref.asset ?? getAsset(assets, ref.assetId);
        if (!asset)
            throw new TypeError('图片引用无法解析：未提供对应的 assets');
        return asset.data;
    }
    return toUint8Array(src);
}
function currentPage(ctx) {
    return ctx.flow ? ctx.flow.page : ctx.page;
}
function topCursor(ctx, elementY = 0) {
    return ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : ctx.originY + elementY;
}
function drawTextStyled(ctx, text, options) {
    const page = currentPage(ctx);
    if (options.bold && ctx.fonts.bold === ctx.fonts.normal) {
        const dx = options.size * 0.042;
        page.drawText(text, {
            x: options.x + dx,
            y: options.y,
            size: options.size,
            font: options.font,
            color: options.color
        });
    }
    page.drawText(text, {
        x: options.x,
        y: options.y,
        size: options.size,
        font: options.font,
        color: options.color
    });
}
function drawTextElement(ctx, element) {
    const isHeading = element.type === 'heading';
    const size = element.size ?? (isHeading ? Math.max(10, 24 - (element.level - 1) * 2.6) : ctx.defaultSize);
    const bold = isHeading || (('bold' in element ? element.bold : false) ?? false);
    const italic = !isHeading && (('italic' in element ? element.italic : false) ?? false);
    const font = bold && italic
        ? ctx.fonts.boldItalic
        : bold
            ? ctx.fonts.bold
            : italic
                ? ctx.fonts.italic
                : ctx.fonts.normal;
    const lineHeight = (element.lineHeight ?? 1.35) * size;
    const width = element.width ?? ctx.contentWidth;
    const lines = wrapText(element.text ?? '', font, size, width);
    const startTop = topCursor(ctx, element.y ?? 0);
    let y = startTop;
    if (ctx.flow && isHeading)
        ctx.flow.ensure(lines.length * lineHeight);
    let localY = y;
    for (const line of lines) {
        let x = ctx.originX + (element.x ?? 0);
        if (element.align === 'center' || element.align === 'right') {
            let lineWidth;
            try {
                lineWidth = font.widthOfTextAtSize(line, size);
            }
            catch {
                lineWidth = line.length * size * 0.5;
            }
            const offset = width - lineWidth;
            x += element.align === 'center' ? offset / 2 : Math.max(0, offset);
        }
        if (ctx.flow) {
            ctx.flow.ensure(lineHeight);
            localY = ctx.flow.marginTop + ctx.flow.cursorY;
        }
        drawTextStyled(ctx, line, {
            x,
            y: ctx.pageHeight - localY - size,
            size,
            font,
            bold,
            color: parseColor(element.color, ctx.defaultColor)
        });
        if (ctx.flow)
            ctx.flow.cursorY += lineHeight;
        localY += lineHeight;
    }
    return (ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : localY) - startTop;
}
function fontForRun(ctx, run, size) {
    const bold = run.bold ?? false;
    const italic = run.italic ?? false;
    void size;
    return bold && italic ? ctx.fonts.boldItalic : bold ? ctx.fonts.bold : italic ? ctx.fonts.italic : ctx.fonts.normal;
}
function measureRun(ctx, run, text, size) {
    try {
        return fontForRun(ctx, run, size).widthOfTextAtSize(text, size);
    }
    catch {
        return text.length * size * 0.5;
    }
}
function drawRunsElement(ctx, element) {
    const size = element.size ?? ctx.defaultSize;
    const indentLeft = element.indentLeft ?? 0;
    const indentRight = element.indentRight ?? 0;
    const firstIndent = element.indentFirstLine ?? 0;
    const width = element.width ?? ctx.contentWidth - indentLeft - indentRight;
    const baseX = ctx.originX + (element.x ?? 0) + indentLeft;
    const pieces = [];
    const isCjk = (ch) => /[　-〿㐀-鿿豈-﫿＀-￯]/.test(ch);
    const isLatin = (ch) => /[A-Za-z0-9]/.test(ch);
    for (const run of element.runs) {
        if (!run.text)
            continue;
        let word = '';
        let wordRun;
        const flushWord = () => {
            if (word) {
                pieces.push({ text: word, run: wordRun });
                word = '';
            }
        };
        for (const ch of Array.from(run.text)) {
            if (ch === '\n') {
                flushWord();
                pieces.push({ text: '\n', run });
                continue;
            }
            if (isCjk(ch) || !/[\x00-\x7F]/.test(ch)) {
                flushWord();
                pieces.push({ text: ch, run });
            }
            else if (ch === '\t') {
                flushWord();
                pieces.push({ text: '\t', run });
            }
            else if (ch === ' ') {
                flushWord();
                pieces.push({ text: ' ', run });
            }
            else if (isLatin(ch)) {
                if (wordRun && wordRun !== run)
                    flushWord();
                word += ch;
                wordRun = run;
            }
            else {
                if (wordRun && wordRun !== run)
                    flushWord();
                word += ch;
                wordRun = run;
            }
        }
        flushWord();
    }
    const lines = [];
    let line = [];
    let lineWidth = 0;
    let lineFirst = true;
    let lineSize = size;
    const pushLine = () => {
        lines.push({ pieces: line, first: lineFirst, size: lineSize });
        line = [];
        lineWidth = 0;
        lineFirst = false;
        lineSize = size;
    };
    const tabStops = element.tabStops?.length
        ? [...element.tabStops].sort((a, b) => (a.pos ?? Infinity) - (b.pos ?? Infinity))
        : undefined;
    for (const piece of pieces) {
        if (piece.text === '\n') {
            pushLine();
            continue;
        }
        const lineAvail = width - (lineFirst ? firstIndent : 0);
        const w = piece.text === '\t' ? (tabStops ? 0 : size * 0.5) : measureRun(ctx, piece.run, piece.text, piece.run.size ?? size);
        if (lineWidth + w > lineAvail && line.length && piece.text !== '\t') {
            pushLine();
        }
        line.push(piece);
        lineWidth += w;
        lineSize = Math.max(lineSize, piece.run.size ?? size);
    }
    pushLine();
    const drawLine = (l, localY) => {
        const lineTextWidth = l.pieces.reduce((sum, p) => (p.text === '\t' ? sum : sum + measureRun(ctx, p.run, p.text, p.run.size ?? size)), 0);
        let x = baseX + (l.first ? firstIndent : 0);
        if (element.align === 'center' || element.align === 'right') {
            const offset = width - lineTextWidth;
            x += element.align === 'center' ? offset / 2 : Math.max(0, offset);
        }
        const segWidths = new Map();
        if (tabStops) {
            for (let i = 0; i < l.pieces.length; i++) {
                if (l.pieces[i].text !== '\t')
                    continue;
                let sum = 0;
                for (let j = i + 1; j < l.pieces.length && l.pieces[j].text !== '\t'; j++) {
                    sum += measureRun(ctx, l.pieces[j].run, l.pieces[j].text, l.pieces[j].run.size ?? size);
                }
                segWidths.set(i, sum);
            }
        }
        for (let i = 0; i < l.pieces.length; i++) {
            const p = l.pieces[i];
            if (p.text === '\t') {
                if (!tabStops) {
                    x += size * 0.5;
                    continue;
                }
                const stop = tabStops.find((s) => (s.pos ?? width) > x - baseX + 0.5);
                if (!stop)
                    continue;
                const stopPos = stop.pos ?? width;
                const segW = segWidths.get(i) ?? 0;
                const segSize = p.run.size ?? size;
                const segColor = p.run.color ? parseColor(p.run.color, ctx.defaultColor) : ctx.defaultColor;
                let contentX = baseX + stopPos;
                if (stop.align === 'right')
                    contentX -= segW;
                else if (stop.align === 'center')
                    contentX -= segW / 2;
                if (stop.leader) {
                    drawLeader(ctx, x, contentX, localY, segSize, stop.leader, segColor);
                }
                x = Math.max(x, contentX);
                continue;
            }
            const rs = p.run.size ?? size;
            const w = measureRun(ctx, p.run, p.text, rs);
            const color = p.run.color ? parseColor(p.run.color, ctx.defaultColor) : ctx.defaultColor;
            if (p.run.code || p.run.link) {
                const bg = p.run.code ? parseColor('#F2F2F2') : undefined;
                if (bg) {
                    currentPage(ctx).drawRectangle({
                        x: x - 1,
                        y: ctx.pageHeight - localY - rs,
                        width: w + 2,
                        height: rs + 2,
                        color: bg
                    });
                }
                drawTextStyled(ctx, p.text, {
                    x,
                    y: ctx.pageHeight - localY - rs,
                    size: rs,
                    font: fontForRun(ctx, p.run, rs),
                    bold: p.run.bold,
                    color: p.run.link ? parseColor('#1565C0', ctx.defaultColor) : color
                });
                if (p.run.link) {
                    currentPage(ctx).drawLine({
                        start: { x, y: ctx.pageHeight - localY - 2 },
                        end: { x: x + w, y: ctx.pageHeight - localY - 2 },
                        thickness: 0.5,
                        color: parseColor('#1565C0', ctx.defaultColor)
                    });
                }
            }
            else {
                drawTextStyled(ctx, p.text, {
                    x,
                    y: ctx.pageHeight - localY - rs,
                    size: rs,
                    font: fontForRun(ctx, p.run, rs),
                    bold: p.run.bold,
                    color
                });
            }
            x += w;
        }
    };
    const startTop = topCursor(ctx, element.y ?? 0);
    let localY = startTop;
    const rotation = element.rotation ?? 0;
    if (rotation && element.height) {
        const rad = (rotation * Math.PI) / 180;
        const cos = Math.cos(rad);
        const sin = Math.sin(rad);
        const cx = ctx.originX + (element.x ?? 0) + indentLeft + width / 2;
        const cy = ctx.pageHeight - startTop - element.height / 2;
        ctx.page.pushOperators(pdfLib.pushGraphicsState(), pdfLib.concatTransformationMatrix(cos, sin, -sin, cos, cx - cx * cos + cy * sin, cy - cx * sin - cy * cos));
    }
    for (let i = 0; i < lines.length; i++) {
        const l = lines[i];
        const lineHeight = (element.lineHeight ?? 1.35) * l.size;
        if (ctx.flow) {
            if (i === 0 && lines.length > 1)
                ctx.flow.ensure(Math.min(2, lines.length) * lineHeight);
            if (i === lines.length - 2 && lines.length > 2)
                ctx.flow.ensure(2 * lineHeight);
            ctx.flow.ensure(lineHeight);
            localY = ctx.flow.marginTop + ctx.flow.cursorY;
        }
        drawLine(l, localY);
        if (ctx.flow)
            ctx.flow.cursorY += lineHeight;
        localY += lineHeight;
    }
    if (rotation && element.height)
        ctx.page.pushOperators(pdfLib.popGraphicsState());
    const used = (ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : localY) - startTop;
    return used + (ctx.flow ? 0 : (element.spacingAfter ?? 4));
}
function drawLeader(ctx, from, to, localY, size, leader, color) {
    if (to - from < 1)
        return;
    const page = currentPage(ctx);
    const y = ctx.pageHeight - localY - size * 0.28;
    switch (leader) {
        case 'dot':
            for (let px = from + 0.5; px < to; px += 3) {
                page.drawLine({ start: { x: px, y }, end: { x: px + 0.9, y }, thickness: 0.8, color });
            }
            break;
        case 'middleDot':
            for (let px = from + 0.5; px < to; px += 3) {
                page.drawLine({ start: { x: px, y: y + 0.8 }, end: { x: px + 0.9, y: y + 0.8 }, thickness: 1.6, color });
            }
            break;
        case 'hyphen':
            for (let px = from + 0.5; px < to; px += 4) {
                page.drawLine({
                    start: { x: px, y },
                    end: { x: Math.min(to - 0.5, px + 2), y },
                    thickness: 0.8,
                    color
                });
            }
            break;
        case 'underscore':
            page.drawLine({ start: { x: from, y }, end: { x: to, y }, thickness: 0.6, color });
            break;
    }
}
function drawTableElement(ctx, element) {
    const fontSize = element.fontSize ?? 10;
    const padding = element.padding ?? 4;
    const lineH = fontSize * 1.35;
    const hasHeader = element.header !== false && element.rows.length > 0;
    if (!element.rows.length)
        return 0;
    const isCellObject = (c) => !!c && typeof c === 'object' && !Array.isArray(c) && 'content' in c;
    const normCell = (cell) => {
        if (isCellObject(cell)) {
            const content = cell.content;
            return {
                text: typeof content === 'string' ? content : cellToPlainText(content),
                colspan: Math.max(1, Math.round(cell.colspan ?? 1)),
                rowspan: Math.max(1, Math.round(cell.rowspan ?? 1)),
                width: cell.width,
                fill: cell.fill,
                borders: cell.borders,
                verticalAlign: cell.verticalAlign
            };
        }
        const text = tableCellText(cell);
        return { text, colspan: 1, rowspan: 1 };
    };
    const rows = element.rows.map((row) => row.map(normCell));
    const placements = [];
    const occupied = [];
    let gridCols = 0;
    rows.forEach((row, r) => {
        if (!occupied[r])
            occupied[r] = [];
        let col = 0;
        for (const cell of row) {
            while (occupied[r][col])
                col++;
            for (let rr = r; rr < r + cell.rowspan && rr < rows.length; rr++) {
                if (!occupied[rr])
                    occupied[rr] = [];
                for (let cc = col; cc < col + cell.colspan; cc++)
                    occupied[rr][cc] = true;
            }
            placements.push({ row: r, col, cell });
            col += cell.colspan;
        }
        gridCols = Math.max(gridCols, col);
    });
    if (!gridCols)
        return 0;
    const widths = element.columnWidths?.length
        ? element.columnWidths
        : new Array(gridCols).fill(ctx.contentWidth / gridCols);
    const colOffset = (col) => {
        let x = 0;
        for (let i = 0; i < col && i < widths.length; i++)
            x += widths[i];
        return x;
    };
    const cellWidth = (cell, col) => {
        if (cell.width !== undefined && cell.width > 0)
            return cell.width;
        let w = 0;
        for (let i = col; i < col + cell.colspan; i++)
            w += widths[i] ?? widths[widths.length - 1] ?? 0;
        if (w <= 0)
            w = widths[widths.length - 1] ?? ctx.contentWidth / gridCols;
        return w;
    };
    const borderColor = element.borderColor === 'none' ? undefined : parseColor(element.borderColor || '#CCCCCC');
    const headerFill = element.headerFill ? parseColor(element.headerFill) : undefined;
    const startTop = topCursor(ctx, element.y ?? 0);
    const startX = ctx.originX + (element.x ?? 0);
    const cellLines = new Map();
    const rowHeights = new Array(rows.length).fill(0);
    for (const p of placements) {
        const isHeader = hasHeader && p.row === 0;
        const font = isHeader ? ctx.fonts.bold : ctx.fonts.normal;
        const lines = wrapText(p.cell.text ?? '', font, fontSize, Math.max(8, cellWidth(p.cell, p.col) - padding * 2));
        cellLines.set(p, lines);
        if (p.cell.rowspan <= 1) {
            rowHeights[p.row] = Math.max(rowHeights[p.row], lines.length * lineH + padding * 2);
        }
    }
    for (const p of placements) {
        if (p.cell.rowspan <= 1)
            continue;
        let sum = 0;
        for (let r = p.row; r < p.row + p.cell.rowspan; r++)
            sum += rowHeights[r];
        const need = (cellLines.get(p)?.length ?? 0) * lineH + padding * 2;
        if (sum < need) {
            const per = (need - sum) / p.cell.rowspan;
            for (let r = p.row; r < p.row + p.cell.rowspan; r++)
                rowHeights[r] += per;
        }
    }
    const heightOf = (r) => element.rowHeight ?? rowHeights[r] ?? 0;
    const drawRowCells = (topY, r) => {
        const page = currentPage(ctx);
        const isHeader = hasHeader && r === 0;
        for (const p of placements) {
            if (p.row !== r)
                continue;
            const c = p.cell;
            const w = cellWidth(c, p.col);
            let h = 0;
            for (let rr = r; rr < r + c.rowspan; rr++)
                h += heightOf(rr);
            const x = startX + colOffset(p.col);
            const fill = c.fill ? parseColor(c.fill) : isHeader ? headerFill : undefined;
            if (fill) {
                page.drawRectangle({
                    x,
                    y: ctx.pageHeight - topY - h,
                    width: w,
                    height: h,
                    color: fill
                });
            }
            const cellB = c.borders;
            const edgeLine = (side, x1, y1, x2, y2) => {
                const spec = cellB?.[side];
                let color = borderColor;
                let width = 0.5;
                if (spec === 'none')
                    return;
                if (spec) {
                    color = parseColor(spec.color || '#000000');
                    width = Math.max(0.3, spec.width ?? 1);
                }
                if (!color)
                    return;
                page.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, color, thickness: width });
            };
            const bottomY = ctx.pageHeight - topY - h;
            const topYpdf = ctx.pageHeight - topY;
            edgeLine('top', x, topYpdf, x + w, topYpdf);
            edgeLine('right', x + w, topYpdf, x + w, bottomY);
            edgeLine('bottom', x, bottomY, x + w, bottomY);
            edgeLine('left', x, topYpdf, x, bottomY);
            const font = isHeader ? ctx.fonts.bold : ctx.fonts.normal;
            const lines = cellLines.get(p) ?? [];
            const textBlock = lines.length * lineH;
            let offsetY = topY + (h - textBlock) / 2;
            if (c.verticalAlign === 'top')
                offsetY = topY + padding / 2;
            else if (c.verticalAlign === 'bottom')
                offsetY = topY + h - textBlock - padding / 2;
            lines.forEach((text, li) => {
                let textWidth;
                try {
                    textWidth = font.widthOfTextAtSize(text, fontSize);
                }
                catch {
                    textWidth = text.length * fontSize * 0.5;
                }
                const drawX = x + Math.max(padding, (w - textWidth) / 2);
                page.drawText(text, {
                    x: drawX,
                    y: ctx.pageHeight - offsetY - li * lineH - fontSize - (lineH - fontSize) / 2,
                    size: fontSize,
                    font,
                    color: ctx.defaultColor
                });
            });
        }
    };
    let localY = startTop;
    rows.forEach((_, rowIndex) => {
        const isHeader = hasHeader && rowIndex === 0;
        const rowHeight = heightOf(rowIndex);
        if (ctx.flow) {
            const bind = isHeader && rows.length > 1 ? rowHeight + heightOf(1) : rowHeight;
            const moved = ctx.flow.ensure(bind);
            if (moved && hasHeader && rowIndex > 0) {
                drawRowCells(ctx.flow.marginTop + ctx.flow.cursorY, 0);
                ctx.flow.cursorY += heightOf(0);
            }
            localY = ctx.flow.marginTop + ctx.flow.cursorY;
        }
        drawRowCells(localY, rowIndex);
        if (ctx.flow)
            ctx.flow.cursorY += rowHeight;
        localY += rowHeight;
    });
    return (ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : localY) - startTop;
}
function cellToPlainText(cell) {
    return cell
        .map((element) => {
        if (element.type === 'text' || element.type === 'heading')
            return element.text ?? '';
        return '';
    })
        .join(' ')
        .trim();
}
function tableCellText(cell) {
    if (typeof cell === 'string')
        return cell;
    if (Array.isArray(cell))
        return cellToPlainText(cell);
    return typeof cell.content === 'string' ? cell.content : cellToPlainText(cell.content);
}
function spacingBeforeOf(element) {
    const el = element;
    return el.spacingBefore ?? 0;
}
function spacingAfterOf(element) {
    if (element.type === 'runs')
        return element.spacingAfter ?? 4;
    const el = element;
    return el.spacingAfter ?? 0;
}
const BEZIER_K$1 = 0.5523;
function roundedRectPath(x, y, w, h, r) {
    const k = BEZIER_K$1 * r;
    return [
        { op: 'm', args: [x + r, y] },
        { op: 'l', args: [x + w - r, y] },
        { op: 'c', args: [x + w - r + k, y, x + w, y + r - k, x + w, y + r] },
        { op: 'l', args: [x + w, y + h - r] },
        { op: 'c', args: [x + w, y + h - r + k, x + w - r + k, y + h, x + w - r, y + h] },
        { op: 'l', args: [x + r, y + h] },
        { op: 'c', args: [x + r - k, y + h, x, y + h - r + k, x, y + h - r] },
        { op: 'l', args: [x, y + r] },
        { op: 'c', args: [x, y + r - k, x + r - k, y, x + r, y] },
        { op: 'h', args: [] }
    ];
}
function ellipsePath(x, y, w, h) {
    const rx = w / 2;
    const ry = h / 2;
    const cx = x + rx;
    const cy = y + ry;
    const kx = BEZIER_K$1 * rx;
    const ky = BEZIER_K$1 * ry;
    return [
        { op: 'm', args: [cx + rx, cy] },
        { op: 'c', args: [cx + rx, cy + ky, cx + kx, cy + ry, cx, cy + ry] },
        { op: 'c', args: [cx - kx, cy + ry, cx - rx, cy + ky, cx - rx, cy] },
        { op: 'c', args: [cx - rx, cy - ky, cx - kx, cy - ry, cx, cy - ry] },
        { op: 'c', args: [cx + kx, cy - ry, cx + rx, cy - ky, cx + rx, cy] },
        { op: 'h', args: [] }
    ];
}
function clipPathOperators(ctx, segments) {
    const ops = [];
    const px = (v) => ctx.originX + v;
    const py = (v) => ctx.pageHeight - ctx.originY - v;
    for (const seg of segments) {
        if (seg.op === 'm')
            ops.push(pdfLib.moveTo(px(seg.args[0]), py(seg.args[1])));
        else if (seg.op === 'l')
            ops.push(pdfLib.lineTo(px(seg.args[0]), py(seg.args[1])));
        else if (seg.op === 'c')
            ops.push(pdfLib.appendBezierCurve(px(seg.args[0]), py(seg.args[1]), px(seg.args[2]), py(seg.args[3]), px(seg.args[4]), py(seg.args[5])));
        else
            ops.push(pdfLib.closePath());
    }
    return ops;
}
function pathSegmentOperators(ctx, segments) {
    const ops = [];
    const px = (v) => ctx.originX + v;
    const py = (v) => ctx.pageHeight - ctx.originY - v;
    for (const seg of segments) {
        if (seg.op === 'm')
            ops.push(pdfLib.moveTo(px(seg.args[0]), py(seg.args[1])));
        else if (seg.op === 'l')
            ops.push(pdfLib.lineTo(px(seg.args[0]), py(seg.args[1])));
        else if (seg.op === 'c')
            ops.push(pdfLib.appendBezierCurve(px(seg.args[0]), py(seg.args[1]), px(seg.args[2]), py(seg.args[3]), px(seg.args[4]), py(seg.args[5])));
        else
            ops.push(pdfLib.closePath());
    }
    return ops;
}
function drawBezierShape(ctx, options, segments) {
    const ops = [];
    for (const seg of segments) {
        if (seg.op === 'm')
            ops.push(pdfLib.moveTo(seg.args[0], seg.args[1]));
        else if (seg.op === 'l')
            ops.push(pdfLib.lineTo(seg.args[0], seg.args[1]));
        else if (seg.op === 'c')
            ops.push(pdfLib.appendBezierCurve(seg.args[0], seg.args[1], seg.args[2], seg.args[3], seg.args[4], seg.args[5]));
        else
            ops.push(pdfLib.closePath());
    }
    const alpha = options.opacity !== undefined ? Math.max(0, Math.min(1, options.opacity)) : 1;
    const hasFill = !!options.color;
    const hasStroke = !!options.borderColor;
    if (alpha < 1) {
        const gs = ctx.doc.context.obj({ Type: 'ExtGState', ca: alpha, CA: alpha });
        const key = ctx.page.node.newExtGState('GS', gs);
        ops.push(pdfLib.pushGraphicsState(), pdfLib.setGraphicsState(key));
    }
    if (options.color)
        ops.push(pdfLib.setFillingColor(options.color));
    if (options.borderColor) {
        ops.push(pdfLib.setStrokingColor(options.borderColor));
        ops.push(pdfLib.setLineWidth(options.borderWidth ?? 1));
    }
    if (hasFill && hasStroke)
        ops.push(pdfLib.fillAndStroke());
    else if (hasFill)
        ops.push(pdfLib.fill());
    else if (hasStroke)
        ops.push(pdfLib.stroke());
    if (alpha < 1)
        ops.push(pdfLib.popGraphicsState());
    ctx.page.pushOperators(...ops);
}
function drawArrowHead(ctx, tipX, tipY, refX, refY, thickness, color) {
    const dx = tipX - refX;
    const dy = tipY - refY;
    const len = Math.hypot(dx, dy);
    if (len < 0.01)
        return;
    const ux = dx / len;
    const uy = dy / len;
    const px = -uy;
    const py = ux;
    const l = thickness * 5;
    const halfH = thickness * 2.5;
    const baseX = tipX - ux * l;
    const baseY = tipY - uy * l;
    ctx.page.pushOperators(pdfLib.pushGraphicsState(), pdfLib.setFillingColor(color), pdfLib.moveTo(tipX, tipY), pdfLib.lineTo(baseX + px * halfH, baseY + py * halfH), pdfLib.lineTo(baseX - px * halfH, baseY - py * halfH), pdfLib.closePath(), pdfLib.fill(), pdfLib.popGraphicsState());
}
async function drawElement(ctx, element) {
    switch (element.type) {
        case 'text':
        case 'heading':
            return drawTextElement(ctx, element);
        case 'line': {
            const thickness = element.lineWidth ?? 1;
            const sx = ctx.originX + element.x1;
            const sy = ctx.pageHeight - element.y1;
            const ex = ctx.originX + element.x2;
            const ey = ctx.pageHeight - element.y2;
            const color = parseColor(element.color, pdfLib.rgb(0.6, 0.6, 0.6));
            ctx.page.drawLine({
                start: { x: sx, y: sy },
                end: { x: ex, y: ey },
                thickness,
                color,
                dashArray: element.dash
            });
            if (element.endArrow)
                drawArrowHead(ctx, ex, ey, sx, sy, thickness, color);
            if (element.startArrow)
                drawArrowHead(ctx, sx, sy, ex, ey, thickness, color);
            return thickness;
        }
        case 'path': {
            const ops = pathSegmentOperators(ctx, element.segments);
            const alpha = element.opacity !== undefined ? Math.max(0, Math.min(1, element.opacity)) : 1;
            const hasFill = !!element.fill;
            const hasStroke = !!element.stroke;
            if (alpha < 1) {
                const gs = ctx.doc.context.obj({ Type: 'ExtGState', ca: alpha, CA: alpha });
                const gsRef = ctx.page.node.newExtGState('GSa', gs);
                ctx.page.pushOperators(pdfLib.pushGraphicsState(), pdfLib.setGraphicsState(gsRef), ...ops);
            }
            else {
                ctx.page.pushOperators(pdfLib.pushGraphicsState(), ...ops);
            }
            if (hasFill)
                ctx.page.pushOperators(pdfLib.setFillingColor(parseColor(element.fill)));
            if (hasStroke) {
                ctx.page.pushOperators(pdfLib.setStrokingColor(parseColor(element.stroke)));
                ctx.page.pushOperators(pdfLib.setLineWidth(element.lineWidth ?? 1));
                if (element.dash?.length)
                    ctx.page.pushOperators(pdfLib.setDashPattern(element.dash, 0));
            }
            if (hasFill && hasStroke)
                ctx.page.pushOperators(pdfLib.fillAndStroke());
            else if (hasFill)
                ctx.page.pushOperators(pdfLib.fill());
            else if (hasStroke)
                ctx.page.pushOperators(pdfLib.stroke());
            ctx.page.pushOperators(pdfLib.popGraphicsState());
            if (element.stroke && (element.startArrow || element.endArrow)) {
                const thickness = element.lineWidth ?? 1;
                const color = parseColor(element.stroke);
                const pts = element.segments
                    .filter((s) => s.op === 'm' || s.op === 'l')
                    .map((s) => ({ x: ctx.originX + s.args[0], y: ctx.pageHeight - ctx.originY - s.args[1] }));
                if (pts.length >= 2) {
                    if (element.endArrow) {
                        const tip = pts[pts.length - 1];
                        const ref = pts[pts.length - 2];
                        drawArrowHead(ctx, tip.x, tip.y, ref.x, ref.y, thickness, color);
                    }
                    if (element.startArrow) {
                        const tip = pts[0];
                        const ref = pts[1];
                        drawArrowHead(ctx, tip.x, tip.y, ref.x, ref.y, thickness, color);
                    }
                }
            }
            return element.lineWidth ?? 0;
        }
        case 'hr': {
            const thickness = element.lineWidth ?? 0.5;
            const before = element.spacingBefore ?? 4;
            const after = element.spacingAfter ?? 4;
            const total = before + thickness + after;
            if (ctx.flow)
                ctx.flow.ensure(total);
            const y = ctx.pageHeight - topCursor(ctx) - before;
            currentPage(ctx).drawLine({
                start: { x: ctx.originX, y },
                end: { x: ctx.originX + ctx.contentWidth, y },
                thickness,
                color: parseColor(element.color, pdfLib.rgb(0.6, 0.6, 0.6))
            });
            if (ctx.flow)
                ctx.flow.cursorY += total;
            return total;
        }
        case 'group': {
            const clipPath = element.clip;
            if (clipPath?.length) {
                ctx.page.pushOperators(pdfLib.pushGraphicsState(), ...clipPathOperators(ctx, clipPath), pdfLib.clip(), pdfLib.endPath());
            }
            for (const child of element.children) {
                await drawElement(ctx, child);
            }
            if (clipPath?.length)
                ctx.page.pushOperators(pdfLib.popGraphicsState());
            return 0;
        }
        case 'rect': {
            const y = ctx.originY + element.y;
            const options = {
                x: ctx.originX + element.x,
                y: ctx.pageHeight - y - element.height,
                width: element.width,
                height: element.height,
                color: element.fill ? parseColor(element.fill) : undefined,
                borderColor: element.stroke ? parseColor(element.stroke) : undefined,
                borderWidth: element.stroke ? element.lineWidth ?? 1 : 0,
                opacity: element.opacity !== undefined ? Math.max(0, Math.min(1, element.opacity)) : undefined
            };
            const rotation = element.rotation ?? 0;
            if (rotation) {
                const rad = (rotation * Math.PI) / 180;
                const cos = Math.cos(rad);
                const sin = Math.sin(rad);
                const cx = options.x + element.width / 2;
                const cy = options.y + element.height / 2;
                ctx.page.pushOperators(pdfLib.pushGraphicsState(), pdfLib.concatTransformationMatrix(cos, sin, -sin, cos, cx - cx * cos + cy * sin, cy - cx * sin - cy * cos));
            }
            if (element.ellipse) {
                const { pageHeight, originX, originY } = ctx;
                const px = originX + element.x;
                const py = pageHeight - originY - element.y - element.height;
                drawBezierShape(ctx, options, ellipsePath(px, py, element.width, element.height));
            }
            else if (element.radius && element.radius > 0) {
                const { pageHeight, originX, originY } = ctx;
                const px = originX + element.x;
                const py = pageHeight - originY - element.y - element.height;
                const r = Math.min(element.radius, element.width / 2, element.height / 2);
                drawBezierShape(ctx, options, roundedRectPath(px, py, element.width, element.height, r));
            }
            else {
                ctx.page.drawRectangle(options);
            }
            if (rotation)
                ctx.page.pushOperators(pdfLib.popGraphicsState());
            return element.height;
        }
        case 'image': {
            try {
                const bytes = await readImageBytes(element.src, ctx.assets);
                const kind = detectImageType(bytes);
                if (!kind)
                    throw new TypeError('仅支持 PNG / JPEG 图片');
                const image = kind === 'png' ? await ctx.doc.embedPng(bytes) : await ctx.doc.embedJpg(bytes);
                const width = element.width ?? image.width;
                const height = element.height ?? image.height;
                if (ctx.flow)
                    ctx.flow.ensure(height);
                const crop = element.crop;
                const cl = Math.max(0, crop?.l ?? 0);
                const cr = Math.max(0, crop?.r ?? 0);
                const ct = Math.max(0, crop?.t ?? 0);
                const cb = Math.max(0, crop?.b ?? 0);
                const cropped = cl + cr < 0.99 && ct + cb < 0.99;
                const drawW = cropped ? width / Math.max(0.01, 1 - cl - cr) : width;
                const drawH = cropped ? height / Math.max(0.01, 1 - ct - cb) : height;
                const drawX = (element.x ?? 0) - (cropped ? cl * drawW : 0);
                const drawTop = topCursor(ctx, element.y ?? 0) - (cropped ? ct * drawH : 0);
                if (cropped) {
                    const clipX = ctx.originX + (element.x ?? 0);
                    const clipY = ctx.pageHeight - topCursor(ctx, element.y ?? 0) - height;
                    ctx.page.pushOperators(pdfLib.pushGraphicsState(), pdfLib.rectangle(clipX, clipY, width, height), pdfLib.clip(), pdfLib.endPath());
                }
                currentPage(ctx).drawImage(image, {
                    x: ctx.originX + drawX,
                    y: ctx.pageHeight - drawTop - drawH,
                    width: drawW,
                    height: drawH
                });
                if (cropped)
                    ctx.page.pushOperators(pdfLib.popGraphicsState());
                if (ctx.flow)
                    ctx.flow.cursorY += height;
                return height;
            }
            catch {
                return 0;
            }
        }
        case 'runs':
            return drawRunsElement(ctx, element);
        case 'table':
            return drawTableElement(ctx, element);
        case 'pageBreak':
            return Number.POSITIVE_INFINITY;
        default:
            return 0;
    }
}
async function generatePdf(spec, options = {}) {
    try {
        if (!spec || typeof spec !== 'object')
            throw new TypeError('缺少文档规格');
        const doc = await pdfLib.PDFDocument.create();
        const fonts = await buildFonts(doc, { ...(spec.fonts || {}), ...(options.fonts || {}) });
        const rawAssets = options.assets ?? spec.assets;
        const assets = Array.isArray(rawAssets) ? toAssetTable(rawAssets) : (rawAssets || {});
        if (!fonts.embedded) {
            for (const element of collectElements$1(spec)) {
                const cjkText = element.type === 'text' || element.type === 'heading'
                    ? element.text ?? ''
                    : element.type === 'runs'
                        ? element.runs.map((r) => r.text).join('')
                        : element.type === 'table'
                            ? element.rows.flat().map(tableCellText).join('')
                            : '';
                if (hasCjk$1(cjkText)) {
                    throw new ConvertError('生成中文 PDF 需要内嵌字体：请通过 fonts.regular 传入中文 TTF/OTF 字节', 'pdf');
                }
            }
        }
        const defaultSize = spec.defaultSize ?? 12;
        const defaultColor = parseColor(spec.defaultColor, pdfLib.rgb(0.12, 0.12, 0.12));
        if (spec.metadata) {
            doc.setTitle(spec.metadata.title || '');
            doc.setAuthor(spec.metadata.author || '');
            doc.setSubject(spec.metadata.subject || '');
            doc.setKeywords(spec.metadata.keywords ? spec.metadata.keywords.split(/[,;\s]+/).filter(Boolean) : []);
            doc.setProducer(spec.metadata.producer || '@fefeding/ppt-parser');
            doc.setCreator(spec.metadata.creator || '@fefeding/ppt-parser');
            if (spec.metadata.created) {
                const date = new Date(spec.metadata.created);
                if (!Number.isNaN(date.getTime()))
                    doc.setCreationDate(date);
            }
            if (spec.metadata.modified) {
                const date = new Date(spec.metadata.modified);
                if (!Number.isNaN(date.getTime()))
                    doc.setModificationDate(date);
            }
        }
        const renderPage = async (elements, pageSpec) => {
            const [pageWidth, pageHeight] = resolvePageSize(pageSpec.size, pageSpec.orientation === 'landscape');
            const margin = resolveMargin(pageSpec.margin);
            const page = doc.addPage([pageWidth, pageHeight]);
            const ctx = {
                doc,
                fonts,
                assets,
                page,
                pageHeight,
                contentWidth: pageWidth - margin.left - margin.right,
                originX: margin.left,
                originY: margin.top,
                defaultSize,
                defaultColor
            };
            let cursorY = 0;
            for (const element of elements) {
                if (element.type === 'pageBreak') {
                    cursorY = Number.POSITIVE_INFINITY;
                    break;
                }
                const absolute = element.y !== undefined;
                const used = await drawElement({ ...ctx, originY: absolute ? margin.top : margin.top + cursorY }, element);
                if (!absolute)
                    cursorY += used;
            }
        };
        if (spec.pages?.length) {
            for (const pageSpec of spec.pages) {
                await renderPage(pageSpec.elements || spec.elements || [], pageSpec);
            }
        }
        else {
            const elements = spec.elements || [];
            const [pageWidth, pageHeight] = resolvePageSize(spec.pageSize, false);
            const margin = resolveMargin(spec.margin);
            const usableHeight = pageHeight - margin.top - margin.bottom;
            let page = doc.addPage([pageWidth, pageHeight]);
            let pageNumber = 1;
            let ctx = {
                doc,
                fonts,
                assets,
                page,
                pageHeight,
                contentWidth: pageWidth - margin.left - margin.right,
                originX: margin.left,
                originY: margin.top,
                defaultSize,
                defaultColor
            };
            const newPage = () => {
                page = doc.addPage([pageWidth, pageHeight]);
                ctx = { ...ctx, page };
                flow.page = page;
                flow.cursorY = 0;
                pageNumber += 1;
            };
            const flow = {
                page,
                marginTop: margin.top,
                cursorY: 0,
                usableHeight,
                ensure(need) {
                    if (this.cursorY + need <= this.usableHeight || this.cursorY === 0)
                        return false;
                    newPage();
                    return true;
                }
            };
            let prevAfter = 0;
            for (const element of elements) {
                if (element.type === 'pageBreak') {
                    newPage();
                    prevAfter = 0;
                    continue;
                }
                const before = spacingBeforeOf(element);
                const after = spacingAfterOf(element);
                const gap = prevAfter + before;
                if (gap > 0) {
                    flow.cursorY += Math.min(gap, Math.max(0, usableHeight - flow.cursorY));
                }
                spec.onElementLayout?.(element, pageNumber);
                await drawElement({ ...ctx, flow }, element);
                prevAfter = after;
            }
        }
        const bytes = await doc.save({ useObjectStreams: options.compress !== false });
        return bytes;
    }
    catch (error) {
        throw wrapError('pdf', '生成', error);
    }
}
function collectElements$1(spec) {
    if (spec.pages?.length)
        return spec.pages.flatMap((page) => page.elements || []);
    return spec.elements || [];
}

const PX_TO_PT = 0.75;
const DEFAULT_FONT_SIZE = 18;
const LINE_HEIGHT_CJK = 1.4;
const LINE_HEIGHT_LATIN = 1.2;
const GRADIENT_BANDS = 48;
function buildSchemeResolver(themeXml, masterXml) {
    const palette = new Map();
    const pick = (xml, tag) => {
        const m = xml.match(new RegExp(`<a:${tag}>[\\s\\S]*?</a:${tag}>`));
        if (!m)
            return undefined;
        const srgb = m[0].match(/srgbClr val="([0-9A-Fa-f]{6})"/);
        if (srgb)
            return `#${srgb[1].toUpperCase()}`;
        const sys = m[0].match(/sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"/);
        if (sys)
            return `#${sys[1].toUpperCase()}`;
        return undefined;
    };
    if (themeXml) {
        const scheme = themeXml.match(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/);
        if (scheme) {
            for (const slot of [
                'dk1', 'lt1', 'dk2', 'lt2',
                'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
                'hlink', 'folHlink'
            ]) {
                const color = pick(scheme[0], slot);
                if (color)
                    palette.set(slot, color);
            }
        }
    }
    const colorMap = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
    if (masterXml) {
        const map = masterXml.match(/<p:clrMap[^>]*\/>/);
        if (map) {
            for (const [k, v] of Object.entries(colorMap)) {
                const m = map[0].match(new RegExp(`${k}="(\\w+)"`));
                if (m)
                    colorMap[k] = m[1];
            }
        }
    }
    return (value) => {
        if (!value)
            return undefined;
        const name = value.replace(/^scheme:/i, '').trim();
        const slot = colorMap[name] ?? name;
        return palette.get(slot);
    };
}
let resolveScheme = () => undefined;
function withThemeColors(doc, fn) {
    const prev = resolveScheme;
    resolveScheme = buildSchemeResolver(doc.themeXmls?.[0] ?? doc.themeXml, doc.masters?.[0]?.__rawXml);
    try {
        return fn();
    }
    finally {
        resolveScheme = prev;
    }
}
function normalizeColor(value) {
    if (!value || value === 'none' || typeof value !== 'string')
        return undefined;
    if (/^scheme:/i.test(value))
        value = resolveScheme(value) ?? undefined;
    if (!value)
        return undefined;
    if (!value.startsWith('#'))
        value = `#${value}`;
    return /^#[0-9a-fA-F]{6}$/.test(value) ? value : undefined;
}
function rgbOf(color) {
    const hex = normalizeColor(color)?.slice(1);
    if (!hex || !/^[0-9a-fA-F]{6}$/.test(hex))
        return null;
    return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}
function toHex(rgb) {
    return `#${rgb.map((v) => Math.round(v).toString(16).padStart(2, '0')).join('')}`;
}
function gradientStopsOf(fill) {
    if (!fill || typeof fill === 'string')
        return null;
    const f = fill;
    if (f.type !== 'gradient' || !f.stops?.length)
        return null;
    const stops = [...f.stops]
        .map((s) => ({ color: normalizeColor(s.color), position: s.position ?? 0, transparency: s.transparency }))
        .filter((s) => !!s.color)
        .sort((a, b) => a.position - b.position);
    if (!stops.length)
        return null;
    return { stops, direction: f.direction || 'vertical', angle: f.angle };
}
function gradientAt(stops, t) {
    const clamped = Math.max(0, Math.min(1, t));
    let lo = stops[0];
    let hi = stops[stops.length - 1];
    for (let i = 0; i < stops.length - 1; i++) {
        const a = stops[i];
        const b = stops[i + 1];
        if (clamped >= a.position && clamped <= b.position) {
            lo = a;
            hi = b;
            break;
        }
    }
    const span = hi.position - lo.position;
    const k = span > 0 ? (clamped - lo.position) / span : 0;
    const c1 = rgbOf(lo.color);
    const c2 = rgbOf(hi.color);
    const color = c1 && c2 ? toHex([0, 1, 2].map((i) => c1[i] + (c2[i] - c1[i]) * k)) : lo.color;
    const a1 = 1 - (lo.transparency ?? 0) / 100;
    const a2 = 1 - (hi.transparency ?? 0) / 100;
    return { color, alpha: a1 + (a2 - a1) * k };
}
function gradientAxis(direction, angle, w, h) {
    let dx;
    let dy;
    if (angle != null) {
        const rad = (angle * Math.PI) / 180;
        dx = Math.cos(rad);
        dy = Math.sin(rad);
    }
    else if (direction === 'horizontal') {
        dx = 1;
        dy = 0;
    }
    else if (direction === 'diagonal') {
        dx = Math.SQRT1_2;
        dy = -Math.SQRT1_2;
    }
    else {
        dx = 0;
        dy = 1;
    }
    const len = Math.hypot(dx, dy) || 1;
    dx /= len;
    dy /= len;
    const corners = [
        [0, 0],
        [w, 0],
        [0, h],
        [w, h]
    ];
    let p1 = corners[0];
    let p2 = corners[0];
    let sMin = Infinity;
    let sMax = -Infinity;
    for (const c of corners) {
        const s = c[0] * dx + c[1] * dy;
        if (s < sMin) {
            sMin = s;
            p1 = c;
        }
        if (s > sMax) {
            sMax = s;
            p2 = c;
        }
    }
    const coverage = Math.abs(w * -dy) + Math.abs(h * dx);
    return { dir: [dx, dy], p1, len: sMax - sMin, coverage, angle: (Math.atan2(dy, dx) * 180) / Math.PI };
}
const BEZIER_K = 0.5523;
function shapeCornerRadii(shapeType, adjust, width, height) {
    const st = (shapeType || '').toLowerCase();
    const min = Math.min(width, height);
    if (min <= 0)
        return [0, 0, 0, 0];
    const adj = adjust || {};
    if (st === 'roundrect') {
        const r = ((adj.adj != null ? adj.adj : 16667) / 100000) * min;
        return [r, r, r, r];
    }
    if (st === 'round1rect' || st === 'round2samerect' || st === 'round2diagrect') {
        const r1 = ((adj.adj1 != null ? adj.adj1 : 33333) / 100000) * min;
        const r2 = ((adj.adj2 != null ? adj.adj2 : 0) / 100000) * min;
        if (st === 'round1rect')
            return [r1, 0, 0, 0];
        if (st === 'round2samerect')
            return [r1, r1, 0, 0];
        return [r1, 0, r1, 0];
    }
    return [0, 0, 0, 0];
}
function solidFillOf(fill) {
    if (!fill || fill === 'none')
        return undefined;
    if (typeof fill === 'string') {
        const color = normalizeColor(fill);
        return color ? { color, alpha: 1 } : undefined;
    }
    const f = fill;
    if (f.type === 'gradient' || f.type === 'image')
        return undefined;
    const color = normalizeColor(f.type === 'pattern' ? f.fg : f.color);
    if (!color)
        return undefined;
    const alpha = 1 - (f.transparency ?? 0) / 100;
    if (alpha < 0.03)
        return undefined;
    return { color, alpha };
}
const DASH_PATTERNS = {
    dash: [5, 5],
    dashDot: [5, 5, 1, 5],
    dot: [1, 5],
    lgDash: [10, 5],
    lgDashDot: [10, 5, 1, 5],
    lgDashDotDot: [10, 5, 1, 5, 1, 5],
    sysDash: [5, 2],
    sysDot: [2, 5],
    sysDashDot: [5, 2, 1, 5],
    sysDashDotDot: [5, 2, 1, 5, 1, 5]
};
function flipOn(value) {
    return value === true || value === '1' || value === 'true';
}
function pathFromPoints(points, closed = false) {
    const segs = points.map(([x, y], i) => ({
        op: i === 0 ? 'm' : 'l',
        args: [x, y]
    }));
    if (closed)
        segs.push({ op: 'h', args: [] });
    return segs;
}
function quadToCubic(p0, c, p1) {
    return [
        p0[0] + (2 / 3) * (c[0] - p0[0]),
        p0[1] + (2 / 3) * (c[1] - p0[1]),
        p1[0] + (2 / 3) * (c[0] - p1[0]),
        p1[1] + (2 / 3) * (c[1] - p1[1]),
        p1[0],
        p1[1]
    ];
}
function connectorToElements(el) {
    const line = el.line && el.line !== 'none' ? el.line : undefined;
    const color = normalizeColor(line?.color);
    if (!color)
        return [];
    const st = (el.shapeType || '').toLowerCase();
    const { x, y, width: w, height: h } = el;
    const flipH = flipOn(el.flipH);
    const flipV = flipOn(el.flipV);
    const dash = DASH_PATTERNS[line?.dashType || '']?.map((d) => d * PX_TO_PT);
    const lineWidth = Math.max(0.75, line?.width ?? 2);
    const isPlainLine = st === 'line';
    const startArrow = !isPlainLine && line?.startArrow && line.startArrow !== 'none' ? true : undefined;
    const endArrow = !isPlainLine && line?.endArrow && line.endArrow !== 'none' ? true : undefined;
    const P = (lx, ly) => [
        (x + (flipH ? w - lx : lx)) * PX_TO_PT,
        (y + (flipV ? h - ly : ly)) * PX_TO_PT
    ];
    if (st.startsWith('bentconnector')) {
        return [
            {
                type: 'path',
                segments: pathFromPoints([P(0, 0), P(w / 2, 0), P(w / 2, h), P(w, h)]),
                stroke: color,
                lineWidth,
                dash,
                startArrow,
                endArrow
            }
        ];
    }
    if (st.startsWith('curvedconnector')) {
        const r = Math.min(Math.max((el.adjust?.adj1 ?? 50000) / 100000, 0), 1);
        const p0 = P(0, 0);
        const c1 = P(w * r, 0);
        const mid = P(w / 2, h / 2);
        const c2 = P(w * (1 - r), h);
        const p1 = P(w, h);
        return [
            {
                type: 'path',
                segments: [
                    { op: 'm', args: [p0[0], p0[1]] },
                    { op: 'c', args: quadToCubic(p0, c1, mid) },
                    { op: 'c', args: quadToCubic(mid, c2, p1) }
                ],
                stroke: color,
                lineWidth,
                dash,
                startArrow,
                endArrow
            }
        ];
    }
    const a = P(0, 0);
    const b = P(w, h);
    return [
        {
            type: 'line',
            x1: a[0],
            y1: a[1],
            x2: b[0],
            y2: b[1],
            lineWidth,
            color,
            dash,
            startArrow,
            endArrow
        }
    ];
}
function blockArrowToElements(el) {
    const fill = solidFillOf(el.fill);
    if (!fill)
        return [];
    const { x, y, width: w, height: h } = el;
    const ss = Math.min(w, h);
    const adj = el.adjust || {};
    const body = ss * ((adj.adj1 ?? 50000) / 100000);
    const head = ss * ((adj.adj2 ?? 50000) / 100000);
    let pts;
    if ((el.shapeType || '').toLowerCase() === 'downarrow') {
        const left = (w - body) / 2;
        const right = (w + body) / 2;
        pts = [
            [left, 0],
            [right, 0],
            [right, h - head],
            [w, h - head],
            [w / 2, h],
            [0, h - head],
            [left, h - head]
        ];
    }
    else {
        const top = (h - body) / 2;
        const bot = (h + body) / 2;
        pts = [
            [0, top],
            [w - head, top],
            [w - head, 0],
            [w, h / 2],
            [w - head, h],
            [w - head, bot],
            [0, bot]
        ];
    }
    const segments = [];
    pts.forEach(([px, py], i) => {
        segments.push({
            op: i === 0 ? 'm' : 'l',
            args: [(x + px) * PX_TO_PT, (y + py) * PX_TO_PT]
        });
    });
    segments.push({ op: 'h', args: [] });
    const stroke = normalizeColor(el.line && el.line !== 'none' ? el.line.color : undefined);
    return [
        {
            type: 'path',
            segments,
            fill: fill.color,
            opacity: fill.alpha < 1 ? fill.alpha : undefined,
            stroke,
            lineWidth: stroke ? Math.max(0.5, el.line.width ?? 1) : undefined
        }
    ];
}
function canToElements(el) {
    const fill = solidFillOf(el.fill);
    if (!fill)
        return [];
    const { x, y, width: w, height: h } = el;
    const ss = Math.min(w, h);
    const ry = (ss * ((el.adjust?.adj ?? 25000) / 100000)) / 2;
    if (ry <= 0 || h <= 2 * ry)
        return [];
    const opacity = fill.alpha < 1 ? fill.alpha : undefined;
    return [
        {
            type: 'rect',
            x: x * PX_TO_PT,
            y: (y + h - 2 * ry) * PX_TO_PT,
            width: w * PX_TO_PT,
            height: 2 * ry * PX_TO_PT,
            fill: fill.color,
            opacity,
            ellipse: true
        },
        {
            type: 'rect',
            x: x * PX_TO_PT,
            y: (y + ry) * PX_TO_PT,
            width: w * PX_TO_PT,
            height: (h - 2 * ry) * PX_TO_PT,
            fill: fill.color,
            opacity
        },
        {
            type: 'rect',
            x: x * PX_TO_PT,
            y: y * PX_TO_PT,
            width: w * PX_TO_PT,
            height: 2 * ry * PX_TO_PT,
            fill: fill.color,
            opacity,
            ellipse: true
        }
    ];
}
function roundedRectPathLocal(w, h, radii) {
    const [tl, tr, br, bl] = radii;
    const k = (v) => BEZIER_K * v;
    const segs = [{ op: 'm', args: [tl, 0] }];
    segs.push({ op: 'l', args: [w - tr, 0] });
    if (tr > 0)
        segs.push({ op: 'c', args: [w - tr + k(tr), 0, w, tr - k(tr), w, tr] });
    segs.push({ op: 'l', args: [w, h - br] });
    if (br > 0)
        segs.push({ op: 'c', args: [w, h - br + k(br), w - br + k(br), h, w - br, h] });
    segs.push({ op: 'l', args: [bl, h] });
    if (bl > 0)
        segs.push({ op: 'c', args: [bl - k(bl), h, 0, h - bl + k(bl), 0, h - bl] });
    segs.push({ op: 'l', args: [0, tl] });
    if (tl > 0)
        segs.push({ op: 'c', args: [0, tl - k(tl), tl - k(tl), 0, tl, 0] });
    segs.push({ op: 'h', args: [] });
    return segs;
}
function ellipsePathLocal(w, h) {
    const cx = w / 2;
    const cy = h / 2;
    const kx = BEZIER_K * (w / 2);
    const ky = BEZIER_K * (h / 2);
    return [
        { op: 'm', args: [cx + w / 2, cy] },
        { op: 'c', args: [cx + w / 2, cy + ky, cx + kx, cy + h / 2, cx, cy + h / 2] },
        { op: 'c', args: [cx - kx, cy + h / 2, cx - w / 2, cy + ky, cx - w / 2, cy] },
        { op: 'c', args: [cx - w / 2, cy - ky, cx - kx, cy - h / 2, cx, cy - h / 2] },
        { op: 'c', args: [cx + kx, cy - h / 2, cx + w / 2, cy - ky, cx + w / 2, cy] },
        { op: 'h', args: [] }
    ];
}
function rotatePoint(p, center, deg) {
    if (!deg)
        return p;
    const rad = (deg * Math.PI) / 180;
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    const dx = p[0] - center[0];
    const dy = p[1] - center[1];
    return [center[0] + dx * cos - dy * sin, center[1] + dx * sin + dy * cos];
}
function pathToPage(segs, originX, originY, center, theta) {
    const out = [];
    for (const seg of segs) {
        if (seg.op === 'h') {
            out.push({ op: 'h', args: [] });
            continue;
        }
        const pointCount = seg.op === 'c' ? 3 : 1;
        const args = [];
        for (let i = 0; i < pointCount; i++) {
            const [rx, ry] = rotatePoint([seg.args[i * 2], seg.args[i * 2 + 1]], center, theta);
            args.push((originX + rx) * PX_TO_PT, (originY + ry) * PX_TO_PT);
        }
        out.push({ op: seg.op, args });
    }
    return out;
}
function shapeToElements(el) {
    if (!el.width && !el.height)
        return [];
    const st = (el.shapeType || '').toLowerCase();
    if (/^(straightconnector|bentconnector|curvedconnector)/.test(st) || st === 'line') {
        return connectorToElements(el);
    }
    if (st === 'rightarrow' || st === 'downarrow')
        return blockArrowToElements(el);
    if (st === 'can' || st === 'cylinder' || st === 'flowchartmagneticdisk') {
        return canToElements(el);
    }
    const x = el.x * PX_TO_PT;
    const y = el.y * PX_TO_PT;
    const w = el.width * PX_TO_PT;
    const h = el.height * PX_TO_PT;
    const line = el.line && el.line !== 'none' ? el.line : undefined;
    const stroke = normalizeColor(line?.color);
    const lineWidth = line?.width ? Math.max(0.5, line.width) : undefined;
    const radii = shapeCornerRadii(el.shapeType, el.adjust, el.width, el.height);
    const uniformRadius = radii[0] === radii[1] && radii[1] === radii[2] && radii[2] === radii[3];
    const ellipse = (el.shapeType || '').toLowerCase() === 'ellipse';
    const theta = el.rotation || 0;
    const rotation = theta ? -theta : undefined;
    const gradient = gradientStopsOf(el.fill);
    if (gradient) {
        const axis = gradientAxis(gradient.direction, gradient.angle, el.width, el.height);
        const n = GRADIENT_BANDS;
        const bands = [];
        for (let i = 0; i < n; i++) {
            const t = (i + 0.5) / n;
            const { color, alpha } = gradientAt(gradient.stops, t);
            if (alpha < 0.02)
                continue;
            const s = t * axis.len - axis.len / 2;
            const cx = el.width / 2 + axis.dir[0] * s;
            const cy = el.height / 2 + axis.dir[1] * s;
            const bandW = axis.len / n;
            const bandH = axis.coverage;
            const [rx, ry] = rotatePoint([cx, cy], [el.width / 2, el.height / 2], theta);
            bands.push({
                type: 'rect',
                x: (el.x + rx - bandW / 2) * PX_TO_PT,
                y: (el.y + ry - bandH / 2) * PX_TO_PT,
                width: bandW * PX_TO_PT,
                height: bandH * PX_TO_PT,
                fill: color,
                opacity: alpha,
                rotation: -(axis.angle + theta)
            });
        }
        if (!bands.length)
            return [];
        const hasOutline = ellipse || radii.some((r) => r > 0);
        if (!hasOutline)
            return bands;
        const local = ellipse
            ? ellipsePathLocal(el.width, el.height)
            : roundedRectPathLocal(el.width, el.height, radii);
        const clip = pathToPage(local, el.x, el.y, [el.width / 2, el.height / 2], theta);
        const out = [{ type: 'group', clip, children: bands }];
        if (stroke && (ellipse || uniformRadius)) {
            out.push({
                type: 'rect',
                x,
                y,
                width: w,
                height: h,
                stroke,
                lineWidth,
                radius: ellipse ? undefined : radii[0] * PX_TO_PT,
                ellipse: ellipse || undefined,
                rotation
            });
        }
        return out;
    }
    const fill = solidFillOf(el.fill);
    if (!fill && !stroke)
        return [];
    return [
        {
            type: 'rect',
            x,
            y,
            width: w,
            height: h,
            fill: fill?.color,
            opacity: fill && fill.alpha < 1 ? fill.alpha : undefined,
            stroke,
            lineWidth,
            radius: uniformRadius && radii[0] > 0 ? radii[0] * PX_TO_PT : undefined,
            ellipse: ellipse || undefined,
            rotation
        }
    ];
}
function runToPdfRun(run, fallbackSize, fontScale = 1) {
    const r = { text: run.text ?? '' };
    if (run.bold)
        r.bold = true;
    if (run.italic)
        r.italic = true;
    if (run.strike)
        r.strike = true;
    if (run.color)
        r.color = normalizeColor(run.color);
    const baseSize = run.fontSize ?? fallbackSize;
    const size = baseSize == null ? undefined : baseSize * fontScale;
    if (size)
        r.size = size;
    if (run.href)
        r.link = run.href;
    return r;
}
function charWidthFactor(ch) {
    if (/[　-〿㐀-鿿豈-﫿＀-￯]/.test(ch))
        return 1;
    if (/[A-Za-z0-9@%]/.test(ch))
        return 0.6;
    return 0.5;
}
function estimateWidth(text, size) {
    let w = 0;
    for (const ch of Array.from(text))
        w += charWidthFactor(ch) * size;
    return w;
}
function hasCjk(text) {
    return /[　-〿㐀-鿿豈-﫿＀-￯]/.test(text);
}
function textToElements(el) {
    const out = [];
    if (el.fill && el.fill !== 'none') {
        if (el.shapeType) {
            const box = {
                type: 'shape',
                shapeType: el.shapeType,
                x: el.x,
                y: el.y,
                width: el.width,
                height: el.height,
                fill: el.fill,
                line: el.line,
                adjust: el.adjust,
                rotation: el.rotation
            };
            out.push(...shapeToElements(box));
        }
        else {
            const fill = solidFillOf(el.fill);
            if (fill) {
                out.push({
                    type: 'rect',
                    x: el.x * PX_TO_PT,
                    y: el.y * PX_TO_PT,
                    width: el.width * PX_TO_PT,
                    height: el.height * PX_TO_PT,
                    fill: fill.color,
                    opacity: fill.alpha < 1 ? fill.alpha : undefined,
                    rotation: el.rotation ? -el.rotation : undefined
                });
            }
        }
    }
    const fontScale = (el.fontScale ?? 100) / 100;
    const paragraphs = el.paragraphs?.length
        ? el.paragraphs
        : el.runs?.length || el.text
            ? [{ runs: el.runs, text: el.text }]
            : [];
    const boxW = el.width * PX_TO_PT;
    const boxH = el.height * PX_TO_PT;
    const lines = [];
    let align;
    let textHeight = 0;
    for (const p of paragraphs) {
        const runs = p.runs?.length
            ? p.runs.map((r) => runToPdfRun(r, el.fontSize, fontScale))
            : p.text
                ? [{ text: p.text, size: (el.fontSize ?? DEFAULT_FONT_SIZE) * fontScale }]
                : [];
        const plain = runs.map((r) => r.text).join('');
        const size = Math.max(...runs.map((r) => r.size ?? 0), el.fontSize ?? 0, 1);
        const lh = size * (hasCjk(plain) ? LINE_HEIGHT_CJK : LINE_HEIGHT_LATIN);
        const wrapped = Math.max(1, Math.ceil(estimateWidth(plain, size) / Math.max(1, boxW) - 0.02));
        textHeight += wrapped * lh;
        if (!plain.trim()) {
            lines.push([{ text: '' }]);
            continue;
        }
        lines.push(runs);
        align = align ?? p.align;
    }
    if (!lines.length)
        return out;
    let offsetY = 0;
    if (el.valign === 'middle')
        offsetY = (boxH - textHeight) / 2;
    else if (el.valign === 'bottom')
        offsetY = boxH - textHeight;
    const runs = [];
    lines.forEach((line, i) => {
        if (i > 0)
            runs.push({ text: '\n' });
        runs.push(...line);
    });
    out.push({
        type: 'runs',
        x: el.x * PX_TO_PT,
        y: el.y * PX_TO_PT + offsetY,
        width: boxW,
        height: boxH,
        runs,
        align,
        lineHeight: LINE_HEIGHT_CJK,
        spacingAfter: 0,
        rotation: el.rotation ? -el.rotation : undefined
    });
    return out;
}
function cellToPdfCell(cell) {
    const text = cell.text ??
        cell.paragraphs
            ?.map((p) => p.text ?? p.runs?.map((r) => r.text ?? '').join('') ?? '')
            .join('\n') ??
        '';
    const bs = cell.borders || (cell.border ? { top: cell.border, right: cell.border, bottom: cell.border, left: cell.border } : undefined);
    if (!cell.colSpan && !cell.rowSpan && !cell.fill && !bs)
        return text;
    const obj = { content: text };
    if (cell.colSpan && cell.colSpan > 1)
        obj.colspan = cell.colSpan;
    if (cell.rowSpan && cell.rowSpan > 1)
        obj.rowspan = cell.rowSpan;
    const fill = normalizeColor(cell.fill);
    if (fill)
        obj.fill = fill;
    if (bs)
        obj.borders = bs;
    return obj;
}
function tableFontSize(el) {
    for (const row of el.rows) {
        for (const cell of row.cells) {
            for (const p of cell.paragraphs || []) {
                for (const r of p.runs || []) {
                    if (r.fontSize)
                        return r.fontSize;
                }
            }
        }
    }
    return 14;
}
function tableToElement(el) {
    return {
        type: 'table',
        x: el.x * PX_TO_PT,
        y: el.y * PX_TO_PT,
        rows: el.rows.map((row) => row.cells.flatMap((cell) => (cell.hMerge || cell.vMerge ? [] : [cellToPdfCell(cell)]))),
        header: false,
        fontSize: tableFontSize(el),
        columnWidths: el.colWidths?.map((w) => w * PX_TO_PT)
    };
}
function collectElements(list, out, offset) {
    const dx = offset?.x ?? 0;
    const dy = offset?.y ?? 0;
    for (const el of list) {
        switch (el.type) {
            case 'text': {
                const t = el;
                out.push(...textToElements({ ...t, x: t.x + dx, y: t.y + dy }));
                break;
            }
            case 'shape': {
                const s = el;
                out.push(...shapeToElements({ ...s, x: s.x + dx, y: s.y + dy }));
                break;
            }
            case 'image': {
                const img = el;
                const src = img.data || img.src;
                if (!src || typeof src !== 'string')
                    break;
                out.push({
                    type: 'image',
                    src,
                    x: (img.x + dx) * PX_TO_PT,
                    y: (img.y + dy) * PX_TO_PT,
                    width: img.width * PX_TO_PT,
                    height: img.height * PX_TO_PT,
                    crop: img.crop
                });
                break;
            }
            case 'table': {
                const t = el;
                out.push(tableToElement({ ...t, x: t.x + dx, y: t.y + dy }));
                break;
            }
            case 'group': {
                const g = el;
                if (g.childrenCoordinates === 'relative') {
                    collectElements(g.children || [], out, { x: (g.x || 0) + dx, y: (g.y || 0) + dy });
                }
                else {
                    collectElements(g.children || [], out);
                }
                break;
            }
            default:
                break;
        }
    }
}
function slideBackgroundColor$1(slide) {
    const bg = slide.background;
    if (typeof bg === 'string')
        return normalizeColor(bg) ?? '#ffffff';
    if (bg) {
        const b = bg;
        const color = b.type === 'gradient' ? normalizeColor(b.stops?.[0]?.color) : normalizeColor(b.color);
        if (color)
            return color;
    }
    return '#ffffff';
}
function slideBackgroundImage$1(slide) {
    const bg = slide.background;
    if (bg && typeof bg === 'object') {
        const b = bg;
        if (b.type === 'image')
            return b.data || b.src;
    }
    return undefined;
}
function slidesHaveCjk(slides) {
    let found = false;
    const walk = (els) => {
        for (const el of els || []) {
            if (found)
                return;
            if (el.type === 'text') {
                const t = el;
                const text = t.text ||
                    (t.paragraphs || []).map((p) => p.text || p.runs?.map((r) => r.text ?? '').join('') || '').join('');
                if (/[一-鿿]/.test(text))
                    found = true;
            }
            else if (el.type === 'group') {
                walk(el.children);
            }
        }
    };
    for (const s of slides)
        walk(s.elements);
    return found;
}
async function pptxToPdf(input, options = {}) {
    const doc = (await pptxToStandard(input, { assets: 'both' }));
    const pageW = doc.slideSize.width * PX_TO_PT;
    const pageH = doc.slideSize.height * PX_TO_PT;
    let fonts = options.fonts;
    if (!fonts && options.autoFont !== false && slidesHaveCjk(doc.slides)) {
        try {
            fonts = await detectSystemFonts();
        }
        catch {
        }
    }
    resolveScheme = buildSchemeResolver(doc.themeXmls?.[0] ?? doc.themeXml, doc.masters?.[0]?.__rawXml);
    const pages = [];
    try {
        for (const slide of doc.slides) {
            if (slide.hidden)
                continue;
            const bg = slideBackgroundColor$1(slide);
            const elements = [{ type: 'rect', x: 0, y: 0, width: pageW, height: pageH, fill: bg }];
            const bgImage = slideBackgroundImage$1(slide);
            if (bgImage) {
                elements.push({ type: 'image', src: bgImage, x: 0, y: 0, width: pageW, height: pageH });
            }
            collectElements(slide.elements || [], elements);
            pages.push({ size: [pageW, pageH], margin: 0, elements });
        }
        if (!pages.length)
            pages.push({ size: [pageW, pageH], margin: 0, elements: [] });
    }
    finally {
        resolveScheme = () => undefined;
    }
    return generatePdf({ pages, fonts, defaultSize: DEFAULT_FONT_SIZE }, {});
}

const FONT_STACK = "'Microsoft YaHei','微软雅黑','PingFang SC','Hiragino Sans GB','Source Han Sans SC'," +
    "'Noto Sans CJK SC',-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif";
function esc(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
function num(value) {
    return String(Math.round(value * 100) / 100);
}
function pathToD(segments) {
    const parts = [];
    for (const seg of segments) {
        if (seg.op === 'm')
            parts.push(`M ${num(seg.args[0])} ${num(seg.args[1])}`);
        else if (seg.op === 'l')
            parts.push(`L ${num(seg.args[0])} ${num(seg.args[1])}`);
        else if (seg.op === 'c')
            parts.push(`C ${num(seg.args[0])} ${num(seg.args[1])} ${num(seg.args[2])} ${num(seg.args[3])} ${num(seg.args[4])} ${num(seg.args[5])}`);
        else if (seg.op === 'h')
            parts.push('Z');
    }
    return parts.join(' ');
}
function gradientStopsSvg(stops) {
    return stops
        .map((s) => {
        const { color, alpha } = gradientAt([s], 0);
        const offset = Math.round((s.position ?? 0) * 1000) / 10;
        const op = alpha < 1 ? ` stop-opacity="${num(alpha)}"` : '';
        return `<stop offset="${num(offset)}%" stop-color="${color}"${op}/>`;
    })
        .join('');
}
function strokeAttrs(el) {
    const line = el.line && el.line !== 'none' ? el.line : undefined;
    const color = normalizeColor(line?.color);
    const width = line?.width ? Math.max(0.5, line.width * (96 / 72)) : 0;
    const dashArr = DASH_PATTERNS[line?.dashType ?? ''];
    return {
        color,
        width,
        dash: dashArr ? ` stroke-dasharray="${dashArr.map((d) => num(d)).join(' ')}"` : undefined
    };
}
function svgBox(x, y, w, h, inner, rotation) {
    const rot = rotation ? ` transform:rotate(${num(rotation)}deg);` : '';
    return (`<svg style="left:${num(x)}px;top:${num(y)}px;width:${num(Math.max(w, 0.1))}px;` +
        `height:${num(Math.max(h, 0.1))}px;${rot}" viewBox="0 0 ${num(Math.max(w, 0.1))} ${num(Math.max(h, 0.1))}" overflow="visible" xmlns="http://www.w3.org/2000/svg">${inner}</svg>`);
}
function connectorToSvg(el) {
    const st = (el.shapeType || '').toLowerCase();
    const { x, y, width: w, height: h } = el;
    const flipH = flipOn(el.flipH);
    const flipV = flipOn(el.flipV);
    const line = el.line && el.line !== 'none' ? el.line : undefined;
    const color = normalizeColor(line?.color) ?? '#000000';
    const lw = Math.max(1, (line?.width ?? 2) * (96 / 72));
    const dashArr = DASH_PATTERNS[line?.dashType ?? ''];
    const dash = dashArr ? ` stroke-dasharray="${dashArr.map((d) => num(d)).join(' ')}"` : '';
    const isPlainLine = st === 'line';
    const hasArrow = !isPlainLine && !!(line?.endArrow || line?.startArrow);
    const P = (lx, ly) => [
        x + (flipH ? w - lx : lx),
        y + (flipV ? h - ly : ly)
    ];
    if (st.startsWith('bentconnector')) {
        const pts = [P(0, 0), P(w / 2, 0), P(w / 2, h), P(w, h)];
        const d = pts.map((p, i) => `${i ? 'L' : 'M'} ${num(p[0] - x)} ${num(p[1] - y)}`).join(' ');
        return svgBox(x, y, w, h, `<path d="${d}" fill="none" stroke="${color}" stroke-width="${num(lw)}"${dash}/>`, el.rotation);
    }
    if (st.startsWith('curvedconnector')) {
        const r = Math.min(Math.max((el.adjust?.adj1 ?? 50000) / 100000, 0), 1);
        const p0 = P(0, 0);
        const c1 = P(w * r, 0);
        const mid = P(w / 2, h / 2);
        const c2 = P(w * (1 - r), h);
        const p1 = P(w, h);
        const q = (a, c, b) => `Q ${num(c[0] - x)} ${num(c[1] - y)} ${num(b[0] - x)} ${num(b[1] - y)}`;
        const d = `M ${num(p0[0] - x)} ${num(p0[1] - y)} ${q(p0, c1, mid)} ${q(mid, c2, p1)}`;
        return svgBox(x, y, w, h, `<path d="${d}" fill="none" stroke="${color}" stroke-width="${num(lw)}"${dash}/>`, el.rotation);
    }
    const a = P(0, 0);
    const b = P(w, h);
    const markerId = `ah${uid()}`;
    const defs = hasArrow
        ? `<defs><marker id="${markerId}" markerWidth="5" markerHeight="5" refX="10" refY="5" ` +
            `orient="auto-start-reverse" markerUnits="strokeWidth" viewBox="0 0 10 10">` +
            `<path d="M 0 0 L 10 5 L 0 10 z" fill="${color}"/></marker></defs>`
        : '';
    const endMarker = hasArrow && line?.endArrow && line.endArrow !== 'none'
        ? ` marker-end="url(#${markerId})"`
        : '';
    const startMarker = hasArrow && line?.startArrow && line.startArrow !== 'none'
        ? ` marker-start="url(#${markerId})"`
        : '';
    const lineEl = `<line x1="${num(a[0] - x)}" y1="${num(a[1] - y)}" x2="${num(b[0] - x)}" y2="${num(b[1] - y)}" stroke="${color}" stroke-width="${num(lw)}"${dash}${endMarker}${startMarker}/>`;
    return svgBox(x, y, w, h, defs + lineEl, el.rotation);
}
let uidCounter = 0;
function uid() {
    uidCounter += 1;
    return String(uidCounter);
}
function blockArrowToSvg(el) {
    const { x, y, width: w, height: h } = el;
    const ss = Math.min(w, h);
    const adj = el.adjust || {};
    const body = ss * ((adj.adj1 ?? 50000) / 100000);
    const head = ss * ((adj.adj2 ?? 50000) / 100000);
    const fill = solidFillOf(el.fill);
    if (!fill)
        return '';
    const stroke = strokeAttrs(el);
    let pts;
    if ((el.shapeType || '').toLowerCase() === 'downarrow') {
        const left = (w - body) / 2;
        const right = (w + body) / 2;
        pts = [
            [left, 0],
            [right, 0],
            [right, h - head],
            [w, h - head],
            [w / 2, h],
            [0, h - head],
            [left, h - head]
        ];
    }
    else {
        const top = (h - body) / 2;
        const bot = (h + body) / 2;
        pts = [
            [0, top],
            [w - head, top],
            [w - head, 0],
            [w, h / 2],
            [w - head, h],
            [w - head, bot],
            [0, bot]
        ];
    }
    const d = pts.map((p, i) => `${i ? 'L' : 'M'} ${num(p[0])} ${num(p[1])}`).join(' ') + ' Z';
    const fillAttr = `fill="${fill.color}"${fill.alpha < 1 ? ` fill-opacity="${num(fill.alpha)}"` : ''}`;
    const strokeAttr = stroke.color
        ? ` stroke="${stroke.color}" stroke-width="${num(stroke.width)}"`
        : '';
    return svgBox(x, y, w, h, `<path d="${d}" ${fillAttr}${strokeAttr}/>`, el.rotation);
}
function canToSvg(el) {
    const { x, y, width: w, height: h } = el;
    const ss = Math.min(w, h);
    const ry = (ss * ((el.adjust?.adj ?? 25000) / 100000)) / 2;
    const fill = solidFillOf(el.fill);
    if (!fill || ry <= 0 || h <= 2 * ry)
        return '';
    const stroke = strokeAttrs(el);
    const fillAttr = `fill="${fill.color}"${fill.alpha < 1 ? ` fill-opacity="${num(fill.alpha)}"` : ''}`;
    const strokeAttr = stroke.color
        ? ` stroke="${stroke.color}" stroke-width="${num(stroke.width)}"`
        : '';
    const bodyTop = ry;
    const bodyH = h - 2 * ry;
    const inner = `<ellipse cx="${num(w / 2)}" cy="${num(h - ry)}" rx="${num(w / 2)}" ry="${num(ry)}" ${fillAttr}${strokeAttr}/>` +
        `<rect x="0" y="${num(bodyTop)}" width="${num(w)}" height="${num(bodyH)}" ${fillAttr}${strokeAttr}/>` +
        `<ellipse cx="${num(w / 2)}" cy="${num(ry)}" rx="${num(w / 2)}" ry="${num(ry)}" ${fillAttr}${strokeAttr}/>`;
    return svgBox(x, y, w, h, inner, el.rotation);
}
function shapeToSvg(el) {
    const { x, y, width: w, height: h } = el;
    if (!w || !h)
        return '';
    const st = (el.shapeType || '').toLowerCase();
    const isEllipse = st === 'ellipse';
    const radii = shapeCornerRadii(el.shapeType, el.adjust, w, h);
    const uniform = radii[0] === radii[1] && radii[1] === radii[2] && radii[2] === radii[3];
    const rx = uniform ? radii[0] : radii[0];
    const stroke = strokeAttrs(el);
    const strokeAttr = stroke.color
        ? ` stroke="${stroke.color}" stroke-width="${num(stroke.width)}"${stroke.dash ?? ''}`
        : '';
    const grad = gradientStopsOf(el.fill);
    let fillAttr = 'fill="none"';
    let defs = '';
    if (grad) {
        const axis = gradientAxis(grad.direction, grad.angle, w, h);
        const x1 = axis.p1[0];
        const y1 = axis.p1[1];
        const x2 = axis.p1[0] + axis.dir[0] * axis.len;
        const y2 = axis.p1[1] + axis.dir[1] * axis.len;
        const gid = `grad${uid()}`;
        defs =
            `<defs><linearGradient id="${gid}" gradientUnits="userSpaceOnUse" ` +
                `x1="${num(x1)}" y1="${num(y1)}" x2="${num(x2)}" y2="${num(y2)}">` +
                gradientStopsSvg(grad.stops) +
                '</linearGradient></defs>';
        fillAttr = `fill="url(#${gid})"`;
    }
    else {
        const fill = solidFillOf(el.fill);
        if (fill) {
            fillAttr = `fill="${fill.color}"${fill.alpha < 1 ? ` fill-opacity="${num(fill.alpha)}"` : ''}`;
        }
    }
    if (!isEllipse && !uniform && radii.some((r) => r > 0)) {
        const d = pathToD(roundedRectPathLocal(w, h, radii));
        return svgBox(x, y, w, h, `${defs}<path d="${d}" ${fillAttr}${strokeAttr}/>`, el.rotation);
    }
    if (isEllipse) {
        return svgBox(x, y, w, h, `${defs}<ellipse cx="${num(w / 2)}" cy="${num(h / 2)}" rx="${num(w / 2)}" ry="${num(h / 2)}" ${fillAttr}${strokeAttr}/>`, el.rotation);
    }
    return svgBox(x, y, w, h, `${defs}<rect x="0" y="0" width="${num(w)}" height="${num(h)}"${rx > 0 ? ` rx="${num(rx)}"` : ''} ${fillAttr}${strokeAttr}/>`, el.rotation);
}
function presetShapeSvg(el) {
    const st = el.shapeType || '';
    const { x, y, width: w, height: h } = el;
    if (!st || !w || !h)
        return '';
    const geo = pptParser.presetShapePath(st, w, h, el.adjust || {});
    if (!geo?.d)
        return '';
    const flipH = flipOn(el.flipH);
    const flipV = flipOn(el.flipV);
    let tf = geo.transform || '';
    if (flipH || flipV) {
        const sx = flipH ? -1 : 1;
        const sy = flipV ? -1 : 1;
        const tx = flipH ? w : 0;
        const ty = flipV ? h : 0;
        tf += (tf ? ' ' : '') + `translate(${tx},${ty}) scale(${sx},${sy})`;
    }
    const tfAttr = tf.trim() ? ` transform="${tf.trim()}"` : '';
    let defs = '';
    let fillAttr = 'fill="none"';
    if (geo.noFill) {
        fillAttr = 'fill="none"';
    }
    else {
        const grad = gradientStopsOf(el.fill);
        if (grad) {
            const axis = gradientAxis(grad.direction, grad.angle, w, h);
            const gid = `grad${uid()}`;
            defs =
                `<defs><linearGradient id="${gid}" gradientUnits="userSpaceOnUse" ` +
                    `x1="${num(axis.p1[0])}" y1="${num(axis.p1[1])}" ` +
                    `x2="${num(axis.p1[0] + axis.dir[0] * axis.len)}" ` +
                    `y2="${num(axis.p1[1] + axis.dir[1] * axis.len)}">` +
                    gradientStopsSvg(grad.stops) +
                    '</linearGradient></defs>';
            fillAttr = `fill="url(#${gid})"`;
        }
        else {
            const fill = solidFillOf(el.fill);
            if (fill) {
                fillAttr = `fill="${fill.color}"${fill.alpha < 1 ? ` fill-opacity="${num(fill.alpha)}"` : ''}`;
            }
        }
    }
    const stroke = strokeAttrs(el);
    const strokeAttr = stroke.color
        ? ` stroke="${stroke.color}" stroke-width="${num(stroke.width)}"${stroke.dash ?? ''}`
        : '';
    const main = `<path d="${geo.d}"${tfAttr} ${fillAttr}${strokeAttr}` +
        `${geo.fillRule ? ` fill-rule="${geo.fillRule}"` : ''}/>`;
    const accent = stroke.color ?? '#5B9BD5';
    const extra = (geo.strokes || [])
        .map((s) => `<path d="${s.d}"${tfAttr} fill="none" stroke="${accent}" ` +
        `stroke-width="${num(s.width ?? Math.max(1, w * 0.05))}" stroke-linecap="round"/>`)
        .join('');
    return svgBox(x, y, w, h, `${defs}${main}${extra}`, el.rotation);
}
function shapeElementToSvg(el) {
    const st = (el.shapeType || '').toLowerCase();
    if (/^(straightconnector|bentconnector|curvedconnector)/.test(st) || st === 'line') {
        return connectorToSvg(el);
    }
    const preset = presetShapeSvg(el);
    if (preset)
        return preset;
    if (st === 'rightarrow' || st === 'downarrow')
        return blockArrowToSvg(el);
    if (st === 'can' || st === 'cylinder' || st === 'flowchartmagneticdisk')
        return canToSvg(el);
    return shapeToSvg(el);
}
function quoteFont(face) {
    return /^[A-Za-z0-9 _-]+$/.test(face) ? face : `'${face.replace(/'/g, "\\'")}'`;
}
function fontFamilyOf(...faces) {
    const out = [];
    for (const face of faces) {
        if (!face)
            continue;
        const quoted = quoteFont(face);
        if (!out.includes(quoted))
            out.push(quoted);
    }
    return out.length ? out.join(', ') : undefined;
}
function themeTextColor() {
    return normalizeColor('scheme:tx1') ?? '#000000';
}
function lineHeightStyle(spacing) {
    if (spacing == null)
        return '';
    if (typeof spacing === 'number') {
        return `line-height:${num(spacing >= 10 ? spacing / 100 : spacing)}`;
    }
    if (spacing.type === 'percent')
        return `line-height:${num(spacing.value / 100)}`;
    if (spacing.type === 'pt')
        return `line-height:${num(spacing.value)}pt`;
    return '';
}
const CN_DIGITS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
const CN_LEGAL = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
const CIRCLED = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩'];
function cnNumber(n, legal) {
    const digits = legal ? CN_LEGAL : CN_DIGITS;
    const ten = legal ? '拾' : '十';
    if (n <= 0)
        return String(n);
    if (n < 10)
        return digits[n];
    if (n === 10)
        return ten;
    if (n < 20)
        return ten + digits[n % 10];
    if (n < 100)
        return digits[Math.floor(n / 10)] + ten + (n % 10 ? digits[n % 10] : '');
    return String(n);
}
function romanNumber(n, upper) {
    const map = [
        [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
        [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']
    ];
    let out = '';
    let rest = n;
    for (const [v, s] of map) {
        while (rest >= v) {
            out += s;
            rest -= v;
        }
    }
    return upper ? out : out.toLowerCase();
}
function alphaNumber(n, upper) {
    let out = '';
    let rest = n;
    while (rest > 0) {
        rest -= 1;
        out = String.fromCharCode(97 + (rest % 26)) + out;
        rest = Math.floor(rest / 26);
    }
    return upper ? out.toUpperCase() : out;
}
function hebrewNumber(n) {
    const letters = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט', 'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ', 'ק', 'ר', 'ש', 'ת'];
    let out = '';
    let rest = Math.max(0, n - 1);
    do {
        out = letters[rest % letters.length] + out;
        rest = Math.floor(rest / letters.length) - 1;
    } while (rest >= 0);
    return out || letters[0];
}
function formatAutoNum(fmt, n) {
    switch (fmt) {
        case 'hebrew1Minus':
        case 'hebrew2Minus':
            return `${hebrewNumber(n)}-`;
        case 'arabicPeriod':
            return `${n}.`;
        case 'arabicParenR':
            return `${n})`;
        case 'arabicParenBoth':
            return `(${n})`;
        case 'arabicPlain':
            return `${n}`;
        case 'alphaLcPeriod':
            return `${alphaNumber(n, false)}.`;
        case 'alphaUcPeriod':
            return `${alphaNumber(n, true)}.`;
        case 'alphaLcParenR':
            return `${alphaNumber(n, false)})`;
        case 'alphaUcParenR':
            return `${alphaNumber(n, true)})`;
        case 'alphaLcParenBoth':
            return `(${alphaNumber(n, false)})`;
        case 'alphaUcParenBoth':
            return `(${alphaNumber(n, true)})`;
        case 'romanLcPeriod':
            return `${romanNumber(n, false)}.`;
        case 'romanUcPeriod':
            return `${romanNumber(n, true)}.`;
        case 'romanLcParenR':
            return `${romanNumber(n, false)})`;
        case 'romanUcParenR':
            return `${romanNumber(n, true)})`;
        case 'romanLcParenBoth':
            return `(${romanNumber(n, false)})`;
        case 'romanUcParenBoth':
            return `(${romanNumber(n, true)})`;
        case 'chineseCounting':
        case 'chineseCountingThousand':
        case 'ea1ChsPeriod':
        case 'ea1ChtPeriod':
            return `${cnNumber(n, false)}、`;
        case 'chineseLegalSimplified':
            return `${cnNumber(n, true)}、`;
        case 'ea1ChsPlain':
        case 'ea1ChtPlain':
            return cnNumber(n, false);
        case 'ideographDigital':
        case 'circleNumDbPlain':
        case 'circleNumWdWhitePlain':
        case 'circleNumWdBlackPlain':
            return CIRCLED[(n - 1) % CIRCLED.length];
        default:
            return `${n}.`;
    }
}
function runStyle(run, el, fallbackSize, fontScale) {
    const size = (run.fontSize ?? fallbackSize) * fontScale;
    const styles = [];
    if (size)
        styles.push(`font-size:${num(size)}pt`);
    const family = fontFamilyOf(run.fontFace, run.fontFaceEa, el.fontFace, el.fontFaceEa);
    if (family)
        styles.push(`font-family:${family},${FONT_STACK}`);
    const color = normalizeColor(run.color) ?? normalizeColor(el.color) ?? themeTextColor();
    if (color)
        styles.push(`color:${color}`);
    if (run.bold ?? el.bold)
        styles.push('font-weight:700');
    if (run.italic ?? el.italic)
        styles.push('font-style:italic');
    const deco = [];
    if (run.underline ?? el.underline)
        deco.push('underline');
    if (run.strike)
        deco.push('line-through');
    if (deco.length)
        styles.push(`text-decoration:${deco.join(' ')}`);
    if (run.baseline)
        styles.push(`vertical-align:${run.baseline > 0 ? 'super' : 'sub'}`);
    if (run.spacing)
        styles.push(`letter-spacing:${num(run.spacing)}px`);
    if (run.highlight)
        styles.push(`background-color:${normalizeColor(run.highlight) ?? 'transparent'}`);
    if (run.smallCaps)
        styles.push('font-variant:small-caps');
    if (run.emphasisMark) {
        const em = run.emphasisMark;
        styles.push(`text-emphasis-style:${em === 'circle' ? 'circle' : em === 'comma' ? 'comma' : 'dot'}`);
        styles.push('text-emphasis-position:under');
    }
    const shadows = [];
    if (run.outline && run.outline !== 'none' && run.outline.color) {
        const oc = normalizeColor(run.outline.color);
        if (oc) {
            const w = Math.max(1, Math.floor(run.outline.width && run.outline.width >= 1 ? run.outline.width : 4 / 3));
            shadows.push(`-${w}px 0 ${oc}, 0 ${w}px ${oc}, ${w}px 0 ${oc}, 0 -${w}px ${oc}`);
        }
    }
    if (run.shadow?.color) {
        const sc = normalizeColor(run.shadow.color);
        if (sc) {
            const alpha = run.shadow.alpha != null ? run.shadow.alpha : 1;
            const blur = run.shadow.blur != null ? run.shadow.blur : 0;
            shadows.push(`${num((run.shadow.x ?? 0) * (96 / 72))}px ${num((run.shadow.y ?? 0) * (96 / 72))}px ` +
                `${num(blur * (96 / 72))}px rgba(${rgbOf(sc)?.join(',') ?? '0,0,0'},${num(alpha)})`);
        }
    }
    if (shadows.length)
        styles.push(`text-shadow:${shadows.join(',')}`);
    return styles.join(';');
}
function bulletToHtml(bullet, el, firstRun, fontSize, counter) {
    const isNumber = bullet === 'number' || (typeof bullet === 'object' && bullet.type === 'number');
    if (isNumber) {
        const obj = typeof bullet === 'object' ? bullet : undefined;
        const fmt = obj?.fmt || 'arabicPeriod';
        if (counter.n == null)
            counter.n = obj?.start || 1;
        const text = formatAutoNum(fmt, counter.n);
        counter.n += 1;
        const color = normalizeColor(obj?.color) ?? normalizeColor(firstRun.color) ?? normalizeColor(el.color) ?? themeTextColor();
        const family = fontFamilyOf(firstRun.fontFace, firstRun.fontFaceEa, el.fontFace, el.fontFaceEa, '微软雅黑');
        return (`<span class="jt-bullet" style="font-size:${num(fontSize)}pt;color:${color};` +
            `font-family:${family ?? FONT_STACK};margin-right:0.3em">${esc(text)}</span>`);
    }
    const obj = typeof bullet === 'object' ? bullet : undefined;
    const color = normalizeColor(obj?.color) ?? normalizeColor(firstRun.color) ?? normalizeColor(el.color) ?? themeTextColor();
    const size = obj?.sizePct ? (fontSize * obj.sizePct) / 100 : fontSize;
    if (obj?.type === 'picture' && obj.data) {
        const px = num((size * 96) / 72);
        return (`<span class="jt-bullet jt-bullet-img" style="margin-right:0.3em">` +
            `<img src='${obj.data}' style="width:${px}px;height:${px}px;vertical-align:-0.15em"/></span>`);
    }
    const char = obj?.char || '•';
    const family = obj?.font
        ? quoteFont(obj.font)
        : fontFamilyOf(firstRun.fontFace, firstRun.fontFaceEa, el.fontFace, el.fontFaceEa, '微软雅黑');
    return (`<span class="jt-bullet" style="font-size:${num(size)}pt;color:${color};` +
        `font-family:${family ?? FONT_STACK};margin-right:0.3em">${esc(char)}</span>`);
}
function paragraphToHtml(p, el, fallbackSize, fontScale, counter) {
    const runs = p.runs?.length
        ? p.runs
        : p.text != null
            ? [{ text: p.text }]
            : [];
    const firstRun = runs[0] || {};
    const effSize = (r) => (r.fontSize ?? fallbackSize) * fontScale;
    const hasContent = runs.some((r) => r.text || r.break);
    const bulletRaw = p.bullet ?? el.bullet;
    const bullet = hasContent ? bulletRaw : null;
    const styles = [`text-align:${p.align ?? el.align ?? 'left'}`];
    const lh = lineHeightStyle(p.lineSpacing ?? el.lineSpacing);
    if (lh)
        styles.push(lh);
    const marLSrc = p.indentLeft ?? el.indentLeft ?? el.indent;
    const indentSrc = p.indent;
    if (marLSrc != null || indentSrc != null) {
        const marL = Math.max(0, marLSrc || 0);
        const firstLine = marL + (indentSrc || 0) < 0 ? -marL : indentSrc || 0;
        styles.push(`padding-left:${num(marL)}px`);
        styles.push(`text-indent:${num(firstLine)}px`);
    }
    if (p.spaceBefore)
        styles.push(`margin-top:${num(p.spaceBefore)}px`);
    if (p.spaceAfter)
        styles.push(`margin-bottom:${num(p.spaceAfter)}px`);
    let bulletHtml = '';
    if (bullet) {
        bulletHtml = bulletToHtml(bullet, el, firstRun, effSize(firstRun), counter);
    }
    else {
        counter.n = null;
    }
    const body = runs
        .map((r) => {
        if (r.break)
            return '<br/>';
        const style = runStyle(r, el, fallbackSize, fontScale);
        const text = esc(r.text ?? '');
        if (r.href) {
            return (`<a href="${esc(r.href)}" target="_blank" rel="noopener noreferrer" ` +
                `style="${style};text-decoration:underline">${text}</a>`);
        }
        return `<span style="${style}">${text}</span>`;
    })
        .join('');
    return `<div class="jt-p" style="${styles.join(';')}">${bulletHtml}${body || '<br/>'}</div>`;
}
function textToHtml(el) {
    const out = [];
    if (el.fill && el.fill !== 'none') {
        if (el.shapeType) {
            const box = {
                type: 'shape',
                shapeType: el.shapeType,
                x: el.x,
                y: el.y,
                width: el.width,
                height: el.height,
                fill: el.fill,
                line: el.line,
                adjust: el.adjust,
                rotation: el.rotation
            };
            const svg = shapeElementToSvg(box);
            if (svg)
                out.push(svg);
        }
        else {
            const fill = solidFillOf(el.fill);
            if (fill) {
                out.push(`<div style="left:${num(el.x)}px;top:${num(el.y)}px;width:${num(el.width)}px;` +
                    `height:${num(el.height)}px;background:${fill.color};` +
                    (fill.alpha < 1 ? `opacity:${num(fill.alpha)};` : '') +
                    `${el.rotation ? `transform:rotate(${num(el.rotation)}deg);` : ''}"></div>`);
            }
        }
    }
    const fontScale = (el.fontScale ?? 100) / 100;
    const paragraphs = el.paragraphs?.length
        ? el.paragraphs
        : el.runs?.length || el.text
            ? [{ runs: el.runs, text: el.text }]
            : [];
    if (!paragraphs.length)
        return out.join('');
    const fallbackSize = el.fontSize ?? 18;
    const justify = el.valign === 'middle' ? 'center' : el.valign === 'bottom' ? 'flex-end' : 'flex-start';
    const counter = { n: null };
    const body = paragraphs
        .map((p) => paragraphToHtml(p, el, fallbackSize, fontScale, counter))
        .join('');
    const inset = el.inset;
    const padding = inset
        ? `padding:${num(inset.t ?? 0)}px ${num(inset.r ?? 0)}px ${num(inset.b ?? 0)}px ${num(inset.l ?? 0)}px;`
        : '';
    const whiteSpace = el.noWrap ? 'nowrap' : 'pre-wrap';
    const dirStyle = el.textDirection && el.textDirection !== 'horz'
        ? 'writing-mode:vertical-rl;text-orientation:upright;'
        : '';
    const rot = el.rotation ? ` transform:rotate(${num(el.rotation)}deg);` : '';
    out.push(`<div class="jt-text" style="left:${num(el.x)}px;top:${num(el.y)}px;width:${num(el.width)}px;` +
        `height:${num(el.height)}px;justify-content:${justify};${padding}white-space:${whiteSpace};` +
        `${dirStyle}${rot}">${body}</div>`);
    return out.join('');
}
function cellBorderCss(edge, side) {
    if (!edge)
        return '';
    if (edge === 'none')
        return `border-${side}:none;`;
    const w = Math.max(0.5, (edge.width ?? 1) * (96 / 72));
    const color = normalizeColor(edge.color) ?? '#000000';
    return `border-${side}:${num(w)}px solid ${color};`;
}
function tableToHtml(el) {
    const rows = el.rows || [];
    if (!rows.length)
        return '';
    const body = rows
        .map((row, ri) => {
        const cells = (row.cells || [])
            .map((cell) => {
            const text = esc(cell.text ?? (cell.paragraphs || []).map((p) => p.text ?? '').join('\n'));
            const fill = cell.fill ? `background:${normalizeColor(cell.fill) ?? 'transparent'};` : '';
            const bs = cell.borders || (cell.border
                ? { top: cell.border, right: cell.border, bottom: cell.border, left: cell.border }
                : undefined);
            const borders = bs
                ? ['top', 'right', 'bottom', 'left']
                    .map((s) => cellBorderCss(bs[s], s))
                    .join('')
                : '';
            const tag = ri === 0 ? 'th' : 'td';
            return `<${tag} style="${fill}${borders}">${text.replace(/\n/g, '<br/>')}</${tag}>`;
        })
            .join('');
        return `<tr>${cells}</tr>`;
    })
        .join('');
    return (`<table class="jt-table" style="left:${num(el.x)}px;top:${num(el.y)}px;width:${num(el.width)}px">` +
        `<tbody>${body}</tbody></table>`);
}
function imageToHtml(el) {
    const src = el.data || el.src;
    if (!src)
        return '';
    const rot = el.rotation ? ` transform:rotate(${num(el.rotation)}deg);` : '';
    const adj = el.imageAdjust || {};
    const filters = [];
    if (adj.brightness)
        filters.push(`brightness(${num(1 + adj.brightness / 100)})`);
    if (adj.contrast)
        filters.push(`contrast(${num(1 + adj.contrast / 100)})`);
    const filter = filters.length ? `filter:${filters.join(' ')};` : '';
    const opacity = adj.transparency ? `opacity:${num(Math.max(0, Math.min(1, 1 - adj.transparency / 100)))};` : '';
    const crop = el.crop;
    const hasCrop = crop && (crop.l || crop.t || crop.r || crop.b);
    if (hasCrop) {
        const cw = Math.max(0.01, 1 - (crop.l || 0) - (crop.r || 0));
        const ch = Math.max(0.01, 1 - (crop.t || 0) - (crop.b || 0));
        const imgStyle = `position:absolute;left:${num((-(crop.l || 0) / cw) * 100)}%;top:${num((-(crop.t || 0) / ch) * 100)}%;` +
            `width:${num(100 / cw)}%;height:${num(100 / ch)}%;max-width:none;${filter}${opacity}`;
        return (`<div class="jt-img-box" style="left:${num(el.x)}px;top:${num(el.y)}px;width:${num(el.width)}px;` +
            `height:${num(el.height)}px;overflow:hidden;${rot}">` +
            `<img class="jt-img" src='${src}' style="${imgStyle}"/></div>`);
    }
    return (`<img class="jt-img" src='${src}' style="left:${num(el.x)}px;top:${num(el.y)}px;` +
        `width:${num(el.width)}px;height:${num(el.height)}px;${filter}${opacity}${rot}"/>`);
}
function collectHtml(list, out, offset, base) {
    const dx = offset?.x ?? 0;
    const dy = offset?.y ?? 0;
    const bx = base?.x ?? 0;
    const by = base?.y ?? 0;
    for (const el of list || []) {
        const ex = el.x + dx - bx;
        const ey = el.y + dy - by;
        switch (el.type) {
            case 'text': {
                const t = el;
                out.push(textToHtml({ ...t, x: ex, y: ey }));
                break;
            }
            case 'shape': {
                const s = el;
                const svg = shapeElementToSvg({ ...s, x: ex, y: ey });
                if (svg)
                    out.push(svg);
                break;
            }
            case 'image': {
                const img = el;
                out.push(imageToHtml({ ...img, x: ex, y: ey }));
                break;
            }
            case 'table': {
                const t = el;
                out.push(tableToHtml({ ...t, x: ex, y: ey }));
                break;
            }
            case 'group': {
                const g = el;
                const relKids = g.childrenCoordinates === 'relative';
                const gx = g.x + dx;
                const gy = g.y + dy;
                const tf = [];
                if (g.rotation)
                    tf.push(`rotate(${num(g.rotation)}deg)`);
                if (flipOn(g.flipH))
                    tf.push('scaleX(-1)');
                if (flipOn(g.flipV))
                    tf.push('scaleY(-1)');
                if (!tf.length) {
                    collectHtml(g.children || [], out, relKids ? { x: gx, y: gy } : { x: 0, y: 0 }, { x: bx, y: by });
                    break;
                }
                const inner = [];
                collectHtml(g.children || [], inner, relKids ? { x: gx, y: gy } : { x: 0, y: 0 }, { x: gx, y: gy });
                out.push(`<div class="jt-group" style="left:${num(gx - bx)}px;top:${num(gy - by)}px;` +
                    `width:${num(g.width || 0)}px;height:${num(g.height || 0)}px;` +
                    `transform:${tf.join(' ')};transform-origin:center center">${inner.join('')}</div>`);
                break;
            }
            default:
                break;
        }
    }
}
function slideBackgroundColor(slide) {
    const bg = slide.background;
    if (typeof bg === 'string')
        return normalizeColor(bg) ?? '#ffffff';
    if (bg) {
        const b = bg;
        const color = b.type === 'gradient' ? normalizeColor(b.stops?.[0]?.color) : normalizeColor(b.color);
        if (color)
            return color;
    }
    return '#ffffff';
}
function slideBackgroundImage(slide) {
    const bg = slide.background;
    if (bg && typeof bg === 'object') {
        const b = bg;
        if (b.type === 'image')
            return b.data || b.src;
    }
    return undefined;
}
function slideToHtml(slide, index, pageW, pageH) {
    const bg = slideBackgroundColor(slide);
    const bgImage = slideBackgroundImage(slide);
    const bgStyle = bgImage
        ?
            `background:#${bg.replace('#', '')} url('${bgImage}') center/cover no-repeat;`
        : `background:${bg};`;
    const parts = [];
    collectHtml(slide.elements || [], parts);
    const notes = slide.notes;
    const notesHtml = notes && String(notes).trim()
        ? `<div class="pptx-notes" data-notes="${esc(String(notes))}" style="display:none;white-space:pre-wrap;">${esc(String(notes))}</div>`
        : '';
    return (`<section class="jt-slide" data-slide="${index}" style="width:${num(pageW)}px;` +
        `height:${num(pageH)}px;${bgStyle}">${parts.join('')}${notesHtml}</section>`);
}
async function pptxToSlideHtml(input, options = {}) {
    try {
        const doc = (await pptxToStandard(input, { assets: 'both' }));
        const pageW = doc.slideSize.width;
        const pageH = doc.slideSize.height;
        const slides = doc.slides.filter((s) => !s.hidden);
        const body = withThemeColors(doc, () => slides
            .map((slide, i) => slideToHtml(slide, i + 1, pageW, pageH))
            .join('\n'));
        const title = options.title ?? 'PPTX';
        return `<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>${esc(title)}</title>
<style>
html,body{margin:0;padding:0;background:#eef0f3}
body{padding:24px 12px;font-family:${FONT_STACK}}
.jt-slides{display:flex;flex-direction:column;align-items:center;gap:24px}
.jt-slide{position:relative;overflow:hidden;background:#fff;box-shadow:0 1px 6px rgba(0,0,0,.2);flex:none}
.jt-slide>svg,.jt-slide>.jt-text,.jt-slide>.jt-img,.jt-slide>.jt-img-box,.jt-slide>.jt-table,.jt-slide>.jt-group,.jt-slide>div{position:absolute}
.jt-group>svg,.jt-group>.jt-text,.jt-group>.jt-img,.jt-group>.jt-img-box,.jt-group>.jt-table,.jt-group>.jt-group,.jt-group>div{position:absolute}
.jt-text{display:flex;flex-direction:column;box-sizing:border-box;overflow:visible;word-break:break-word}
.jt-text>.jt-p{margin:0;min-height:1em;flex-shrink:0}
.jt-bullet{display:inline-block;min-width:1.2em;text-align:left;white-space:pre}
.jt-bullet-img img{display:inline-block;max-width:none;object-fit:contain}
.jt-img{object-fit:fill}
.jt-table{border-collapse:collapse;table-layout:fixed}
.jt-table td,.jt-table th{border:1px solid #d0d0d0;padding:2px 6px;font-size:12pt;text-align:left;vertical-align:top}
@media print{html,body{background:#fff;padding:0}.jt-slides{gap:0}.jt-slide{box-shadow:none;page-break-after:always;margin:0}}
</style>
</head>
<body>
<div class="jt-slides">
${body}
</div>
</body>
</html>
`;
    }
    catch (error) {
        throw wrapError('pptx', '转 HTML', error);
    }
}

function runToRun(run) {
    const r = { text: run.text ?? '' };
    if (run.bold)
        r.bold = true;
    if (run.italic)
        r.italic = true;
    if (run.underline)
        r.underline = true;
    if (run.strike)
        r.strike = true;
    if (run.color)
        r.color = normalizeColor(run.color) || run.color;
    if (run.fontFaceEa || run.fontFace)
        r.font = run.fontFaceEa || run.fontFace;
    if (run.fontSize)
        r.size = run.fontSize;
    if (run.href)
        r.link = run.href;
    return r;
}
function paragraphToRuns(p) {
    const runs = p.runs?.length ? p.runs.map(runToRun) : p.text != null ? [{ text: p.text }] : [];
    return mergeRuns$2(runs);
}
function mergeRuns$2(runs) {
    const out = [];
    for (const r of runs) {
        const prev = out[out.length - 1];
        if (prev &&
            (prev.bold ?? false) === (r.bold ?? false) &&
            (prev.italic ?? false) === (r.italic ?? false) &&
            (prev.underline ?? false) === (r.underline ?? false) &&
            (prev.strike ?? false) === (r.strike ?? false) &&
            (prev.subscript ?? false) === (r.subscript ?? false) &&
            (prev.superscript ?? false) === (r.superscript ?? false) &&
            (prev.color ?? undefined) === (r.color ?? undefined) &&
            (prev.font ?? undefined) === (r.font ?? undefined) &&
            (prev.size ?? undefined) === (r.size ?? undefined) &&
            (prev.link ?? undefined) === (r.link ?? undefined)) {
            prev.text += r.text;
        }
        else {
            out.push({ ...r });
        }
    }
    return out;
}
function paragraphBullet(p) {
    if (!p.bullet)
        return false;
    if (p.bullet === true)
        return true;
    if (p.bullet === 'number')
        return 'number';
    if (typeof p.bullet === 'object')
        return p.bullet.type === 'number' ? 'number' : true;
    return true;
}
function pptxCellToBlockCell(cell) {
    if (cell.hMerge || cell.vMerge)
        return null;
    const runs = cell.paragraphs?.length ? cell.paragraphs.flatMap(paragraphToRuns) : cell.text ? [{ text: cell.text }] : [];
    const text = cell.text ?? cell.paragraphs?.map((p) => p.text ?? '').join('') ?? '';
    const detail = { type: 'cell', text };
    if (runs.length)
        detail.runs = runs;
    if (cell.colSpan && cell.colSpan > 1)
        detail.colspan = cell.colSpan;
    if (cell.rowSpan && cell.rowSpan > 1)
        detail.rowspan = cell.rowSpan;
    if (cell.fill)
        detail.fill = normalizeColor(cell.fill);
    const edge = (e) => {
        if (e === 'none')
            return { style: 'nil' };
        if (!e)
            return undefined;
        return { style: 'single', size: e.width, color: normalizeColor(e.color) };
    };
    const bs = cell.borders || (cell.border ? { top: cell.border, right: cell.border, bottom: cell.border, left: cell.border } : undefined);
    if (bs) {
        detail.borders = {
            top: edge(bs.top), right: edge(bs.right), bottom: edge(bs.bottom), left: edge(bs.left)
        };
    }
    return detail;
}
function pptxTableToBlock(table) {
    const rows = [];
    for (const row of table.rows) {
        const cells = [];
        for (const cell of row.cells) {
            const blockCell = pptxCellToBlockCell(cell);
            if (blockCell)
                cells.push(blockCell);
        }
        rows.push(cells);
    }
    return { type: 'table', header: rows.length > 1, rows };
}
function textCell(text) {
    return { type: 'cell', text };
}
function pptxChartToBlocks(chart) {
    const series = chart.series ?? [];
    const categories = chart.categories ?? [];
    const header = [textCell(''), ...series.map((s) => textCell(s.name ?? ''))];
    const rows = categories.map((cat, i) => [
        textCell(cat),
        ...series.map((s) => textCell(String(s.values?.[i] ?? '')))
    ]);
    const fallback = { type: 'table', header: true, rows: [header, ...rows] };
    return [
        {
            type: 'chart',
            spec: pptxChartToSpec(chart),
            width: Number(chart.width) || undefined,
            height: Number(chart.height) || undefined,
            caption: chart.title,
            fallback
        }
    ];
}
function flattenDiagram(nodes) {
    const out = [];
    const walk = (list) => {
        for (const node of list || []) {
            if (node.text)
                out.push(node.text);
            walk(node.children);
        }
    };
    walk(nodes);
    return out;
}
function firstTextOf(el) {
    for (const p of el.paragraphs ?? []) {
        const t = p.text ?? p.runs?.map((r) => r.text).join('') ?? '';
        if (t.trim())
            return t.trim();
    }
    if (el.text?.trim())
        return el.text.trim();
    if (el.runs?.length) {
        const t = el.runs.map((r) => r.text).join('').trim();
        if (t)
            return t;
    }
    return '';
}
function maxFontSizeOf(el) {
    let max = el.fontSize ?? 0;
    for (const p of el.paragraphs ?? []) {
        for (const r of p.runs ?? [])
            max = Math.max(max, r.fontSize ?? 0);
    }
    return max;
}
function textElementToItems(el) {
    const items = [];
    const paragraphs = el.paragraphs ?? [];
    if (paragraphs.length) {
        for (const p of paragraphs) {
            const runs = paragraphToRuns(p);
            if (!runs.some((r) => r.text))
                continue;
            items.push({ runs, bullet: paragraphBullet(p), align: p.align });
        }
    }
    if (!items.length) {
        if (el.runs?.length) {
            const runs = mergeRuns$2(el.runs.map(runToRun));
            if (runs.some((r) => r.text))
                items.push({ runs, bullet: false });
        }
        else if (el.text) {
            items.push({ runs: [{ text: el.text }], bullet: false });
        }
    }
    return items;
}
function elementToItems(el) {
    switch (el.type) {
        case 'text':
            return textElementToItems(el).map((item) => ({ kind: 'para', item }));
        case 'table':
            return [{ kind: 'block', block: pptxTableToBlock(el) }];
        case 'image': {
            const img = el;
            const src = img.data || img.src;
            if (!src)
                return [];
            const block = {
                type: 'image',
                src,
                width: img.width,
                height: img.height,
                caption: img.name
            };
            return [{ kind: 'block', block }];
        }
        case 'chart':
            return pptxChartToBlocks(el).map((block) => ({ kind: 'block', block }));
        case 'diagram': {
            const diagram = el;
            const texts = diagram.texts?.length ? diagram.texts : flattenDiagram(diagram.nodes);
            return texts
                .filter((t) => t && t.trim())
                .map((t) => ({ kind: 'para', item: { runs: [{ text: t }], bullet: true } }));
        }
        case 'group': {
            const out = [];
            for (const child of el.children ?? [])
                out.push(...elementToItems(child));
            return out;
        }
        case 'math': {
            const t = el.text?.trim();
            return t ? [{ kind: 'para', item: { runs: [{ text: t }], bullet: false } }] : [];
        }
        case 'gauge': {
            const g = el;
            const label = g.label ?? '值';
            return [{ kind: 'para', item: { runs: [{ text: `${label}: ${g.value}` }], bullet: false } }];
        }
        default:
            return [];
    }
}
function groupItems(items) {
    const blocks = [];
    let i = 0;
    while (i < items.length) {
        const item = items[i];
        if (item.bullet) {
            const ordered = item.bullet === 'number';
            const listItems = [];
            while (i < items.length && items[i].bullet) {
                listItems.push(items[i]);
                i++;
            }
            blocks.push({
                type: 'list',
                ordered,
                items: listItems.map((li) => li.runs)
            });
        }
        else {
            blocks.push({ type: 'paragraph', align: item.align, runs: item.runs });
            i++;
        }
    }
    return blocks;
}
function flowToBlocks(flow) {
    const blocks = [];
    let pending = [];
    const flush = () => {
        if (pending.length)
            blocks.push(...groupItems(pending));
        pending = [];
    };
    for (const item of flow) {
        if (item.kind === 'heading') {
            flush();
            blocks.push({ type: 'heading', level: item.level, text: item.text });
            continue;
        }
        pending.push(item.item);
    }
    flush();
    return blocks;
}
function dominantFontSize(elements) {
    const counter = new Map();
    for (const el of elements) {
        if (el.type !== 'text')
            continue;
        const size = maxFontSizeOf(el);
        if (!size)
            continue;
        counter.set(size, (counter.get(size) ?? 0) + 1);
    }
    let best = 0;
    let bestCount = 0;
    for (const [size, count] of counter) {
        if (count > bestCount || (count === bestCount && size < best)) {
            best = size;
            bestCount = count;
        }
    }
    return best;
}
function isSubHeading(el, bodySize) {
    if (!bodySize)
        return false;
    if (el.shapeType)
        return false;
    if (el.fill && el.fill !== 'none')
        return false;
    const paragraphs = el.paragraphs ?? [];
    if (paragraphs.length !== 1)
        return false;
    const p = paragraphs[0];
    if (p.bullet)
        return false;
    const runs = p.runs?.length ? p.runs : p.text != null ? [{ text: p.text }] : [];
    if (!runs.length)
        return false;
    if (!runs.every((r) => r.bold))
        return false;
    const text = runs.map((r) => r.text ?? '').join('').trim();
    if (!text || text.length > 30)
        return false;
    return maxFontSizeOf(el) > bodySize;
}
function slideToBlocks(slide, index, options) {
    const blocks = [];
    const elements = slide.elements ?? [];
    const textEls = elements.filter((e) => e.type === 'text');
    let titleEl = null;
    let best = -1;
    for (const el of textEls) {
        const fs = maxFontSizeOf(el);
        if (fs > best) {
            best = fs;
            titleEl = el;
        }
    }
    const titleText = titleEl ? firstTextOf(titleEl) : '';
    let titleRuns;
    if (titleEl) {
        for (const p of titleEl.paragraphs ?? []) {
            const rs = paragraphToRuns(p);
            if (rs.some((r) => r.text)) {
                titleRuns = rs;
                break;
            }
        }
    }
    if (options.slideHeading) {
        blocks.push({
            type: 'heading',
            level: options.headingLevel,
            text: titleText || `幻灯片 ${index + 1}`,
            runs: titleRuns
        });
    }
    const bodyEls = elements.filter((e) => e !== titleEl);
    bodyEls.sort((a, b) => (a.y ?? 0) - (b.y ?? 0) || (a.x ?? 0) - (b.x ?? 0));
    const flow = [];
    const direct = [];
    const bodySize = options.subHeadings ? dominantFontSize(bodyEls) : 0;
    for (const el of bodyEls) {
        if (bodySize && el.type === 'text' && isSubHeading(el, bodySize)) {
            flow.push({
                kind: 'heading',
                level: (options.headingLevel + 1),
                text: firstTextOf(el)
            });
            continue;
        }
        const elementItems = elementToItems(el);
        for (const ei of elementItems) {
            if (ei.kind === 'para')
                flow.push(ei);
            else
                direct.push(ei.block);
        }
    }
    blocks.push(...flowToBlocks(flow));
    blocks.push(...direct);
    if (options.notes && slide.notes?.trim()) {
        blocks.push({ type: 'quote', runs: [{ text: slide.notes.trim() }] });
    }
    return blocks;
}
function toDocxMetadata(metadata) {
    if (!metadata)
        return undefined;
    const result = {};
    const put = (key, value) => {
        if (value === undefined || value === null || value === '')
            return;
        result[key] = String(value);
    };
    put('title', metadata.title);
    put('subject', metadata.subject);
    put('author', metadata.author);
    put('keywords', metadata.keywords);
    put('description', metadata.description);
    put('lastModifiedBy', metadata.lastModifiedBy);
    return Object.keys(result).length ? result : undefined;
}
function pptxDocumentToBlocks(doc, options = {}) {
    const { slideHeading = true, headingLevel = 1, notes = false, slideSeparator = false, slidePageBreak = false, subHeadings = true } = options;
    const blocks = withThemeColors(doc, () => {
        const out = [];
        for (let i = 0; i < doc.slides.length; i++) {
            if (slideSeparator && i > 0)
                out.push({ type: 'hr' });
            if (slidePageBreak && i > 0)
                out.push({ type: 'pageBreak' });
            out.push(...slideToBlocks(doc.slides[i], i, { slideHeading, headingLevel, notes, subHeadings }));
        }
        return out;
    });
    return {
        blocks,
        metadata: toDocxMetadata(doc.metadata)
    };
}
async function pptxToBlocks(input, options = {}) {
    const { assets = 'both', ...rest } = options;
    const doc = (await pptxToStandard(input, { assets }));
    const result = pptxDocumentToBlocks(doc, rest);
    const parsedAssets = doc.assets;
    if (parsedAssets?.length)
        result.assets = parsedAssets;
    return result;
}

const DIAGRAM_LANGS = new Set(['mermaid']);
const MATH_LANGS = new Set(['math', 'tex', 'latex', 'katex']);
function attrValue(value) {
    return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function escapeText(value) {
    return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function mathBlockHtml(tex) {
    const trimmed = tex.trim();
    return `<pre class="pp-math-block" data-tex="${attrValue(trimmed)}"><code>${escapeText(trimmed)}</code></pre>`;
}
function mathInlineHtml(tex) {
    const trimmed = tex.trim();
    return `<code class="pp-math-inline" data-tex="${attrValue(trimmed)}">${escapeText(trimmed)}</code>`;
}
function markdownExtensions(options = {}) {
    const useMath = options.enableMath !== false;
    const useDiagrams = options.enableDiagrams !== false;
    const fenced = {
        name: 'ppDiagramCode',
        level: 'block',
        tokenizer(src) {
            if (!useDiagrams && !useMath)
                return undefined;
            const m = /^(?:```|~~~)([a-zA-Z0-9_-]*)[ \t]*\n([\s\S]*?)\n?(?:```|~~~)[ \t]*(?:\n|$)/.exec(src);
            if (!m)
                return undefined;
            const lang = (m[1] || '').toLowerCase();
            if (useDiagrams && DIAGRAM_LANGS.has(lang)) {
                return { type: 'ppDiagramCode', raw: m[0], lang, text: m[2] || '' };
            }
            if (useMath && MATH_LANGS.has(lang)) {
                return { type: 'ppDiagramCode', raw: m[0], lang, text: m[2] || '' };
            }
            return undefined;
        },
        renderer(token) {
            const generic = token;
            const text = generic.text || '';
            if (generic.lang && DIAGRAM_LANGS.has(generic.lang)) {
                return `<pre class="pp-mermaid" data-diagram="mermaid"><code class="language-mermaid">${escapeText(text)}</code></pre>\n`;
            }
            return `${mathBlockHtml(text)}\n`;
        }
    };
    const mathBlock = {
        name: 'ppMathBlock',
        level: 'block',
        start(src) {
            if (!useMath)
                return undefined;
            const candidates = [src.indexOf('$$'), src.indexOf('\\[')].filter((i) => i >= 0);
            return candidates.length ? Math.min(...candidates) : undefined;
        },
        tokenizer(src) {
            if (!useMath)
                return undefined;
            const dollar = /^\$\$([\s\S]+?)\$\$[ \t]*(?:\n|$)/.exec(src);
            if (dollar)
                return { type: 'ppMathBlock', raw: dollar[0], text: dollar[1] };
            const bracket = /^\\\[([\s\S]+?)\\][ \t]*(?:\n|$)/.exec(src);
            if (bracket)
                return { type: 'ppMathBlock', raw: bracket[0], text: bracket[1] };
            return undefined;
        },
        renderer(token) {
            return `${mathBlockHtml(token.text || '')}\n`;
        }
    };
    const mathInline = {
        name: 'ppMathInline',
        level: 'inline',
        start(src) {
            if (!useMath)
                return undefined;
            const candidates = [src.search(/\$(?!\s)/), src.search(/\\\(/)]
                .map((i) => (i >= 0 ? i : -1))
                .filter((i) => i >= 0);
            return candidates.length ? Math.min(...candidates) : undefined;
        },
        tokenizer(src) {
            if (!useMath)
                return undefined;
            const dollar = /^\$(?!\s)([^$\n]+?)(?<!\s)\$/.exec(src);
            if (dollar)
                return { type: 'ppMathInline', raw: dollar[0], text: dollar[1] };
            const paren = /^\\\(([^)\n]+?)\\\)/.exec(src);
            if (paren)
                return { type: 'ppMathInline', raw: paren[0], text: paren[1] };
            return undefined;
        },
        renderer(token) {
            return mathInlineHtml(token.text || '');
        }
    };
    return { extensions: [fenced, mathBlock, mathInline] };
}

const INLINE_TAGS = new Set(['strong', 'b', 'em', 'i', 'u', 's', 'strike', 'del', 'code', 'a', 'span', 'br', 'sub', 'sup', 'small', 'mark']);
const SKIP_TAGS = new Set(['script', 'style', 'noscript', 'meta', 'link', 'head', 'title', 'template', 'textarea']);
function attr(node, name) {
    if (node.nodeName === '#text' || node.nodeName === '#comment')
        return undefined;
    const attrs = node.attrs;
    return attrs?.find((a) => a.name === name)?.value;
}
function textContent(node) {
    if (node.nodeName === '#text')
        return node.value;
    const children = node.childNodes || [];
    return children.map(textContent).join('');
}
function mergeRuns$1(runs) {
    const out = [];
    for (const r of runs) {
        const last = out[out.length - 1];
        const same = !!last &&
            !!last.bold === !!r.bold &&
            !!last.italic === !!r.italic &&
            !!last.strike === !!r.strike &&
            !!last.underline === !!r.underline &&
            !!last.code === !!r.code &&
            (last.link ?? '') === (r.link ?? '') &&
            (last.color ?? '') === (r.color ?? '');
        if (same)
            last.text += r.text;
        else
            out.push({ ...r });
    }
    return out;
}
function inlineRuns(nodes, inherited = {}) {
    const runs = [];
    const push = (text, style) => {
        if (!text)
            return;
        runs.push({ text, ...inherited, ...style });
    };
    for (const node of nodes) {
        if (node.nodeName === '#text') {
            push(node.value, {});
        }
        else if (node.nodeName === 'br') {
            push('\n', {});
        }
        else if (SKIP_TAGS.has(node.nodeName)) {
        }
        else if (INLINE_TAGS.has(node.nodeName)) {
            const children = node.childNodes || [];
            const style = { ...inherited };
            switch (node.nodeName) {
                case 'strong':
                case 'b':
                    style.bold = true;
                    break;
                case 'em':
                case 'i':
                    style.italic = true;
                    break;
                case 'u':
                    style.underline = true;
                    break;
                case 's':
                case 'strike':
                case 'del':
                    style.strike = true;
                    break;
                case 'code':
                    style.code = true;
                    break;
                case 'a': {
                    const href = attr(node, 'href');
                    if (href)
                        style.link = href;
                    break;
                }
                case 'span': {
                    const css = attr(node, 'style') || '';
                    const colorMatch = /(?:^|;)\s*color\s*:\s*([^;]+)/i.exec(css);
                    if (colorMatch)
                        style.color = colorMatch[1].trim();
                    break;
                }
                default:
                    break;
            }
            runs.push(...inlineRuns(children, style));
        }
        else {
            push(textContent(node), inherited);
        }
    }
    return mergeRuns$1(runs);
}
function childrenToBlocks(nodes) {
    const blocks = [];
    let pending = [];
    const flush = () => {
        const runs = mergeRuns$1(pending);
        if (runs.length && runs.some((r) => r.text.trim())) {
            blocks.push({ type: 'paragraph', runs });
        }
        pending = [];
    };
    for (const node of nodes) {
        if (node.nodeName === '#text') {
            const text = node.value;
            if (text.trim())
                pending.push(...inlineRuns([node]));
            continue;
        }
        if (node.nodeName === '#comment')
            continue;
        if (SKIP_TAGS.has(node.nodeName))
            continue;
        const children = node.childNodes || [];
        const tag = node.nodeName;
        const level = /^h([1-6])$/.exec(tag);
        if (level) {
            flush();
            blocks.push({
                type: 'heading',
                level: Number(level[1]),
                runs: inlineRuns(children),
                text: textContent(node).trim()
            });
        }
        else if (tag === 'p') {
            flush();
            if (children.some((c) => c.nodeName === 'img')) {
                let buffer = [];
                for (const c of children) {
                    if (c.nodeName === 'img') {
                        if (buffer.length) {
                            blocks.push({ type: 'paragraph', runs: buffer });
                            buffer = [];
                        }
                        const img = imageFromNode(c);
                        if (img.src)
                            blocks.push(img);
                    }
                    else {
                        buffer.push(...inlineRuns([c]));
                    }
                }
                if (buffer.length)
                    blocks.push({ type: 'paragraph', runs: buffer });
            }
            else {
                blocks.push({ type: 'paragraph', runs: inlineRuns(children), text: textContent(node).trim() });
            }
        }
        else if (tag === 'ul' || tag === 'ol') {
            flush();
            blocks.push(listToBlock$1(node));
        }
        else if (tag === 'table') {
            flush();
            blocks.push(...tableToBlocks(node));
        }
        else if (tag === 'img') {
            flush();
            blocks.push(imageFromNode(node));
        }
        else if (tag === 'hr') {
            flush();
            blocks.push({ type: 'hr' });
        }
        else if (tag === 'pre') {
            flush();
            blocks.push(codeFromPre(node));
        }
        else if (tag === 'blockquote') {
            flush();
            blocks.push({
                type: 'quote',
                runs: inlineRuns(children),
                text: textContent(node).trim()
            });
        }
        else if (tag === 'br') {
            pending.push({ text: '\n' });
        }
        else if (tag === 'div' || tag === 'section' || tag === 'article' || tag === 'main' || tag === 'figure' || tag === 'header' || tag === 'footer') {
            blocks.push(...childrenToBlocks(children));
        }
        else {
            blocks.push(...childrenToBlocks(children));
        }
    }
    flush();
    return blocks;
}
function listToBlock$1(node) {
    const ordered = node.nodeName === 'ol';
    const items = [];
    for (const child of node.childNodes || []) {
        if (child.nodeName !== 'li')
            continue;
        const liChildren = child.childNodes || [];
        const inlineNodes = liChildren.filter((n) => n.nodeName !== 'ul' && n.nodeName !== 'ol');
        const nested = liChildren.find((n) => n.nodeName === 'ul' || n.nodeName === 'ol');
        const runs = inlineRuns(inlineNodes);
        if (nested)
            runs.children = listToBlock$1(nested).items;
        items.push(runs);
    }
    return { type: 'list', ordered, items };
}
function tableToBlocks(node) {
    const rows = [];
    let header = false;
    const spanRows = new Map();
    const addRow = (rowNode, isHead) => {
        const occupied = new Set();
        for (const [col, rest] of spanRows) {
            if (rest > 0)
                occupied.add(col);
        }
        const cells = cellsFromRow(rowNode, occupied);
        let col = 0;
        for (const cell of cells) {
            const detail = cell;
            const colspan = detail.colspan ?? 1;
            const rowspan = detail.rowspan ?? 1;
            if (rowspan > 1) {
                for (let c = col; c < col + colspan; c++)
                    spanRows.set(c, rowspan - 1);
            }
            col += colspan;
        }
        for (const [c, rest] of spanRows) {
            const left = rest - 1;
            if (left > 0)
                spanRows.set(c, left);
            else
                spanRows.delete(c);
        }
        if (isHead || (!rows.length && cells.some((c) => isHeaderCell(c))))
            header = true;
        rows.push(cells);
    };
    for (const child of node.childNodes || []) {
        if (child.nodeName === 'tr') {
            addRow(child, false);
        }
        else if (child.nodeName === 'thead' || child.nodeName === 'tbody' || child.nodeName === 'tfoot') {
            const isHead = child.nodeName === 'thead';
            for (const sub of child.childNodes || []) {
                if (sub.nodeName === 'tr')
                    addRow(sub, isHead);
            }
        }
    }
    if (!rows.length)
        return [];
    return [{ type: 'table', header, rows }];
}
function cellsFromRow(row, occupied) {
    const cells = [];
    let col = 0;
    for (const child of row.childNodes || []) {
        if (child.nodeName !== 'td' && child.nodeName !== 'th')
            continue;
        while (occupied.has(col))
            col++;
        const colspan = Math.max(1, parseInt(attr(child, 'colspan') || '1', 10) || 1);
        const rowspan = Math.max(1, parseInt(attr(child, 'rowspan') || '1', 10) || 1);
        const children = child.childNodes || [];
        const inherited = child.nodeName === 'th' ? { bold: true } : {};
        const runs = inlineRuns(children, inherited);
        const detail = { type: 'cell', text: runs.map((r) => r.text).join('') };
        if (runs.length)
            detail.runs = runs;
        if (colspan > 1)
            detail.colspan = colspan;
        if (rowspan > 1)
            detail.rowspan = rowspan;
        cells.push(detail);
        col += colspan;
    }
    return cells;
}
function isHeaderCell(cell) {
    if (Array.isArray(cell))
        return cell.some((r) => r.bold);
    const detail = cell;
    return !!detail.runs?.some((r) => r.bold);
}
function imageFromNode(node) {
    const src = attr(node, 'src') || '';
    const alt = attr(node, 'alt');
    const width = attr(node, 'width');
    const height = attr(node, 'height');
    const toPt = (v) => (v ? Math.round(Number(v) * 0.75 * 100) / 100 : undefined);
    return {
        type: 'image',
        src,
        alt: alt || undefined,
        width: toPt(width),
        height: toPt(height)
    };
}
function codeFromPre(node) {
    const codeChild = (node.childNodes || []).find((n) => n.nodeName === 'code');
    const target = codeChild || node;
    const text = textContent(target).replace(/\n$/, '');
    const className = attr(target, 'class') || '';
    const langMatch = /(?:language-|lang-)([a-zA-Z0-9_-]+)/.exec(className);
    return { type: 'code', text, language: langMatch?.[1] };
}
function htmlToBlocks(html) {
    const doc = parse5.parse(html);
    const root = doc.childNodes?.find((n) => n.nodeName === 'html');
    const body = root
        ? (root.childNodes || []).find((n) => n.nodeName === 'body')
        : doc;
    const target = body || root || doc;
    return childrenToBlocks(target.childNodes || []);
}

const markedCache = new Map();
function getMarked(options) {
    const key = `${options.gfm !== false}|${options.math !== false}|${options.diagrams !== false}`;
    let instance = markedCache.get(key);
    if (!instance) {
        instance = new marked.Marked({ gfm: options.gfm !== false, breaks: false });
        instance.use(markdownExtensions({ enableMath: options.math !== false, enableDiagrams: options.diagrams !== false }));
        markedCache.set(key, instance);
    }
    return instance;
}
function inlineToRuns(tokens) {
    const runs = [];
    const pushStyled = (source, style) => {
        for (const r of source)
            runs.push({ ...r, ...style });
    };
    for (const t of tokens || []) {
        switch (t.type) {
            case 'text':
            case 'escape':
                runs.push({ text: t.text });
                break;
            case 'codespan':
                runs.push({ text: t.text, code: true });
                break;
            case 'strong':
                pushStyled(inlineToRuns(t.tokens), { bold: true });
                break;
            case 'em':
                pushStyled(inlineToRuns(t.tokens), { italic: true });
                break;
            case 'del':
                pushStyled(inlineToRuns(t.tokens), { strike: true });
                break;
            case 'link':
                pushStyled(inlineToRuns(t.tokens), { link: t.href });
                break;
            case 'ppMathInline':
                runs.push({ text: `$${t.text ?? ''}$` });
                break;
            case 'br':
                runs.push({ text: '\n' });
                break;
            case 'image':
                runs.push({ text: `![${t.text ?? ''}](${t.href})` });
                break;
            case 'html':
                runs.push({ text: t.text });
                break;
            case 'space':
            default:
                break;
        }
    }
    return mergeRuns(runs);
}
function mergeRuns(runs) {
    const out = [];
    for (const r of runs) {
        const last = out[out.length - 1];
        if (last && sameStyle(last, r))
            last.text += r.text;
        else
            out.push({ ...r });
    }
    return out;
}
function sameStyle(a, b) {
    return (!!a.bold === !!b.bold &&
        !!a.italic === !!b.italic &&
        !!a.strike === !!b.strike &&
        !!a.code === !!b.code &&
        (a.link ?? '') === (b.link ?? '') &&
        (a.color ?? '') === (b.color ?? '') &&
        (a.font ?? '') === (b.font ?? ''));
}
function cellToRuns(cell) {
    return inlineToRuns(cell?.tokens);
}
function listItemInline(item) {
    const out = [];
    for (const t of item.tokens || []) {
        if (t.type === 'list')
            continue;
        if (t.type === 'paragraph') {
            const inner = t.tokens;
            if (inner?.length)
                out.push(...inner);
            else
                out.push(t);
            continue;
        }
        if (t.type === 'text') {
            const inner = t.tokens;
            if (inner?.length) {
                out.push(...inner);
                continue;
            }
        }
        out.push(t);
    }
    return out;
}
function listToBlock(token) {
    const itemRuns = (item) => {
        const inline = listItemInline(item);
        const nested = (item.tokens || []).find((t) => t.type === 'list');
        const runs = inlineToRuns(inline);
        if (nested)
            runs.children = nested.items.map(itemRuns);
        return runs;
    };
    return { type: 'list', ordered: token.ordered, items: token.items.map(itemRuns) };
}
const TOC_PAGE_PATTERN = /<!--\s*page:(\d+)\s*-->/;
function listToTocEntries(list, level = 1) {
    const entries = [];
    for (const item of list.items) {
        const inline = listItemInline(item);
        const nested = (item.tokens || []).find((t) => t.type === 'list');
        const raw = tokensToText(inline);
        const pageMatch = TOC_PAGE_PATTERN.exec(raw);
        const text = raw.replace(TOC_PAGE_PATTERN, '').trim();
        if (text) {
            entries.push({
                level,
                text,
                page: pageMatch ? Number(pageMatch[1]) : undefined
            });
        }
        if (nested)
            entries.push(...listToTocEntries(nested, level + 1));
    }
    return entries;
}
function paragraphToBlocks(token) {
    const inline = token.tokens || [];
    if (!inline.some((t) => t.type === 'image')) {
        const runs = inlineToRuns(inline);
        return [{ type: 'paragraph', runs, text: token.text }];
    }
    const blocks = [];
    let buffer = [];
    const flush = () => {
        if (!buffer.length)
            return;
        const runs = inlineToRuns(buffer);
        if (runs.length) {
            blocks.push({ type: 'paragraph', runs, text: runs.map((r) => r.text).join('') });
        }
        buffer = [];
    };
    for (const t of inline) {
        if (t.type === 'image') {
            flush();
            blocks.push(imageBlockFromToken(t));
            continue;
        }
        buffer.push(t);
    }
    flush();
    if (!blocks.length)
        blocks.push({ type: 'paragraph', runs: [], text: token.text });
    return blocks;
}
function imageBlockFromToken(token) {
    const title = token.title || '';
    const w = /(?:^|,)w=([\d.]+)/.exec(title);
    const h = /(?:^|,)h=([\d.]+)/.exec(title);
    return {
        type: 'image',
        src: token.href || '',
        alt: token.text || undefined,
        width: w ? Number(w[1]) : undefined,
        height: h ? Number(h[1]) : undefined
    };
}
function tokensToText(tokens) {
    let text = '';
    for (const t of tokens || []) {
        if ('text' in t && typeof t.text === 'string')
            text += t.text;
        else if ('tokens' in t && Array.isArray(t.tokens)) {
            text += tokensToText(t.tokens);
        }
        else if ('items' in t && Array.isArray(t.items)) {
            text += t.items.map((i) => tokensToText(i.tokens)).join(' ');
        }
    }
    return text;
}
function markdownToBlocks(markdown, options = {}) {
    const text = typeof markdown === 'string' ? markdown : new TextDecoder('utf-8').decode(markdown);
    const tokens = getMarked(options).lexer(text);
    const blocks = [];
    let expectToc = false;
    for (const token of tokens) {
        switch (token.type) {
            case 'heading': {
                const t = token;
                const runs = t.tokens ? inlineToRuns(t.tokens) : undefined;
                blocks.push({
                    type: 'heading',
                    level: Math.min(6, Math.max(1, t.depth || 1)),
                    text: t.text,
                    runs
                });
                break;
            }
            case 'paragraph': {
                blocks.push(...paragraphToBlocks(token));
                break;
            }
            case 'list': {
                const listToken = token;
                if (expectToc) {
                    expectToc = false;
                    const entries = listToTocEntries(listToken);
                    if (entries.length) {
                        blocks.push({ type: 'toc', entries });
                        break;
                    }
                }
                blocks.push(listToBlock(listToken));
                break;
            }
            case 'table': {
                const t = token;
                const rows = [];
                const header = (t.header || []).map((c) => [cellToRuns(c)]);
                rows.push(header);
                for (const row of t.rows || []) {
                    const cells = (row.cells || row.tokens || []);
                    rows.push(cells.map((c) => cellToRuns(c)));
                }
                blocks.push({ type: 'table', header: true, rows });
                break;
            }
            case 'code': {
                const t = token;
                blocks.push({ type: 'code', text: t.text, language: t.lang || undefined });
                break;
            }
            case 'hr':
                blocks.push({ type: 'hr' });
                break;
            case 'blockquote': {
                const t = token;
                blocks.push({ type: 'quote', runs: inlineToRuns(t.tokens), text: tokensToText(t.tokens) });
                break;
            }
            case 'ppDiagramCode': {
                const t = token;
                blocks.push({ type: 'code', text: t.text || '', language: t.lang || 'mermaid' });
                break;
            }
            case 'ppMathBlock': {
                const t = token;
                blocks.push({ type: 'code', text: t.text || '', language: 'math' });
                break;
            }
            case 'html': {
                const html = token.text;
                if (/^\s*<!--\s*toc\s*-->\s*$/i.test(html)) {
                    expectToc = true;
                    break;
                }
                blocks.push(...htmlToBlocks(html));
                break;
            }
            case 'space':
            default:
                break;
        }
    }
    return blocks;
}
function escapeMarkdownText(text) {
    return text.replace(/([\\`*_[\]~<>])/g, '\\$1');
}
function escapeLineStart(line) {
    return line.replace(/^(\s*)(#{1,6}\s|[-+]\s|\d+[.)]\s|>\s)/, (_all, space, mark) => space + mark.replace(/([#\-+.>])/g, '\\$1'));
}
function plainTextOf(block) {
    const runsText = (runs, text) => runs ? runs.map((r) => r.text ?? '').join('') : (text ?? '');
    switch (block.type) {
        case 'paragraph':
            return runsText(block.runs, block.text);
        case 'heading':
            return runsText(block.runs, block.text);
        case 'quote':
            return runsText(block.runs, block.text);
        case 'code':
            return block.text ?? '';
        case 'list':
            return block.items
                .map((item) => (Array.isArray(item) ? item.map((r) => r.text ?? '').join('') : String(item)))
                .join('\n');
        case 'textBox':
            return block.paragraphs.map(plainTextOf).join('\n');
        default:
            return '';
    }
}
function hasIndentedCode(text) {
    return /^(?: {4,}|\t)/.test(text);
}
function fencedCode(text) {
    const marks = text.match(/`+/g) || [];
    let longest = 0;
    for (const mark of marks)
        longest = Math.max(longest, mark.length);
    const fence = '`'.repeat(Math.max(3, longest + 1));
    return `${fence}\n${text.replace(/\n$/, '')}\n${fence}`;
}
function mdStyleKey(r) {
    return [r.bold ? 'B' : '', r.italic ? 'I' : '', r.strike ? 'S' : '', r.code ? 'C' : '', r.link ?? ''].join('|');
}
function runsToMarkdown(runs, inlineCode) {
    const merged = [];
    for (const r of runs) {
        const last = merged[merged.length - 1];
        if (last && mdStyleKey(last) === mdStyleKey(r))
            last.text += r.text;
        else
            merged.push({ ...r });
    }
    let out = '';
    for (const r of merged) {
        if (!r.text)
            continue;
        let text = r.text;
        if (r.code && inlineCode) {
            text = '`' + text + '`';
        }
        else {
            text = escapeMarkdownText(text);
            if (r.bold)
                text = `**${text}**`;
            if (r.italic)
                text = `*${text}*`;
            if (r.strike)
                text = `~~${text}~~`;
        }
        if (r.link)
            text = `[${text}](${r.link})`;
        out += text;
    }
    return out;
}
function quoteToMarkdown(runs, text, inlineCode) {
    const content = runs ? runsToMarkdown(runs, inlineCode) : text ?? '';
    if (!content.trim())
        return '>';
    return content
        .split('\n')
        .map((line) => (line.trim() ? '> ' + line : '>'))
        .join('\n');
}
function cellToText(cell, inlineCode) {
    if (cell == null)
        return '';
    if (typeof cell === 'string')
        return cell;
    if (Array.isArray(cell))
        return runsToMarkdown(cell, inlineCode);
    const detail = cell;
    if (detail.blocks)
        return detail.blocks.map((b) => blockToMarkdownLine(b, inlineCode)).join(' ');
    if (detail.runs)
        return runsToMarkdown(detail.runs, inlineCode);
    return detail.text ?? '';
}
function blockToMarkdownLine(block, inlineCode, imageSizes = true) {
    switch (block.type) {
        case 'heading':
            return `${'#'.repeat(block.level)} ${block.runs ? runsToMarkdown(block.runs, inlineCode) : block.text ?? ''}`;
        case 'paragraph': {
            const raw = plainTextOf(block);
            if (hasIndentedCode(raw))
                return fencedCode(raw);
            const line = block.runs ? runsToMarkdown(block.runs, inlineCode) : block.text ?? '';
            return escapeLineStart(line);
        }
        case 'code':
            return '```' + (block.language || '') + '\n' + block.text + '\n```';
        case 'image': {
            const size = imageSizes
                ? [
                    block.width ? `w=${block.width}` : '',
                    block.height ? `h=${block.height}` : ''
                ]
                    .filter(Boolean)
                    .join(',')
                : '';
            return `![${block.alt ?? ''}](${imageSrc(block)}${size ? ` "${size}"` : ''})`;
        }
        case 'hr':
            return '---';
        case 'pageBreak':
            return '<!-- pagebreak -->';
        case 'quote':
            return quoteToMarkdown(block.runs, block.text, inlineCode);
        case 'textBox': {
            const paragraphs = block.paragraphs.map(plainTextOf);
            if (paragraphs.length > 1)
                return fencedCode(paragraphs.join('\n'));
            const text = paragraphs[0] ?? '';
            if (hasIndentedCode(text))
                return fencedCode(text);
            return escapeLineStart(text);
        }
        case 'list':
            return renderList(block, 0, inlineCode);
        case 'table':
            return renderTable(block, inlineCode);
        case 'toc':
            return renderToc(block);
        default:
            return '';
    }
}
function renderToc(toc) {
    const lines = ['<!-- toc -->'];
    for (const entry of toc.entries) {
        const level = Math.min(Math.max(entry.level || 1, 1), 6);
        const indent = '  '.repeat(level - 1);
        const page = entry.page !== undefined ? ` <!-- page:${entry.page} -->` : '';
        const text = escapeLineStart(escapeMarkdownText(entry.text));
        lines.push(`${indent}- ${text}${page}`);
    }
    return lines.join('\n');
}
function renderList(list, depth, inlineCode) {
    const indent = '  '.repeat(depth);
    const lines = [];
    let ordinal = 0;
    list.items.forEach((item) => {
        const runs = typeof item === 'string' ? [{ text: item }] : item;
        const children = runs.children;
        const text = runsToMarkdown(runs, inlineCode);
        if (!text.trim() && !children?.length)
            return;
        ordinal += 1;
        const prefix = list.ordered ? `${ordinal}. ` : '- ';
        lines.push(indent + prefix + text);
        if (children?.length) {
            const childList = {
                type: 'list',
                ordered: list.ordered,
                items: children
            };
            lines.push(renderList(childList, depth + 1, inlineCode));
        }
    });
    return lines.join('\n');
}
function cellDetail(cell) {
    if (!cell || typeof cell === 'string' || Array.isArray(cell))
        return undefined;
    return cell;
}
function tableHasMerges(table) {
    return table.rows.some((row) => row.some((cell) => {
        const detail = cellDetail(cell);
        return !!detail && ((detail.colspan ?? 1) > 1 || (detail.rowspan ?? 1) > 1);
    }));
}
function renderTable(table, inlineCode) {
    if (!table.rows.length)
        return '';
    if (tableHasMerges(table))
        return renderTableHtml(table, inlineCode);
    const cell = (c) => cellToText(c, inlineCode).replace(/\|/g, '\\|').replace(/\n/g, ' ');
    const width = table.rows.reduce((max, row) => Math.max(max, row.length), 0);
    const line = (row) => {
        const cells = row.slice(0, width);
        while (cells.length < width)
            cells.push('');
        return '| ' + cells.map(cell).join(' | ') + ' |';
    };
    const header = line(table.rows[0]);
    const divider = '| ' + Array.from({ length: width }, () => '---').join(' | ') + ' |';
    return [header, divider, ...table.rows.slice(1).map(line)].join('\n');
}
function cellPlainText(cell) {
    if (cell == null)
        return '';
    if (typeof cell === 'string')
        return cell;
    if (Array.isArray(cell))
        return cell.map((r) => r.text ?? '').join('');
    const detail = cell;
    if (detail.blocks)
        return detail.blocks.map((b) => plainTextOf(b)).join(' ');
    if (detail.runs)
        return detail.runs.map((r) => r.text ?? '').join('');
    return detail.text ?? '';
}
function renderTableHtml(table, inlineCode) {
    const esc = (text) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const lines = ['<table>'];
    table.rows.forEach((row, rowIndex) => {
        const tag = table.header && rowIndex === 0 ? 'th' : 'td';
        lines.push('  <tr>');
        for (const cell of row) {
            const detail = cellDetail(cell);
            const colspan = detail?.colspan && detail.colspan > 1 ? ` colspan="${detail.colspan}"` : '';
            const rowspan = detail?.rowspan && detail.rowspan > 1 ? ` rowspan="${detail.rowspan}"` : '';
            lines.push(`    <${tag}${colspan}${rowspan}>${esc(cellPlainText(cell))}</${tag}>`);
        }
        lines.push('  </tr>');
    });
    lines.push('</table>');
    return lines.join('\n');
}
let assetLookup;
function imageSrc(image, assets) {
    const table = assets ?? assetLookup;
    if (typeof image.src === 'string')
        return image.src;
    if (image.assetId) {
        const asset = table ? getAsset(table, image.assetId) : undefined;
        if (asset)
            return assetToDataUrl(asset);
    }
    return '';
}
function normalizeAssets(assets) {
    if (!assets)
        return undefined;
    if (Array.isArray(assets))
        return toAssetTable(assets);
    const explicit = assets.assets;
    if (explicit)
        return toAssetTable(explicit);
    return assets;
}
function blocksToMarkdown(input, options = {}) {
    const blocks = Array.isArray(input) ? input : input.blocks || [];
    const inlineCode = options.inlineCode !== false;
    const imageSizes = options.imageSizes !== false;
    const assetTable = normalizeAssets(options.assets);
    assetLookup = assetTable;
    try {
        const parts = [];
        for (const block of blocks) {
            const text = blockToMarkdownLine(block, inlineCode, imageSizes);
            if (parts.length)
                parts.push('');
            parts.push(text);
        }
        return parts.join('\n').replace(/\n{3,}/g, '\n\n').trim();
    }
    finally {
        assetLookup = undefined;
    }
}

const DXA_PER_PT = 20;
const HALF_PT_PER_PT = 2;
const PAGE_SIZES = {
    A3: { width: 16838, height: 23811 },
    A4: { width: 11906, height: 16838 },
    A5: { width: 8391, height: 11906 },
    Letter: { width: 12240, height: 15840 },
    Legal: { width: 12240, height: 20160 }
};
const QUOTE_STYLE_ID = 'PptQuote';
const tocStyleId = (level) => `PptToc${level}`;
const ORDERED_REF = 'ppt-ordered';
const UNORDERED_REF = 'ppt-unordered';
const LIST_INDENT = [
    { left: 720, hanging: 360 },
    { left: 1440, hanging: 360 },
    { left: 2160, hanging: 360 },
    { left: 2880, hanging: 360 }
];
function bareColor(value) {
    if (!value)
        return undefined;
    const hex = value.replace(/^#/, '').trim();
    return /^[0-9a-fA-F]{6}$/.test(hex) ? hex.toUpperCase() : undefined;
}
function buildTextRun(run, style) {
    const options = {
        text: run.text,
        bold: run.bold,
        italics: run.italic,
        underline: run.underline ? {} : undefined,
        strike: run.strike,
        subScript: run.subscript,
        superScript: run.superscript,
        color: bareColor(run.color),
        size: run.size !== undefined ? Math.round(run.size * HALF_PT_PER_PT) : style.size !== undefined ? Math.round(style.size * HALF_PT_PER_PT) : undefined,
        font: run.font || style.font,
        highlight: run.highlight ? bareColor(run.highlight) : undefined
    };
    const textRun = new docx.TextRun(options);
    if (run.link)
        return new docx.ExternalHyperlink({ children: [textRun], link: run.link });
    return textRun;
}
function buildChildren(text, runs, style) {
    if (runs && runs.length)
        return runs.map((run) => buildTextRun(run, style));
    if (text !== undefined && text !== '')
        return [new docx.TextRun({ text, font: style.font, size: style.size !== undefined ? style.size * HALF_PT_PER_PT : undefined })];
    return [];
}
function buildParagraphLayout(block) {
    const layout = {};
    const para = block;
    if (para.align)
        layout.alignment = docx.AlignmentType[para.align.toUpperCase()];
    const indent = {};
    if (para.indentLeft)
        indent.left = Math.round(para.indentLeft * DXA_PER_PT);
    if (para.indentRight)
        indent.right = Math.round(para.indentRight * DXA_PER_PT);
    if (para.indentFirstLine !== undefined) {
        if (para.indentFirstLine < 0)
            indent.hanging = Math.round(-para.indentFirstLine * DXA_PER_PT);
        else
            indent.firstLine = Math.round(para.indentFirstLine * DXA_PER_PT);
    }
    if (Object.keys(indent).length)
        layout.indent = indent;
    const spacing = {};
    if (para.spacingBefore)
        spacing.before = Math.round(para.spacingBefore * DXA_PER_PT);
    if (para.spacingAfter)
        spacing.after = Math.round(para.spacingAfter * DXA_PER_PT);
    if (para.lineSpacing && para.lineSpacing !== 1)
        spacing.line = Math.round(para.lineSpacing * 240);
    if (Object.keys(spacing).length)
        layout.spacing = spacing;
    return layout;
}
async function decodeImageSource(src, assets) {
    if (isAssetRef(src)) {
        const ref = src;
        const asset = ref.asset ?? getAsset(assets, ref.assetId);
        if (!asset)
            throw new ConvertError('生成 DOCX 失败：资源引用无法解析（未提供对应的 assets）', 'docx');
        return { bytes: asset.data, mime: asset.mime || 'png' };
    }
    if (typeof src !== 'string') {
        if (src instanceof Uint8Array || src instanceof ArrayBuffer) {
            const bytes = await toUint8Array(src);
            return { bytes, mime: sniffMime(bytes) };
        }
        if (src && typeof src.arrayBuffer === 'function') {
            const bytes = new Uint8Array(await src.arrayBuffer());
            return { bytes, mime: src.type || sniffMime(bytes) };
        }
        throw new TypeError('不支持的图片数据类型');
    }
    const match = /^data:([^;,]+)(;base64)?,([\s\S]*)$/.exec(src.trim());
    if (match) {
        const mime = match[1] || 'image/png';
        const payload = match[3] || '';
        const bytes = match[2] ? decodeBase64(payload) : new TextEncoder().encode(decodeURIComponent(payload));
        return { bytes, mime };
    }
    return { bytes: decodeBase64(src), mime: 'png' };
}
function decodeBase64(value) {
    const clean = value.replace(/\s+/g, '');
    const atobFn = globalThis.atob;
    const binary = atobFn ? atobFn(clean) : Buffer.from(clean, 'base64').toString('binary');
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++)
        bytes[i] = binary.charCodeAt(i) & 0xff;
    return bytes;
}
function toImageType(mime) {
    const m = mime.toLowerCase();
    if (m.includes('jpeg') || m.includes('jpg'))
        return 'jpg';
    if (m.includes('gif'))
        return 'gif';
    if (m.includes('bmp'))
        return 'bmp';
    if (m.includes('svg'))
        return 'svg';
    return 'png';
}
function ptToPx(pt, fallback) {
    if (!pt || pt <= 0)
        return fallback;
    return Math.round((pt * 96) / 72);
}
async function buildImageParagraphs(block, style, assets) {
    const paragraphs = [];
    try {
        const { bytes, mime } = await decodeImageSource(block.src, assets);
        const width = ptToPx(block.width, 300);
        const height = ptToPx(block.height, Math.round((300 / 4) * 3));
        const imageRun = new docx.ImageRun({
            type: toImageType(mime),
            data: bytes,
            transformation: { width, height },
            altText: block.alt ? { name: block.alt, title: block.alt, description: block.alt } : undefined
        });
        paragraphs.push(new docx.Paragraph({ children: [imageRun], alignment: docx.AlignmentType.CENTER }));
    }
    catch {
        paragraphs.push(new docx.Paragraph({ children: [new docx.TextRun({ text: '[图片无法解析]', italics: true, color: '999999' })] }));
    }
    if (block.caption) {
        paragraphs.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: block.caption, size: 18, color: '666666' })],
            alignment: docx.AlignmentType.CENTER
        }));
    }
    return paragraphs;
}
function buildListParagraphs(block, level = 0) {
    const paragraphs = [];
    const indent = LIST_INDENT[Math.min(level, LIST_INDENT.length - 1)];
    for (const item of block.items) {
        const children = typeof item === 'string' ? undefined : item;
        const childItems = typeof item === 'string' ? [] : (item.children || []);
        paragraphs.push(new docx.Paragraph({
            children: buildChildren(typeof item === 'string' ? item : undefined, children, {}),
            numbering: {
                reference: block.ordered ? ORDERED_REF : UNORDERED_REF,
                level: Math.min(level, 8)
            },
            indent
        }));
        if (childItems.length) {
            paragraphs.push(...buildListParagraphs({ type: 'list', ordered: block.ordered, items: childItems }, level + 1));
        }
    }
    return paragraphs;
}
function toBorderStyle(style, fallback) {
    switch ((style || '').toLowerCase()) {
        case 'nil':
        case 'none':
            return docx.BorderStyle.NONE;
        case 'double':
            return docx.BorderStyle.DOUBLE;
        case 'dashed':
            return docx.BorderStyle.DASHED;
        case 'dotted':
            return docx.BorderStyle.DOTTED;
        case 'single':
            return docx.BorderStyle.SINGLE;
        default:
            return fallback;
    }
}
function toBorderSpec(border, fallbackStyle, fallbackColor) {
    if (!border)
        return { style: fallbackStyle, size: 2, color: fallbackColor };
    return {
        style: toBorderStyle(border.style, fallbackStyle),
        size: Math.max(2, Math.round((border.size !== undefined ? border.size : 1) * 8)),
        color: bareColor(border.color) || '000000'
    };
}
async function buildTableCell(cell, isHeader, borderSize, assets) {
    const fallbackStyle = borderSize > 0 ? docx.BorderStyle.SINGLE : docx.BorderStyle.NONE;
    const fallbackColor = 'CCCCCC';
    const fallbackBorder = () => ({
        style: fallbackStyle,
        size: Math.max(2, Math.round(borderSize * 8)),
        color: fallbackColor
    });
    const detail = typeof cell === 'object' && cell !== null && !Array.isArray(cell) ? cell : undefined;
    let children;
    if (detail) {
        const blocks = detail.blocks?.length
            ? detail.blocks
            : detail.runs?.length
                ? [{ type: 'paragraph', runs: detail.runs }]
                : detail.text
                    ? [{ type: 'paragraph', text: detail.text }]
                    : [];
        children = [];
        for (const block of blocks)
            children.push(...(await buildBlockParagraphs(block, {}, assets)));
    }
    else {
        const base = typeof cell === 'string'
            ? cell
                ? [{ text: cell }]
                : []
            : Array.isArray(cell)
                ? cell.filter(Boolean)
                : [];
        const styled = isHeader ? base.map((run) => ({ ...run, bold: true })) : base;
        children = [
            new docx.Paragraph({
                children: styled.length
                    ? styled.map((run) => buildTextRun(run, {}))
                    : [new docx.TextRun({ text: '' })]
            })
        ];
    }
    const options = {
        children: children.length ? children : [new docx.Paragraph({ children: [] })],
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        borders: {
            top: toBorderSpec(detail?.borders?.top, fallbackStyle, fallbackColor),
            bottom: toBorderSpec(detail?.borders?.bottom, fallbackStyle, fallbackColor),
            left: toBorderSpec(detail?.borders?.left, fallbackStyle, fallbackColor),
            right: toBorderSpec(detail?.borders?.right, fallbackStyle, fallbackColor)
        }
    };
    if (!detail) {
        options.borders = {
            top: fallbackBorder(),
            bottom: fallbackBorder(),
            left: fallbackBorder(),
            right: fallbackBorder()
        };
    }
    if (detail?.colspan && detail.colspan > 1)
        options.columnSpan = detail.colspan;
    if (detail?.rowspan && detail.rowspan > 1)
        options.rowSpan = detail.rowspan;
    if (detail?.fill)
        options.shading = { type: 'clear', fill: bareColor(detail.fill) || 'FFFFFF', color: 'auto' };
    if (detail?.verticalAlign) {
        options.verticalAlign =
            detail.verticalAlign === 'middle'
                ? docx.VerticalAlign.CENTER
                : detail.verticalAlign === 'bottom'
                    ? docx.VerticalAlign.BOTTOM
                    : docx.VerticalAlign.TOP;
    }
    if (detail?.width)
        options.width = { size: Math.round(detail.width * DXA_PER_PT), type: docx.WidthType.DXA };
    return new docx.TableCell(options);
}
async function buildTable(block, assets) {
    const borderSize = block.borderSize === undefined ? 1 : block.borderSize;
    const hasHeader = block.header !== false && block.rows.length > 0;
    const rows = await Promise.all(block.rows.map(async (row, rowIndex) => {
        const isHeaderRow = hasHeader && rowIndex === 0;
        const cells = await Promise.all(row.map((cell) => buildTableCell(cell ?? '', isHeaderRow, borderSize, assets)));
        return new docx.TableRow({
            tableHeader: isHeaderRow,
            children: cells
        });
    }));
    return new docx.Table({
        rows,
        width: block.columnWidths
            ? { size: Math.round(block.columnWidths.reduce((sum, w) => sum + w, 0) * DXA_PER_PT), type: docx.WidthType.DXA }
            : { size: 100, type: docx.WidthType.PERCENTAGE }
    });
}
const HEADING_MAP = {
    1: docx.HeadingLevel.HEADING_1,
    2: docx.HeadingLevel.HEADING_2,
    3: docx.HeadingLevel.HEADING_3,
    4: docx.HeadingLevel.HEADING_4,
    5: docx.HeadingLevel.HEADING_5,
    6: docx.HeadingLevel.HEADING_6
};
async function buildBlockParagraphs(block, style, assets) {
    switch (block.type) {
        case 'heading': {
            const layout = buildParagraphLayout(block);
            const heading = HEADING_MAP[block.level] || docx.HeadingLevel.HEADING_1;
            const children = buildChildren(block.text, block.runs, style);
            return [new docx.Paragraph({ children, heading, ...layout })];
        }
        case 'paragraph': {
            const layout = buildParagraphLayout(block);
            return [new docx.Paragraph({ children: buildChildren(block.text, block.runs, style), ...layout })];
        }
        case 'quote': {
            const base = block.runs && block.runs.length
                ? block.runs
                : block.text !== undefined
                    ? [{ text: block.text }]
                    : [];
            return [
                new docx.Paragraph({
                    style: QUOTE_STYLE_ID,
                    children: base.map((run) => buildTextRun({ ...run, italic: true }, style)),
                    indent: { left: 360 },
                    border: { left: { style: docx.BorderStyle.SINGLE, size: 12, color: 'CCCCCC', space: 8 } }
                })
            ];
        }
        case 'code': {
            const lines = String(block.text ?? '').split('\n');
            const font = 'Consolas';
            return lines.map((line) => new docx.Paragraph({
                children: [new docx.TextRun({ text: line || ' ', font, size: 20 })],
                shading: { type: 'clear', fill: 'F5F5F5', color: 'auto' },
                spacing: { before: 0, after: 0 }
            }));
        }
        case 'list':
            return buildListParagraphs(block);
        case 'table':
            return [await buildTable(block, assets)];
        case 'textBox': {
            const inner = [];
            for (const paragraph of block.paragraphs) {
                inner.push(...(await buildBlockParagraphs(paragraph, style, assets)));
            }
            return inner.length ? inner : [new docx.Paragraph({ children: [] })];
        }
        case 'image':
            return buildImageParagraphs(block, style, assets);
        case 'pageBreak':
            return [new docx.Paragraph({ children: [new docx.PageBreak()] })];
        case 'toc':
            return block.entries.map((entry) => {
                const level = Math.min(Math.max(entry.level, 1), 3);
                const text = entry.page !== undefined ? `${entry.text}\t${entry.page}` : entry.text;
                return new docx.Paragraph({
                    style: tocStyleId(level),
                    children: [new docx.TextRun({ text })],
                    indent: { left: (level - 1) * 320 },
                    tabStops: entry.page !== undefined
                        ? [{ type: docx.TabStopType.RIGHT, position: 8300, leader: docx.LeaderType.DOT }]
                        : undefined
                });
            });
        case 'hr':
            return [
                new docx.Paragraph({
                    children: [],
                    border: { bottom: { style: docx.BorderStyle.SINGLE, size: 6, color: 'DDDDDD', space: 1 } },
                    spacing: { before: 120, after: 120 }
                })
            ];
        default:
            return [];
    }
}
function buildNumberingConfig() {
    const levels = (format, text) => LIST_INDENT.map((indent, level) => ({
        level,
        format,
        text: format === docx.LevelFormat.BULLET ? text : `%${level + 1}.`,
        alignment: docx.AlignmentType.START,
        style: { paragraph: { indent } }
    }));
    return [
        { reference: ORDERED_REF, levels: levels(docx.LevelFormat.DECIMAL, '%1.') },
        { reference: UNORDERED_REF, levels: levels(docx.LevelFormat.BULLET, '•') }
    ];
}
function buildMargin(margin) {
    const DEFAULT = 1440;
    if (typeof margin === 'number') {
        const value = Math.round(margin * DXA_PER_PT);
        return { top: value, right: value, bottom: value, left: value };
    }
    if (margin && typeof margin === 'object') {
        return {
            top: Math.round((margin.top ?? 72) * DXA_PER_PT),
            right: Math.round((margin.right ?? 72) * DXA_PER_PT),
            bottom: Math.round((margin.bottom ?? 72) * DXA_PER_PT),
            left: Math.round((margin.left ?? 72) * DXA_PER_PT)
        };
    }
    return { top: DEFAULT, right: DEFAULT, bottom: DEFAULT, left: DEFAULT };
}
async function generateDocx(input, options = {}) {
    const doc = Array.isArray(input) ? { blocks: input } : input;
    if (!doc || !Array.isArray(doc.blocks)) {
        throw new ConvertError('生成 DOCX 失败：缺少 blocks 内容', 'docx');
    }
    const rawAssets = options.assets ?? doc.assets;
    const assets = Array.isArray(rawAssets) ? toAssetTable(rawAssets) : (rawAssets || {});
    const style = {
        font: options.defaultFont,
        size: options.defaultSize !== undefined ? Math.round(options.defaultSize * HALF_PT_PER_PT) : undefined
    };
    const children = [];
    for (const block of doc.blocks) {
        if (!block || typeof block !== 'object')
            continue;
        children.push(...(await buildBlockParagraphs(block, style, assets)));
    }
    if (!children.length)
        children.push(new docx.Paragraph({ children: [] }));
    const landscape = (options.orientation ?? doc.section?.orientation) === 'landscape';
    const sizeKey = typeof options.pageSize === 'string' ? options.pageSize : 'A4';
    const rawSize = typeof options.pageSize === 'object' && options.pageSize
        ? {
            width: Math.round(options.pageSize.width * DXA_PER_PT),
            height: Math.round(options.pageSize.height * DXA_PER_PT)
        }
        : options.pageSize === undefined && doc.section
            ?
                {
                    width: Math.round(doc.section.width * DXA_PER_PT),
                    height: Math.round(doc.section.height * DXA_PER_PT)
                }
            : PAGE_SIZES[sizeKey] || PAGE_SIZES.A4;
    const size = landscape
        ? { width: rawSize.height, height: rawSize.width, orientation: 'landscape' }
        : { width: rawSize.width, height: rawSize.height, orientation: 'portrait' };
    const document = new docx.Document({
        styles: {
            paragraphStyles: [
                {
                    id: 'Heading1',
                    name: 'Heading 1',
                    basedOn: 'Normal',
                    next: 'Normal',
                    quickFormat: true,
                    run: { size: 44, bold: true, color: '1F2937' },
                    paragraph: { spacing: { before: 240, after: 160 } }
                },
                {
                    id: 'Heading2',
                    name: 'Heading 2',
                    basedOn: 'Normal',
                    next: 'Normal',
                    quickFormat: true,
                    run: { size: 32, bold: true, color: '1F2937' },
                    paragraph: { spacing: { before: 200, after: 120 } }
                },
                {
                    id: QUOTE_STYLE_ID,
                    name: 'Quote',
                    basedOn: 'Normal',
                    quickFormat: true,
                    run: { italics: true, color: '4A4A4A' },
                    paragraph: { indent: { left: 360 }, spacing: { before: 120, after: 120 } }
                },
                ...[1, 2, 3].map((level) => ({
                    id: tocStyleId(level),
                    name: `toc ${level}`,
                    basedOn: 'Normal',
                    quickFormat: true,
                    run: { size: 22 },
                    paragraph: { spacing: { before: 0, after: 0, line: 300 } }
                }))
            ]
        },
        creator: doc.metadata?.creator || doc.metadata?.author || 'ppt-parser',
        title: doc.metadata?.title,
        subject: doc.metadata?.subject,
        description: doc.metadata?.description,
        keywords: doc.metadata?.keywords,
        lastModifiedBy: doc.metadata?.lastModifiedBy,
        numbering: { config: buildNumberingConfig() },
        sections: [
            {
                properties: {
                    page: {
                        size,
                        margin: buildMargin(options.margin ??
                            (doc.section
                                ? {
                                    top: doc.section.marginTop,
                                    right: doc.section.marginRight,
                                    bottom: doc.section.marginBottom,
                                    left: doc.section.marginLeft
                                }
                                : undefined))
                    }
                },
                children
            }
        ]
    });
    const arrayBuffer = await docx.Packer.toArrayBuffer(document);
    return new Uint8Array(arrayBuffer);
}

class ConverterRegistry {
    constructor() {
        this.map = new Map();
    }
    register(converter) {
        this.map.set(converter.id, converter);
        return this;
    }
    get(id) {
        return this.map.get(id);
    }
    has(id) {
        return this.map.has(id);
    }
    list() {
        return Array.from(this.map.values());
    }
    remove(id) {
        return this.map.delete(id);
    }
}
const converters = new ConverterRegistry();
function registerConverter(converter) {
    converters.register(converter);
}
function formatOfTarget(target) {
    switch (target) {
        case 'docx':
        case 'pdf':
        case 'md':
            return target;
        case 'html':
            return 'txt';
        default:
            return 'txt';
    }
}
async function convert(input, target, options = {}) {
    const converter = converters.get(target);
    if (!converter) {
        throw new TypeError(`未注册的转换目标：${target}。已注册：${converters.list().map((c) => c.id).join(', ')}`);
    }
    try {
        return await converter.convert(input, options);
    }
    catch (error) {
        throw wrapError(formatOfTarget(target), `转换(${target})`, error);
    }
}
const docxConverter = {
    id: 'docx',
    ext: 'docx',
    label: 'Word Document',
    async convert(input, options = {}) {
        const { blocks, metadata, assets } = await pptxToBlocks(input, { slideSeparator: false, slidePageBreak: true });
        const plain = blocks.some((b) => b.type === 'chart')
            ? blocks.flatMap((b) => (b.type === 'chart' ? (b.fallback ? [b.fallback] : []) : [b]))
            : blocks;
        const pool = [...(options.assets || []), ...(assets || [])];
        const doc = { metadata, blocks: plain, assets: pool.length ? pool : undefined };
        return generateDocx(doc);
    }
};
const mdConverter = {
    id: 'md',
    ext: 'md',
    label: 'Markdown',
    async convert(input, options = {}) {
        const { blocks, metadata, assets } = await pptxToBlocks(input, {
            slideSeparator: true
        });
        const plain = blocks.some((b) => b.type === 'chart')
            ? blocks.flatMap((b) => (b.type === 'chart' ? (b.fallback ? [b.fallback] : []) : [b]))
            : blocks;
        const pool = [...(options.assets || []), ...(assets || [])];
        const doc = { metadata, blocks: plain, assets: pool.length ? pool : undefined };
        return blocksToMarkdown(doc, { assets: pool, imageSizes: false });
    }
};
const pdfConverter = {
    id: 'pdf',
    ext: 'pdf',
    label: 'PDF Document',
    async convert(input, options = {}) {
        return pptxToPdf(input, {
            fonts: options.fonts,
            autoFont: options.autoFont ?? true
        });
    }
};
const htmlConverter = {
    id: 'html',
    ext: 'html',
    label: 'HTML Web Page',
    async convert(input, options = {}) {
        return pptxToSlideHtml(input, { title: options.title });
    }
};
converters
    .register(docxConverter)
    .register(mdConverter)
    .register(pdfConverter)
    .register(htmlConverter);
async function convertPptx(input, target, options = {}) {
    return convert(input, target, options);
}

exports.ConvertError = ConvertError;
exports.convert = convert;
exports.convertPptx = convertPptx;
exports.converters = converters;
exports.registerConverter = registerConverter;
exports.wrapError = wrapError;
//# sourceMappingURL=convert.cjs.map
