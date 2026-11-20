/**
 * 二进制资源（图片及其它内嵌文件）统一模型
 *
 * 设计要点：
 *   1. **纯数据**：资源是可 JSON 序列化的普通对象（不含方法/类实例），
 *      因此解析结果可以直接 JSON.stringify 落库、传输、放进前端 state。
 *   2. **多形态输出**：同一份字节，按运行环境取用最合适的形式——
 *      浏览器要 Blob / data URL 直接渲染，Node 要 Buffer / 流直接落盘或转存对象存储。
 *   3. **位置锚点**：每个资源都带 anchor，指明它在文档中的归属
 *      （第几个块 / 哪张表哪个单元格 / 第几页幻灯片），实现「资源 ↔ 文档位置」双向定位。
 *   4. **引用优先**：文档结构（block/cell/slide）用 assetId 引用资源，
 *      字节只存一份，重复引用不重复占用内存。
 *
 * @module core/asset
 */

import { base64ToBytes, bytesToBase64, toUint8Array, type DocInput } from './input';

/**
 * 解析时二进制资源的处理方式
 *
 * - `inline`：内联到结构里（data URL / base64），浏览器直接可用，JSON 体积大
 * - `extract`：抽出到 `document.assets`，结构里只留 `assetId` 引用，适合后台存储与另存为
 * - `both`：两者都给，前端显示与后端取字节互不干扰
 * - `ignore`：完全丢弃
 */
export type DocAssetMode = 'inline' | 'extract' | 'both' | 'ignore';

/** 资源归属的文档范围 */
export type AssetScope = 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'bundle';

/** 锚点类型：资源挂在文档结构的哪一层 */
export type AssetAnchorKind =
    /** 文档正文块（DOCX block / PDF 页内元素） */
    | 'block'
    /** 表格单元格（XLSX / DOCX 表格） */
    | 'cell'
    /** 幻灯片（PPTX） */
    | 'slide'
    /** 页面（PDF） */
    | 'page'
    /** 其它（页眉页脚、批注、附件等） */
    | 'other';

/** 资源在文档中的位置 */
export interface AssetAnchor {
    /** 所属格式 */
    scope: AssetScope;
    /** 锚点层级 */
    kind: AssetAnchorKind;
    /** 该层级内的下标：block 下标 / 单元格行号 / 幻灯片序号（1 起）/ 页码（1 起） */
    index: number;
    /** 更细的定位串，如单元格 `Sheet1!B3`、页内 `p3/fig2` */
    path?: string;
    /** 供人类阅读的位置描述，如「第 3 页 / 第 2 个图」 */
    label?: string;
}

/**
 * 文档内的二进制资源
 *
 * 注意 `data` 是完整字节而非引用——它与文档结构分离存放，
 * 由 `assetId` 关联；这样同一张图被引用多次也只存一份。
 */
export interface DocAsset {
    /** 文档内唯一的稳定 ID，被 block.assetId 等字段引用 */
    id: string;
    /** 建议文件名（另存为时使用），含扩展名 */
    name: string;
    /** MIME 类型，如 image/png */
    mime: string;
    /** 原始字节（不复制则与解析缓存共享内存） */
    data: Uint8Array;
    /** 来源在容器内的部件路径，如 word/media/image1.png */
    part?: string;
    /** 在文档中的位置 */
    anchor: AssetAnchor;
    /** 显示宽度（磅），仅图片有意义 */
    width?: number;
    /** 显示高度（磅），仅图片有意义 */
    height?: number;
    /** 替代文本（alt），仅图片有意义 */
    alt?: string;
    /** 备注 */
    meta?: Record<string, unknown>;
}

/** 扩展名 -> MIME 兜底表（无法从部件路径判断时使用） */
const MIME_BY_EXT: Record<string, string> = {
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

/** 由部件路径或文件名推断 MIME */
export function mimeOf(nameOrPath: string, fallback = 'application/octet-stream'): string {
    const ext = (nameOrPath.split('.').pop() || '').toLowerCase();
    return MIME_BY_EXT[ext] || fallback;
}

/** 常见扩展名清单，用于 bundle 打包时补全文件名 */
const EXT_BY_MIME: Record<string, string> = {
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

/** 由 MIME 取建议扩展名 */
export function extOf(mime: string, fallback = ''): string {
    return EXT_BY_MIME[mime] ?? fallback;
}

/**
 * 生成文档内唯一的资源 ID
 * @param prefix - 业务前缀，通常是格式名，便于人读
 * @param index - 序号
 */
export function createAssetId(prefix: string, index: number): string {
    return `${prefix}-${String(index).padStart(4, '0')}`;
}

/**
 * 由字节与元信息构造资源对象
 * @param params - 资源字段
 */
export function createAsset(params: {
    id: string;
    data: Uint8Array;
    mime?: string;
    name?: string;
    part?: string;
    anchor: AssetAnchor;
    width?: number;
    height?: number;
    alt?: string;
    meta?: Record<string, unknown>;
}): DocAsset {
    // MIME 推断：显式值 > 部件路径扩展名 > 字节魔数
    const mime =
        params.mime ||
        mimeOf(params.part || params.name || '', '') ||
        sniffMime(params.data);
    const ext = extOf(mime, params.part?.split('.').pop() ? `.${params.part.split('.').pop()}` : '');
    const asset: DocAsset = {
        id: params.id,
        name: params.name || params.id + ext,
        mime,
        data: params.data,
        anchor: params.anchor
    };
    if (params.part) asset.part = params.part;
    if (params.width != null) asset.width = params.width;
    if (params.height != null) asset.height = params.height;
    if (params.alt) asset.alt = params.alt;
    if (params.meta) asset.meta = params.meta;
    return asset;
}

// ---------------------------------------------------------------------------
// 取用方式：同一份字节，按端取最合适的形态
// ---------------------------------------------------------------------------

/**
 * 解析 data URL，取出字节与 MIME
 * @param dataUrl - data URL 字符串，如 `data:image/png;base64,…`
 * @returns 字节与 MIME；不是合法 data URL 返回 null
 */
export function parseDataUrl(dataUrl: string): { bytes: Uint8Array; mime: string } | null {
    const match = /^data:([^;,]*)(;base64)?,([\s\S]*)$/.exec(dataUrl.trim());
    if (!match) return null;
    const mime = match[1] || 'application/octet-stream';
    const payload = match[3] || '';
    if (match[2]) {
        return { bytes: base64ToBytes(payload), mime };
    }
    // 非 base64 的 data URL 按 UTF-8 文本解码
    return { bytes: new TextEncoder().encode(decodeURIComponent(payload)), mime };
}

/**
 * 取 base64（不带 data URL 前缀）
 * @param asset - 资源对象或原始字节
 */
export function assetToBase64(asset: DocAsset | Uint8Array): string {
    return bytesToBase64(toBytes(asset));
}

/**
 * 取 data URL，可直接塞进 `<img src>` / CSS / 浏览器存储
 * @param asset - 资源对象或原始字节
 * @param mime - 覆盖 MIME（仅传字节时需要）
 */
export function assetToDataUrl(asset: DocAsset | Uint8Array, mime?: string): string {
    const bytes = toBytes(asset);
    const type = mime || ('mime' in asset ? asset.mime : undefined) || mimeOf('', 'application/octet-stream');
    return `data:${type};base64,${bytesToBase64(bytes)}`;
}

/**
 * 取 Blob，浏览器端用于 `<img>`、预览、下载
 * @param asset - 资源对象或原始字节
 */
export function assetToBlob(asset: DocAsset | Uint8Array, mime?: string): Blob {
    const bytes = toBytes(asset);
    const type = mime || ('mime' in asset ? asset.mime : undefined) || mimeOf('', 'application/octet-stream');
    // 复制一份，避免 Blob 引用后原字节被转移/复用导致内容错乱
    return new Blob([bytes.slice().buffer as ArrayBuffer], { type });
}

/**
 * 取 Node.js 可读流，直接 pipe 到文件、HTTP 响应或对象存储
 * @param asset - 资源对象或原始字节
 * @returns Node.js `Readable`；浏览器环境返回 `ReadableStream`
 *
 * @example
 * ```js
 * const { Readable } = await import('node:stream');
 * assetToNodeStream(asset).pipe(createWriteStream('out.png'));
 * ```
 */
export function assetToNodeStream(asset: DocAsset | Uint8Array): unknown {
    const bytes = toBytes(asset);
    const Ctor = (globalThis as { Readable?: new (b: Uint8Array) => unknown }).Readable;
    if (Ctor) return new Ctor(bytes);
    // 浏览器：ReadableStream 是 web 标准，可直接喂给 fetch/WHATWG Response
    const RS = (globalThis as { ReadableStream?: new (b: Uint8Array) => unknown }).ReadableStream;
    if (RS) return new RS(bytes.slice());
    throw new TypeError('jdoc: 当前环境既无 Node Readable 也无 ReadableStream，无法生成流');
}

/**
 * 取 Web ReadableStream（浏览器 / 现代 Node 18+ 均有）
 * @param asset - 资源对象或原始字节
 */
export function assetToWebStream(asset: DocAsset | Uint8Array): ReadableStream<Uint8Array> {
    const bytes = toBytes(asset).slice();
    const RS = (globalThis as { ReadableStream?: new (b: Uint8Array) => ReadableStream<Uint8Array> })
        .ReadableStream;
    // Node 18+ 与现代浏览器均有全局 ReadableStream
    if (RS) return new RS(bytes);
    throw new TypeError('jdoc: 当前环境不支持 ReadableStream，请改用 assetToNodeStream 或 assetToArrayBuffer');
}

/**
 * 取 ArrayBuffer（便于交给底层库 / postMessage 传输）
 * @param asset - 资源对象或原始字节
 */
export function assetToArrayBuffer(asset: DocAsset | Uint8Array): ArrayBuffer {
    const bytes = toBytes(asset);
    return bytes.slice().buffer as ArrayBuffer;
}

/** 内部：统一取字节 */
function toBytes(asset: DocAsset | Uint8Array): Uint8Array {
    return asset instanceof Uint8Array ? asset : asset.data;
}

// ---------------------------------------------------------------------------
// 资源表：文档与资源的一体化容器
// ---------------------------------------------------------------------------

/** 资源表：id -> 资源 */
export type AssetTable = Record<string, DocAsset>;

/**
 * 由资源数组构建 id -> 资源 的索引表
 * @param assets - 资源数组
 */
export function toAssetTable(assets: DocAsset[] | undefined): AssetTable {
    const table: AssetTable = {};
    for (const asset of assets || []) table[asset.id] = asset;
    return table;
}

/**
 * 按 id 取资源
 * @param table - 资源表，或带 assets 字段的解析结果
 * @param id - 资源 ID
 */
export function getAsset(
    source: AssetTable | DocAsset[] | { assets?: DocAsset[] } | undefined,
    id: string | undefined
): DocAsset | undefined {
    if (!source || !id) return undefined;
    // 数组形式：解析结果里的 assets 直接传进来
    if (Array.isArray(source)) return source.find((asset) => asset.id === id);
    // 带 assets 字段的解析结果
    if (Array.isArray((source as { assets?: DocAsset[] }).assets)) {
        return (source as { assets?: DocAsset[] }).assets!.find((asset) => asset.id === id);
    }
    return (source as AssetTable)[id];
}

/**
 * 汇总资源体积，用于日志与限额控制
 * @param assets - 资源数组
 */
export function summarizeAssets(assets: DocAsset[] | undefined): { count: number; bytes: number } {
    let bytes = 0;
    for (const asset of assets || []) bytes += asset.data?.byteLength || 0;
    return { count: assets?.length || 0, bytes };
}

// ---------------------------------------------------------------------------
// 生成侧：把「引用」解析成字节
// ---------------------------------------------------------------------------

/** 生成侧图片/二进制字段可接受的形态：字节、data URL、base64、资源 ID、资源对象 */
export type AssetSource = DocInput | DocAsset | { assetId: string } | { asset: DocAsset };

/**
 * 判断值是否为资源引用对象
 * @param value - 待判断值
 */
export function isAssetRef(value: unknown): value is { assetId: string } | { asset: DocAsset } {
    if (!value || typeof value !== 'object') return false;
    const record = value as Record<string, unknown>;
    return typeof record.assetId === 'string' || (!!record.asset && typeof record.asset === 'object');
}

/**
 * 判断值是否为资源对象（含字节与元信息）
 * @param value - 待判断值
 */
export function isDocAsset(value: unknown): value is DocAsset {
    if (!value || typeof value !== 'object') return false;
    const record = value as Record<string, unknown>;
    return typeof record.id === 'string' && record.data instanceof Uint8Array;
}

/**
 * 把生成侧的资源字段解析为字节
 *
 * 支持三种写法，让「解析 → 改 → 生成」与「手写 JSON」都能喂字节：
 *   - `{ type:'image', src: 'data:image/png;base64,…' }`  （内联，最通用）
 *   - `{ type:'image', src: { assetId:'docx-0001' } }`    （引用，字节复用）
 *   - `{ type:'image', src: <DocAsset 实例> }`           （直接给资源对象）
 *
 * @param value - 资源字段值
 * @param table - 可用资源表，用于解析 assetId 引用
 * @returns 字节与推断出的 MIME；无法解析返回 null
 *
 * @example
 * ```js
 * const bytes = await jsonToDocx({ blocks: [{ type:'image', src:{ assetId:'docx-0001' } }] }, { assets });
 * ```
 */
export async function resolveAssetBytes(
    value: unknown,
    table?: AssetTable | DocAsset[] | { assets?: DocAsset[] }
): Promise<{ bytes: Uint8Array; mime: string } | null> {
    if (value == null) return null;

    // 资源对象本身（带 Uint8Array 的 data 字段）
    if (isDocAsset(value)) return { bytes: value.data, mime: value.mime || sniffMime(value.data) };

    // 引用对象
    if (isAssetRef(value)) {
        const ref = value as { assetId?: string; asset?: DocAsset };
        const asset = ref.asset ?? getAsset(table, ref.assetId);
        if (!asset) return null;
        return { bytes: asset.data, mime: asset.mime };
    }

    // 资源对象
    if (value instanceof Uint8Array || value instanceof ArrayBuffer) {
        const bytes = await toUint8Array(value as DocInput);
        // 裸字节靠魔数识别类型，省去调用方手工标注
        return { bytes, mime: sniffMime(bytes) };
    }
    if (typeof value === 'string') {
        const bytes = await toUint8Array(value);
        // data URL 能带出 MIME
        const match = /^data:([^;,]+)/.exec(value.trim());
        return { bytes, mime: match ? match[1] : mimeOf(value, 'application/octet-stream') };
    }
    // Blob / File
    if (typeof (value as Blob)?.arrayBuffer === 'function') {
        const bytes = new Uint8Array(await (value as Blob).arrayBuffer());
        return { bytes, mime: (value as Blob).type || 'application/octet-stream' };
    }
    return null;
}

/** 判断字节是否以任一候选前缀开头 */
function startsWith(bytes: Uint8Array, ...prefixes: number[][]): boolean {
    return prefixes.some((prefix) => prefix.every((b, i) => bytes[i] === b));
}

/**
 * 猜测字节的 MIME（用于 generate 时用户直接传 Uint8Array 的场景）
 *
 * 依据 PNG / JPEG / GIF / WebP / BMP / PDF 的文件头魔数判断，避免要求用户手工标注。
 * @param bytes - 文件头字节
 */
export function sniffMime(bytes: Uint8Array): string {
    if (bytes.length >= 8 &&
        bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return 'image/png';
    if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg';
    if (bytes.length >= 6 && startsWith(bytes, [0x47, 0x49, 0x46, 0x38, 0x37, 0x61], [0x47, 0x49, 0x46, 0x38, 0x39, 0x61])) {
        return 'image/gif';
    }
    if (bytes.length >= 12 && startsWith(bytes, [0x52, 0x49, 0x46, 0x46]) &&
        bytes[8] === 0x57 && bytes[9] === 0x45 && bytes[10] === 0x42 && bytes[11] === 0x50) return 'image/webp';
    if (bytes.length >= 2 && bytes[0] === 0x42 && bytes[1] === 0x4d) return 'image/bmp';
    if (bytes.length >= 5 &&
        bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46 && bytes[4] === 0x2d) {
        return 'application/pdf';
    }
    return 'application/octet-stream';
}