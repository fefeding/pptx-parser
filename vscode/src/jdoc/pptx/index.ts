/**
 * PPTX 解析与生成公开入口
 *
 * PPTX 能力由 @fefeding/ppt-parser 提供，本模块负责：
 * 1) 把 jdoc 的宽松入参（Blob / base64 / TypedArray）转成底层库要求的 ArrayBuffer
 * 2) 把底层库的宽松出参（string | Buffer | Blob | ...）统一为 Uint8Array
 * 3) 导出标准文档契约 PptxDocument，供 jsonToPptx 往返
 *
 * @module pptx
 */

import {
    PPTXComposer as BaseComposer,
    editPptx as baseEditPptx,
    jsonToPptx as baseJsonToPptx,
    pptxToFiles as basePptxToFiles,
    pptxToHtml as basePptxToHtml,
    pptxToJson as basePptxToJson,
    pptxToStandard as basePptxToStandard
} from '@fefeding/ppt-parser';
import { wrapError } from '../core/errors';
import { toArrayBuffer, type DocInput } from '../core/input';
import type { DocAsset, DocAssetMode } from '../core/asset';
import { extractPptxAssets } from './assets';
import { resolvePptxAssetRefs } from './asset-refs';

export type { PptxDocument } from '@fefeding/ppt-parser';
export { PPTXShapeUtils, presetShapePath, createStore, createActions, renderChartSVG } from '@fefeding/ppt-parser';
export { normalizeDoc, normalizeElement } from '@fefeding/ppt-parser';
// 编辑核心与编辑器 UI 常量：供 VS Code 插件（及任何需要直接编辑 PPTX 的消费者）
// 复用 ppt-parser 的编辑器能力。docFromPptx/docToPptx 是编辑器文档模型的序列化入口，
// 其余为 ribbon/dialogs 所需的展示常量。
export { pptxChartToSpec, type PptxChartTheme } from './chart';
export {
  docFromPptx,
  docToPptx,
  SHAPES,
  FONT_LIST,
  THEMES,
  LAYOUTS,
  SLIDE_SIZES,
  TRANSITIONS,
  ANIM_CLASSES,
  ANIM_TYPES,
  ANIM_DIRECTIONS,
  CHART_TYPES
} from '@fefeding/ppt-parser';
// 渲染辅助（pptx-editor 的 render.ts/geometry 用到）
export {
  FONT_SIZES,
  getTheme,
  ptToPx,
  normalizeColor,
  withAlpha,
  hexToRgb,
  absoluteElementRect,
  effectMargin,
  elementRect,
  rotatedRect
} from '@fefeding/ppt-parser';

/** 底层库可能返回多种类型，这里统一收敛为 Uint8Array */
function toUint8(result: unknown): Uint8Array {
    if (result instanceof Uint8Array) return result;
    if (typeof result === 'string') {
        const binary = (globalThis as { atob?: (s: string) => string }).atob
            ? (globalThis as { atob: (s: string) => string }).atob(result)
            : Buffer.from(result, 'base64').toString('binary');
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i) & 0xff;
        return bytes;
    }
    if (result instanceof ArrayBuffer) return new Uint8Array(result);
    if (Array.isArray(result)) return new Uint8Array(result);
    throw new TypeError('PPTX 生成结果类型无法识别');
}

/** PPTX 生成选项 */
export interface PptxGenerateOptions {
    [key: string]: unknown;
    /**
     * 可用资源表：用于把文档里的 `{ assetId }` 图片引用还原成 data URL
     *
     * ppt-parser 只认 data URL / 远程 URL，这里在调用底层库前完成解析，
     * 因此「解析 → 改 → 生成」时可以直接复用资源表，字节不必再转成字符串搬运。
     *
     * @example
     * ```js
     * const deck = await pptxToStandard(file, { assets: 'extract' });
     * const bytes = await jsonToPptx(deck, { assets: deck.assets });
     * ```
     */
    assets?: DocAsset[] | Record<string, DocAsset>;
}

/**
 * PPTX 解析结果的宽松类型。
 * ppt-parser 的返回结构引用了其包内未导出的类型（HtmlSlideResult / XmlNode / PptxMetadata 等），
 * 直接透传会让声明文件无法生成，因此对外以宽松类型暴露，运行时结构保持不变。
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type PptxLooseResult = any;

/**
 * 解析 PPTX 为 HTML 幻灯片
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 透传 ppt-parser 的解析选项
 */
export async function pptxToHtml(
    input: DocInput,
    options: Record<string, unknown> = {}
): Promise<PptxLooseResult> {
    try {
        return await basePptxToHtml(await toArrayBuffer(input), options as never);
    } catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}

/**
 * 解析 PPTX 为结构化 JSON
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 透传 ppt-parser 的解析选项
 */
export async function pptxToJson(
    input: DocInput,
    options: Record<string, unknown> = {}
): Promise<PptxLooseResult> {
    try {
        return await basePptxToJson(await toArrayBuffer(input), options as never);
    } catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}

/** PPTX 解析选项（在 ppt-parser 选项基础上补充资源控制） */
export interface PptxParseOptions extends Record<string, unknown> {
    /**
     * 图片资源处理方式，默认 `both`
     *
     * ppt-parser 本身把内嵌图片以 data URL 放在元素上，这里做后处理：
     * - `both`（默认）：保留元素上的 data URL 供直接显示，同时抽出资源表
     * - `extract`：元素上只留 `assetId`，字节进资源表（JSON 体积小，适合后台）
     * - `ignore`：不抽取资源
     */
    assets?: DocAssetMode;
}

/**
 * 解析 PPTX 为标准文档结构（与 jsonToPptx 同源，可直接往返）
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 透传 ppt-parser 的解析选项，另支持 `assets` 控制图片资源
 * @returns 标准文档结构；当 `assets` 非 `ignore` 时附带 `assets` 资源表
 *
 * @example
 * ```js
 * const doc = await pptxToStandard(file, { assets: 'extract' });
 * doc.assets[0].anchor; // 第几张幻灯片
 * ```
 */
export async function pptxToStandard(
    input: DocInput,
    options: PptxParseOptions = {}
): Promise<PptxLooseResult> {
    try {
        const { assets: assetOption, ...parserOptions } = options;
        const document = await basePptxToStandard(await toArrayBuffer(input), parserOptions as never);
        if (assetOption === 'ignore') return document;

        const assets: DocAsset[] = extractPptxAssets(document, assetOption ?? 'both');
        if (assets.length) (document as { assets?: DocAsset[] }).assets = assets;
        return document;
    } catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}

/** 列出 PPTX 内部文件及内容 */
export async function pptxToFiles(input: DocInput): Promise<PptxLooseResult> {
    try {
        return await basePptxToFiles(await toArrayBuffer(input));
    } catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}

/**
 * 生成 PPTX 文件
 * @param presentation - 标准文档结构（或 ppt-parser 兼容的 JSON）
 * @returns PPTX 字节数据
 */
export async function jsonToPptx(
    presentation: unknown,
    options: PptxGenerateOptions = {}
): Promise<Uint8Array> {
    try {
        // assets 不是底层库的选项，取出后不往下传
        const { assets, ...rest } = options;
        const { document } = resolvePptxAssetRefs(presentation, {
            assets: assets ?? (presentation as { assets?: DocAsset[] })?.assets
        });
        const result = await baseJsonToPptx(document, { ...rest, outputType: 'uint8array' });
        return toUint8(result);
    } catch (error) {
        throw wrapError('pptx', '生成', error);
    }
}

/** 链式构建 PPTX（save() 统一返回 Uint8Array） */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export const PPTXComposer = BaseComposer as unknown as new (
    options?: Record<string, unknown>
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
) => any;

/**
 * 打开已有 PPTX 进行编辑（增删页、改元数据等）
 * @param input - 文件二进制
 */
export async function editPptx(input: DocInput): Promise<PptxLooseResult> {
    try {
        return await baseEditPptx(await toArrayBuffer(input));
    } catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}
