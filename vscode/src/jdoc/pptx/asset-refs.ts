/**
 * PPTX 生成侧的资源引用解析
 *
 * ppt-parser 只接受 data URL 或远程 URL，遇到 `{ assetId }` 这类引用会直接报错。
 * 因此在调用底层库之前，把引用统一还原成 data URL：
 * 字节只存在于资源表里，文档结构保持轻量，便于存储与传输。
 *
 * @module pptx/asset-refs
 */

import { assetToDataUrl, getAsset, isAssetRef, type AssetTable, type DocAsset } from '../core/asset';

/** 生成选项里可携带的资源表 */
export interface PptxAssetOptions {
    assets?: DocAsset[] | Record<string, DocAsset>;
}

/** 资源表归一化 */
function normalizeAssets(assets: DocAsset[] | Record<string, DocAsset> | undefined): AssetTable {
    if (!assets) return {};
    if (Array.isArray(assets)) {
        const table: AssetTable = {};
        for (const asset of assets) table[asset.id] = asset;
        return table;
    }
    return assets;
}

/**
 * 把文档里的资源引用还原成 data URL（就地改写）
 *
 * @param document - 标准文档结构
 * @param options - 携带资源表的选项
 * @returns 改写后的文档；无法解析的引用记入 `messages`
 *
 * @example
 * ```js
 * const deck = { slides: [{ elements: [{ type: 'image', data: { assetId: 'pptx-0001' } }] }] };
 * const bytes = await jsonToPptx(deck, { assets: parsed.assets });
 * ```
 */
export function resolvePptxAssetRefs(
    document: unknown,
    options: PptxAssetOptions | undefined
): { document: unknown; messages: string[] } {
    const messages: string[] = [];
    const table = normalizeAssets(
        options?.assets ?? (document as { assets?: DocAsset[] })?.assets
    );
    if (!Object.keys(table).length) return { document, messages };

    const walk = (node: unknown): void => {
        if (!node || typeof node !== 'object') return;
        if (Array.isArray(node)) {
            for (const item of node) walk(item);
            return;
        }

        const record = node as Record<string, unknown>;

        // extract 模式的元素只留顶层 assetId，data 已被移除——按资源表补回
        if (typeof record.assetId === 'string' && record.data === undefined) {
            const asset = getAsset(table, record.assetId);
            if (asset) {
                record.data = assetToDataUrl(asset);
            } else {
                messages.push(`资源引用无法解析：${record.assetId}`);
            }
        }

        // 图片字段两种约定：ppt-parser 用 data，这里兼容 assetId 引用
        for (const field of ['data', 'src']) {
            const value = record[field];
            if (isAssetRef(value)) {
                const ref = value as { assetId?: string; asset?: DocAsset };
                const asset = ref.asset ?? getAsset(table, ref.assetId);
                if (asset) {
                    record[field] = assetToDataUrl(asset);
                } else {
                    messages.push(`资源引用无法解析：${ref.assetId || '(缺少 assetId)'}`);
                    delete record[field];
                }
            }
        }
        // assetId 只是线索，data 已是 data URL 时可移除
        if (typeof record.assetId === 'string' && typeof record.data === 'string') {
            delete record.assetId;
        }

        for (const value of Object.values(record)) walk(value);
    };

    walk(document);
    return { document, messages };
}