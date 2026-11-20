/**
 * PPTX 资源抽取
 *
 * ppt-parser 把内嵌图片统一塞在元素上（`element.data` 为 data URL），
 * 因此这里做「后处理」：遍历幻灯片元素，把 data URL 换成资源引用，
 * 字节收集到资源表，并记录它属于第几张幻灯片。
 *
 * 这样 PPTX 与 DOCX/XLSX 共享同一套资源语义：
 *   前端用 `data` 直接显示，后台用 `assetId` 取字节或流。
 *
 * @module pptx/assets
 */

import {
    createAsset,
    createAssetId,
    parseDataUrl,
    type DocAsset,
    type DocAssetMode
} from '../core/asset';

/** 判断元素是否携带图片（内嵌 data URL） */
function isImageElement(element: unknown): element is { type?: string; data?: unknown } {
    if (!element || typeof element !== 'object') return false;
    const record = element as Record<string, unknown>;
    return typeof record.data === 'string' && record.data.startsWith('data:');
}

/**
 * 从标准文档结构中抽取图片资源
 *
 * @param document - ppt-parser 的标准文档结构（`{ slides: [...] }`）
 * @param mode - 资源处理方式，默认 `both`（元素上保留 data URL 供显示，同时抽出资源）
 * @returns 抽取出的资源数组；已就地改写文档元素
 */
export function extractPptxAssets(document: unknown, mode: DocAssetMode = 'both'): DocAsset[] {
    if (mode === 'ignore') return [];
    const assets: DocAsset[] = [];
    const slides = (document as { slides?: unknown[] })?.slides;
    if (!Array.isArray(slides)) return assets;

    const walk = (node: unknown, slideIndex: number): void => {
        if (!node || typeof node !== 'object') return;
        if (Array.isArray(node)) {
            for (const item of node) walk(item, slideIndex);
            return;
        }

        if (isImageElement(node)) {
            const parsed = parseDataUrl(node.data as string);
            if (parsed) {
                const id = createAssetId('pptx', assets.length + 1);
                const alt = (node as { alt?: unknown }).alt;
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
                const element = node as Record<string, unknown>;
                // both / extract 都写入 assetId；extract 模式下清掉 data 以免 JSON 膨胀
                element.assetId = id;
                if (mode === 'extract') delete element.data;
            }
            return;
        }

        // 继续下钻：形状填充图、子元素数组等
        for (const value of Object.values(node as Record<string, unknown>)) {
            walk(value, slideIndex);
        }
    };

    slides.forEach((slide, index) => walk(slide, index));
    return assets;
}

/** 由 MIME 取扩展名 */
function extOfMime(mime: string): string {
    const map: Record<string, string> = {
        'image/png': '.png',
        'image/jpeg': '.jpg',
        'image/gif': '.gif',
        'image/webp': '.webp',
        'image/svg+xml': '.svg'
    };
    return map[mime] || '.bin';
}