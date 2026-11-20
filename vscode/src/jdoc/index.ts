/**
 * PPTX 格式转换门面。
 *
 * 这里只保留 jdoc `src/convert/index.ts` 里 `from === 'pptx'` 的那条分支：
 *
 *   pptx --(pptxToPdf)-------------> pdf   幻灯片版式直排，每页一张幻灯片
 *   pptx --(pptxToSlideHtml)------> html  同样按画布绝对定位还原
 *   pptx --(pptxToBlocks)--------> blocks --(generateDocx)---------> docx
 *                                        \--(blocksToMarkdown)------> md
 *
 * 其余源文件逐字拷贝自 jdoc（见同级目录），只裁剪了 PDF 的 barrel（pdf/index.ts），
 * 因为本插件不做「PDF 作为源」的转换，无需引入 pdfjs worker / cMap / wasm 资源。
 *
 * 与 jdoc 同步时：覆盖 src/jdoc/ 下除 pdf/index.ts 与本文件之外的所有文件，然后重新打上
 * 两处本地补丁（均已带注释，搜索「与 jdoc 上游的差异」「本仓库的 TS 7」即可定位）：
 *   1. convert/pptx-pdf.ts: gradientAt 的循环上界笔误（单停点渐变会让 HTML 转换崩溃）
 *   2. convert/markdown.ts: fencedCode 的 reduce 在 TypeScript 7 下推断成 never[]
 */
import { pptxToPdf } from './convert/pptx-pdf';
import { pptxToSlideHtml } from './convert/pptx-html';
import { pptxToBlocks } from './convert/pptx';
import { blocksToMarkdown } from './convert/markdown';
import { generateDocx } from './docx/generate';
import { wrapError } from './core/errors';
import type { DocAsset } from './core/asset';
import type { DocxBlock, DocxDocument } from './docx/types';
import type { PdfFontSet } from './pdf/types';
import type { DocInput } from './core/input';

/** 可转换的目标格式 */
export type ConvertTarget = 'docx' | 'pdf' | 'md' | 'html';

export interface ConvertPptxOptions {
    /** 输出标题（HTML 的 <title>、文档属性），缺省用文件名 */
    title?: string;
    /** 内嵌字体（生成中文 PDF 时可由 autoFont 自动探测系统字体） */
    fonts?: PdfFontSet;
    /** 自动探测系统字体（仅扩展宿主 Node 端有效），默认 true */
    autoFont?: boolean;
    /** 调用方额外提供的资源表 */
    assets?: DocAsset[];
}

/**
 * 把 PPTX 字节转换为 docx / pdf / md / html。
 *
 * @returns docx / pdf 返回 Uint8Array；md / html 返回字符串
 */
export async function convertPptx(
    input: DocInput,
    target: ConvertTarget,
    options: ConvertPptxOptions = {}
): Promise<Uint8Array | string> {
    try {
        // PDF / HTML 走幻灯片版式直排：经 blocks 中转的流式排版会把 16:9 画布
        // 重排成 A4 竖版，绝对定位元素互相叠压、大图溢出成多页。
        if (target === 'pdf') {
            return await pptxToPdf(input, {
                fonts: options.fonts,
                autoFont: options.autoFont ?? true
            });
        }
        if (target === 'html') {
            return await pptxToSlideHtml(input, { title: options.title });
        }

        const { blocks, metadata, assets } = await pptxToBlocks(input, {
            // Markdown 没有「页」的概念，多页 PPT 之间插分隔线
            slideSeparator: target === 'md'
        });

        // 图表块只在 HTML 目标能渲染成矢量 SVG；docx / md 退化为随块携带的数据表
        const plain: DocxBlock[] = blocks.some((b) => b.type === 'chart')
            ? blocks.flatMap((b) => (b.type === 'chart' ? (b.fallback ? [b.fallback] : []) : [b]))
            : blocks;
        const pool = [...(options.assets || []), ...(assets || [])];
        const doc: DocxDocument = {
            metadata,
            blocks: plain,
            assets: pool.length ? pool : undefined
        };

        if (target === 'docx') return await generateDocx(doc);
        if (target === 'md') {
            // PPTX 的图片宽高是画布坐标，写进 md 的 title 只是噪音
            return blocksToMarkdown(doc, { assets: pool, imageSizes: false });
        }
        throw new TypeError(`不支持的目标格式：${target}`);
    } catch (error) {
        throw wrapError('pptx', '转换', error);
    }
}
