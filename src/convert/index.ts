/**
 * PPTX 格式转换：插件化转换引擎。
 *
 * 每种目标格式是一个可注册的 Converter，统一通过 registerConverter / convert 使用。
 * 内置的 docx / pdf / md / html 四个转换器在模块加载时自动注册；消费者也可以
 * registerConverter 自己的实现来扩展目标格式（例如 docx -> html）。
 *
 *   pptx --(pptxToPdf)-------------> pdf   幻灯片版式直排，每页一张幻灯片
 *   pptx --(pptxToSlideHtml)------> html  同样按画布绝对定位还原
 *   pptx --(pptxToBlocks)--------> blocks --(generateDocx)---------> docx
 *                                        \--(blocksToMarkdown)------> md
 *
 * 本模块是主库 src/convert 的一部分，随主库一起构建发布；
 * 不依赖任何外部转换包，也不需要手工同步补丁。
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

/** 转换选项：所有转换器共用的上下文 */
export interface ConvertOptions {
    /** 输出标题（HTML 的 <title>、文档属性），缺省用文件名 */
    title?: string;
    /** 内嵌字体（生成中文 PDF 时可由 autoFont 自动探测系统字体） */
    fonts?: PdfFontSet;
    /** 自动探测系统字体（仅 Node 端有效），默认 true */
    autoFont?: boolean;
    /** 调用方额外提供的资源表 */
    assets?: DocAsset[];
}

/** 单个转换器：把一种源格式的输入转换成目标格式字节/文本 */
export interface Converter<Id extends string = string> {
    /** 目标格式标识，如 docx / pdf / md / html */
    readonly id: Id;
    /** 输出文件扩展名（不含点） */
    readonly ext: string;
    /** 可选的人类可读名称 */
    readonly label?: string;
    /** 执行转换；二进制目标返回 Uint8Array，文本目标返回 string */
    convert(input: DocInput, options?: ConvertOptions): Promise<Uint8Array | string>;
}

/** 转换器注册表：支持覆盖内置实现与新增自定义格式 */
class ConverterRegistry {
    private readonly map = new Map<string, Converter>();

    register(converter: Converter): this {
        this.map.set(converter.id, converter);
        return this;
    }

    get(id: string): Converter | undefined {
        return this.map.get(id);
    }

    has(id: string): boolean {
        return this.map.has(id);
    }

    list(): Converter[] {
        return Array.from(this.map.values());
    }

    remove(id: string): boolean {
        return this.map.delete(id);
    }
}

/** 全局转换注册表（内置转换器在下方自动注册） */
export const converters = new ConverterRegistry();

/** 可转换的目标格式（内置集合；自定义转换器可用任意 string） */
export type ConvertTarget = 'docx' | 'pdf' | 'md' | 'html';

/** 兼容旧调用名 */
export type ConvertPptxOptions = ConvertOptions;

/** 注册一个转换器（可覆盖同 id 的内置转换器） */
export function registerConverter(converter: Converter): void {
    converters.register(converter);
}

/** 统一转换入口：按 target 分派到已注册的转换器 */
export async function convert(
    input: DocInput,
    target: string,
    options: ConvertOptions = {}
): Promise<Uint8Array | string> {
    const converter = converters.get(target);
    if (!converter) {
        throw new TypeError(`未注册的转换目标：${target}。已注册：${converters.list().map((c) => c.id).join(', ')}`);
    }
    try {
        return await converter.convert(input, options);
    } catch (error) {
        throw wrapError('pptx', `转换(${target})`, error);
    }
}

// ---- 内置转换器 ----

/** docx：PPTX -> blocks -> DOCX */
const docxConverter: Converter<'docx'> = {
    id: 'docx',
    ext: 'docx',
    label: 'Word Document',
    async convert(input, options = {}) {
        const { blocks, metadata, assets } = await pptxToBlocks(input, { slideSeparator: false });
        const plain: DocxBlock[] = blocks.some((b) => b.type === 'chart')
            ? blocks.flatMap((b) => (b.type === 'chart' ? (b.fallback ? [b.fallback] : []) : [b]))
            : blocks;
        const pool = [...(options.assets || []), ...(assets || [])];
        const doc: DocxDocument = { metadata, blocks: plain, assets: pool.length ? pool : undefined };
        return generateDocx(doc);
    }
};

/** md：PPTX -> blocks -> Markdown */
const mdConverter: Converter<'md'> = {
    id: 'md',
    ext: 'md',
    label: 'Markdown',
    async convert(input, options = {}) {
        const { blocks, metadata, assets } = await pptxToBlocks(input, {
            // Markdown 没有「页」的概念，多页 PPT 之间插分隔线
            slideSeparator: true
        });
        const plain: DocxBlock[] = blocks.some((b) => b.type === 'chart')
            ? blocks.flatMap((b) => (b.type === 'chart' ? (b.fallback ? [b.fallback] : []) : []))
            : blocks;
        const pool = [...(options.assets || []), ...(assets || [])];
        const doc: DocxDocument = { metadata, blocks: plain, assets: pool.length ? pool : undefined };
        // PPTX 的图片宽高是画布坐标，写进 md 的 title 只是噪音
        return blocksToMarkdown(doc, { assets: pool, imageSizes: false });
    }
};

/** pdf：幻灯片版式直排（每页 = 一张幻灯片） */
const pdfConverter: Converter<'pdf'> = {
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

/** html：幻灯片版式还原（绝对定位画布） */
const htmlConverter: Converter<'html'> = {
    id: 'html',
    ext: 'html',
    label: 'HTML Web Page',
    async convert(input, options = {}) {
        return pptxToSlideHtml(input, { title: options.title });
    }
};

// 注册内置转换器（后注册覆盖同 id，便于消费者在导入后覆盖）
converters
    .register(docxConverter)
    .register(mdConverter)
    .register(pdfConverter)
    .register(htmlConverter);

/** 兼容旧门面：convertPptx(bytes, target, opts) */
export async function convertPptx(
    input: DocInput,
    target: ConvertTarget,
    options: ConvertOptions = {}
): Promise<Uint8Array | string> {
    return convert(input, target, options);
}

export type { DocInput } from './core/input';
export type { DocAsset } from './core/asset';
export type { PdfFontSet } from './pdf/types';
