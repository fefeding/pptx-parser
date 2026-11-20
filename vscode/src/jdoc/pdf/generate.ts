/**
 * PDF 生成
 *
 * 基于 pdf-lib 输出 PDF。坐标系对调用方保持「顶部原点」（与 HTML 心智模型一致），
 * 内部换算为 pdf-lib 的左下原点。
 *
 * 中文字体：pdf-lib 的 14 种标准字体不含 CJK 字形，需要通过 fonts 传入 TTF/OTF 字节；
 * 未传入时使用标准字体，此时中文会抛出明确错误而不是静默输出空白。
 *
 * @module pdf/generate
 */

import {
    PDFDocument,
    StandardFonts,
    appendBezierCurve,
    clip,
    closePath,
    concatTransformationMatrix,
    endPath,
    fill,
    fillAndStroke,
    lineTo,
    moveTo,
    popGraphicsState,
    pushGraphicsState,
    rectangle,
    rgb,
    setDashPattern,
    setFillingColor,
    setGraphicsState,
    setLineWidth,
    setStrokingColor,
    stroke,
    type PDFFont,
    type PDFImage,
    type RGB
} from 'pdf-lib';
import { JdocError, wrapError } from '../core/errors';
import {
    getAsset,
    isAssetRef,
    toAssetTable,
    type AssetTable,
    type DocAsset
} from '../core/asset';
import { toUint8Array } from '../core/input';
import type {
    PdfDocumentSpec,
    PdfElement,
    PdfFontSet,
    PdfGenerateOptions,
    PdfImageElement,
    PdfPageSize,
    PdfPathSegment,
    PdfRun,
    PdfRunsElement,
    PdfTableCell,
    PdfTableCellObject,
    PdfTableElement,
    PdfTabStop,
    PdfTextElement
} from './types';
import { registerPdfFontkit } from './fonts';

/** 纸张尺寸（pt，1pt = 1/72 英寸） */
const PAGE_SIZES: Record<string, [number, number]> = {
    A3: [841.89, 1190.55],
    A4: [595.28, 841.89],
    A5: [419.53, 595.28],
    Letter: [612, 792],
    Legal: [612, 1008]
};

/** 命名颜色 */
const NAMED_COLORS: Record<string, string> = {
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

/** 颜色字符串 -> RGB；无法解析时返回黑色 */
export function parseColor(value: string | undefined, fallback: RGB = rgb(0, 0, 0)): RGB {
    if (!value) return fallback;
    let hex = value.trim();

    const named = NAMED_COLORS[hex.toLowerCase()];
    if (named) hex = named;
    if (hex.startsWith('#')) hex = hex.slice(1);
    if (hex.length === 3) hex = hex.split('').map((ch) => ch + ch).join('');
    if (!/^[0-9a-fA-F]{6}$/.test(hex)) return fallback;

    return rgb(
        parseInt(hex.slice(0, 2), 16) / 255,
        parseInt(hex.slice(2, 4), 16) / 255,
        parseInt(hex.slice(4, 6), 16) / 255
    );
}

/** 纸张尺寸 -> [宽, 高]（已考虑横竖屏） */
export function resolvePageSize(
    size: PdfPageSize | undefined,
    landscape: boolean
): [number, number] {
    let width: number;
    let height: number;

    if (Array.isArray(size)) {
        [width, height] = size;
    } else if (size && typeof size === 'object') {
        width = size.width;
        height = size.height;
    } else {
        const key = typeof size === 'string' ? size : 'A4';
        [width, height] = PAGE_SIZES[key] || PAGE_SIZES.A4;
    }

    return landscape ? [height, width] : [width, height];
}

/** 页边距归一化 */
function resolveMargin(margin: PdfPageSpecLike['margin']): { top: number; right: number; bottom: number; left: number } {
    const DEFAULT = 56; // 约 2cm
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

/** 页边距类型别名 */
type PdfPageSpecLike = { margin?: number | { top?: number; right?: number; bottom?: number; left?: number } };

/** 字体集合：按 normal/bold/italic/boldItalic 分类 */
export interface FontBundle {
    normal: PDFFont;
    bold: PDFFont;
    italic: PDFFont;
    boldItalic: PDFFont;
    /** 是否为自定义内嵌字体（决定中文是否可用） */
    embedded: boolean;
}

/** 标准字体族名 -> pdf-lib 常量 */
const STANDARD_FAMILIES = [
    { key: 'normal', constant: StandardFonts.Helvetica },
    { key: 'bold', constant: StandardFonts.HelveticaBold },
    { key: 'italic', constant: StandardFonts.HelveticaOblique },
    { key: 'boldItalic', constant: StandardFonts.HelveticaBoldOblique }
] as const;

/** 内嵌字体字节归一化 */
async function toFontBytes(source: string | ArrayBuffer | Uint8Array): Promise<Uint8Array> {
    if (typeof source !== 'string') return toUint8Array(source);
    // base64 或 data URL
    const match = /^data:[^;,]*(;base64)?,(.*)$/s.exec(source.trim());
    const payload = match ? match[2] : source;
    const clean = payload.replace(/\s+/g, '');
    const atobFn = (globalThis as { atob?: (s: string) => string }).atob;
    const binary = atobFn ? atobFn(clean) : Buffer.from(clean, 'base64').toString('binary');
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i) & 0xff;
    return bytes;
}

/** 构建字体集合 */
export async function buildFonts(
    doc: PDFDocument,
    fonts: PdfFontSet | undefined
): Promise<FontBundle> {
    if (fonts && (fonts.regular || fonts.bold)) {
        // 内嵌自定义字体必须注册 fontkit，否则 embedFont(TTF) 会抛 "Unknown font format"
        doc.registerFontkit((await registerPdfFontkit()) as never);

        const embed = async (source: string | ArrayBuffer | Uint8Array | undefined, fallback?: string) => {
            const bytes = source ? await toFontBytes(source) : fallback ? await toFontBytes(fallback) : undefined;
            // 必须显式 subset: true：pdf-lib 的 embedFont 对该选项的默认值是 **false**
            // （见 PDFDocument.embedFont 内 `subset = options.subset ?? false`），不传就整份
            // 嵌入字体。Arial Unicode.ttf 有 22MB，不子集化会让一行中文的 PDF 膨胀到 15MB，
            // 且体积过大的 PDF 会让 pdfjs + canvas 栅格化直接失败。
            return bytes ? doc.embedFont(bytes, { subset: true }) : undefined;
        };

        const regular = await embed(fonts.regular);
        const normal = regular || (await doc.embedFont(StandardFonts.Helvetica));
        const bold = (await embed(fonts.bold, typeof fonts.regular === 'string' ? fonts.regular : undefined)) || normal;
        const italic = (await embed(fonts.italic, typeof fonts.regular === 'string' ? fonts.regular : undefined)) || normal;
        const boldItalic =
            (await embed(fonts.boldItalic, typeof fonts.bold === 'string' ? fonts.bold : undefined)) || bold;
        return { normal, bold, italic, boldItalic, embedded: true };
    }

    // 标准 14 字体：Helvetica + Times + Courier 共 4 组，取 Helvetica 族做默认
    const normal = await doc.embedFont(StandardFonts.Helvetica);
    const bold = await doc.embedFont(StandardFonts.HelveticaBold);
    const italic = await doc.embedFont(StandardFonts.HelveticaOblique);
    const boldItalic = await doc.embedFont(StandardFonts.HelveticaBoldOblique);
    return { normal, bold, italic, boldItalic, embedded: false };
}

/** 检测文本中是否含有 CJK 字符（标准字体无法渲染） */
function hasCjk(text: string): boolean {
    return /[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF]/.test(text);
}

/** 文本按宽度换行（中文按字断行，西文优先按空格断行） */
export function wrapText(text: string, font: PDFFont, size: number, maxWidth: number): string[] {
    if (!text) return [];
    if (!maxWidth || maxWidth <= 0) return text.split('\n');

    const lines: string[] = [];
    for (const paragraph of String(text).split('\n')) {
        if (!paragraph) {
            lines.push('');
            continue;
        }

        let line = '';
        for (const char of Array.from(paragraph)) {
            const candidate = line + char;
            let width: number;
            try {
                width = font.widthOfTextAtSize(candidate, size);
            } catch {
                // 标准字体遇到缺字形会抛错，降级为「放得下就继续」
                width = candidate.length * size * 0.5;
            }

            if (width <= maxWidth || !line) {
                line = candidate;
                continue;
            }

            // 超宽：在最后一个空格处断行，避免西文单词被劈开
            const lastSpace = line.lastIndexOf(' ');
            if (lastSpace > 0 && /[\x20-\x7E]/.test(char)) {
                lines.push(line.slice(0, lastSpace));
                line = line.slice(lastSpace + 1) + char;
            } else {
                lines.push(line);
                line = char;
            }
        }
        if (line) lines.push(line);
    }
    return lines;
}

/** 图片字节类型识别 */
function detectImageType(bytes: Uint8Array): 'png' | 'jpg' | undefined {
    if (bytes.length > 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return 'png';
    if (bytes.length > 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'jpg';
    return undefined;
}

/**
 * 图片源归一化为字节
 *
 * 支持内联 data URL / base64 / 二进制，以及资源引用 `{ assetId }`
 * （引用需在 `jsonToPdf` 的 `assets` 里给出）。
 */
async function readImageBytes(
    src: PdfImageElement['src'],
    assets?: AssetTable
): Promise<Uint8Array> {
    if (isAssetRef(src)) {
        const ref = src as { assetId?: string; asset?: DocAsset };
        const asset = ref.asset ?? getAsset(assets, ref.assetId);
        if (!asset) throw new TypeError('图片引用无法解析：未提供对应的 assets');
        return asset.data;
    }
    return toUint8Array(src as string);
}

/**
 * 自动分页的流式状态：绘制函数逐行询问剩余空间，不足即换页。
 *
 * 由 generatePdf 的自动分页循环创建并注入 DrawContext；显式 pages 模式无 flow，
 * 绘制函数保持原有「一次画完、返回总高」的行为。
 */
interface FlowState {
    /** 当前页（换页时被循环替换） */
    page: ReturnType<PDFDocument['addPage']>;
    /** 页顶边距（顶部原点的基准） */
    marginTop: number;
    /** 当前页内容光标（自内容区顶部起算，pt） */
    cursorY: number;
    /** 单页可用高度（pageHeight - 上边距 - 下边距） */
    usableHeight: number;
    /** 确保剩余空间足够容纳 need；不足则换页。返回是否发生了换页 */
    ensure: (need: number) => boolean;
}

/** 绘制上下文 */
interface DrawContext {
    doc: PDFDocument;
    fonts: FontBundle;
    /** 可用资源表：用于解析图片元素的 { assetId } 引用 */
    assets: AssetTable;
    page: ReturnType<PDFDocument['addPage']>;
    pageHeight: number;
    contentWidth: number;
    originX: number;
    originY: number;
    defaultSize: number;
    defaultColor: RGB;
    /** 自动分页流状态：存在时绘制函数逐行换页而非整块挪页 */
    flow?: FlowState;
}

/** 当前绘制目标页（flow 模式下随换页更新） */
function currentPage(ctx: DrawContext): DrawContext['page'] {
    return ctx.flow ? ctx.flow.page : ctx.page;
}

/** 顶部原点光标（不含元素 y 偏移；flow 模式取流光标） */
function topCursor(ctx: DrawContext, elementY = 0): number {
    return ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : ctx.originY + elementY;
}

/**
 * 绘制一段文本（必要时合成粗体）
 *
 * 只内嵌了 regular 字形时没有真正的粗体（PDF 也没有描边加粗的能力），
 * 这里按 1/24 em 把同一字形偏移重画一次，等价于浏览器对无粗体字形字体
 * 的合成粗体（font-weight: 700），字重观感更接近原稿。
 */
function drawTextStyled(
    ctx: DrawContext,
    text: string,
    options: { x: number; y: number; size: number; font: PDFFont; color: RGB; bold?: boolean }
): void {
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

/** 在页面上绘制文本（含换行与对齐），返回占用高度 */
function drawTextElement(
    ctx: DrawContext,
    element: PdfTextElement | (Omit<PdfTextElement, 'type'> & { type: 'heading'; level: number })
): number {
    const isHeading = element.type === 'heading';
    const size = element.size ?? (isHeading ? Math.max(10, 24 - ((element.level - 1) as number) * 2.6) : ctx.defaultSize);

    // 标题与加粗走 bold 字体，斜体走 italic，二者同时命中用 boldItalic
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

    // 标题保持整块不拆页（Word 标题样式不跨页拆行）
    if (ctx.flow && isHeading) ctx.flow.ensure(lines.length * lineHeight);

    let localY = y;
    for (const line of lines) {
        let x = ctx.originX + (element.x ?? 0);
        if (element.align === 'center' || element.align === 'right') {
            let lineWidth: number;
            try {
                lineWidth = font.widthOfTextAtSize(line, size);
            } catch {
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
            // 顶部原点 -> pdf-lib 左下原点
            y: ctx.pageHeight - localY - size,
            size,
            font,
            bold,
            color: parseColor(element.color, ctx.defaultColor)
        });
        if (ctx.flow) ctx.flow.cursorY += lineHeight;
        localY += lineHeight;
    }

    return (ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : localY) - startTop;
}

/** 取某 run 对应的 pdf-lib 字体（bold / italic 组合） */
function fontForRun(ctx: DrawContext, run: PdfRun, size: number): PDFFont {
    const bold = run.bold ?? false;
    const italic = run.italic ?? false;
    void size;
    return bold && italic ? ctx.fonts.boldItalic : bold ? ctx.fonts.bold : italic ? ctx.fonts.italic : ctx.fonts.normal;
}

/** 测一段 run 文本在给定字号下的像素宽度（含字体回退容错） */
function measureRun(ctx: DrawContext, run: PdfRun, text: string, size: number): number {
    try {
        return fontForRun(ctx, run, size).widthOfTextAtSize(text, size);
    } catch {
        return text.length * size * 0.5;
    }
}

/**
 * 绘制富文本段落（内联样式的多 run 混排）
 *
 * 先把 runs 切成语义「片段」：CJK 字符逐字成段（按字断行），西文按词成段
 * （词内不拆），样式切换/空白也作为片段边界。随后贪心地把片段塞进当前行，
 * 放不下就在该片段前换行，因此加粗/彩色/行内代码可落在不连续的 x 偏移上。
 */
function drawRunsElement(ctx: DrawContext, element: PdfRunsElement): number {
    const size = element.size ?? ctx.defaultSize;
    const indentLeft = element.indentLeft ?? 0;
    const indentRight = element.indentRight ?? 0;
    const firstIndent = element.indentFirstLine ?? 0;
    const width = element.width ?? ctx.contentWidth - indentLeft - indentRight;
    const baseX = ctx.originX + (element.x ?? 0) + indentLeft;

    /** 渲染片段：一段同一样式的文本（可能是单字、一个西文词、一段空白） */
    interface Piece {
        text: string;
        run: PdfRun;
    }
    const pieces: Piece[] = [];
    const isCjk = (ch: string) => /[　-〿㐀-鿿豈-﫿＀-￯]/.test(ch);
    const isLatin = (ch: string) => /[A-Za-z0-9]/.test(ch);

    for (const run of element.runs) {
        if (!run.text) continue;
        let word = '';
        let wordRun: PdfRun | undefined;
        const flushWord = () => {
            if (word) {
                pieces.push({ text: word, run: wordRun! });
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
                // CJK 与任何非 ASCII：逐字成段
                flushWord();
                pieces.push({ text: ch, run });
            } else if (ch === '\t') {
                // 制表符单独成段：绘制时按 tabStops 跳位（无 tabStops 时退化为空格）
                flushWord();
                pieces.push({ text: '\t', run });
            } else if (ch === ' ') {
                flushWord();
                pieces.push({ text: ' ', run });
            } else if (isLatin(ch)) {
                // 西文词：连同标点尽量聚合，但样式切换时先收尾
                if (wordRun && wordRun !== run) flushWord();
                word += ch;
                wordRun = run;
            } else {
                // 其它 ASCII 标点：并入西文词缓冲（避免 "hello," 被拆开）
                if (wordRun && wordRun !== run) flushWord();
                word += ch;
                wordRun = run;
            }
        }
        flushWord();
    }

    /** 一行：片段序列 + 是否首行（首行缩进收窄可用宽）+ 该行字号 */
    interface Line {
        pieces: Piece[];
        first: boolean;
        /** 该行的最大字号（run 自带字号优先），行高按它计算 */
        size: number;
    }
    const lines: Line[] = [];
    let line: Piece[] = [];
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
        // tab 的落点在绘制时按制表位确定，宽度不预占（有 tabStops 时按 0 计）
        const w =
            piece.text === '\t' ? (tabStops ? 0 : size * 0.5) : measureRun(ctx, piece.run, piece.text, piece.run.size ?? size);
        if (lineWidth + w > lineAvail && line.length && piece.text !== '\t') {
            pushLine();
        }
        line.push(piece);
        lineWidth += w;
        lineSize = Math.max(lineSize, piece.run.size ?? size);
    }
    pushLine();

    const drawLine = (l: Line, localY: number) => {
        const lineTextWidth = l.pieces.reduce(
            (sum, p) => (p.text === '\t' ? sum : sum + measureRun(ctx, p.run, p.text, p.run.size ?? size)),
            0
        );
        let x = baseX + (l.first ? firstIndent : 0);
        if (element.align === 'center' || element.align === 'right') {
            const offset = width - lineTextWidth;
            x += element.align === 'center' ? offset / 2 : Math.max(0, offset);
        }

        // 制表位：预扫每个 tab 到「下一个 tab / 行末」之间的片段总宽（右对齐需要）
        const segWidths = new Map<number, number>();
        if (tabStops) {
            for (let i = 0; i < l.pieces.length; i++) {
                if (l.pieces[i].text !== '\t') continue;
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
                    x += size * 0.5; // 无制表位定义：退化为半角空格
                    continue;
                }
                // 跳到第一个位于当前位置右侧的制表位（pos 缺省 = 段落右端）
                const stop = tabStops.find((s) => (s.pos ?? width) > x - baseX + 0.5);
                if (!stop) continue;
                const stopPos = stop.pos ?? width;
                const segW = segWidths.get(i) ?? 0;
                const segSize = p.run.size ?? size;
                const segColor = p.run.color ? parseColor(p.run.color, ctx.defaultColor) : ctx.defaultColor;
                let contentX = baseX + stopPos;
                if (stop.align === 'right') contentX -= segW;
                else if (stop.align === 'center') contentX -= segW / 2;
                if (stop.leader) {
                    // 右/居中对齐时前导符只填「内容与制表位之间」的空白
                    drawLeader(ctx, x, contentX, localY, segSize, stop.leader, segColor);
                }
                x = Math.max(x, contentX); // 内容超宽时不倒退
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
            } else {
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

    // 文本框整体旋转（PPTX 的文本框 rot）：绕框中心旋转，用变换矩阵实现。
    // 框高未知时按首行位置退化处理（不旋转）。
    const rotation = element.rotation ?? 0;
    if (rotation && element.height) {
        const rad = (rotation * Math.PI) / 180;
        const cos = Math.cos(rad);
        const sin = Math.sin(rad);
        const cx = ctx.originX + (element.x ?? 0) + indentLeft + width / 2;
        const cy = ctx.pageHeight - startTop - element.height / 2;
        ctx.page.pushOperators(
            pushGraphicsState(),
            concatTransformationMatrix(
                cos,
                sin,
                -sin,
                cos,
                cx - cx * cos + cy * sin,
                cy - cx * sin - cy * cos
            )
        );
    }

    for (let i = 0; i < lines.length; i++) {
        const l = lines[i];
        // 行高按该行实际字号计算（run 自带字号时不能用元素默认字号，
        // 否则大字号折行后的行距会过小、行与行叠在一起）
        const lineHeight = (element.lineHeight ?? 1.35) * l.size;
        if (ctx.flow) {
            // widow/orphan：首行不独留页尾（orphan）、末行不独留页首（widow）
            if (i === 0 && lines.length > 1) ctx.flow.ensure(Math.min(2, lines.length) * lineHeight);
            if (i === lines.length - 2 && lines.length > 2) ctx.flow.ensure(2 * lineHeight);
            ctx.flow.ensure(lineHeight);
            localY = ctx.flow.marginTop + ctx.flow.cursorY;
        }
        drawLine(l, localY);
        if (ctx.flow) ctx.flow.cursorY += lineHeight;
        localY += lineHeight;
    }
    if (rotation && element.height) ctx.page.pushOperators(popGraphicsState());

    const used = (ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : localY) - startTop;
    // 段后间距在 flow 模式下由自动分页循环统一进给，不占元素内容高
    return used + (ctx.flow ? 0 : (element.spacingAfter ?? 4));
}

/** 绘制制表位前导符（目录的点线导引等） */
function drawLeader(
    ctx: DrawContext,
    from: number,
    to: number,
    localY: number,
    size: number,
    leader: NonNullable<PdfTabStop['leader']>,
    color: RGB
): void {
    if (to - from < 1) return;
    const page = currentPage(ctx);
    // 前导符位于基线略下方
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

/**
 * 绘制表格，返回占用高度
 *
 * 按 Word 的 tblGrid 网格布局：columnWidths 是「网格列」宽，而 rows 里每个 cell
 * 只占其中连续若干列（gridSpan），纵向合并的 cell 会占住后续行的同一列。因此
 * 先把 rows 展开成网格占位（placement），每个 cell 拿到自己的起始网格列，再按
 * 该列累加宽度——不能按「第 N 个 cell」取宽，否则合并表头会拿到某个数字列的
 * 窄宽度，文字被逐字折行。
 *
 * 单元格文本按实际宽度换行、行高按内容自适应；flow 模式下按行跨页，跨页后重复
 * 表头行（与 Word「标题行重复」行为一致）。
 */
function drawTableElement(ctx: DrawContext, element: PdfTableElement): number {
    const fontSize = element.fontSize ?? 10;
    const padding = element.padding ?? 4;
    const lineH = fontSize * 1.35;
    const hasHeader = element.header !== false && element.rows.length > 0;
    if (!element.rows.length) return 0;

    /** 归一化后的单元格：内容一律转成纯文本，另带合并/底纹元数据 */
    interface NCell {
        text: string;
        colspan: number;
        rowspan: number;
        width?: number;
        fill?: string;
        verticalAlign?: 'top' | 'middle' | 'bottom';
    }
    const isCellObject = (c: PdfTableCell): c is PdfTableCellObject =>
        !!c && typeof c === 'object' && !Array.isArray(c) && 'content' in c;
    const normCell = (cell: PdfTableCell): NCell => {
        if (isCellObject(cell)) {
            const content = cell.content;
            return {
                text: typeof content === 'string' ? content : cellToPlainText(content),
                colspan: Math.max(1, Math.round(cell.colspan ?? 1)),
                rowspan: Math.max(1, Math.round(cell.rowspan ?? 1)),
                width: cell.width,
                fill: cell.fill,
                verticalAlign: cell.verticalAlign
            };
        }
        const text = tableCellText(cell);
        return { text, colspan: 1, rowspan: 1 };
    };
    const rows = element.rows.map((row) => row.map(normCell));

    // ---- 展开网格：确定每个 cell 的起始网格列，并标记 rowspan/colspan 占位 ----
    interface Placement {
        row: number;
        col: number;
        cell: NCell;
    }
    const placements: Placement[] = [];
    const occupied: boolean[][] = [];
    let gridCols = 0;
    rows.forEach((row, r) => {
        if (!occupied[r]) occupied[r] = [];
        let col = 0;
        for (const cell of row) {
            // 跳过被上方 rowspan 占住的网格列
            while (occupied[r][col]) col++;
            for (let rr = r; rr < r + cell.rowspan && rr < rows.length; rr++) {
                if (!occupied[rr]) occupied[rr] = [];
                for (let cc = col; cc < col + cell.colspan; cc++) occupied[rr][cc] = true;
            }
            placements.push({ row: r, col, cell });
            col += cell.colspan;
        }
        gridCols = Math.max(gridCols, col);
    });
    if (!gridCols) return 0;

    const widths = element.columnWidths?.length
        ? element.columnWidths
        : new Array(gridCols).fill(ctx.contentWidth / gridCols);

    /** 网格列 col 之前的累计宽度 */
    const colOffset = (col: number): number => {
        let x = 0;
        for (let i = 0; i < col && i < widths.length; i++) x += widths[i];
        return x;
    };
    /** cell 实际宽度 = 它跨越的那些网格列之和（tcW 显式宽度优先） */
    const cellWidth = (cell: NCell, col: number): number => {
        if (cell.width !== undefined && cell.width > 0) return cell.width;
        let w = 0;
        for (let i = col; i < col + cell.colspan; i++) w += widths[i] ?? widths[widths.length - 1] ?? 0;
        if (w <= 0) w = widths[widths.length - 1] ?? ctx.contentWidth / gridCols;
        return w;
    };

    const borderColor = element.borderColor === 'none' ? undefined : parseColor(element.borderColor || '#CCCCCC');
    const headerFill = element.headerFill ? parseColor(element.headerFill) : undefined;

    const startTop = topCursor(ctx, element.y ?? 0);
    const startX = ctx.originX + (element.x ?? 0);

    // ---- 预排版：按各自实际宽度换行，并求出每行行高（含 rowspan 分摊） ----
    const cellLines = new Map<Placement, string[]>();
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
    // 纵向合并的 cell 若比它跨越的那些行加起来还高，把差额均摊到这些行
    for (const p of placements) {
        if (p.cell.rowspan <= 1) continue;
        let sum = 0;
        for (let r = p.row; r < p.row + p.cell.rowspan; r++) sum += rowHeights[r];
        const need = (cellLines.get(p)?.length ?? 0) * lineH + padding * 2;
        if (sum < need) {
            const per = (need - sum) / p.cell.rowspan;
            for (let r = p.row; r < p.row + p.cell.rowspan; r++) rowHeights[r] += per;
        }
    }
    const heightOf = (r: number): number => element.rowHeight ?? rowHeights[r] ?? 0;

    /** 画第 r 行的所有 cell（rowspan>1 的 cell 只在起始行画一次，高度含其跨越的行） */
    const drawRowCells = (topY: number, r: number): void => {
        const page = currentPage(ctx);
        const isHeader = hasHeader && r === 0;
        for (const p of placements) {
            if (p.row !== r) continue;
            const c = p.cell;
            const w = cellWidth(c, p.col);
            let h = 0;
            for (let rr = r; rr < r + c.rowspan; rr++) h += heightOf(rr);
            const x = startX + colOffset(p.col);
            const fill = c.fill ? parseColor(c.fill) : isHeader ? headerFill : undefined;

            if (borderColor || fill) {
                page.drawRectangle({
                    x,
                    y: ctx.pageHeight - topY - h,
                    width: w,
                    height: h,
                    borderColor,
                    borderWidth: 0.5,
                    color: fill
                });
            }

            const font = isHeader ? ctx.fonts.bold : ctx.fonts.normal;
            const lines = cellLines.get(p) ?? [];
            const textBlock = lines.length * lineH;
            let offsetY = topY + (h - textBlock) / 2;
            if (c.verticalAlign === 'top') offsetY = topY + padding / 2;
            else if (c.verticalAlign === 'bottom') offsetY = topY + h - textBlock - padding / 2;
            lines.forEach((text, li) => {
                let textWidth: number;
                try {
                    textWidth = font.widthOfTextAtSize(text, fontSize);
                } catch {
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
            // 表头与首条数据行保持同页（避免表头孤悬页尾）
            const bind = isHeader && rows.length > 1 ? rowHeight + heightOf(1) : rowHeight;
            const moved = ctx.flow.ensure(bind);
            if (moved && hasHeader && rowIndex > 0) {
                // 跨页：重复表头行
                drawRowCells(ctx.flow.marginTop + ctx.flow.cursorY, 0);
                ctx.flow.cursorY += heightOf(0);
            }
            localY = ctx.flow.marginTop + ctx.flow.cursorY;
        }

        drawRowCells(localY, rowIndex);

        if (ctx.flow) ctx.flow.cursorY += rowHeight;
        localY += rowHeight;
    });

    return (ctx.flow ? ctx.flow.marginTop + ctx.flow.cursorY : localY) - startTop;
}

/** 表格单元格 -> 纯文本（嵌套元素取其文本拼接） */
function cellToPlainText(cell: PdfElement[]): string {
    return cell
        .map((element) => {
            if (element.type === 'text' || element.type === 'heading') return element.text ?? '';
            return '';
        })
        .join(' ')
        .trim();
}

/** 表格单元格 -> 纯文本（兼容 纯文本 / 元素序列 / 带合并信息对象 三种形态） */
function tableCellText(cell: PdfTableCell): string {
    if (typeof cell === 'string') return cell;
    if (Array.isArray(cell)) return cellToPlainText(cell);
    return typeof cell.content === 'string' ? cell.content : cellToPlainText(cell.content);
}

/** 元素段前间距（pt，自动分页时由循环层进给） */
function spacingBeforeOf(element: PdfElement): number {
    const el = element as { spacingBefore?: number };
    return el.spacingBefore ?? 0;
}

/** 元素段后间距（pt；runs 元素默认 4 与原行为一致） */
function spacingAfterOf(element: PdfElement): number {
    if (element.type === 'runs') return element.spacingAfter ?? 4;
    const el = element as { spacingAfter?: number };
    return el.spacingAfter ?? 0;
}

/**
 * 圆角矩形 / 椭圆的路径（PDF 坐标系，y 向上）
 *
 * pdf-lib 没有圆角矩形与椭圆 API（drawSvgPath 在 1.17 里也不可用），
 * 这里用三次贝塞尔曲线逼近圆弧：控制点系数 k = 4/3·(√2−1) ≈ 0.5523。
 */
const BEZIER_K = 0.5523;

/** 圆角矩形路径段 */
function roundedRectPath(x: number, y: number, w: number, h: number, r: number): PdfPathSegment[] {
    const k = BEZIER_K * r;
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

/** 椭圆路径段（顺时针四段四分之一圆弧） */
function ellipsePath(x: number, y: number, w: number, h: number): PdfPathSegment[] {
    const rx = w / 2;
    const ry = h / 2;
    const cx = x + rx;
    const cy = y + ry;
    const kx = BEZIER_K * rx;
    const ky = BEZIER_K * ry;
    return [
        { op: 'm', args: [cx + rx, cy] },
        { op: 'c', args: [cx + rx, cy + ky, cx + kx, cy + ry, cx, cy + ry] },
        { op: 'c', args: [cx - kx, cy + ry, cx - rx, cy + ky, cx - rx, cy] },
        { op: 'c', args: [cx - rx, cy - ky, cx - kx, cy - ry, cx, cy - ry] },
        { op: 'c', args: [cx + kx, cy - ry, cx + rx, cy - ky, cx + rx, cy] },
        { op: 'h', args: [] }
    ];
}

/**
 * 裁剪路径 -> PDF 操作符
 *
 * 路径坐标为元素坐标系（原点在内容区左上、y 向下），这里翻转到 PDF 的 y 向上坐标系。
 */
function clipPathOperators(ctx: DrawContext, segments: PdfPathSegment[]) {
    const ops = [];
    const px = (v: number) => ctx.originX + v;
    const py = (v: number) => ctx.pageHeight - ctx.originY - v;
    for (const seg of segments) {
        if (seg.op === 'm') ops.push(moveTo(px(seg.args[0]), py(seg.args[1])));
        else if (seg.op === 'l') ops.push(lineTo(px(seg.args[0]), py(seg.args[1])));
        else if (seg.op === 'c')
            ops.push(
                appendBezierCurve(
                    px(seg.args[0]),
                    py(seg.args[1]),
                    px(seg.args[2]),
                    py(seg.args[3]),
                    px(seg.args[4]),
                    py(seg.args[5])
                )
            );
        else ops.push(closePath());
    }
    return ops;
}

/**
 * path 元素：内容区左上原点、y 向下、pt -> PDF 操作符
 *
 * 坐标约定与 group 的裁剪路径一致（都来自同一处路径构造），因此同样要
 * 叠加 originX / originY。
 */
function pathSegmentOperators(ctx: DrawContext, segments: PdfPathSegment[]) {
    const ops = [];
    const px = (v: number) => ctx.originX + v;
    const py = (v: number) => ctx.pageHeight - ctx.originY - v;
    for (const seg of segments) {
        if (seg.op === 'm') ops.push(moveTo(px(seg.args[0]), py(seg.args[1])));
        else if (seg.op === 'l') ops.push(lineTo(px(seg.args[0]), py(seg.args[1])));
        else if (seg.op === 'c')
            ops.push(
                appendBezierCurve(
                    px(seg.args[0]),
                    py(seg.args[1]),
                    px(seg.args[2]),
                    py(seg.args[3]),
                    px(seg.args[4]),
                    py(seg.args[5])
                )
            );
        else ops.push(closePath());
    }
    return ops;
}

/** 用路径段绘制形状（填充 / 描边） */
function drawBezierShape(
    ctx: DrawContext,
    options: { color?: RGB; borderColor?: RGB; borderWidth?: number; opacity?: number },
    segments: PdfPathSegment[]
): void {
    const ops = [];
    for (const seg of segments) {
        if (seg.op === 'm') ops.push(moveTo(seg.args[0], seg.args[1]));
        else if (seg.op === 'l') ops.push(lineTo(seg.args[0], seg.args[1]));
        else if (seg.op === 'c') ops.push(appendBezierCurve(seg.args[0], seg.args[1], seg.args[2], seg.args[3], seg.args[4], seg.args[5]));
        else ops.push(closePath());
    }
    const alpha = options.opacity !== undefined ? Math.max(0, Math.min(1, options.opacity)) : 1;
    const hasFill = !!options.color;
    const hasStroke = !!options.borderColor;
    // 半透明填充：PDF 图形状态没有直接的透明度参数，用 ExtGState（/ca）
    if (alpha < 1) {
        const gs = ctx.doc.context.obj({ Type: 'ExtGState', ca: alpha, CA: alpha });
        const key = ctx.page.node.newExtGState('GS', gs);
        ops.push(pushGraphicsState(), setGraphicsState(key));
    }
    if (hasFill) ops.push(setFillingColor(options.color));
    if (hasStroke) {
        ops.push(setStrokingColor(options.borderColor));
        ops.push(setLineWidth(options.borderWidth ?? 1));
    }
    // 填充（f）会终结当前路径，之后再描边（S）将无路径可描 —— 边框会静默丢失，
    // 所以「填充 + 描边」必须用 B 操作符一次完成。
    if (hasFill && hasStroke) ops.push(fillAndStroke());
    else if (hasFill) ops.push(fill());
    else if (hasStroke) ops.push(stroke());
    if (alpha < 1) ops.push(popGraphicsState());
    ctx.page.pushOperators(...ops);
}

/**
 * 绘制箭头三角（PPT 连接线的 startArrow / endArrow）
 *
 * 尺寸与编辑器渲染端的 SVG marker 对齐：markerWidth=5 + markerUnits=strokeWidth，
 * viewBox 为 0 0 10 10 的三角形 -> 实际长度 5×线宽、半高 2.5×线宽。
 *
 * @param tipX 箭头尖端 x（PDF 坐标）
 * @param tipY 箭头尖端 y（PDF 坐标）
 * @param refX 参考点 x（箭头方向的来源）
 * @param refY 参考点 y
 */
function drawArrowHead(
    ctx: DrawContext,
    tipX: number,
    tipY: number,
    refX: number,
    refY: number,
    thickness: number,
    color: RGB
): void {
    const dx = tipX - refX;
    const dy = tipY - refY;
    const len = Math.hypot(dx, dy);
    if (len < 0.01) return;
    const ux = dx / len;
    const uy = dy / len;
    const px = -uy;
    const py = ux;
    const l = thickness * 5;
    const halfH = thickness * 2.5;
    const baseX = tipX - ux * l;
    const baseY = tipY - uy * l;
    ctx.page.pushOperators(
        pushGraphicsState(),
        setFillingColor(color),
        moveTo(tipX, tipY),
        lineTo(baseX + px * halfH, baseY + py * halfH),
        lineTo(baseX - px * halfH, baseY - py * halfH),
        closePath(),
        fill(),
        popGraphicsState()
    );
}

/** 绘制单个元素，返回占用高度 */
async function drawElement(ctx: DrawContext, element: PdfElement): Promise<number> {
    switch (element.type) {
        case 'text':
        case 'heading':
            return drawTextElement(ctx, element as never);

        case 'line': {
            const thickness = element.lineWidth ?? 1;
            const sx = ctx.originX + element.x1;
            const sy = ctx.pageHeight - element.y1;
            const ex = ctx.originX + element.x2;
            const ey = ctx.pageHeight - element.y2;
            const color = parseColor(element.color, rgb(0.6, 0.6, 0.6));
            ctx.page.drawLine({
                start: { x: sx, y: sy },
                end: { x: ex, y: ey },
                thickness,
                color,
                dashArray: element.dash
            });
            // 箭头（PPT 连接线的 endArrow / startArrow）
            if (element.endArrow) drawArrowHead(ctx, ex, ey, sx, sy, thickness, color);
            if (element.startArrow) drawArrowHead(ctx, sx, sy, ex, ey, thickness, color);
            return thickness;
        }

        case 'path': {
            // 任意路径（PPTX 的块箭头 / 圆柱等预设形状）
            const ops = pathSegmentOperators(ctx, element.segments);
            const alpha =
                element.opacity !== undefined ? Math.max(0, Math.min(1, element.opacity)) : 1;
            const hasFill = !!element.fill;
            const hasStroke = !!element.stroke;
            if (alpha < 1) {
                const gs = ctx.doc.context.obj({ Type: 'ExtGState', ca: alpha, CA: alpha });
                const gsRef = ctx.page.node.newExtGState('GSa', gs);
                ctx.page.pushOperators(pushGraphicsState(), setGraphicsState(gsRef), ...ops);
            } else {
                ctx.page.pushOperators(pushGraphicsState(), ...ops);
            }
            if (hasFill) ctx.page.pushOperators(setFillingColor(parseColor(element.fill)));
            if (hasStroke) {
                ctx.page.pushOperators(setStrokingColor(parseColor(element.stroke)));
                ctx.page.pushOperators(setLineWidth(element.lineWidth ?? 1));
                if (element.dash?.length) ctx.page.pushOperators(setDashPattern(element.dash, 0));
            }
            // 同 drawBezierShape：f 会终结路径，填充+描边需用 B 一次完成
            if (hasFill && hasStroke) ctx.page.pushOperators(fillAndStroke());
            else if (hasFill) ctx.page.pushOperators(fill());
            else if (hasStroke) ctx.page.pushOperators(stroke());
            ctx.page.pushOperators(popGraphicsState());

            // 箭头：起点取路径第一段的方向，终点取最后一段的方向
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
            // 流式分隔线：在当前内容位置画一条横贯内容区的细线
            const thickness = element.lineWidth ?? 0.5;
            const before = element.spacingBefore ?? 4;
            const after = element.spacingAfter ?? 4;
            const total = before + thickness + after;
            if (ctx.flow) ctx.flow.ensure(total);
            const y = ctx.pageHeight - topCursor(ctx) - before;
            currentPage(ctx).drawLine({
                start: { x: ctx.originX, y },
                end: { x: ctx.originX + ctx.contentWidth, y },
                thickness,
                color: parseColor(element.color, rgb(0.6, 0.6, 0.6))
            });
            if (ctx.flow) ctx.flow.cursorY += total;
            return total;
        }

        case 'group': {
            // 裁剪路径 + 子元素：渐变条带限制在形状轮廓内（PDF 无原生渐变）
            const clipPath = element.clip;
            if (clipPath?.length) {
                ctx.page.pushOperators(
                    pushGraphicsState(),
                    ...clipPathOperators(ctx, clipPath),
                    clip(),
                    endPath()
                );
            }
            for (const child of element.children) {
                await drawElement(ctx, child);
            }
            if (clipPath?.length) ctx.page.pushOperators(popGraphicsState());
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
            // 绕自身中心旋转：PDF 图形状态没有直接的旋转参数，用变换矩阵实现
            const rotation = element.rotation ?? 0;
            if (rotation) {
                const rad = (rotation * Math.PI) / 180;
                const cos = Math.cos(rad);
                const sin = Math.sin(rad);
                const cx = options.x + element.width / 2;
                const cy = options.y + element.height / 2;
                ctx.page.pushOperators(
                    pushGraphicsState(),
                    concatTransformationMatrix(
                        cos,
                        sin,
                        -sin,
                        cos,
                        cx - cx * cos + cy * sin,
                        cy - cx * sin - cy * cos
                    )
                );
            }
            if (element.ellipse) {
                // 椭圆：pdf-lib 没有椭圆 API，用贝塞尔曲线拼四段圆弧
                const { pageHeight, originX, originY } = ctx;
                const px = originX + element.x;
                const py = pageHeight - originY - element.y - element.height;
                drawBezierShape(ctx, options, ellipsePath(px, py, element.width, element.height));
            } else if (element.radius && element.radius > 0) {
                // 圆角矩形：pdf-lib 没有圆角 API，同样用贝塞尔曲线拼四段圆角
                const { pageHeight, originX, originY } = ctx;
                const px = originX + element.x;
                const py = pageHeight - originY - element.y - element.height;
                const r = Math.min(element.radius, element.width / 2, element.height / 2);
                drawBezierShape(ctx, options, roundedRectPath(px, py, element.width, element.height, r));
            } else {
                ctx.page.drawRectangle(options);
            }
            if (rotation) ctx.page.pushOperators(popGraphicsState());
            return element.height;
        }

        case 'image': {
            try {
                const bytes = await readImageBytes(element.src, ctx.assets);
                const kind = detectImageType(bytes);
                if (!kind) throw new TypeError('仅支持 PNG / JPEG 图片');
                const image: PDFImage = kind === 'png' ? await ctx.doc.embedPng(bytes) : await ctx.doc.embedJpg(bytes);
                const width = element.width ?? image.width;
                const height = element.height ?? image.height;
                if (ctx.flow) ctx.flow.ensure(height);
                // 裁剪（PPTX a:srcRect）：把源图按裁剪窗口放大绘制，再用裁剪路径
                // 限制在原框内 —— 不用裁剪就只能看到源图全貌，边缘内容会多出来
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
                    // 裁剪窗口 = 原框
                    const clipX = ctx.originX + (element.x ?? 0);
                    const clipY = ctx.pageHeight - topCursor(ctx, element.y ?? 0) - height;
                    ctx.page.pushOperators(
                        pushGraphicsState(),
                        rectangle(clipX, clipY, width, height),
                        clip(),
                        endPath()
                    );
                }
                currentPage(ctx).drawImage(image, {
                    x: ctx.originX + drawX,
                    y: ctx.pageHeight - drawTop - drawH,
                    width: drawW,
                    height: drawH
                });
                if (cropped) ctx.page.pushOperators(popGraphicsState());
                if (ctx.flow) ctx.flow.cursorY += height;
                return height;
            } catch {
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

/** 生成 PDF
 * @param spec - 文档规格（页面或元素序列）
 * @param options - 生成选项
 * @returns PDF 字节数据
 */
export async function generatePdf(
    spec: PdfDocumentSpec,
    options: PdfGenerateOptions = {}
): Promise<Uint8Array> {
    try {
        if (!spec || typeof spec !== 'object') throw new TypeError('缺少文档规格');

        const doc = await PDFDocument.create();
        const fonts = await buildFonts(doc, { ...(spec.fonts || {}), ...(options.fonts || {}) });

        // 资源表：显式传入优先，否则沿用规格自带的 assets（往返零搬运）
        const rawAssets = options.assets ?? spec.assets;
        const assets: AssetTable = Array.isArray(rawAssets) ? toAssetTable(rawAssets) : (rawAssets || {});

        // 未内嵌字体时提前拦下中文，避免输出空白方块
        if (!fonts.embedded) {
            for (const element of collectElements(spec)) {
                const cjkText =
                    element.type === 'text' || element.type === 'heading'
                        ? element.text ?? ''
                        : element.type === 'runs'
                          ? element.runs.map((r) => r.text).join('')
                          : element.type === 'table'
                            ? element.rows.flat().map(tableCellText).join('')
                            : '';
                if (hasCjk(cjkText)) {
                    throw new JdocError(
                        '生成中文 PDF 需要内嵌字体：请通过 fonts.regular 传入中文 TTF/OTF 字节',
                        'pdf'
                    );
                }
            }
        }

        const defaultSize = spec.defaultSize ?? 12;
        const defaultColor = parseColor(spec.defaultColor, rgb(0.12, 0.12, 0.12));

        // 元数据
        if (spec.metadata) {
            doc.setTitle(spec.metadata.title || '');
            doc.setAuthor(spec.metadata.author || '');
            doc.setSubject(spec.metadata.subject || '');
            doc.setKeywords(spec.metadata.keywords ? spec.metadata.keywords.split(/[,;\s]+/).filter(Boolean) : []);
            doc.setProducer(spec.metadata.producer || '@fefeding/jdoc');
            doc.setCreator(spec.metadata.creator || '@fefeding/jdoc');
            if (spec.metadata.created) {
                const date = new Date(spec.metadata.created);
                if (!Number.isNaN(date.getTime())) doc.setCreationDate(date);
            }
            if (spec.metadata.modified) {
                const date = new Date(spec.metadata.modified);
                if (!Number.isNaN(date.getTime())) doc.setModificationDate(date);
            }
        }

        /** 显式分页：每页独立坐标系，元素超出页高不再自动拆页 */
        const renderPage = async (elements: PdfElement[], pageSpec: PdfPageSpecLike & { size?: PdfPageSize; orientation?: 'portrait' | 'landscape' }) => {
            const [pageWidth, pageHeight] = resolvePageSize(pageSpec.size, pageSpec.orientation === 'landscape');
            const margin = resolveMargin(pageSpec.margin);
            const page = doc.addPage([pageWidth, pageHeight]);
            const ctx: DrawContext = {
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
                // 带显式 y 的元素按「距页顶」的绝对坐标绘制（PPTX 幻灯片版式直排
                // 依赖此语义），不参与流式光标推进；未带 y 的元素保持流式排版。
                const absolute = (element as { y?: number }).y !== undefined;
                const used = await drawElement(
                    { ...ctx, originY: absolute ? margin.top : margin.top + cursorY },
                    element
                );
                if (!absolute) cursorY += used;
            }
        };

        if (spec.pages?.length) {
            for (const pageSpec of spec.pages) {
                await renderPage(pageSpec.elements || spec.elements || [], pageSpec);
            }
        } else {
            // 自动分页：行级流式排版，内容逐行溢出即另起一页
            const elements = spec.elements || [];
            const [pageWidth, pageHeight] = resolvePageSize(spec.pageSize, false);
            const margin = resolveMargin(spec.margin);
            const usableHeight = pageHeight - margin.top - margin.bottom;

            let page = doc.addPage([pageWidth, pageHeight]);
            let pageNumber = 1; // 从 1 开始，供 onElementLayout 回调使用
            let ctx: DrawContext = {
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
            const flow: FlowState = {
                page,
                marginTop: margin.top,
                cursorY: 0,
                usableHeight,
                ensure(need: number) {
                    // 页首不换页（元素高于整页时避免死循环）
                    if (this.cursorY + need <= this.usableHeight || this.cursorY === 0) return false;
                    newPage();
                    return true;
                }
            };

            // 段间距：Word 规则为「前段 spacingAfter + 本段 spacingBefore」相加
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
    } catch (error) {
        throw wrapError('pdf', '生成', error);
    }
}

/** 收集文档中所有元素（供前置校验使用） */
function collectElements(spec: PdfDocumentSpec): PdfElement[] {
    if (spec.pages?.length) return spec.pages.flatMap((page) => page.elements || []);
    return spec.elements || [];
}
