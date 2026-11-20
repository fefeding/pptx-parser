/**
 * Markdown <-> 文档块模型（DocxBlock）互转
 *
 * 以 `markdownToBlocks` 把 Markdown 解析为库级统一 IR（`DocxBlock`），
 * 以 `blocksToMarkdown` 把 IR 渲染回 Markdown 文本。两条路径与 docx/html/pdf
 * 共用同一套块模型，因此任何格式都能经 `DocxBlock` 中转互相转换。
 *
 * 解析端直接用 marked 的词法器（lexer），保留行内样式（加粗/斜体/代码/链接）
 * 为 `DocxRun`，并对 mermaid / 数学公式走与 `markdownToHtml` 一致的扩展，
 * 保证 `md -> blocks -> md` 可无损往返。
 *
 * @module convert/markdown
 */

import { Marked, type Token, type Tokens } from 'marked';
import { markdownExtensions } from '../markdown/extensions';
import { getAsset, assetToDataUrl, toAssetTable, type DocAsset } from '../core/asset';
import { htmlToBlocks } from './html';
import type {
    DocxBlock,
    DocxCell,
    DocxCellDetail,
    DocxImage,
    DocxListItem,
    DocxRun
} from '../docx/types';

/** Markdown 解析选项（与 markdown 模块对齐） */
export interface MarkdownToBlocksOptions {
    /** 是否启用 GFM（表格、删除线、任务列表、自动链接），默认 true */
    gfm?: boolean;
    /** 是否启用公式语法，默认 true */
    math?: boolean;
    /** 是否启用 mermaid 图表语法，默认 true */
    diagrams?: boolean;
}

/** blocksToMarkdown 选项 */
export interface BlocksToMarkdownOptions {
    /** 资源表，用于把图片块上的 assetId 解析为可内联的 data URL */
    assets?: DocAsset[] | Record<string, DocAsset> | { assets?: DocAsset[] };
    /** 是否把行内代码包成反引号，默认 true */
    inlineCode?: boolean;
    /**
     * 是否把图片显示尺寸写进 title（`"w=..,h=.."`），默认 true。
     *
     * Markdown 没有图片尺寸语法，尺寸只能借 title 承载 —— 这是 md -> blocks -> md
     * 往返保尺寸的约定。但对 PPTX 这类「版式来源」没意义（宽高是幻灯片画布坐标），
     * 转出的 md 里是噪音，可关掉。
     */
    imageSizes?: boolean;
}

// ---------------------------------------------------------------------------
// Markdown -> blocks
// ---------------------------------------------------------------------------

/** 实例缓存：按选项组合分别构建，避免第一次调用的选项“焊死”后续所有调用 */
const markedCache = new Map<string, Marked>();
function getMarked(options: MarkdownToBlocksOptions): Marked {
    const key = `${options.gfm !== false}|${options.math !== false}|${options.diagrams !== false}`;
    let instance = markedCache.get(key);
    if (!instance) {
        instance = new Marked({ gfm: options.gfm !== false, breaks: false });
        instance.use(
            markdownExtensions({ enableMath: options.math !== false, enableDiagrams: options.diagrams !== false })
        );
        markedCache.set(key, instance);
    }
    return instance;
}

/** 把若干行内 token 转成 run 序列 */
function inlineToRuns(tokens: Token[] | undefined): DocxRun[] {
    const runs: DocxRun[] = [];
    const pushStyled = (source: DocxRun[], style: Partial<DocxRun>) => {
        for (const r of source) runs.push({ ...r, ...style });
    };

    for (const t of tokens || []) {
        switch (t.type) {
            case 'text':
            case 'escape':
                runs.push({ text: (t as Tokens.Text | Tokens.Escape).text });
                break;
            case 'codespan':
                runs.push({ text: (t as Tokens.Codespan).text, code: true });
                break;
            case 'strong':
                pushStyled(inlineToRuns((t as Tokens.Strong).tokens), { bold: true });
                break;
            case 'em':
                pushStyled(inlineToRuns((t as Tokens.Em).tokens), { italic: true });
                break;
            case 'del':
                pushStyled(inlineToRuns((t as Tokens.Del).tokens), { strike: true });
                break;
            case 'link':
                pushStyled(inlineToRuns((t as Tokens.Link).tokens), { link: (t as Tokens.Link).href });
                break;
            case 'jdocMathInline':
                runs.push({ text: `$${(t as Token & { text?: string }).text ?? ''}$` });
                break;
            case 'br':
                runs.push({ text: '\n' });
                break;
            case 'image':
                runs.push({ text: `![${(t as Tokens.Image).text ?? ''}](${(t as Tokens.Image).href})` });
                break;
            case 'html':
                runs.push({ text: (t as Tokens.HTML).text });
                break;
            case 'space':
            default:
                break;
        }
    }
    return mergeRuns(runs);
}

/** 合并连续的同样式 run，避免碎片段 */
function mergeRuns(runs: DocxRun[]): DocxRun[] {
    const out: DocxRun[] = [];
    for (const r of runs) {
        const last = out[out.length - 1];
        if (last && sameStyle(last, r)) last.text += r.text;
        else out.push({ ...r });
    }
    return out;
}

/** 两个 run 是否样式相同（仅比较影响渲染的标志） */
function sameStyle(a: DocxRun, b: DocxRun): boolean {
    return (
        !!a.bold === !!b.bold &&
        !!a.italic === !!b.italic &&
        !!a.strike === !!b.strike &&
        !!a.code === !!b.code &&
        (a.link ?? '') === (b.link ?? '') &&
        (a.color ?? '') === (b.color ?? '') &&
        (a.font ?? '') === (b.font ?? '')
    );
}

/** 单元格 -> run 序列 */
function cellToRuns(cell: { tokens?: Token[] } | undefined): DocxRun[] {
    return inlineToRuns(cell?.tokens);
}

/** 列表项的行内 token（展开 loose list 的 paragraph 包装与 tight list 的 text 包装） */
function listItemInline(item: Tokens.ListItem): Token[] {
    const out: Token[] = [];
    for (const t of item.tokens || []) {
        if (t.type === 'list') continue;
        if (t.type === 'paragraph') {
            // loose list（项之间有空行）时 marked 把项内容包在 paragraph 里
            const inner = (t as Tokens.Paragraph).tokens;
            if (inner?.length) out.push(...inner);
            else out.push(t);
            continue;
        }
        if (t.type === 'text') {
            // tight list 时 marked 把项内容包在 text token 里，其 text 是未解码的
            // 原文（`1\. 背景` 里的转义不会还原），要用内层已解码的 token
            const inner = (t as Tokens.Text).tokens;
            if (inner?.length) {
                out.push(...inner);
                continue;
            }
        }
        out.push(t);
    }
    return out;
}

/** 把列表 token（含嵌套）转成 DocxList */
function listToBlock(token: Tokens.List): DocxBlock {
    const itemRuns = (item: Tokens.ListItem): DocxListItem => {
        const inline = listItemInline(item);
        const nested = (item.tokens || []).find((t) => t.type === 'list') as Tokens.List | undefined;
        const runs: DocxRun[] & { children?: DocxListItem[] } = inlineToRuns(inline);
        if (nested) runs.children = nested.items.map(itemRuns);
        return runs;
    };
    return { type: 'list', ordered: token.ordered, items: token.items.map(itemRuns) };
}

/** 目录条目的页码标记（写在条目文本尾部，往返时还原成页码） */
const TOC_PAGE_PATTERN = /<!--\s*page:(\d+)\s*-->/;

/** 列表 -> 目录条目（层级按嵌套深度，页码从条目尾部的标记里取） */
function listToTocEntries(list: Tokens.List, level = 1): { level: number; text: string; page?: number }[] {
    const entries: { level: number; text: string; page?: number }[] = [];
    for (const item of list.items) {
        const inline = listItemInline(item);
        const nested = (item.tokens || []).find((t) => t.type === 'list') as Tokens.List | undefined;
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
        if (nested) entries.push(...listToTocEntries(nested, level + 1));
    }
    return entries;
}

/**
 * 段落 -> 块序列
 *
 * 图片（`![alt](src "w=..,h=..")`）无法塞进 run，单独拆成图片块，
 * 文字部分照常成段。
 */
function paragraphToBlocks(token: Tokens.Paragraph): DocxBlock[] {
    const inline = token.tokens || [];
    if (!inline.some((t) => t.type === 'image')) {
        const runs = inlineToRuns(inline);
        return [{ type: 'paragraph', runs, text: token.text }];
    }
    const blocks: DocxBlock[] = [];
    let buffer: Token[] = [];
    const flush = () => {
        if (!buffer.length) return;
        const runs = inlineToRuns(buffer);
        if (runs.length) {
            blocks.push({ type: 'paragraph', runs, text: runs.map((r) => r.text).join('') });
        }
        buffer = [];
    };
    for (const t of inline) {
        if (t.type === 'image') {
            flush();
            blocks.push(imageBlockFromToken(t as Tokens.Image));
            continue;
        }
        buffer.push(t);
    }
    flush();
    if (!blocks.length) blocks.push({ type: 'paragraph', runs: [], text: token.text });
    return blocks;
}

/** Markdown 图片 token -> 图片块（尺寸写在 title 里，往返保留显示大小） */
function imageBlockFromToken(token: Tokens.Image): DocxImage {
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

/** 取一个 token 流里的全部纯文本（用于引用块降级等） */
function tokensToText(tokens: Token[] | undefined): string {
    let text = '';
    for (const t of tokens || []) {
        if ('text' in t && typeof (t as { text?: string }).text === 'string') text += (t as { text: string }).text;
        else if ('tokens' in t && Array.isArray((t as { tokens?: Token[] }).tokens)) {
            text += tokensToText((t as { tokens: Token[] }).tokens);
        } else if ('items' in t && Array.isArray((t as { items?: Tokens.ListItem[] }).items)) {
            text += (t as { items: Tokens.ListItem[] }).items.map((i) => tokensToText(i.tokens)).join(' ');
        }
    }
    return text;
}

/**
 * 把 Markdown 解析为文档块模型
 * @param markdown - Markdown 文本（字符串）或 UTF-8 字节
 * @returns 块序列
 */
export function markdownToBlocks(markdown: string | Uint8Array, options: MarkdownToBlocksOptions = {}): DocxBlock[] {
    const text = typeof markdown === 'string' ? markdown : new TextDecoder('utf-8').decode(markdown);
    const tokens = getMarked(options).lexer(text);
    const blocks: DocxBlock[] = [];
    // `<!-- toc -->` 标记：后随的列表按目录还原（docx -> md 写入的往返约定）
    let expectToc = false;

    for (const token of tokens) {
        switch (token.type) {
            case 'heading': {
                const t = token as Tokens.Heading;
                const runs = t.tokens ? inlineToRuns(t.tokens) : undefined;
                blocks.push({
                    type: 'heading',
                    level: Math.min(6, Math.max(1, t.depth || 1)) as 1 | 2 | 3 | 4 | 5 | 6,
                    text: t.text,
                    runs
                });
                break;
            }
            case 'paragraph': {
                // 段落里可能夹着图片（`![alt](data:...)`），要拆成独立的图片块，
                // 否则图片会退化成一行字面量 `![alt](...)`，转回 docx 就没有图了
                blocks.push(...paragraphToBlocks(token as Tokens.Paragraph));
                break;
            }
            case 'list': {
                const listToken = token as Tokens.List;
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
                // marked 的表格 token 结构随版本略有差异，这里用宽松断言访问
                const t = token as unknown as {
                    header?: { tokens?: Token[] }[];
                    rows?: { cells?: { tokens?: Token[] }[]; tokens?: { tokens?: Token[] }[] }[];
                };
                const rows: DocxCell[][] = [];
                const header = (t.header || []).map((c) => [cellToRuns(c)] as DocxCell[]);
                rows.push(header as DocxCell[]);
                for (const row of t.rows || []) {
                    const cells = (row.cells || row.tokens || []) as { tokens?: Token[] }[];
                    rows.push(cells.map((c) => cellToRuns(c)) as DocxCell[]);
                }
                blocks.push({ type: 'table', header: true, rows });
                break;
            }
            case 'code': {
                const t = token as Tokens.Code;
                blocks.push({ type: 'code', text: t.text, language: t.lang || undefined });
                break;
            }
            case 'hr':
                blocks.push({ type: 'hr' });
                break;
            case 'blockquote': {
                const t = token as Tokens.Blockquote;
                blocks.push({ type: 'quote', runs: inlineToRuns(t.tokens), text: tokensToText(t.tokens) });
                break;
            }
            case 'jdocDiagramCode': {
                const t = token as Token & { lang?: string; text?: string };
                blocks.push({ type: 'code', text: t.text || '', language: t.lang || 'mermaid' });
                break;
            }
            case 'jdocMathBlock': {
                const t = token as Token & { text?: string };
                blocks.push({ type: 'code', text: t.text || '', language: 'math' });
                break;
            }
            case 'html': {
                const html = (token as Tokens.HTML).text;
                // 目录标记：docx -> md 用 `<!-- toc -->` + 列表表示目录，这里识别回来
                if (/^\s*<!--\s*toc\s*-->\s*$/i.test(html)) {
                    expectToc = true;
                    break;
                }
                // 原始 HTML 块：交给 htmlToBlocks 还原，保持 HTML 互转能力
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

// ---------------------------------------------------------------------------
// blocks -> Markdown
// ---------------------------------------------------------------------------

/**
 * 转义 Markdown 行内语法字符
 *
 * 原文里的 `*` `_` `[` `]` 会被当成语法：`U*` 与 `U*` 之间的内容会被解析成斜体，
 * 文字直接消失（docx -> md -> docx 往返后语义就变了）。
 */
function escapeMarkdownText(text: string): string {
    // 用反斜杠转义而不是 HTML 实体：marked 解析时会把 `\<` 还原成 `<`，
    // 但不会解码 `&lt;`（原样保留，往返后文本就变了）
    return text.replace(/([\\`*_[\]~<>])/g, '\\$1');
}

/** 转义行首的块级语法（`# `、`- `、`1. `、`> `），否则会被解析成标题 / 列表 / 引用 */
function escapeLineStart(line: string): string {
    return line.replace(
        /^(\s*)(#{1,6}\s|[-+]\s|\d+[.)]\s|>\s)/,
        (_all, space: string, mark: string) => space + mark.replace(/([#\-+.>])/g, '\\$1')
    );
}

/** 块的原始文本（不含 Markdown 标记），用于代码块等场景 */
function plainTextOf(block: DocxBlock): string {
    const runsText = (runs: DocxRun[] | undefined, text: string | undefined) =>
        runs ? runs.map((r) => r.text ?? '').join('') : (text ?? '');
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

/**
 * 首行是否带 4 空格以上缩进
 *
 * CommonMark 里这种行会被当成「缩进代码块」：内容原样保留（连反斜杠转义都不解码），
 * 往返后文本就变了。缩进段落（伪代码、代码清单）应该用围栏代码块表达。
 */
function hasIndentedCode(text: string): boolean {
    return /^(?: {4,}|\t)/.test(text);
}

/** 围栏代码块（内容原样保留，不做 Markdown 转义） */
function fencedCode(text: string): string {
    // 内容里出现围栏就加长围栏，避免提前闭合
    // 本仓库的 TS 7 会把 `|| []` 的 reduce 推断成 never[] 的 reduce，改成显式类型的循环。
    const marks: string[] = text.match(/`+/g) || [];
    let longest = 0;
    for (const mark of marks) longest = Math.max(longest, mark.length);
    const fence = '`'.repeat(Math.max(3, longest + 1));
    return `${fence}\n${text.replace(/\n$/, '')}\n${fence}`;
}

/** Markdown 能表达的样式键（font/size/color 在 md 里无对应语法，忽略） */
function mdStyleKey(r: DocxRun): string {
    return [r.bold ? 'B' : '', r.italic ? 'I' : '', r.strike ? 'S' : '', r.code ? 'C' : '', r.link ?? ''].join('|');
}

/** 行内 run 序列 -> Markdown 文本 */
function runsToMarkdown(runs: DocxRun[], inlineCode: boolean): string {
    // 先合并仅在 font/size/color 上不同的相邻 run（md 无法表达这些，否则会
    // 产生 **a****b** 这类相邻加粗标记，渲染虽正确但源文件很脏）
    const merged: DocxRun[] = [];
    for (const r of runs) {
        const last = merged[merged.length - 1];
        if (last && mdStyleKey(last) === mdStyleKey(r)) last.text += r.text;
        else merged.push({ ...r });
    }
    let out = '';
    for (const r of merged) {
        if (!r.text) continue;
        let text = r.text;
        if (r.code && inlineCode) {
            // 代码片段内容保持原样（反引号内不再转义）
            text = '`' + text + '`';
        } else {
            text = escapeMarkdownText(text);
            if (r.bold) text = `**${text}**`;
            if (r.italic) text = `*${text}*`;
            if (r.strike) text = `~~${text}~~`;
        }
        if (r.link) text = `[${text}](${r.link})`;
        out += text;
    }
    return out;
}

/** 引用块 -> Markdown（每行补 `> ` 前缀，无引用内容时至少输出 `>`） */
function quoteToMarkdown(runs: DocxRun[] | undefined, text: string | undefined, inlineCode: boolean): string {
    const content = runs ? runsToMarkdown(runs, inlineCode) : text ?? '';
    if (!content.trim()) return '>';
    return content
        .split('\n')
        .map((line) => (line.trim() ? '> ' + line : '>'))
        .join('\n');
}

/** 单元格 -> 纯文本（用于 Markdown 表格） */
function cellToText(cell: DocxCell | undefined, inlineCode: boolean): string {
    if (cell == null) return '';
    if (typeof cell === 'string') return cell;
    if (Array.isArray(cell)) return runsToMarkdown(cell, inlineCode);
    const detail = cell as { blocks?: DocxBlock[]; runs?: DocxRun[]; text?: string };
    if (detail.blocks) return detail.blocks.map((b) => blockToMarkdownLine(b, inlineCode)).join(' ');
    if (detail.runs) return runsToMarkdown(detail.runs, inlineCode);
    return detail.text ?? '';
}

/** 单个块转成单行 Markdown（不含结尾换行，供表格/引用等复用） */
function blockToMarkdownLine(block: DocxBlock, inlineCode: boolean, imageSizes = true): string {
    switch (block.type) {
        case 'heading':
            return `${'#'.repeat(block.level)} ${block.runs ? runsToMarkdown(block.runs, inlineCode) : block.text ?? ''}`;
        case 'paragraph': {
            const raw = plainTextOf(block);
            // 缩进段落（伪代码、代码清单）用围栏代码块表达：否则 Markdown 会当成
            // 缩进代码块，转义不还原，往返后文本就变了
            if (hasIndentedCode(raw)) return fencedCode(raw);
            const line = block.runs ? runsToMarkdown(block.runs, inlineCode) : block.text ?? '';
            return escapeLineStart(line);
        }
        case 'code':
            return '```' + (block.language || '') + '\n' + block.text + '\n```';
        case 'image': {
            // 尺寸写进 title（Markdown 没有图片尺寸语法），往返时还原成显示大小；
            // imageSizes=false 时不写（PPTX 等版式来源的宽高是画布坐标，写进 md 只是噪音）
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
            // 引用块：保持 `> ` 前缀（内部换行按引用语法逐行补前缀）
            return quoteToMarkdown(block.runs, block.text, inlineCode);
        case 'textBox': {
            // Markdown 没有文本框概念：多段文本框（多为伪代码）用围栏代码块保留
            // 内容与缩进，单段的按普通段落输出
            const paragraphs = block.paragraphs.map(plainTextOf);
            if (paragraphs.length > 1) return fencedCode(paragraphs.join('\n'));
            const text = paragraphs[0] ?? '';
            if (hasIndentedCode(text)) return fencedCode(text);
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

/**
 * 目录块 -> Markdown
 *
 * Markdown 没有目录语法，这里用 `<!-- toc -->` 标记 + 层级列表表示：
 * 标记让 md -> blocks 能把它还原成目录块（往返闭合），列表本身也便于人读。
 */
function renderToc(toc: Extract<DocxBlock, { type: 'toc' }>): string {
    const lines: string[] = ['<!-- toc -->'];
    for (const entry of toc.entries) {
        const level = Math.min(Math.max(entry.level || 1, 1), 6);
        const indent = '  '.repeat(level - 1);
        const page = entry.page !== undefined ? ` <!-- page:${entry.page} -->` : '';
        // 条目文本自带编号（Word 目录把「1. 背景」写进文本），要转义掉行首的
        // 「1. 」，否则回读时会被当成有序列表标记，编号就没了
        const text = escapeLineStart(escapeMarkdownText(entry.text));
        lines.push(`${indent}- ${text}${page}`);
    }
    return lines.join('\n');
}

/** 列表渲染（支持嵌套，每层 2 空格缩进） */
function renderList(list: Extract<DocxBlock, { type: 'list' }>, depth: number, inlineCode: boolean): string {
    const indent = '  '.repeat(depth);
    const lines: string[] = [];
    let ordinal = 0;
    list.items.forEach((item) => {
        // DocxListItem 允许 string 形态（等价于 [{text}]），统一归一为 runs
        const runs: DocxRun[] & { children?: DocxListItem[] } = typeof item === 'string' ? [{ text: item }] : (item as DocxRun[]);
        const children = runs.children;
        const text = runsToMarkdown(runs, inlineCode);
        // 空项跳过：Word 里「编号在列表项、文本在下一段」的写法会产生空列表项，
        // 输出成 "- " 只会得到一堆空行和空项（还会把相邻列表拆成 loose list）
        if (!text.trim() && !children?.length) return;
        ordinal += 1;
        const prefix = list.ordered ? `${ordinal}. ` : '- ';
        lines.push(indent + prefix + text);
        if (children?.length) {
            const childList: Extract<DocxBlock, { type: 'list' }> = {
                type: 'list',
                ordered: list.ordered,
                items: children
            };
            lines.push(renderList(childList, depth + 1, inlineCode));
        }
    });
    return lines.join('\n');
}

/** 取单元格的 detail 形态（合并信息只挂在 detail 上） */
function cellDetail(cell: DocxCell | undefined): DocxCellDetail | undefined {
    if (!cell || typeof cell === 'string' || Array.isArray(cell)) return undefined;
    return cell as DocxCellDetail;
}

/** 表格里是否存在合并单元格 */
function tableHasMerges(table: Extract<DocxBlock, { type: 'table' }>): boolean {
    return table.rows.some((row) =>
        row.some((cell) => {
            const detail = cellDetail(cell);
            return !!detail && ((detail.colspan ?? 1) > 1 || (detail.rowspan ?? 1) > 1);
        })
    );
}

/**
 * 表格渲染
 *
 * 有合并单元格的表输出 HTML 表格：Markdown 管道表没有 colspan/rowspan，
 * 硬压成管道表会出现「表头 4 列、数据行 18 列」的畸形表（渲染器直接错位）。
 * 普通表格仍走 GFM 管道表，并保证各行列数一致。
 */
function renderTable(table: Extract<DocxBlock, { type: 'table' }>, inlineCode: boolean): string {
    if (!table.rows.length) return '';
    if (tableHasMerges(table)) return renderTableHtml(table, inlineCode);
    const cell = (c: DocxCell | undefined) => cellToText(c, inlineCode).replace(/\|/g, '\\|').replace(/\n/g, ' ');
    const width = table.rows.reduce((max, row) => Math.max(max, row.length), 0);
    const line = (row: DocxCell[]) => {
        const cells = row.slice(0, width);
        while (cells.length < width) cells.push('');
        return '| ' + cells.map(cell).join(' | ') + ' |';
    };
    const header = line(table.rows[0]);
    const divider = '| ' + Array.from({ length: width }, () => '---').join(' | ') + ' |';
    return [header, divider, ...table.rows.slice(1).map(line)].join('\n');
}

/** 单元格 -> 纯文本（不含 Markdown 标记，HTML 表格里用） */
function cellPlainText(cell: DocxCell | undefined): string {
    if (cell == null) return '';
    if (typeof cell === 'string') return cell;
    if (Array.isArray(cell)) return cell.map((r) => r.text ?? '').join('');
    const detail = cell as { blocks?: DocxBlock[]; runs?: DocxRun[]; text?: string };
    if (detail.blocks) return detail.blocks.map((b) => plainTextOf(b)).join(' ');
    if (detail.runs) return detail.runs.map((r) => r.text ?? '').join('');
    return detail.text ?? '';
}

/** 含合并单元格的表格 -> HTML 表格（保留 colspan / rowspan） */
function renderTableHtml(table: Extract<DocxBlock, { type: 'table' }>, inlineCode: boolean): string {
    const esc = (text: string) =>
        text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const lines: string[] = ['<table>'];
    table.rows.forEach((row, rowIndex) => {
        const tag = table.header && rowIndex === 0 ? 'th' : 'td';
        lines.push('  <tr>');
        for (const cell of row) {
            const detail = cellDetail(cell);
            const colspan = detail?.colspan && detail.colspan > 1 ? ` colspan="${detail.colspan}"` : '';
            const rowspan = detail?.rowspan && detail.rowspan > 1 ? ` rowspan="${detail.rowspan}"` : '';
            // HTML 表格里不能带 Markdown 标记（`**` 会变成字面文本）
            lines.push(`    <${tag}${colspan}${rowspan}>${esc(cellPlainText(cell))}</${tag}>`);
        }
        lines.push('  </tr>');
    });
    lines.push('</table>');
    return lines.join('\n');
}

/** 当前渲染上下文的资源表（blocksToMarkdown 时设置） */
let assetLookup: Record<string, DocAsset> | undefined;

/** 图片源解析（assetId -> data URL） */
function imageSrc(image: DocxImage, assets?: Record<string, DocAsset>): string {
    const table = assets ?? assetLookup;
    if (typeof image.src === 'string') return image.src;
    if (image.assetId) {
        const asset = table ? getAsset(table, image.assetId) : undefined;
        if (asset) return assetToDataUrl(asset);
    }
    return '';
}

/** 把多种资源表形态归一为 id -> 资源 的索引表 */
function normalizeAssets(assets: BlocksToMarkdownOptions['assets']): Record<string, DocAsset> | undefined {
    if (!assets) return undefined;
    if (Array.isArray(assets)) return toAssetTable(assets);
    const explicit = (assets as { assets?: DocAsset[] }).assets;
    if (explicit) return toAssetTable(explicit);
    return assets as Record<string, DocAsset>;
}

/**
 * 把文档块模型渲染为 Markdown 文本
 * @param blocks - 块序列或完整文档
 * @param options - 资源表等
 */
export function blocksToMarkdown(
    input: DocxBlock[] | { blocks?: DocxBlock[] },
    options: BlocksToMarkdownOptions = {}
): string {
    const blocks = Array.isArray(input) ? input : input.blocks || [];
    const inlineCode = options.inlineCode !== false;
    const imageSizes = options.imageSizes !== false;
    const assetTable = normalizeAssets(options.assets);
    assetLookup = assetTable;
    try {
        const parts: string[] = [];

        for (const block of blocks) {
            const text = blockToMarkdownLine(block, inlineCode, imageSizes);
            // 相邻块之间一律空行：列表后的段落缺空行会被 CommonMark 当作
            // 列表项的 lazy continuation，相邻列表会合并、相邻段落会软换行合并，
            // 统一空行才能保证块边界在往返后不丢。
            if (parts.length) parts.push('');
            parts.push(text);
        }
        return parts.join('\n').replace(/\n{3,}/g, '\n\n').trim();
    } finally {
        assetLookup = undefined;
    }
}
