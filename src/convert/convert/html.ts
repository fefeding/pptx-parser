/**
 * HTML <-> 文档块模型（DocxBlock）互转
 *
 * 解析端用 parse5 把（良构或不严整的）HTML 解析为 DOM 树，再映射为 `DocxBlock`：
 * 标题/段落/列表/表格/图片/代码块/引用/分隔线都能还原。行内样式（加粗/斜体/
 * 下划线/删除线/行内代码/超链接）保留为 `DocxRun`，因此 HTML 可作为 md/docx/pdf
 * 互转路径中的一环。
 *
 * 生成端（blocks -> HTML）已由 `docx/blocksToHtml` 提供，本模块只负责 HTML -> blocks。
 *
 * @module convert/html
 */

import { parse } from 'parse5';
import type {
    DocxBlock,
    DocxCell,
    DocxCellDetail,
    DocxListItem,
    DocxRun
} from '../docx/types';

/** 宽松的 HTML 节点结构（parse5 的 Node 是联合类型，这里用最小接口访问） */
interface HtmlNode {
    nodeName: string;
    value?: string;
    attrs?: { name: string; value: string }[];
    childNodes?: HtmlNode[];
}

/** 行内级标签（不单独成块，作为 run 的一部分） */
const INLINE_TAGS = new Set(['strong', 'b', 'em', 'i', 'u', 's', 'strike', 'del', 'code', 'a', 'span', 'br', 'sub', 'sup', 'small', 'mark']);
/** 跳过的内容标签 */
const SKIP_TAGS = new Set(['script', 'style', 'noscript', 'meta', 'link', 'head', 'title', 'template', 'textarea']);

/** 取节点的属性值 */
function attr(node: HtmlNode, name: string): string | undefined {
    if (node.nodeName === '#text' || node.nodeName === '#comment') return undefined;
    const attrs = (node as { attrs?: { name: string; value: string }[] }).attrs;
    return attrs?.find((a) => a.name === name)?.value;
}

/** 取节点全部文本（递归拼接，去除多余空白） */
function textContent(node: HtmlNode): string {
    if (node.nodeName === '#text') return (node as { value: string }).value;
    const children = (node as { childNodes?: HtmlNode[] }).childNodes || [];
    return children.map(textContent).join('');
}

/** 合并连续同样式 run */
function mergeRuns(runs: DocxRun[]): DocxRun[] {
    const out: DocxRun[] = [];
    for (const r of runs) {
        const last = out[out.length - 1];
        const same =
            !!last &&
            !!last.bold === !!r.bold &&
            !!last.italic === !!r.italic &&
            !!last.strike === !!r.strike &&
            !!last.underline === !!r.underline &&
            !!last.code === !!r.code &&
            (last.link ?? '') === (r.link ?? '') &&
            (last.color ?? '') === (r.color ?? '');
        if (same) last.text += r.text;
        else out.push({ ...r });
    }
    return out;
}

/** 把一组子节点转成行内 run 序列（含嵌套样式） */
function inlineRuns(nodes: HtmlNode[], inherited: Partial<DocxRun> = {}): DocxRun[] {
    const runs: DocxRun[] = [];
    const push = (text: string, style: Partial<DocxRun>) => {
        if (!text) return;
        runs.push({ text, ...inherited, ...style });
    };

    for (const node of nodes) {
        if (node.nodeName === '#text') {
            push((node as { value: string }).value, {});
        } else if (node.nodeName === 'br') {
            push('\n', {});
        } else if (SKIP_TAGS.has(node.nodeName)) {
            // 跳过脚本/样式
        } else if (INLINE_TAGS.has(node.nodeName)) {
            const children = (node as { childNodes?: HtmlNode[] }).childNodes || [];
            const style: Partial<DocxRun> = { ...inherited };
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
                    if (href) style.link = href;
                    break;
                }
                case 'span': {
                    const css = attr(node, 'style') || '';
                    // 只认独立的 color 声明，避免 background-color 误命中
                    const colorMatch = /(?:^|;)\s*color\s*:\s*([^;]+)/i.exec(css);
                    if (colorMatch) style.color = colorMatch[1].trim();
                    break;
                }
                default:
                    break;
            }
            runs.push(...inlineRuns(children, style));
        } else {
            // 其它块级元素内的文本也按普通文本收集（保证不丢字）
            push(textContent(node), inherited);
        }
    }
    return mergeRuns(runs);
}

/** 把一组子节点转成块序列（处理行内文本与块级元素的混合） */
function childrenToBlocks(nodes: HtmlNode[]): DocxBlock[] {
    const blocks: DocxBlock[] = [];
    let pending: DocxRun[] = [];

    const flush = () => {
        const runs = mergeRuns(pending);
        if (runs.length && runs.some((r) => r.text.trim())) {
            blocks.push({ type: 'paragraph', runs });
        }
        pending = [];
    };

    for (const node of nodes) {
        if (node.nodeName === '#text') {
            const text = (node as { value: string }).value;
            if (text.trim()) pending.push(...inlineRuns([node]));
            continue;
        }
        if (node.nodeName === '#comment') continue;
        if (SKIP_TAGS.has(node.nodeName)) continue;

        const children = (node as { childNodes?: HtmlNode[] }).childNodes || [];
        const tag = node.nodeName;
        const level = /^h([1-6])$/.exec(tag);
        if (level) {
            flush();
            blocks.push({
                type: 'heading',
                level: Number(level[1]) as 1 | 2 | 3 | 4 | 5 | 6,
                runs: inlineRuns(children),
                text: textContent(node).trim()
            });
        } else if (tag === 'p') {
            flush();
            // p 内直接嵌 img：拆成 段落/图片 序列，避免图片作为行内节点被丢弃
            if (children.some((c) => c.nodeName === 'img')) {
                let buffer: DocxRun[] = [];
                for (const c of children) {
                    if (c.nodeName === 'img') {
                        if (buffer.length) {
                            blocks.push({ type: 'paragraph', runs: buffer });
                            buffer = [];
                        }
                        const img = imageFromNode(c);
                        if (img.src) blocks.push(img);
                    } else {
                        buffer.push(...inlineRuns([c]));
                    }
                }
                if (buffer.length) blocks.push({ type: 'paragraph', runs: buffer });
            } else {
                blocks.push({ type: 'paragraph', runs: inlineRuns(children), text: textContent(node).trim() });
            }
        } else if (tag === 'ul' || tag === 'ol') {
            flush();
            blocks.push(listToBlock(node));
        } else if (tag === 'table') {
            flush();
            blocks.push(...tableToBlocks(node));
        } else if (tag === 'img') {
            flush();
            blocks.push(imageFromNode(node));
        } else if (tag === 'hr') {
            flush();
            blocks.push({ type: 'hr' });
        } else if (tag === 'pre') {
            flush();
            blocks.push(codeFromPre(node));
        } else if (tag === 'blockquote') {
            flush();
            blocks.push({
                type: 'quote',
                runs: inlineRuns(children),
                text: textContent(node).trim()
            });
        } else if (tag === 'br') {
            pending.push({ text: '\n' });
        } else if (tag === 'div' || tag === 'section' || tag === 'article' || tag === 'main' || tag === 'figure' || tag === 'header' || tag === 'footer') {
            // 容器：递归，不新增段落边界
            blocks.push(...childrenToBlocks(children));
        } else {
            // 未知标签：当作容器
            blocks.push(...childrenToBlocks(children));
        }
    }
    flush();
    return blocks;
}

/** 列表节点 -> DocxList */
function listToBlock(node: HtmlNode): Extract<DocxBlock, { type: 'list' }> {
    const ordered = node.nodeName === 'ol';
    const items: DocxListItem[] = [];
    for (const child of (node as { childNodes?: HtmlNode[] }).childNodes || []) {
        if (child.nodeName !== 'li') continue;
        const liChildren = (child as { childNodes?: HtmlNode[] }).childNodes || [];
        const inlineNodes = liChildren.filter((n) => n.nodeName !== 'ul' && n.nodeName !== 'ol');
        const nested = liChildren.find((n) => n.nodeName === 'ul' || n.nodeName === 'ol');
        const runs: DocxRun[] & { children?: DocxListItem[] } = inlineRuns(inlineNodes);
        if (nested) runs.children = listToBlock(nested).items;
        items.push(runs);
    }
    return { type: 'list', ordered, items };
}

/** 表格节点 -> 一个或多个表格块 */
function tableToBlocks(node: HtmlNode): DocxBlock[] {
    const rows: DocxCell[][] = [];
    let header = false;
    // 列 -> 该列还被上方的 rowspan 占用多少行
    const spanRows = new Map<number, number>();

    const addRow = (rowNode: HtmlNode, isHead: boolean) => {
        // 本行开始前就存在的跨行占用
        const occupied = new Set<number>();
        for (const [col, rest] of spanRows) {
            if (rest > 0) occupied.add(col);
        }
        const cells = cellsFromRow(rowNode, occupied);
        // 记录本行新产生的跨行占用（rowspan=N 时还要占住后面 N-1 行）
        let col = 0;
        for (const cell of cells) {
            const detail = cell as DocxCellDetail;
            const colspan = detail.colspan ?? 1;
            const rowspan = detail.rowspan ?? 1;
            if (rowspan > 1) {
                for (let c = col; c < col + colspan; c++) spanRows.set(c, rowspan - 1);
            }
            col += colspan;
        }
        // 本行结束：老占用各减一行
        for (const [c, rest] of spanRows) {
            const left = rest - 1;
            if (left > 0) spanRows.set(c, left);
            else spanRows.delete(c);
        }
        if (isHead || (!rows.length && cells.some((c) => isHeaderCell(c)))) header = true;
        rows.push(cells);
    };

    for (const child of (node as { childNodes?: HtmlNode[] }).childNodes || []) {
        if (child.nodeName === 'tr') {
            addRow(child, false);
        } else if (child.nodeName === 'thead' || child.nodeName === 'tbody' || child.nodeName === 'tfoot') {
            const isHead = child.nodeName === 'thead';
            for (const sub of (child as { childNodes?: HtmlNode[] }).childNodes || []) {
                if (sub.nodeName === 'tr') addRow(sub, isHead);
            }
        }
    }
    if (!rows.length) return [];
    return [{ type: 'table', header, rows }];
}

/** 取一行内的单元格（th/td），识别 colspan / rowspan */
function cellsFromRow(row: HtmlNode, occupied: Set<number>): DocxCell[] {
    const cells: DocxCell[] = [];
    let col = 0;
    for (const child of (row as { childNodes?: HtmlNode[] }).childNodes || []) {
        if (child.nodeName !== 'td' && child.nodeName !== 'th') continue;
        // 跳过被上方 rowspan 占住的列
        while (occupied.has(col)) col++;
        const colspan = Math.max(1, parseInt(attr(child, 'colspan') || '1', 10) || 1);
        const rowspan = Math.max(1, parseInt(attr(child, 'rowspan') || '1', 10) || 1);
        const children = (child as { childNodes?: HtmlNode[] }).childNodes || [];
        const inherited: Partial<DocxRun> = child.nodeName === 'th' ? { bold: true } : {};
        const runs = inlineRuns(children, inherited) as DocxRun[];
        const detail: DocxCellDetail = { type: 'cell', text: runs.map((r) => r.text).join('') };
        if (runs.length) detail.runs = runs;
        if (colspan > 1) detail.colspan = colspan;
        if (rowspan > 1) detail.rowspan = rowspan;
        cells.push(detail);
        col += colspan;
    }
    return cells;
}

/** 单元格是否表头单元格（th 默认加粗，这里以加粗 run 近似） */
function isHeaderCell(cell: DocxCell): boolean {
    if (Array.isArray(cell)) return cell.some((r) => r.bold);
    const detail = cell as DocxCellDetail;
    return !!detail.runs?.some((r) => r.bold);
}

/** 图片节点 -> DocxImage */
function imageFromNode(node: HtmlNode): Extract<DocxBlock, { type: 'image' }> {
    const src = attr(node, 'src') || '';
    const alt = attr(node, 'alt');
    const width = attr(node, 'width');
    const height = attr(node, 'height');
    // 像素 -> 磅（1px ≈ 0.75pt）
    const toPt = (v: string | undefined) => (v ? Math.round(Number(v) * 0.75 * 100) / 100 : undefined);
    return {
        type: 'image',
        src,
        alt: alt || undefined,
        width: toPt(width),
        height: toPt(height)
    };
}

/** pre 节点 -> 代码块 */
function codeFromPre(node: HtmlNode): DocxBlock {
    const codeChild = ((node as { childNodes?: HtmlNode[] }).childNodes || []).find(
        (n) => n.nodeName === 'code'
    );
    const target = codeChild || node;
    const text = textContent(target).replace(/\n$/, '');
    const className = attr(target, 'class') || '';
    const langMatch = /(?:language-|lang-)([a-zA-Z0-9_-]+)/.exec(className);
    return { type: 'code', text, language: langMatch?.[1] };
}

/**
 * 把 HTML 解析为文档块模型
 * @param html - HTML 字符串
 * @returns 块序列
 */
export function htmlToBlocks(html: string): DocxBlock[] {
    const doc = parse(html) as HtmlNode;
    // 跳过 document 顶层，直接进入 <html>/<body>
    const root = (doc as { childNodes?: HtmlNode[] }).childNodes?.find(
        (n) => n.nodeName === 'html'
    ) as HtmlNode | undefined;
    const body = root
        ? ((root as { childNodes?: HtmlNode[] }).childNodes || []).find(
              (n) => n.nodeName === 'body'
          ) as HtmlNode | undefined
        : doc;
    const target = body || root || doc;
    return childrenToBlocks((target as { childNodes?: HtmlNode[] }).childNodes || []);
}
