/**
 * SmartArt（图示）文字编辑
 *
 * 解析端把图示文字抽到两处：
 * - `el.texts`：数据部件（ppt/diagrams/dataN.xml）中**非空** `<a:t>` 的顺序列表
 *   （取号规则见 json-from-pptx 的 collectTexts：空白文本不占号）
 * - `el.shapes`：缓存绘图部件（ppt/diagrams/drawingN.xml）解析出的形状树，带 text 字段
 *
 * 导出时，带 `__raw` 的图示走「原样回写」分支：data/drawing 部件按整串 XML 落盘。
 * 因此只改 `el.texts` 不会有任何效果 —— 必须同步改写 `__raw.parts` 里两个部件的文本节点。
 *
 * 本模块负责这次同步改写，并回填 `el.shapes[].text`（画布渲染读的是形状缓存）。
 */
import { escapeXml } from '../serializer/xml-builder';

/** 数据部件（语义数据，PowerPoint 编辑 SmartArt 时真正使用的那份） */
const DATA_PART_RE = /diagrams\/data\d*\.xml$/i;
/** 缓存绘图部件（各渲染器实际显示的那份） */
const DRAWING_PART_RE = /diagrams\/drawing\d*\.xml$/i;

/** `<a:t>文字</a:t>` 或自闭合 `<a:t/>`（属性位保留原样） */
const A_T_RE = /<a:t(\s[^>]*?)?>([\s\S]*?)<\/a:t>|<a:t(\s[^>]*?)?\/>/g;

/** XML 实体解码（只需覆盖解析端 decodeEntities 支持的几种） */
function decodeXml(s: string): string {
    return String(s)
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&#(\d+);/g, (_m, d) => String.fromCharCode(Number(d)))
        .replace(/&#x([0-9a-f]+);/gi, (_m, h) => String.fromCharCode(parseInt(h, 16)))
        .replace(/&amp;/g, '&');
}

/** 改写结果 */
export interface DiagramTextUpdate {
    /** 生效后的文本列表（长度与原 el.texts 一致） */
    texts: string[];
    /** 是否真的改动了内容（false 表示无需回写） */
    changed: boolean;
}

/**
 * 按顺序替换部件里的非空 `<a:t>` 文本。
 * 空文本节点保持原样（不占号），与解析端 collectTexts 的取号规则一致。
 * @param xml 部件原始 XML
 * @param texts 新的文本列表
 * @returns 改写后的 XML 与消耗的文本个数
 */
function patchTexts(xml: string, texts: string[]): { xml: string; used: number } {
    let used = 0;
    const out = xml.replace(A_T_RE, (full: string, attrs1: string | undefined, inner: string | undefined, attrs2: string | undefined) => {
        if (inner === undefined) return full;                 // 自闭合 <a:t/>：无文字
        if (!decodeXml(inner).trim()) return full;           // 空白文本不占号
        if (used >= texts.length) return full;
        const t = texts[used++];
        const attrs = attrs1 || '';
        return `<a:t${attrs}>${escapeXml(t)}</a:t>`;
    });
    return { xml: out, used };
}

/**
 * 把新文本回填到形状缓存（el.shapes）。
 *
 * 形状的 text 是其各段落文字的拼接（如 "ABC\nabc"），与数据部件的文本条目一一对应。
 * 这里按「原值顺序匹配」把每行文字换成新值：逐行在原始文本里向后查找匹配项，
 * 命中即替换并推进游标 —— 同一段落里出现重复文字也不会串位。
 * @param el 图示元素
 * @param origTexts 修改前的文本列表
 * @param newTexts 修改后的文本列表
 */
function syncShapeTexts(el: any, origTexts: string[], newTexts: string[]): void {
    const shapes: any[] = Array.isArray(el.shapes) ? el.shapes : [];
    let cursor = 0;
    for (const s of shapes) {
        if (!s || s.connector) continue;
        if (typeof s.text !== 'string' || !s.text) continue;
        const lines: string[] = String(s.text).split('\n');
        let hit = false;
        const next = lines.map((line: string) => {
            for (let i = cursor; i < origTexts.length; i++) {
                if (origTexts[i] === line) {
                    cursor = i + 1;
                    hit = true;
                    return newTexts[i] ?? line;
                }
            }
            return line;
        });
        if (hit) s.text = next.join('\n');
    }
}

/**
 * 修改 SmartArt 图示的文字。
 *
 * 会同步改写：数据部件与缓存绘图部件的 `<a:t>` 文本节点、el.texts、以及形状缓存 el.shapes[].text。
 * @param el 编辑模型里的图示元素（会被就地修改）
 * @param texts 新的文字列表（按 el.texts 的顺序）
 */
export function updateDiagramTexts(el: any, texts: string[]): DiagramTextUpdate {
    if (!el || el.type !== 'diagram') return { texts: [], changed: false };
    const orig: string[] = Array.isArray(el.texts) ? el.texts.map((t: any) => String(t ?? '')) : [];
    const next: string[] = (Array.isArray(texts) ? texts : []).map((t: any) => String(t ?? ''));
    // 长度对齐：多出的截断、不足的补原值（避免序号错位导致文字串位）
    while (next.length < orig.length) next.push(orig[next.length]);
    next.length = orig.length;

    let changed = next.some((t, i) => t !== orig[i]);
    const parts: any[] = (el.__raw && Array.isArray(el.__raw.parts)) ? el.__raw.parts : [];
    const dataPart = parts.find((p) => p && DATA_PART_RE.test(String(p.path || '')));
    const drawPart = parts.find((p) => p && DRAWING_PART_RE.test(String(p.path || '')));

    if (dataPart && typeof dataPart.content === 'string') {
        const r = patchTexts(dataPart.content, next);
        if (r.xml !== dataPart.content) changed = true;
        dataPart.content = r.xml;
    }
    if (drawPart && typeof drawPart.content === 'string') {
        const r = patchTexts(drawPart.content, next);
        syncShapeTexts(el, orig, next);
        if (r.xml !== drawPart.content) changed = true;
        drawPart.content = r.xml;
    }

    el.texts = next.slice();
    return { texts: next, changed };
}
