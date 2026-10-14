/**
 * SmartArt 重布局引擎（P2-9）
 *
 * 面向已解析的 PptxDiagramElement（或创作端元素）提供「按目标布局类别重排节点」的能力：
 * - relayoutDiagram(el, opts)：对 el.nodes（或从 __raw 数据部件提取的节点树）执行布局，
 *   产出新的缓存绘图形状（PptxDiagramShape[]，px、相对图示框），并置 el.relayout=true；
 *   生成端（buildDiagramElement）据此重建 dataN.xml / drawingN.xml，
 *   layout/colors/quickStyle 复用 __raw 原件以保留 SmartArt 身份与样式。
 * - inferDiagramType(el)：从 __raw 的 layoutDef 分类（cat@type / alg@type）、
 *   diagramType 标记与缓存形状几何推断布局类别。
 * - extractDiagramNodes(el)：从 __raw 数据部件（dataN.xml）提取节点层级树，
 *   命名空间前缀无关（同时兼容标准 dgm: 与本库旧版 dsdgm: 格式）。
 *
 * 坐标约定与 extractDiagramShapes 一致：px、相对图示框左上角；
 * 连接线用退化包围盒表示（水平线 height=0 / 垂直线 width=0），prst='straightConnector1'。
 */

import type {
    PptxDiagramElement,
    PptxDiagramNode,
    PptxDiagramShape,
    DiagramLayoutType
} from '../types/pptx-document';

/** 重布局选项 */
export interface DiagramRelayoutOptions {
    /** 目标布局类别；缺省按 el.layoutType → inferDiagramType(el) 推断 */
    type?: DiagramLayoutType;
    /** 显式节点树（优先于 el.nodes / __raw 数据部件 / 缓存形状文本） */
    nodes?: PptxDiagramNode[];
    /** 画布内边距（px，默认 12） */
    padding?: number;
}

/** 全部合法布局类别 */
const LAYOUT_TYPES: ReadonlySet<string> = new Set([
    'list', 'process', 'hierarchy', 'cycle', 'pyramid', 'matrix', 'venn'
]);

/** 无样式模板时的默认色板（Office accent 系） */
const DEFAULT_PALETTE: readonly string[] = [
    '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'
];

const round2 = (v: number) => Math.round(v * 100) / 100;
const clamp = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, v));

/** 最小 XML 实体解码（数据部件文本用） */
function decodeXmlEntities(s: string): string {
    if (!s || s.indexOf('&') === -1) return s;
    return s
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&#(\d+);/g, (_, d) => {
            try { return String.fromCodePoint(Number(d)); } catch { return _; }
        })
        .replace(/&amp;/g, '&');
}

/** 从属性串中取值：<tag a="1" b="2"> → attr(attrs, 'a') === '1' */
function readAttr(attrs: string, name: string): string | undefined {
    const m = new RegExp(`\\b${name}\\s*=\\s*"([^"]*)"`).exec(attrs);
    return m ? m[1] : undefined;
}

/** __raw 载荷形态（与 json-from-pptx attachRawDeps 产出一致，做宽松防御） */
interface RawPartLike { path?: string; content?: string }
interface RawLike { parts?: RawPartLike[] }

/** 在 __raw.parts 中按类别查找原始部件内容（diagrams/<kind>N.xml），找不到返回 null */
function findRawPartContent(el: PptxDiagramElement, kind: 'data' | 'layout' | 'colors' | 'quickStyle'): string | null {
    const raw = (el as unknown as { __raw?: RawLike }).__raw;
    if (!raw || !Array.isArray(raw.parts)) return null;
    const dataPath = el.dataPath ? String(el.dataPath) : '';
    if (kind === 'data' && dataPath) {
        for (const p of raw.parts) {
            if (String(p.path || '') === dataPath && typeof p.content === 'string' && p.content) return p.content;
        }
    }
    const re = new RegExp(`(?:^|/)diagrams/${kind}\\d*\\.xml$`, 'i');
    for (const p of raw.parts) {
        const path = String(p.path || '');
        if (re.test(path) && typeof p.content === 'string' && p.content) return p.content;
    }
    return null;
}

/** 深度优先前序扁平化（与 buildDiagramDataModel 的 walk 顺序一致） */
function flattenNodes(nodes: PptxDiagramNode[], out: PptxDiagramNode[] = []): PptxDiagramNode[] {
    for (const n of nodes) {
        if (!n) continue;
        out.push(n);
        if (n.children && n.children.length) flattenNodes(n.children, out);
    }
    return out;
}

/**
 * 解析 dataModel XML 为节点树（命名空间前缀无关）。
 *
 * 兼容两种形态：
 * - 标准 OOXML：<dgm:ptLst>/<dgm:pt modelId type="node"><dgm:t><a:p><a:r><a:t>文本
 *   连接关系 <dgm:cxn type="parOf" srcId destId srcOrd destOrd>
 * - 本库旧版生成格式：<dsdgm:pt ...><dsdgm:t><dsdgm:str val="文本"/>
 *   连接关系 <dsdgm:cxn type="parOf" srcId destId>（无序号属性，按出现顺序）
 */
function parseDataModelNodes(xml: string): PptxDiagramNode[] | null {
    interface PtInfo { type: string; text: string }
    const pts = new Map<string, PtInfo>();

    const ptRe = /<([A-Za-z_][\w.-]*):pt\s+([^>]*?)(?:\/>|>((?:(?!<\/\1:pt>)[\s\S])*?)<\/\1:pt>)/g;
    let m: RegExpExecArray | null;
    while ((m = ptRe.exec(xml)) !== null) {
        const attrs = m[2];
        const body = m[3] || '';
        const modelId = readAttr(attrs, 'modelId');
        if (modelId === undefined) continue;
        pts.set(modelId, { type: readAttr(attrs, 'type') || 'node', text: extractPtText(body) });
    }
    if (!pts.size) return null;

    interface CxnInfo { type: string; srcId: string; destId: string; ord: number }
    const cxns: CxnInfo[] = [];
    const cxnRe = /<([A-Za-z_][\w.-]*):cxn\s+([^>]*?)(?:\/>|>[\s\S]*?<\/\1:cxn>)/g;
    while ((m = cxnRe.exec(xml)) !== null) {
        const attrs = m[2];
        const srcId = readAttr(attrs, 'srcId');
        const destId = readAttr(attrs, 'destId');
        if (srcId === undefined || destId === undefined) continue;
        const ordRaw = readAttr(attrs, 'destOrd') ?? readAttr(attrs, 'srcOrd');
        // destOrd 从 0 开始是合法且常见的（Number('0') 为 falsy，不能短路回退）
        const ordNum = ordRaw === undefined ? Number.NaN : Number(ordRaw);
        cxns.push({
            type: readAttr(attrs, 'type') || 'parOf',
            srcId,
            destId,
            ord: Number.isFinite(ordNum) ? ordNum : cxns.length
        });
    }

    const nodeIds = new Set<string>();
    const docIds: string[] = [];
    for (const [id, info] of pts) {
        if (info.type === 'node') nodeIds.add(id);
        else if (info.type === 'doc') docIds.push(id);
    }
    if (!nodeIds.size) return null;

    const kids = new Map<string, Array<{ id: string; ord: number }>>();
    const referenced = new Set<string>();
    for (const c of cxns) {
        if (c.type !== 'parOf') continue;
        const srcIsParent = nodeIds.has(c.srcId) || docIds.includes(c.srcId);
        if (!srcIsParent || !nodeIds.has(c.destId)) continue;
        if (!kids.has(c.srcId)) kids.set(c.srcId, []);
        kids.get(c.srcId)!.push({ id: c.destId, ord: c.ord });
        referenced.add(c.destId);
    }

    const build = (id: string, seen: Set<string>): PptxDiagramNode | null => {
        const info = pts.get(id);
        if (!info || info.type !== 'node' || seen.has(id)) return null;
        seen.add(id);
        const node: PptxDiagramNode = { text: info.text || '' };
        const children = (kids.get(id) || [])
            .slice().sort((a, b) => a.ord - b.ord)
            .map((k) => build(k.id, seen))
            .filter((x): x is PptxDiagramNode => x !== null);
        if (children.length) node.children = children;
        return node;
    };

    let roots: PptxDiagramNode[] = [];
    const seen = new Set<string>();
    for (const docId of docIds) {
        for (const k of kids.get(docId) || []) {
            const nd = build(k.id, seen);
            if (nd) roots.push(nd);
        }
    }
    if (!roots.length) {
        // 无 doc 节点：未被 parOf 引用为 dest 的 node 视为根
        for (const id of nodeIds) {
            if (!referenced.has(id)) {
                const nd = build(id, seen);
                if (nd) roots.push(nd);
            }
        }
    }
    if (!roots.length) {
        // 异常兜底：取任一节点整树
        for (const id of nodeIds) {
            const nd = build(id, new Set());
            if (nd) { roots = [nd]; break; }
        }
    }
    return roots.length ? roots : null;
}

/** 提取单个 dgm:pt 正文中的文本（标准 a:t 段落 + 旧版 *:str@val） */
function extractPtText(body: string): string {
    const strRe = /<[A-Za-z_][\w.-]*:str\s+[^>]*\bval\s*=\s*"([^"]*)"/g;
    let hasStr = false;
    const vals: string[] = [];
    let sm: RegExpExecArray | null;
    while ((sm = strRe.exec(body)) !== null) {
        hasStr = true;
        const v = decodeXmlEntities(sm[1]);
        if (v) vals.push(v);
    }
    if (hasStr) return vals.join('\n');
    // 标准 DrawingML 文本：按段落（a:p）组织，段内 run 文本（a:t）直接拼接
    const lines: string[] = [];
    for (const para of body.split(/<\/a:p>/)) {
        const runRe = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g;
        let line = '';
        let rm: RegExpExecArray | null;
        while ((rm = runRe.exec(para)) !== null) line += decodeXmlEntities(rm[1]);
        if (line) lines.push(line);
    }
    return lines.length ? lines.join('\n') : '';
}

/**
 * 从 __raw 数据部件（ppt/diagrams/dataN.xml）提取节点层级树。
 * @returns 节点树；无 __raw / 无数据部件 / 解析失败时返回 null
 */
export function extractDiagramNodes(el: PptxDiagramElement): PptxDiagramNode[] | null {
    const content = findRawPartContent(el, 'data');
    if (!content) return null;
    try {
        return parseDataModelNodes(content);
    } catch {
        return null;
    }
}

// ============================================================
// 布局类别推断
// ============================================================

/**
 * 推断 SmartArt 的布局类别（relayoutDiagram 未显式指定 type 时使用）。
 *
 * 优先级：
 * 1. __raw 的 layoutDef 分类（cat@type，其次 alg@type 启发）——最可靠；
 * 2. el.diagramType / el.layoutType 创作端标记；
 * 3. 缓存形状几何启发（垂直排列→list，水平排列→process，环形分布→cycle）；
 * 4. 兜底 'list'。
 */
export function inferDiagramType(el: PptxDiagramElement): DiagramLayoutType {
    const layout = findRawPartContent(el, 'layout');
    if (layout) {
        const cat = /<[A-Za-z_][\w.-]*:cat\b[^>]*\btype\s*=\s*"([^"]+)"/i.exec(layout);
        if (cat) {
            const t = cat[1].toLowerCase();
            if (t === 'hier' || t.includes('hierarch') || t.includes('orgchart')) return 'hierarchy';
            if (t === 'process' || t === 'flowchart') return 'process';
            if (t === 'cycle') return 'cycle';
            if (t === 'pyramid' || t === 'pyra') return 'pyramid';
            if (t === 'matrix') return 'matrix';
            if (t === 'venn') return 'venn';
            if (t === 'list') return 'list';
        }
        const algOf = (t: string) => new RegExp(`<[A-Za-z_][\\w.-]*:alg\\b[^>]*\\btype\\s*=\\s*"${t}"`, 'i').test(layout);
        if (algOf('hierRoot') || algOf('hierChild') || algOf('hierLn')) return 'hierarchy';
        if (algOf('cyCl')) return 'cycle';
        if (algOf('pyra')) return 'pyramid';
        if (algOf('snake')) return 'list';
        if (algOf('conn')) return 'process';
    }
    for (const tag of [el.diagramType, el.layoutType]) {
        const t = String(tag || '').toLowerCase();
        if (LAYOUT_TYPES.has(t)) return t as DiagramLayoutType;
    }
    // 缓存形状几何启发
    const shapes = (el.shapes || []).filter((s) => !s.connector);
    if (shapes.length >= 3) {
        const centers = shapes.map((s) => ({ x: (s.x || 0) + (s.width || 0) / 2, y: (s.y || 0) + (s.height || 0) / 2 }));
        const cx = centers.reduce((a, p) => a + p.x, 0) / centers.length;
        const cy = centers.reduce((a, p) => a + p.y, 0) / centers.length;
        const dists = centers.map((p) => Math.hypot(p.x - cx, p.y - cy));
        const dMean = dists.reduce((a, d) => a + d, 0) / dists.length;
        const dMaxDev = dists.reduce((a, d) => Math.max(a, Math.abs(d - dMean)), 0);
        if (dMean > 0 && dMaxDev / dMean < 0.28) return 'cycle';
        const xs = shapes.map((s) => s.x || 0);
        const ys = shapes.map((s) => s.y || 0);
        const spreadX = Math.max(...xs) - Math.min(...xs);
        const spreadY = Math.max(...ys) - Math.min(...ys);
        if (spreadY > spreadX * 1.4) return 'list';
        if (spreadX > spreadY * 1.4) return 'process';
    }
    return 'list';
}

// ============================================================
// 样式模板（重布局保真：沿用原缓存形状的外观）
// ============================================================

/** 从既有缓存形状提取的样式模板（节点外观/色板/连接线样式） */
interface StyleTemplate {
    /** 节点预设几何（roundRect 等） */
    nodePrst: string;
    /** 节点填充色板（按文档序收集，重布局时循环取色） */
    palette: string[];
    nodeLineColor?: string;
    nodeLineWidth?: number;
    /** 基础字号（px，文本超框时自适应收缩） */
    fontSize?: number;
    /** 文字颜色 */
    color?: string;
    bold?: boolean;
    align?: string;
    anchor?: string;
    connLineColor?: string;
    connLineWidth?: number;
    connEndArrow?: string;
}

/** 从既有缓存形状提取样式模板：节点样式取首个带文字的形状，色板按文档顺序去重收集 */
function collectStyleTemplate(el: PptxDiagramElement): StyleTemplate {
    const tmpl: StyleTemplate = { nodePrst: 'roundRect', palette: [] };
    const shapes = (el.shapes || []) as PptxDiagramShape[];
    for (const s of shapes) {
        if (s.connector) continue;
        const fill = String(s.fill || '');
        if (/^#?[0-9A-Fa-f]{6}$/.test(fill) && !tmpl.palette.includes(fill)) tmpl.palette.push(fill);
    }
    const first = shapes.find((s) => !s.connector && s.text) || shapes.find((s) => !s.connector);
    if (first) {
        if (first.prst) tmpl.nodePrst = first.prst;
        if (first.lineColor) tmpl.nodeLineColor = first.lineColor;
        if (first.lineWidth) tmpl.nodeLineWidth = first.lineWidth;
        if (first.fontSize) tmpl.fontSize = first.fontSize;
        if (first.color) tmpl.color = first.color;
        if (first.bold != null) tmpl.bold = first.bold;
        if (first.align) tmpl.align = first.align;
        if (first.anchor) tmpl.anchor = first.anchor;
    }
    const conn = shapes.find((s) => s.connector);
    if (conn) {
        if (conn.lineColor) tmpl.connLineColor = conn.lineColor;
        if (conn.lineWidth) tmpl.connLineWidth = conn.lineWidth;
        if (conn.endArrow) tmpl.connEndArrow = conn.endArrow;
    }
    return tmpl;
}

/** 文本自适应：按框宽估算每行字符数，超出框高则收缩字号（下限 9） */
function fitFontSize(text: string, w: number, h: number, base: number): number {
    if (!text) return base;
    const charW = Math.max(6, base * 0.62);
    const perLine = Math.max(3, Math.floor((w - 10) / charW));
    const lines = Math.max(1, Math.ceil(String(text).length / perLine));
    const need = lines * base * 1.4 + 6;
    return need <= h ? base : Math.max(9, Math.floor((base * h) / need));
}

/** 构建节点形状（按文档序循环取色板） */
function nodeShape(x: number, y: number, w: number, h: number, text: string, tmpl: StyleTemplate, colorIdx: number): PptxDiagramShape {
    const palette = tmpl.palette.length ? tmpl.palette : DEFAULT_PALETTE;
    const fontSize = fitFontSize(text || '', w, h, tmpl.fontSize || 14);
    const shape: PptxDiagramShape = {
        x: round2(x),
        y: round2(y),
        width: round2(Math.max(w, 1)),
        height: round2(Math.max(h, 1)),
        prst: tmpl.nodePrst || 'roundRect',
        fill: palette[colorIdx % palette.length],
        lineColor: tmpl.nodeLineColor || '#FFFFFF',
        lineWidth: tmpl.nodeLineWidth || 1,
        text: text || '',
        fontSize,
        color: tmpl.color || '#FFFFFF',
        bold: tmpl.bold ?? true,
        align: tmpl.align || 'ctr',
        anchor: tmpl.anchor || 'ctr'
    };
    return shape;
}

/** 构建连接线形状（轴对齐：水平线 height=0 / 垂直线 width=0） */
function connShape(x1: number, y1: number, x2: number, y2: number, tmpl: StyleTemplate, arrow?: boolean): PptxDiagramShape {
    const shape: PptxDiagramShape = {
        x: round2(Math.min(x1, x2)),
        y: round2(Math.min(y1, y2)),
        width: round2(Math.abs(x2 - x1)),
        height: round2(Math.abs(y2 - y1)),
        prst: 'straightConnector1',
        connector: true,
        lineColor: tmpl.connLineColor || '#A6A6A6',
        lineWidth: tmpl.connLineWidth || 1
    };
    // 包围盒默认表达 TL→BR 对角；斜向/反向线段须以 flipH/flipV 表达实际方向
    // （cycle 圆周连线的反斜线、箭头指向都依赖此编码，与编辑器 lineEndpoints 语义一致）
    if (x2 < x1) shape.flipH = true;
    if (y2 < y1) shape.flipV = true;
    if (arrow) shape.endArrow = tmpl.connEndArrow || 'triangle';
    return shape;
}

// ============================================================
// 布局算法（扁平类布局使用前序扁平化；树形类保留层级）
// ============================================================

/** 垂直列表：全宽节点框 + 框间竖直连接线 */
function layoutList(flat: PptxDiagramNode[], W: number, H: number, pad: number, tmpl: StyleTemplate): PptxDiagramShape[] {
    const n = flat.length;
    if (!n) return [];
    const gap = n > 1 ? 10 : 0;
    const boxW = Math.max(20, W - 2 * pad);
    const boxH = Math.max(18, (H - 2 * pad - gap * (n - 1)) / n);
    const shapes: PptxDiagramShape[] = [];
    let prevBottom = -1;
    for (let i = 0; i < n; i++) {
        const y = pad + i * (boxH + gap);
        shapes.push(nodeShape(pad, y, boxW, boxH, flat[i].text || '', tmpl, i));
        if (prevBottom >= 0) shapes.push(connShape(W / 2, prevBottom, W / 2, y, tmpl));
        prevBottom = y + boxH;
    }
    return shapes;
}

/**
 * 流程：横排节点框 + 框间箭头连接线；竖版画布（W < H 且节点 > 2）自动转垂直。
 */
function layoutProcess(flat: PptxDiagramNode[], W: number, H: number, pad: number, tmpl: StyleTemplate): PptxDiagramShape[] {
    const n = flat.length;
    if (!n) return [];
    const vertical = W < H && n > 2;
    if (vertical) {
        // 垂直流程 = 带箭头的列表
        const list = layoutList(flat, W, H, pad, tmpl);
        for (const s of list) if (s.connector) s.endArrow = tmpl.connEndArrow || 'triangle';
        return list;
    }
    const gap = n > 1 ? Math.max(8, Math.min(18, (W - 2 * pad) / (n * 4))) : 0;
    const boxW = Math.max(20, (W - 2 * pad - gap * (n - 1)) / n);
    const boxH = Math.max(18, H - 2 * pad);
    const midY = pad + boxH / 2;
    const shapes: PptxDiagramShape[] = [];
    let prevRight = -1;
    for (let i = 0; i < n; i++) {
        const x = pad + i * (boxW + gap);
        shapes.push(nodeShape(x, pad, boxW, boxH, flat[i].text || '', tmpl, i));
        if (prevRight >= 0) shapes.push(connShape(prevRight, midY, x, midY, tmpl, true));
        prevRight = x + boxW;
    }
    return shapes;
}

/** 树形深度 */
function treeDepth(nodes: PptxDiagramNode[], depth = 1): number {
    let max = depth;
    for (const n of nodes) {
        if (n.children && n.children.length) max = Math.max(max, treeDepth(n.children, depth + 1));
    }
    return max;
}

/**
 * 层级/组织结构：全深度 tidy tree。
 * 叶子自左向右占位，父节点居中于子节点跨度上方；连接线为肘形三段
 * （父底竖线 → 水平汇流线 → 各子竖线），与 buildHierarchyDiag 风格一致。
 */
function layoutHierarchy(nodes: PptxDiagramNode[], W: number, H: number, pad: number, tmpl: StyleTemplate): PptxDiagramShape[] {
    if (!nodes.length) return [];
    const availW = W - 2 * pad;
    const availH = H - 2 * pad;
    let boxW = clamp(W * 0.18, 80, 200);
    let boxH = clamp(H * 0.16, 32, 72);
    let gapX = 16;
    const levels = treeDepth(nodes);
    // 深树纵向溢出保护：层数过多时收缩框高
    if (levels * boxH > availH) boxH = Math.max(20, availH / levels);
    const gapY = levels > 1 ? clamp((availH - levels * boxH) / (levels - 1), 12, 48) : 0;

    // 子树宽度（叶子=boxW，父=max(自身, 子树和+间隔)）
    const widthCache = new Map<PptxDiagramNode, number>();
    const subtreeW = (node: PptxDiagramNode): number => {
        const hit = widthCache.get(node);
        if (hit !== undefined) return hit;
        let w = boxW;
        if (node.children && node.children.length) {
            let s = 0;
            for (const c of node.children) s += subtreeW(c);
            w = Math.max(boxW, s + gapX * (node.children.length - 1));
        }
        widthCache.set(node, w);
        return w;
    };
    // 溢出收缩：总宽超出可用宽度时按比例缩框
    const totalW = (list: PptxDiagramNode[]) => {
        let s = 0;
        for (const r of list) s += subtreeW(r);
        return s + gapX * Math.max(0, list.length - 1);
    };
    if (totalW(nodes) > availW) {
        const factor = availW / totalW(nodes);
        boxW = Math.max(56, boxW * factor);
        gapX = Math.max(6, gapX * factor);
        widthCache.clear();
    }

    const shapes: PptxDiagramShape[] = [];
    let nodeIdx = 0;
    const place = (node: PptxDiagramNode, left: number, depth: number): number => {
        const myIdx = nodeIdx++;
        const y = pad + depth * (boxH + gapY);
        if (!node.children || !node.children.length) {
            const x = left + (subtreeW(node) - boxW) / 2;
            shapes.push(nodeShape(x, y, boxW, boxH, node.text || '', tmpl, myIdx));
            return x + boxW / 2;
        }
        const centers: number[] = [];
        let cur = left;
        for (const c of node.children) {
            const cw = subtreeW(c);
            centers.push(place(c, cur, depth + 1));
            cur += cw + gapX;
        }
        const c0 = centers[0];
        const c1 = centers[centers.length - 1];
        const cx = (c0 + c1) / 2;
        shapes.push(nodeShape(cx - boxW / 2, y, boxW, boxH, node.text || '', tmpl, myIdx));
        // 肘形连接：父底竖线 → 汇流横线 → 各子竖线（单子节点时汇流线退化为零长度，跳过）
        const busY = y + boxH + gapY / 2;
        shapes.push(connShape(cx, y + boxH, cx, busY, tmpl));
        if (c1 > c0) shapes.push(connShape(c0, busY, c1, busY, tmpl));
        for (let i = 0; i < centers.length; i++) {
            shapes.push(connShape(centers[i], busY, centers[i], y + boxH + gapY, tmpl));
        }
        return cx;
    };

    let cursor = pad;
    for (const root of nodes) {
        const w = subtreeW(root);
        place(root, cursor, 0);
        cursor += w + gapX;
    }
    return shapes;
}

/** 环形：节点沿圆周分布 + 相邻节点连线；单节点居中 */
function layoutCycle(flat: PptxDiagramNode[], W: number, H: number, pad: number, tmpl: StyleTemplate): PptxDiagramShape[] {
    const n = flat.length;
    if (!n) return [];
    if (n === 1) {
        const bs = clamp(Math.min(W, H) * 0.4, 40, 180);
        return [nodeShape((W - bs) / 2, (H - bs) / 2, bs, bs, flat[0].text || '', tmpl, 0)];
    }
    const bs = clamp(Math.min(W, H) * 0.26, 48, 130);
    const cx = W / 2;
    const cy = H / 2;
    const R = Math.max(10, Math.min(W, H) / 2 - pad - bs / 2);
    const centers: Array<{ x: number; y: number }> = [];
    for (let i = 0; i < n; i++) {
        const ang = -Math.PI / 2 + (i * 2 * Math.PI) / n;
        centers.push({ x: cx + R * Math.cos(ang), y: cy + R * Math.sin(ang) });
    }
    const shapes: PptxDiagramShape[] = [];
    for (let i = 0; i < n; i++) {
        const c = centers[i];
        shapes.push(nodeShape(c.x - bs / 2, c.y - bs / 2, bs, bs, flat[i].text || '', tmpl, i));
        const next = centers[(i + 1) % n];
        shapes.push(connShape(c.x, c.y, next.x, next.y, tmpl));
    }
    return shapes;
}

/** 金字塔：自顶向底梯形分层（顶层最窄，逐层加宽），无连接线 */
function layoutPyramid(flat: PptxDiagramNode[], W: number, H: number, pad: number, tmpl: StyleTemplate): PptxDiagramShape[] {
    const n = flat.length;
    if (!n) return [];
    const availW = W - 2 * pad;
    const availH = H - 2 * pad;
    const gap = n > 1 ? 6 : 0;
    const layerH = Math.max(14, (availH - gap * (n - 1)) / n);
    const shapes: PptxDiagramShape[] = [];
    for (let i = 0; i < n; i++) {
        const w = (availW * (i + 1)) / n;
        const y = pad + i * (layerH + gap);
        const s = nodeShape((W - w) / 2, y, w, layerH, flat[i].text || '', tmpl, i);
        s.prst = 'trapezoid';
        shapes.push(s);
    }
    return shapes;
}

/** 矩阵：≤4 节点 2×2，更多按近似方形网格；无连接线 */
function layoutMatrix(flat: PptxDiagramNode[], W: number, H: number, pad: number, tmpl: StyleTemplate): PptxDiagramShape[] {
    const n = flat.length;
    if (!n) return [];
    const cols = n <= 4 ? 2 : Math.ceil(Math.sqrt(n));
    const rows = Math.max(1, Math.ceil(n / cols));
    const gapX = 10;
    const gapY = 10;
    const cellW = Math.max(30, (W - 2 * pad - gapX * (cols - 1)) / cols);
    const cellH = Math.max(24, (H - 2 * pad - gapY * (rows - 1)) / rows);
    const shapes: PptxDiagramShape[] = [];
    for (let i = 0; i < n; i++) {
        const r = Math.floor(i / cols);
        const c = i % cols;
        const s = nodeShape(pad + c * (cellW + gapX), pad + r * (cellH + gapY), cellW, cellH, flat[i].text || '', tmpl, i);
        s.prst = 'rect';
        shapes.push(s);
    }
    return shapes;
}

/** 维恩图：≤3 圆交叠（ellipse，无连接线）；>3 节点退化为矩阵 */
function layoutVenn(flat: PptxDiagramNode[], W: number, H: number, pad: number, tmpl: StyleTemplate): PptxDiagramShape[] {
    const n = flat.length;
    if (n > 3) return layoutMatrix(flat, W, H, pad, tmpl);
    if (!n) return [];
    const r = clamp(Math.min(W, H) * 0.3, 40, 160);
    let centers: Array<{ x: number; y: number }>;
    if (n === 1) {
        centers = [{ x: W / 2, y: H / 2 }];
    } else if (n === 2) {
        centers = [
            { x: W / 2 - r * 0.55, y: H / 2 },
            { x: W / 2 + r * 0.55, y: H / 2 }
        ];
    } else {
        centers = [
            { x: W / 2, y: H / 2 - r * 0.5 },
            { x: W / 2 - r * 0.55, y: H / 2 + r * 0.35 },
            { x: W / 2 + r * 0.55, y: H / 2 + r * 0.35 }
        ];
    }
    const shapes: PptxDiagramShape[] = [];
    for (let i = 0; i < n; i++) {
        const s = nodeShape(centers[i].x - r, centers[i].y - r, r * 2, r * 2, flat[i].text || '', tmpl, i);
        s.prst = 'ellipse';
        shapes.push(s);
    }
    return shapes;
}

/** 按布局类别分派 */
function layoutShapes(type: DiagramLayoutType, nodes: PptxDiagramNode[], W: number, H: number, pad: number, tmpl: StyleTemplate): PptxDiagramShape[] {
    switch (type) {
        case 'process':
            return layoutProcess(flattenNodes(nodes), W, H, pad, tmpl);
        case 'hierarchy':
            return layoutHierarchy(nodes, W, H, pad, tmpl);
        case 'cycle':
            return layoutCycle(flattenNodes(nodes), W, H, pad, tmpl);
        case 'pyramid':
            return layoutPyramid(flattenNodes(nodes), W, H, pad, tmpl);
        case 'matrix':
            return layoutMatrix(flattenNodes(nodes), W, H, pad, tmpl);
        case 'venn':
            return layoutVenn(flattenNodes(nodes), W, H, pad, tmpl);
        case 'list':
        default:
            return layoutList(flattenNodes(nodes), W, H, pad, tmpl);
    }
}

// ============================================================
// 主 API
// ============================================================

/**
 * SmartArt 重布局：按目标布局类别重排节点树，生成新的缓存绘图形状。
 *
 * 节点来源优先级：options.nodes > el.nodes > __raw 数据部件（extractDiagramNodes）
 * > 缓存形状文本 > el.texts。**注意**：extractDiagramNodes 每次返回新树，
 * 编辑节点后必须把树传回（options.nodes）或先调用一次 relayoutDiagram 建立 el.nodes
 * 再就地编辑 el.nodes 并再次 relayoutDiagram，否则编辑会被忽略。
 *
 * 典型用法（编辑已解析的 SmartArt）：
 * ```ts
 * const nodes = extractDiagramNodes(el);        // 取可编辑节点树（无则用 el.texts/shapes 兜底）
 * nodes[0].children.push({ text: '新节点' });     // 增删改节点
 * relayoutDiagram(el, { nodes });               // 重布局：el.shapes 立即更新
 * // 导出时生成端自动重建 data/drawing，layout/colors/quickStyle 复用 __raw 原件
 * ```
 *
 * 就地更新元素并返回：
 * - el.nodes / el.texts：规范化的节点树与扁平文本（与 dataN.xml 重建顺序一致）
 * - el.shapes：重布局结果（编辑器/渲染端立即可用）
 * - el.relayout = true：生成端据此改走「重生成 diagrams 部件」路径
 *   （data/drawing 按新内容重建，layout/colors/quickStyle 复用 __raw 原件）
 */
export function relayoutDiagram(el: PptxDiagramElement, options: DiagramRelayoutOptions = {}): PptxDiagramElement {
    const type = options.type || el.layoutType || inferDiagramType(el);
    let nodes: PptxDiagramNode[] | null =
        (options.nodes && options.nodes.length ? options.nodes : null) ||
        (el.nodes && el.nodes.length ? el.nodes : null) ||
        extractDiagramNodes(el);
    if (!nodes) nodes = nodesFromShapes(el.shapes);
    if (!nodes && el.texts && el.texts.length) nodes = el.texts.map((t) => ({ text: t }));
    if (!nodes || !nodes.length) nodes = [{ text: el.name || '标题' }];

    const W = el.width && el.width > 0 ? el.width : 480;
    const H = el.height && el.height > 0 ? el.height : 320;
    const pad = typeof options.padding === 'number' && options.padding >= 0 ? options.padding : 12;
    const tmpl = collectStyleTemplate(el);
    const shapes = layoutShapes(type, nodes, W, H, pad, tmpl);

    el.nodes = nodes;
    el.shapes = shapes;
    el.texts = flattenNodes(nodes).map((nd) => nd.text || '');
    el.layoutType = type;
    el.diagramType = type;
    el.relayout = true;
    return el;
}

/** 从缓存形状文本重建扁平节点树（无数据部件时的兜底来源） */
function nodesFromShapes(shapes: PptxDiagramShape[] | undefined): PptxDiagramNode[] | null {
    const texts = (shapes || []).filter((s) => !s.connector && s.text).map((s) => ({ text: String(s.text) }));
    return texts.length ? texts : null;
}

