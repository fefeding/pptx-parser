/**
 * 编辑器元素级几何（纯计算，零 DOM）。
 *
 * 从 examples/editor/src/render.js 抽出的纯几何部分：这些函数只做数值运算，
 * 却被埋在 DOM 渲染文件里，导致任何非 DOM 消费者（内核 actions、命中测试、
 * 第二个编辑器的渲染层）都无法复用——下沉后渲染层反向引用本模块。
 */
import { ptToPx } from '../utils/units';
import { rotatePoint } from '../utils/geometry';

/**
 * 元素外接矩形。
 * 组合元素取子元素并集；但子元素为组合相对坐标时，其并集原点随内容浮动，
 * 不能代表组合位置，故改用声明的 x/y/w/h。
 */
export function elementRect(el: any) {
    if (el.type === 'group') {
        const kids = el.children || [];
        if (!kids.length || el.childrenCoordinates === 'relative') {
            return { x: el.x, y: el.y, width: el.width || 0, height: el.height || 0 };
        }
        let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
        for (const c of kids) {
            const r = rotatedRect(c);
            x0 = Math.min(x0, r.x0); y0 = Math.min(y0, r.y0);
            x1 = Math.max(x1, r.x1); y1 = Math.max(y1, r.y1);
        }
        return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
    }
    return { x: el.x, y: el.y, width: el.width, height: el.height };
}

/** 旋转后的轴对齐外接矩形（x0/y0/x1/y1） */
export function rotatedRect(el: any) {
    const w = el.width || 0, hh = el.height || 0;
    if (!el.rotation) return { x0: el.x, y0: el.y, x1: el.x + w, y1: el.y + hh };
    const cx = el.x + w / 2, cy = el.y + hh / 2;
    const rad = (el.rotation * Math.PI) / 180;
    const cos = Math.abs(Math.cos(rad)), sin = Math.abs(Math.sin(rad));
    const nw = w * cos + hh * sin, nh = w * sin + hh * cos;
    return { x0: cx - nw / 2, y0: cy - nh / 2, x1: cx + nw / 2, y1: cy + nh / 2 };
}

/**
 * 元素在幻灯片中的绝对外接矩形（计入所有祖先组合的偏移）。
 * 组合子元素坐标为「相对组合」(childrenCoordinates==='relative') 或「页面绝对」(默认)，
 * 二者渲染时的换算逻辑不同，故这里沿祖先链累加偏移，得到页面坐标，
 * 供选择框、移动包围盒、吸附统一使用（避免相对坐标直接绘到幻灯片坐标系而错位）。
 */
export function absoluteElementRect(el: any, elements: any[]): any {
    const ancestors: any[] = [];
    const walk = (list: any[]): boolean => {
        for (const e of list) {
            if (e === el) return true;
            if (e.children && walk(e.children)) { ancestors.push(e); return true; }
        }
        return false;
    };
    walk(elements);
    let x = el.x || 0, y = el.y || 0;
    for (const g of ancestors) {
        if (g.childrenCoordinates === 'relative') { x += g.x || 0; y += g.y || 0; }
    }
    return { x, y, width: el.width, height: el.height };
}

/** 阴影/发光等效果导致的选择框外扩边距（左右/上下，单位 px） */
export function effectMargin(el: any) {
    let left = 0, top = 0, right = 0, bottom = 0;
    if (el.shadow) {
        const s = el.shadow;
        const rad = ((s.angle ?? 45) * Math.PI) / 180;
        const dist = ptToPx(s.distance ?? 4);
        const dx = dist * Math.cos(rad);
        const dy = dist * Math.sin(rad);
        const blur = ptToPx(s.blur ?? 8);
        const EPS = 1e-6;
        if (dx > EPS) right += dx + blur; else if (dx < -EPS) left += -dx + blur;
        else { right += blur; left += blur; }
        if (dy > EPS) bottom += dy + blur; else if (dy < -EPS) top += -dy + blur;
        else { bottom += blur; top += blur; }
    }
    if (el.glow && el.glow.blur) {
        const blur = ptToPx(el.glow.blur);
        left = Math.max(left, blur); right = Math.max(right, blur);
        top = Math.max(top, blur); bottom = Math.max(bottom, blur);
    }
    return { left, top, right, bottom };
}

/* ======================= 线条：端点 / 顶点 / 连接点 ======================= */

function clamp(v: number, lo: number, hi: number) {
    return Math.max(lo, Math.min(hi, v));
}

/** 在 doc 的指定页递归查找元素（含组合子元素） */
function findInDoc(doc: any, id: string, slideIndex = 0): any {
    const list = (doc.slides && doc.slides[slideIndex] && doc.slides[slideIndex].elements) || [];
    const walk = (arr: any[]): any => {
        for (const el of arr) {
            if (el.id === id) return el;
            if (el.children) {
                const f = walk(el.children);
                if (f) return f;
            }
        }
        return null;
    };
    return walk(list);
}

/** 是否可端点/顶点编辑的线条元素（连接符预设或开放自定义曲线） */
export function isLineElement(el: any) {
    if (!el || el.type !== 'shape') return false;
    if (/^(curvedConnector|bentConnector|straightConnector|line)/.test(el.shapeType || '')) return true;
    if (el.custGeom && Array.isArray(el.custGeom.paths)) {
        return el.custGeom.paths.every((p: any) => !p.closed);
    }
    return false;
}

/** 线条两个端点的幻灯片绝对坐标（考虑 flip；rotation 为 0 时即对角点） */
export function lineEndpoints(el: any) {
    const x = el.x || 0, y = el.y || 0, w = el.width || 0, h = el.height || 0;
    let a = { x: el.flipH ? x + w : x, y: el.flipV ? y + h : y };
    let b = { x: el.flipH ? x : x + w, y: el.flipV ? y : y + h };
    if (el.rotation) {
        const cx = x + w / 2, cy = y + h / 2, deg = -el.rotation;
        a = rotatePoint(a.x, a.y, cx, cy, deg);
        b = rotatePoint(b.x, b.y, cx, cy, deg);
    }
    return { a, b };
}

/** 用两个绝对坐标端点重设线条包围盒与 flip（rotation 为 0 情形） */
export function setLineEndpoints(el: any, a: any, b: any) {
    const minX = Math.min(a.x, b.x), minY = Math.min(a.y, b.y);
    el.x = Math.round(minX); el.y = Math.round(minY);
    el.width = Math.round(Math.abs(b.x - a.x)); el.height = Math.round(Math.abs(b.y - a.y));
    el.flipH = b.x < a.x; el.flipV = b.y < a.y;
}

/** 把线条（直线 / 肘形 / 曲线）转成可顶点编辑的 custGeom 局部坐标 */
export function toVertexEditable(el: any) {
    const W = el.width || 100, H = el.height || 100;
    const { a, b } = lineEndpoints(el);
    const ax = a.x - el.x, ay = a.y - el.y;
    const bx = b.x - el.x, by = b.y - el.y;
    const st = el.shapeType || 'line';
    let commands: any[];
    if (st.startsWith('bentConnector')) {
        const mx = (ax + bx) / 2;
        commands = [
            { type: 'moveTo', x: ax, y: ay },
            { type: 'lnTo', x: mx, y: ay },
            { type: 'lnTo', x: mx, y: by },
            { type: 'lnTo', x: bx, y: by }
        ];
    } else if (st.startsWith('curvedConnector')) {
        const mx = (ax + bx) / 2, my = (ay + by) / 2;
        const c1x = (ax + mx) / 2, c1y = ay;
        const c2x = (mx + bx) / 2, c2y = by;
        commands = [
            { type: 'moveTo', x: ax, y: ay },
            { type: 'quadBezTo', x1: c1x, y1: c1y, x: mx, y: my },
            { type: 'quadBezTo', x1: c2x, y1: c2y, x: bx, y: by }
        ];
    } else {
        commands = [
            { type: 'moveTo', x: ax, y: ay },
            { type: 'lnTo', x: bx, y: by }
        ];
    }
    return { w: W, h: H, closed: false, commands };
}

/** 取 custGeom 第一路径的全部可拖点（顶点 + 贝塞尔控制点），返回绝对坐标 */
export function geomPoints(el: any) {
    const p = (el.custGeom && el.custGeom.paths && el.custGeom.paths[0]) || null;
    if (!p) return [];
    const W = el.width || 100, H = el.height || 100;
    const fx = el.flipH, fy = el.flipV;
    const toAbs = (lx: number, ly: number) => ({ x: el.x + (fx ? W - lx : lx), y: el.y + (fy ? H - ly : ly) });
    const pts: any[] = [];
    p.commands.forEach((c: any, i: number) => {
        if (c.type === 'moveTo' || c.type === 'lnTo') {
            pts.push({ cmdIndex: i, kind: 'point', abs: toAbs(c.x, c.y) });
        } else if (c.type === 'cubicBezTo') {
            pts.push({ cmdIndex: i, kind: 'c1', abs: toAbs(c.x1, c.y1) });
            pts.push({ cmdIndex: i, kind: 'c2', abs: toAbs(c.x2, c.y2) });
            pts.push({ cmdIndex: i, kind: 'point', abs: toAbs(c.x, c.y) });
        } else if (c.type === 'quadBezTo') {
            pts.push({ cmdIndex: i, kind: 'c1', abs: toAbs(c.x1, c.y1) });
            pts.push({ cmdIndex: i, kind: 'point', abs: toAbs(c.x, c.y) });
        }
    });
    return pts;
}

/** 标准连接点（site）：8 方位 + 中心，返回绝对坐标 */
export function connectionPoints(el: any) {
    const r = elementRect(el);
    const cx = r.x + r.width / 2, cy = r.y + r.height / 2;
    return [
        { site: 'top', x: cx, y: r.y },
        { site: 'bottom', x: cx, y: r.y + r.height },
        { site: 'left', x: r.x, y: cy },
        { site: 'right', x: r.x + r.width, y: cy },
        { site: 'topLeft', x: r.x, y: r.y },
        { site: 'topRight', x: r.x + r.width, y: r.y },
        { site: 'bottomLeft', x: r.x, y: r.y + r.height },
        { site: 'bottomRight', x: r.x + r.width, y: r.y + r.height },
        { site: 'center', x: cx, y: cy }
    ];
}

/** 根据 site 取某形状当前连接点绝对坐标 */
export function connectionPointOf(el: any, site: string) {
    const cps = connectionPoints(el);
    const c = cps.find((p: any) => p.site === site) || cps.find((p: any) => p.site === 'center') || cps[0];
    return { x: (c as any).x, y: (c as any).y };
}

/**
 * 命中最近连接点（用于端点拖拽时吸附成 glue）。
 * @param p 鼠标绝对坐标
 * @param excludeId 自身线条 id（不吸附自己）
 * @param els 候选形状数组（通常 store.slide.elements 顶层）
 * @param tol 吸附阈值（幻灯片坐标）；调用方需按 zoom 换算
 */
export function findGlueTarget(p: any, excludeId: string, els: any[], tol = 14) {
    let best: any = null;
    let bestD = tol;
    for (const el of els) {
        if (el.id === excludeId) continue;
        if (isLineElement(el) || el.type === 'group') continue;
        for (const cp of connectionPoints(el)) {
            const d = Math.hypot(p.x - cp.x, p.y - cp.y);
            if (d < bestD) { bestD = d; best = { shapeId: el.id, site: cp.site, point: { x: cp.x, y: cp.y } }; }
        }
    }
    return best;
}

/**
 * 被移动的形状集合变化后，重算所有 glue 到这些形状的线条端点。
 * 在同一 draft 内调用（store.update 的 mutator 内）。
 * @param doc 文档草稿
 * @param slideIndex 当前页索引
 * @param shapeIds 发生位移的形状 id 集合
 */
export function resyncGlue(doc: any, slideIndex: number, shapeIds: string[]) {
    const slide = doc.slides[slideIndex];
    if (!slide) return;
    for (const s of slide.elements) {
        if (!isLineElement(s)) continue;
        const gb = s.begin && shapeIds.includes(s.begin.shapeId);
        const ge = s.end && shapeIds.includes(s.end.shapeId);
        if (!gb && !ge) continue;
        const a = gb ? connectionPointOf(findInDoc(doc, s.begin.shapeId, slideIndex), s.begin.site) : lineEndpoints(s).a;
        const b = ge ? connectionPointOf(findInDoc(doc, s.end.shapeId, slideIndex), s.end.site) : lineEndpoints(s).b;
        setLineEndpoints(s, a, b);
    }
}

/**
 * 几何 API 聚合对象。集中引用本模块所有导出函数，供 actions / 渲染层复用，
 * 同时作为 index.ts 具名导出（isLineElement 等）的取值来源——
 * 由于该对象被 createActions 实际引用，rollup 不会将其整体摇除，
 * 从而让依赖它的具名导出也得以保留。
 */
export const geometryApi = {
    elementRect, rotatedRect, absoluteElementRect, effectMargin,
    isLineElement, lineEndpoints, setLineEndpoints, toVertexEditable, geomPoints,
    connectionPoints, connectionPointOf, findGlueTarget, resyncGlue
};
