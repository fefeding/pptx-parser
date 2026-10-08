/**
 * 编辑器元素级几何（纯计算，零 DOM）。
 *
 * 从 examples/editor/src/render.js 抽出的纯几何部分：这些函数只做数值运算，
 * 却被埋在 DOM 渲染文件里，导致任何非 DOM 消费者（内核 actions、命中测试、
 * 第二个编辑器的渲染层）都无法复用——下沉后渲染层反向引用本模块。
 */
import { ptToPx } from '../utils/units';

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
