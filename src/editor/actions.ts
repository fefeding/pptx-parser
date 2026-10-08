/**
 * 文档操作集合（增删改 / 层级 / 对齐 / 组合 / 幻灯片管理 / 动画）
 *
 * 从 examples/editor/src/actions.js 下沉并解耦：原实现直接 import 单例 store，
 * 这里改为 createActions(store) 工厂——同一套操作逻辑可被任意编辑器实例复用。
 */
import { clone, uid } from '../utils/misc';
import { unionBBox } from '../utils/geometry';
import { elementRect } from './geometry';
import { buildSlideFromLayout, getTheme, syncTableStyle, createGroupElement, applyTextStyle } from './model';

export function createActions(store: any) {
    /* ======================= 元素增删改 ======================= */
    function cloneElement(el: any): any {
        const c = clone(el);
        c.id = uid();
        if (c.children) c.children = c.children.map(cloneElement);
        return c;
    }

    function addElement(el: any, opts: any = {}) {
        store.update((doc: any) => {
            const slide = doc.slides[store.slideIndex];
            if (opts.center) {
                el.x = Math.round((doc.slideSize.width - el.width) / 2);
                el.y = Math.round((doc.slideSize.height - el.height) / 2);
            }
            slide.elements.push(el);
        });
        store.setSel([el.id]);
        return el;
    }

    function deleteSelected() {
        const ids = store.sel.slice();
        if (!ids.length) return;
        store.update((doc: any) => {
            const slide = doc.slides[store.slideIndex];
            slide.elements = slide.elements.filter((e: any) => !ids.includes(e.id));
        });
        store.setSel([]);
    }

    function duplicateSelected() {
        const els = store.selected();
        if (!els.length) return;
        const copies = els.map((el: any) => {
            const c = cloneElement(el);
            c.x += 16; c.y += 16;
            return c;
        });
        store.update((doc: any) => {
            doc.slides[store.slideIndex].elements.push(...copies);
        });
        store.setSel(copies.map((c: any) => c.id));
    }

    function copySelected(cut = false) {
        const els = store.selected();
        if (!els.length) return;
        store.clipboard = els.map((el: any) => clone(el));
        if (cut) deleteSelected();
    }

    function paste() {
        const items = store.clipboard || [];
        if (!items.length) return;
        const copies = items.map((el: any) => {
            const c = cloneElement(el);
            c.x += 20; c.y += 20;
            return c;
        });
        store.update((doc: any) => {
            doc.slides[store.slideIndex].elements.push(...copies);
        });
        store.setSel(copies.map((c: any) => c.id));
    }

    function selectAll() {
        store.setSel(store.elements().filter((e: any) => !e.locked).map((e: any) => e.id));
    }

    function nudge(dx: number, dy: number) {
        const els = store.selected();
        if (!els.length) return;
        store.update((doc: any) => {
            for (const el of doc.slides[store.slideIndex].elements) {
                if (!store.sel.includes(el.id) || el.locked) continue;
                el.x += dx; el.y += dy;
                if (el.children) for (const c of el.children) { c.x += dx; c.y += dy; }
            }
        }, { coalesce: 'nudge' });
    }

    /* ======================= 层级 ======================= */
    /**
     * 在文档中定位某个 id 所在的列表（顶层 slide.elements 或某个 group 的 children），
     * 并返回该数组引用。组合内的子元素被选中时，层级操作应在其所属 group 内部进行，
     * 否则顶层 elements 找不到该 id 会导致 zOrder 静默失效。
     */
    function locateList(doc: any, id: string): any[] | null {
        const slide = doc.slides[store.slideIndex];
        if (slide.elements.some((e: any) => e.id === id)) return slide.elements;
        for (const g of slide.elements) {
            if (g.type === 'group' && g.children && g.children.some((c: any) => c.id === id)) return g.children;
        }
        return null;
    }

    function zOrder(op: string) {
        const ids = store.sel.slice();
        if (!ids.length) return;
        store.update((doc: any) => {
            // 同一列表内的选中项一起重排；不同容器（顶层 vs 各 group）分别处理
            const lists = new Map<any[], Set<string>>();
            for (const id of ids) {
                const list = locateList(doc, id);
                if (!list) continue;
                if (!lists.has(list)) lists.set(list, new Set());
                (lists.get(list) as Set<string>).add(id);
            }
            for (const [list, selSet] of lists) {
                const selIds = Array.from(selSet);
                if (op === 'front') {
                    const picked = list.filter((e: any) => selSet.has(e.id));
                    const rest = list.filter((e: any) => !selSet.has(e.id));
                    list.length = 0;
                    list.push(...rest, ...picked);
                } else if (op === 'back') {
                    const picked = list.filter((e: any) => selSet.has(e.id));
                    const rest = list.filter((e: any) => !selSet.has(e.id));
                    list.length = 0;
                    list.push(...picked, ...rest);
                } else if (op === 'forward') {
                    // 每个选中项上移一位；从右往左处理，相邻选中项作为整体移动
                    for (let i = list.length - 2; i >= 0; i--) {
                        if (selSet.has(list[i].id) && !selSet.has(list[i + 1].id)) {
                            [list[i], list[i + 1]] = [list[i + 1], list[i]];
                        }
                    }
                } else {
                    // backward：每个选中项下移一位；从左往右处理
                    for (let i = 1; i < list.length; i++) {
                        if (selSet.has(list[i].id) && !selSet.has(list[i - 1].id)) {
                            [list[i], list[i - 1]] = [list[i - 1], list[i]];
                        }
                    }
                }
            }
        });
        // 重排后重新绘制选择框（元素 DOM 已重建，旧的 .is-sel / sel-box 会失效）
        store.emit('sel');
    }

    function moveTo(el: any, x: any, y: any) {
        const dx = x == null ? 0 : x - el.x;
        const dy = y == null ? 0 : y - el.y;
        el.x += dx; el.y += dy;
        if (el.children) for (const c of el.children) { c.x += dx; c.y += dy; }
    }

    /* ======================= 对齐 / 分布 ======================= */
    function alignElements(op: string) {
        const els = store.selected();
        if (!els.length) return;
        const box: any = els.length > 1
            ? unionBBox(els.map((e: any) => elementRect(e)))
            : { x: 0, y: 0, width: store.doc.slideSize.width, height: store.doc.slideSize.height };
        store.update((doc: any) => {
            for (const el of doc.slides[store.slideIndex].elements) {
                if (!store.sel.includes(el.id) || el.locked) continue;
                const r = elementRect(el);
                const dx = r.x - el.x, dy = r.y - el.y;
                switch (op) {
                    case 'left': moveTo(el, box.x - dx, null); break;
                    case 'hcenter': moveTo(el, box.x + box.width / 2 - (r.width / 2) - dx, null); break;
                    case 'right': moveTo(el, box.x + box.width - r.width - dx, null); break;
                    case 'top': moveTo(el, null, box.y - dy); break;
                    case 'vcenter': moveTo(el, null, box.y + box.height / 2 - r.height / 2 - dy); break;
                    case 'bottom': moveTo(el, null, box.y + box.height - r.height - dy); break;
                }
            }
        });
    }

    function distribute(op: string) {
        const els = store.selected();
        if (els.length < 3) return;
        store.update((doc: any) => {
            const work = doc.slides[store.slideIndex].elements.filter((e: any) => store.sel.includes(e.id) && !e.locked);
            const rects = work.map((e: any) => ({ el: e, r: elementRect(e) }));
            if (op === 'h') {
                rects.sort((a: any, b: any) => a.r.x - b.r.x);
                const first = rects[0], last = rects[rects.length - 1];
                const span = (last.r.x + last.r.width) - first.r.x;
                const total = rects.reduce((s: number, it: any) => s + it.r.width, 0);
                const gap = (span - total) / (rects.length - 1);
                let cursor = first.r.x;
                for (const it of rects) {
                    moveTo(it.el, cursor - (it.r.x - it.el.x), null);
                    cursor += it.r.width + gap;
                }
            } else {
                rects.sort((a: any, b: any) => a.r.y - b.r.y);
                const first = rects[0], last = rects[rects.length - 1];
                const span = (last.r.y + last.r.height) - first.r.y;
                const total = rects.reduce((s: number, it: any) => s + it.r.height, 0);
                const gap = (span - total) / (rects.length - 1);
                let cursor = first.r.y;
                for (const it of rects) {
                    moveTo(it.el, null, cursor - (it.r.y - it.el.y));
                    cursor += it.r.height + gap;
                }
            }
        });
    }

    /* ======================= 组合 ======================= */
    function groupSelection() {
        const els = store.selected();
        if (els.length < 2) return;
        const group = createGroupElement(els.map((el: any) => clone(el)));
        const rect: any = unionBBox(els.map((e: any) => elementRect(e)));
        group.x = rect.x; group.y = rect.y; group.width = rect.width; group.height = rect.height;
        const ids = els.map((e: any) => e.id);
        store.update((doc: any) => {
            const slide = doc.slides[store.slideIndex];
            const minIdx = Math.min(...ids.map((id: string) => slide.elements.findIndex((e: any) => e.id === id)));
            slide.elements = slide.elements.filter((e: any) => !ids.includes(e.id));
            slide.elements.splice(Math.min(minIdx, slide.elements.length), 0, group);
        });
        store.setSel([group.id]);
    }

    function ungroupSelection() {
        const els = store.selected().filter((e: any) => e.type === 'group');
        if (!els.length) return;
        const newIds: string[] = [];
        store.update((doc: any) => {
            const slide = doc.slides[store.slideIndex];
            const out: any[] = [];
            for (const el of slide.elements) {
                if (els.some((g: any) => g.id === el.id)) {
                    for (const c of (el.children || [])) {
                        const cc = cloneElement(c);
                        newIds.push(cc.id);
                        out.push(cc);
                    }
                } else out.push(el);
            }
            slide.elements = out;
        });
        store.setSel(newIds);
    }

    function toggleLock(force?: boolean) {
        const els = store.selected();
        if (!els.length) return;
        const val = force === undefined ? !els[0].locked : force;
        store.update((doc: any) => {
            for (const el of doc.slides[store.slideIndex].elements) {
                if (store.sel.includes(el.id)) el.locked = val;
            }
        });
        if (val) store.setSel([]);
    }

    function toggleHidden(force?: boolean) {
        const els = store.selected();
        if (!els.length) return;
        const val = force === undefined ? !els[0].hidden : force;
        store.update((doc: any) => {
            for (const el of doc.slides[store.slideIndex].elements) {
                if (store.sel.includes(el.id)) el.hidden = val;
            }
        });
    }

    /* ======================= 幻灯片 ======================= */
    function addSlide(layoutId?: string, opts: any = {}) {
        const doc = store.doc;
        const slide = buildSlideFromLayout(layoutId || 'titleBody', doc.theme, doc.slideSize);
        store.update((d: any) => {
            d.slides.splice(store.slideIndex + 1, 0, slide);
        });
        store.setSlide(store.slideIndex + 1);
        store.setSel([]);
        return slide;
    }

    function duplicateSlide(index = store.slideIndex) {
        const copy = clone(store.doc.slides[index]);
        copy.id = uid('s');
        copy.elements = copy.elements.map(cloneElement);
        store.update((d: any) => { d.slides.splice(index + 1, 0, copy); });
        store.setSlide(index + 1);
    }

    function deleteSlide(index = store.slideIndex) {
        if (store.doc.slides.length <= 1) return;
        store.update((d: any) => { d.slides.splice(index, 1); });
        store.setSlide(Math.min(index, store.doc.slides.length - 1), { force: true });
    }

    function moveSlide(from: number, to: number) {
        if (from === to || from < 0 || to < 0) return;
        store.update((d: any) => {
            const [s] = d.slides.splice(from, 1);
            d.slides.splice(to, 0, s);
        });
        store.setSlide(to, { force: true });
    }

    /** 隐藏 / 显示幻灯片（隐藏页在放映时跳过，导入的 PPTX show="0" 也会带此标记） */
    function toggleSlideHidden(index = store.slideIndex) {
        const idx = Math.max(0, Math.min(index, store.doc.slides.length - 1));
        store.update((d: any) => { d.slides[idx].hidden = !d.slides[idx].hidden; });
    }

    function applyLayout(layoutId: string) {
        const elements = buildSlideFromLayout(layoutId, store.doc.theme, store.doc.slideSize).elements;
        store.update((doc: any) => {
            const slide = doc.slides[store.slideIndex];
            slide.elements = elements;
        }, { history: true });
        store.setSel([]);
    }

    function setBackground(bg: any) {
        store.update((doc: any) => { doc.slides[store.slideIndex].background = bg; }, { coalesce: 'bg' });
    }

    function setSlideSize(size: any) {
        store.update((doc: any) => { doc.slideSize = { ...size }; });
    }

    function applyTheme(themeId: string, all = false) {
        const theme = getTheme(themeId);
        store.update((doc: any) => {
            doc.theme = themeId;
            if (all) {
                for (const s of doc.slides) {
                    if (!s.background || (typeof s.background === 'object' && s.background.type === 'solid')) {
                        s.background = { type: 'solid', color: theme.bg };
                    }
                }
            } else {
                const s = doc.slides[store.slideIndex];
                if (s && (!s.background || (typeof s.background === 'object' && s.background.type === 'solid'))) {
                    s.background = { type: 'solid', color: theme.bg };
                }
            }
        });
    }

    function setNotes(text: string) {
        store.update((doc: any) => { doc.slides[store.slideIndex].notes = text; }, { coalesce: 'notes' });
    }

    /** 对选中文本元素应用样式（同步元素默认 + 所有 run），并同步到组内子元素 */
    function applyTextStyleSel(patch: any, opts: any = {}) {
        store.update((doc: any) => {
            for (const el of doc.slides[store.slideIndex].elements) {
                if (!store.sel.includes(el.id)) continue;
                if (el.type !== 'text') continue;
                applyTextStyle(el, patch);
                if (el.children) for (const c of el.children) if (c.type === 'text') applyTextStyle(c, patch);
            }
        }, opts);
    }

    function setBackgroundImage(data: any) {
        store.update((doc: any) => { doc.slides[store.slideIndex].background = { type: 'image', data }; }, { coalesce: 'bg' });
    }

    function findInDoc(doc: any, id: string): any {
        for (const el of doc.slides[store.slideIndex].elements) {
            if (el.id === id) return el;
            if (el.children) {
                const f = el.children.find((c: any) => c.id === id);
                if (f) return f;
            }
        }
        return null;
    }

    /** 设置元素几何（支持组合：子元素同步移动/缩放） */
    function setElementGeo(id: string, geo: any) {
        store.update((doc: any) => {
            const el = findInDoc(doc, id);
            if (!el) return;
            if (geo.x != null || geo.y != null) {
                const dx = (geo.x == null ? el.x : geo.x) - el.x;
                const dy = (geo.y == null ? el.y : geo.y) - el.y;
                el.x += dx; el.y += dy;
                if (el.children) for (const c of el.children) { c.x += dx; c.y += dy; }
            }
            if (geo.w != null) {
                const w = Math.max(4, geo.w);
                if (el.type === 'group') {
                    const fx = w / el.width;
                    for (const c of el.children || []) { c.x = el.x + (c.x - el.x) * fx; c.width = Math.max(4, c.width * fx); }
                }
                el.width = w;
            }
            if (geo.h != null) {
                const hh = Math.max(4, geo.h);
                if (el.type === 'group') {
                    const fy = hh / el.height;
                    for (const c of el.children || []) { c.y = el.y + (c.y - el.y) * fy; c.height = Math.max(4, c.height * fy); }
                }
                el.height = hh;
            }
            if (geo.rot != null) el.rotation = ((geo.rot + 180) % 360) - 180;
        }, { coalesce: 'insp' });
    }

    function resizeTable(el: any, rows: number, cols: number) {
        store.update((doc: any) => {
            const target = findInDoc(doc, el.id);
            if (!target || target.type !== 'table') return;
            const curRows = target.rows.length;
            const curCols = target.rows[0] ? target.rows[0].cells.length : 0;
            const cellTpl = () => ({ text: '', fill: null, align: null, valign: null });
            for (let r = 0; r < curRows; r++) {
                let cells = target.rows[r].cells;
                while (cells.length < cols) cells.push(cellTpl());
                if (cells.length > cols) cells.length = cols;
            }
            while (target.rows.length < rows) {
                target.rows.push({ height: Math.round((el.height || 200) / Math.max(1, rows)), cells: new Array(cols).fill(0).map(cellTpl) });
            }
            if (target.rows.length > rows) target.rows.length = rows;
            const w = (el.width || 400) / Math.max(1, cols);
            target.colWidths = new Array(cols).fill(Math.round(w));
            target.rows.forEach((r: any) => { r.height = Math.round((el.height || 200) / Math.max(1, rows)); });
            syncTableStyle(target);
        });
    }

    function updateElement(id: string, patch: any, opts: any = {}) {
        store.update((doc: any) => {
            const el = findInDoc(doc, id);
            if (el) Object.assign(el, patch);
        }, opts);
    }

    /* ======================= 过渡 / 动画 ======================= */
    function setTransition(trans: any) {
        store.update((doc: any) => {
            doc.slides[store.slideIndex].transition = trans;
        }, { coalesce: 'trans' });
    }

    function addAnimation(anim: any) {
        store.update((doc: any) => {
            const slide = doc.slides[store.slideIndex];
            if (!slide.animations) slide.animations = [];
            slide.animations.push({ ...anim });
        });
    }

    function updateAnimation(index: number, patch: any) {
        store.update((doc: any) => {
            const slide = doc.slides[store.slideIndex];
            if (slide.animations && slide.animations[index]) {
                Object.assign(slide.animations[index], patch);
            }
        });
    }

    function removeAnimation(index: number) {
        store.update((doc: any) => {
            const slide = doc.slides[store.slideIndex];
            if (slide.animations) slide.animations.splice(index, 1);
        });
    }

    function moveAnimation(index: number, dir: number) {
        store.update((doc: any) => {
            const slide = doc.slides[store.slideIndex];
            if (!slide.animations) return;
            const ni = index + dir;
            if (ni < 0 || ni >= slide.animations.length) return;
            const [a] = slide.animations.splice(index, 1);
            slide.animations.splice(ni, 0, a);
        });
    }

    return {
        cloneElement, addElement, deleteSelected, duplicateSelected, copySelected, paste, selectAll, nudge,
        zOrder, alignElements, distribute, groupSelection, ungroupSelection, toggleLock, toggleHidden,
        addSlide, duplicateSlide, deleteSlide, moveSlide, toggleSlideHidden, applyLayout,
        setBackground, setSlideSize, applyTheme, setNotes, applyTextStyleSel, setBackgroundImage,
        setElementGeo, resizeTable, updateElement, findInDoc,
        setTransition, addAnimation, updateAnimation, removeAnimation, moveAnimation
    };
}

export type EditorActions = ReturnType<typeof createActions>;
