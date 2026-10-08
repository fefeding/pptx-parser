/**
 * 编辑器状态仓库：文档 + 选中 + 撤销栈 + 事件总线。
 *
 * 从 examples/editor/src/store.js 下沉并解耦：
 * - 去掉 `persist` 里的 document.getElementById（改为 emit('persist')，由 UI 层订阅）
 * - 去掉硬编码 localStorage（改为可注入 StorageAdapter，默认内存），故内核可在 Node/测试运行
 * - 由单例改为工厂 createStore()，第二个编辑器可各自持有实例
 */
import { clone } from '../utils/misc';
import { createStarterDoc } from './model';

const LS_KEY = 'pptx-editor:doc:v1';
const LS_UI = 'pptx-editor:ui:v1';
const MAX_HISTORY = 120;

/** 持久化适配：浏览器传 localStorage，Node/测试可用默认内存实现 */
export interface StorageAdapter {
    getItem(key: string): string | null;
    setItem(key: string, value: string): void;
    removeItem(key: string): void;
}

class MemoryStorage implements StorageAdapter {
    private m = new Map<string, string>();
    getItem(k: string) { return this.m.has(k) ? (this.m.get(k) as string) : null; }
    setItem(k: string, v: string) { this.m.set(k, v); }
    removeItem(k: string) { this.m.delete(k); }
}

function debounce(fn: (...a: any[]) => void, wait = 300) {
    let t: any = 0;
    return (...args: any[]) => {
        clearTimeout(t);
        t = setTimeout(() => fn(...args), wait);
    };
}

/** 补齐缺失字段，防止旧数据/导入数据缺字段导致崩溃 */
export function normalizeDoc(doc: any) {
    if (!doc || !Array.isArray(doc.slides)) throw new Error('文档数据无效');
    doc.title = doc.title || '未命名演示文稿';
    doc.theme = doc.theme || 'blue';
    doc.slideSize = doc.slideSize && doc.slideSize.width ? doc.slideSize : { width: 1280, height: 720 };
    for (const s of doc.slides) {
        s.id = s.id || ('s_' + Math.random().toString(36).slice(2, 8));
        s.elements = Array.isArray(s.elements) ? s.elements : [];
        s.notes = s.notes || '';
        s.animations = Array.isArray(s.animations) ? s.animations : [];
        if (!s.transition) s.transition = null;
        for (const el of s.elements) normalizeElement(el);
    }
    return doc;
}

export function normalizeElement(el: any) {
    if (!el.id) el.id = 'e_' + Math.random().toString(36).slice(2, 9);
    el.x = Number(el.x) || 0; el.y = Number(el.y) || 0;
    // 连接线/线段的 width 或 height 为 0 是合法几何（垂直线/水平线），
    // 不能套 `|| 40` 兜底——否则垂直线全被改成 40px 宽的斜线。
    // 仅在缺值/NaN 时才回退默认尺寸；其余元素 0 尺寸视为无效照旧兜底。
    const isLinear = el.type === 'shape' && /^(curvedConnector|bentConnector|straightConnector|line)/.test(el.shapeType || '');
    const numOr = (v: any, dflt: number) => (v != null && !Number.isNaN(Number(v))) ? Number(v) : dflt;
    el.width = isLinear ? numOr(el.width, 40) : (Number(el.width) || 40);
    el.height = isLinear ? numOr(el.height, 40) : (Number(el.height) || 40);
    el.rotation = Number(el.rotation) || 0;
    if (el.type === 'text' && !el.paragraphs) el.paragraphs = [{ runs: [{ text: el.text || '' }] }];
    if (el.type === 'group') (el.children || []).forEach(normalizeElement);
}

export class EditorStore {
    doc: any = null;
    slideIndex = 0;
    sel: string[] = [];            // 选中的元素 id（当前页顶层）
    editingId: string | null = null; // 正在内联编辑的文本元素
    clipboard: any[] = [];
    zoom = 1;
    showGrid = false;
    snap = true;
    showNotes = false;
    fitted = false;
    undoStack: string[] = [];
    redoStack: string[] = [];
    private _lastCoalesce: { key: string | null; time: number } = { key: null, time: 0 };
    private _listeners = new Map<string, Set<Function>>();
    private _loading = true;
    private storage: StorageAdapter;

    constructor(opts: { storage?: StorageAdapter } = {}) {
        this.storage = opts.storage || new MemoryStorage();
    }

    /* ---------- 事件 ---------- */
    on(evt: string, fn: Function) {
        if (!this._listeners.has(evt)) this._listeners.set(evt, new Set());
        (this._listeners.get(evt) as Set<Function>).add(fn);
        return () => (this._listeners.get(evt) as Set<Function>).delete(fn);
    }
    emit(evt: string, payload?: any) {
        const set = this._listeners.get(evt);
        if (set) for (const fn of Array.from(set)) fn(payload);
    }

    /* ---------- 初始化 ---------- */
    init() {
        // 不再从存储恢复文档：打开过的 PPTX 不做持久化缓存，每次启动从默认文档开始
        try { this.storage.removeItem(LS_KEY); } catch { /* ignore */ }
        this.doc = createStarterDoc();
        try {
            const ui = JSON.parse(this.storage.getItem(LS_UI) || '{}');
            if (ui.showGrid != null) this.showGrid = !!ui.showGrid;
            if (ui.snap != null) this.snap = !!ui.snap;
            if (ui.showNotes != null) this.showNotes = !!ui.showNotes;
        } catch { /* ignore */ }
        this._loading = false;
    }

    /* ---------- 文档访问 ---------- */
    get slide() { return this.doc.slides[this.slideIndex] || this.doc.slides[0]; }
    get slideCount() { return this.doc.slides.length; }

    /** 当前页顶层元素 */
    elements() { return this.slide ? this.slide.elements : []; }
    findElement(id: string) {
        for (const el of this.elements()) if (el.id === id) return el;
        return null;
    }
    selected() {
        const els = this.elements();
        return this.sel.map((id) => els.find((e: any) => e.id === id)).filter(Boolean);
    }
    /** 递归查找（含组内子元素） */
    deepFind(id: string, list: any = this.elements()): any {
        for (const el of list) {
            if (el.id === id) return el;
            if (el.children) {
                const f = this.deepFind(id, el.children);
                if (f) return f;
            }
        }
        return null;
    }

    /* ---------- 选择 ---------- */
    setSel(ids: any, opts: any = {}) {
        const next = Array.isArray(ids) ? ids : [ids];
        const same = next.length === this.sel.length && next.every((v, i) => v === this.sel[i]);
        if (same && !opts.force) return;
        this.sel = next;
        if (this.editingId && !next.includes(this.editingId)) this.editingId = null;
        this.emit('sel');
    }
    clearSel() { this.setSel([]); }
    toggleSel(id: string) {
        this.setSel(this.sel.includes(id) ? this.sel.filter((v) => v !== id) : this.sel.concat(id));
    }

    setSlide(i: number, opts: any = {}) {
        const idx = Math.max(0, Math.min(this.doc.slides.length - 1, i));
        if (idx === this.slideIndex && !opts.force) return;
        this.slideIndex = idx;
        this.editingId = null;
        this.setSel([]);
        this.emit('slide');
        this.emit('sel');
    }

    /* ---------- 修改 / 历史 ---------- */
    /**
     * @param mutator 直接修改草稿；返回 false 可取消本次修改
     * @param opts { history?:boolean, coalesce?:string, silent?:boolean }
     */
    update(mutator: (draft: any, before: any) => any, opts: any = {}) {
        const before = this.doc;
        const draft = clone(before);
        const ret = mutator(draft, before);
        if (ret === false) return;      // 允许 mutator 取消
        this.doc = draft;
        const useHistory = opts.history !== false;
        if (useHistory) {
            const now = Date.now();
            const c = opts.coalesce;
            const canCoalesce = c && this._lastCoalesce.key === c && now - this._lastCoalesce.time < 700;
            if (!canCoalesce) {
                this.undoStack.push(JSON.stringify(before));
                if (this.undoStack.length > MAX_HISTORY) this.undoStack.shift();
            }
            this._lastCoalesce = { key: c || null, time: now };
            this.redoStack.length = 0;
        }
        this.persist();
        this.emit('doc', opts);
        return draft;
    }

    /** 连续操作（拖拽等）前手动记录一次快照 */
    snapshot() {
        this.undoStack.push(JSON.stringify(this.doc));
        if (this.undoStack.length > MAX_HISTORY) this.undoStack.shift();
        this.redoStack.length = 0;
        this._lastCoalesce = { key: null, time: 0 };
    }

    undo() {
        if (!this.undoStack.length) return false;
        const cur = JSON.stringify(this.doc);
        const prev = this.undoStack.pop() as string;
        this.redoStack.push(cur);
        this.doc = normalizeDoc(JSON.parse(prev));
        if (this.slideIndex >= this.doc.slides.length) this.slideIndex = this.doc.slides.length - 1;
        this.sel = this.sel.filter((id) => !!this.findElement(id));
        this.editingId = null;
        this.persist();
        this.emit('doc', { undo: true });
        this.emit('sel');
        return true;
    }
    redo() {
        if (!this.redoStack.length) return false;
        const cur = JSON.stringify(this.doc);
        const next = this.redoStack.pop() as string;
        this.undoStack.push(cur);
        this.doc = normalizeDoc(JSON.parse(next));
        if (this.slideIndex >= this.doc.slides.length) this.slideIndex = this.doc.slides.length - 1;
        this.sel = this.sel.filter((id) => !!this.findElement(id));
        this.persist();
        this.emit('doc', { redo: true });
        this.emit('sel');
        return true;
    }
    canUndo() { return this.undoStack.length > 0; }
    canRedo() { return this.redoStack.length > 0; }

    setDoc(doc: any, opts: any = {}) {
        if (!opts.noHistory && this.doc) {
            this.undoStack.push(JSON.stringify(this.doc));
            this.redoStack.length = 0;
        }
        this.doc = normalizeDoc(doc);
        this.slideIndex = Math.min(this.slideIndex, this.doc.slides.length - 1);
        if (this.slideIndex < 0) this.slideIndex = 0;
        this.sel = [];
        this.editingId = null;
        this.persist();
        this.emit('doc', { full: true });
        this.emit('slide');
        this.emit('sel');
    }

    /* ---------- 视图 ---------- */
    setZoom(z: number) {
        this.zoom = Math.max(0.15, Math.min(4, z));
        this.emit('zoom');
    }
    setView(patch: any) {
        Object.assign(this, patch);
        this.saveUI();
        this.emit('view');
    }

    /* ---------- 持久化 ---------- */
    saveUI() {
        try {
            this.storage.setItem(LS_UI, JSON.stringify({ showGrid: this.showGrid, snap: this.snap, showNotes: this.showNotes }));
        } catch { /* ignore */ }
    }
    /** 文档变更后触发（UI 层订阅以更新“已修改”提示）；文档本身不再写入存储 */
    persist = debounce(() => {
        if (this._loading) return;
        this.emit('persist');
    }, 800);
}

/** 创建编辑器状态仓库实例 */
export function createStore(opts: { storage?: StorageAdapter } = {}) {
    return new EditorStore(opts);
}
