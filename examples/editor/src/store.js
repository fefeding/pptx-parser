/**
 * 全局状态仓库：文档 + 选中 + 撤销栈 + 本地持久化
 */
import { clone, debounce } from './util.js';
import { createStarterDoc } from './model.js';

const LS_KEY = 'pptx-editor:doc:v1';
const LS_UI = 'pptx-editor:ui:v1';
const MAX_HISTORY = 120;

class Store {
  constructor() {
    this.doc = null;
    this.slideIndex = 0;
    this.sel = [];               // 选中的元素 id（当前页顶层）
    this.editingId = null;       // 正在内联编辑的文本元素
    this.clipboard = [];
    this.zoom = 1;
    this.showGrid = false;
    this.snap = true;
    this.showNotes = false;
    this.fitted = false;
    this.undoStack = [];
    this.redoStack = [];
    this._lastCoalesce = { key: null, time: 0 };
    this._listeners = new Map();
    this._loading = true;
  }

  /* ---------- 事件 ---------- */
  on(evt, fn) {
    if (!this._listeners.has(evt)) this._listeners.set(evt, new Set());
    this._listeners.get(evt).add(fn);
    return () => this._listeners.get(evt).delete(fn);
  }
  emit(evt, payload) {
    const set = this._listeners.get(evt);
    if (set) for (const fn of Array.from(set)) fn(payload);
  }

  /* ---------- 初始化 ---------- */
  init() {
    let doc = null;
    try {
      const raw = localStorage.getItem(LS_KEY);
      if (raw) doc = normalizeDoc(JSON.parse(raw));
    } catch { doc = null; }
    this.doc = doc && doc.slides && doc.slides.length ? doc : createStarterDoc();
    try {
      const ui = JSON.parse(localStorage.getItem(LS_UI) || '{}');
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
  findElement(id) {
    for (const el of this.elements()) if (el.id === id) return el;
    return null;
  }
  selected() {
    const els = this.elements();
    return this.sel.map((id) => els.find((e) => e.id === id)).filter(Boolean);
  }
  /** 递归查找（含组内子元素） */
  deepFind(id, list = this.elements()) {
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
  setSel(ids, opts = {}) {
    const next = Array.isArray(ids) ? ids : [ids];
    const same = next.length === this.sel.length && next.every((v, i) => v === this.sel[i]);
    if (same && !opts.force) return;
    this.sel = next;
    if (this.editingId && !next.includes(this.editingId)) this.editingId = null;
    this.emit('sel');
  }
  clearSel() { this.setSel([]); }
  toggleSel(id) {
    this.setSel(this.sel.includes(id) ? this.sel.filter((v) => v !== id) : this.sel.concat(id));
  }

  setSlide(i, opts = {}) {
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
   * @param {(draft:any)=>void} mutator 直接修改草稿
   * @param {{history?:boolean, coalesce?:string, silent?:boolean}} opts
   */
  update(mutator, opts = {}) {
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
    const prev = this.undoStack.pop();
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
    const next = this.redoStack.pop();
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

  setDoc(doc, opts = {}) {
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
  setZoom(z) {
    this.zoom = Math.max(0.15, Math.min(4, z));
    this.emit('zoom');
  }
  setView(patch) {
    Object.assign(this, patch);
    this.saveUI();
    this.emit('view');
  }

  /* ---------- 持久化 ---------- */
  saveUI() {
    try {
      localStorage.setItem(LS_UI, JSON.stringify({ showGrid: this.showGrid, snap: this.snap, showNotes: this.showNotes }));
    } catch { /* ignore */ }
  }
  persist = debounce(() => {
    if (this._loading) return;
    try {
      localStorage.setItem(LS_KEY, JSON.stringify(this.doc));
      const st = document.getElementById('saveState');
      if (st) { st.textContent = '已自动保存 ' + new Date().toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' }); }
    } catch { /* ignore quota */ }
  }, 800);
}

/** 补齐缺失字段，防止旧数据/导入数据缺字段导致崩溃 */
function normalizeDoc(doc) {
  if (!doc || !Array.isArray(doc.slides)) throw new Error('文档数据无效');
  doc.title = doc.title || '未命名演示文稿';
  doc.theme = doc.theme || 'blue';
  doc.slideSize = doc.slideSize && doc.slideSize.width ? doc.slideSize : { width: 1280, height: 720 };
  for (const s of doc.slides) {
    s.id = s.id || ('s_' + Math.random().toString(36).slice(2, 8));
    s.elements = Array.isArray(s.elements) ? s.elements : [];
    s.notes = s.notes || '';
    for (const el of s.elements) normalizeElement(el);
  }
  return doc;
}
function normalizeElement(el) {
  if (!el.id) el.id = 'e_' + Math.random().toString(36).slice(2, 9);
  el.x = Number(el.x) || 0; el.y = Number(el.y) || 0;
  el.width = Number(el.width) || 40; el.height = Number(el.height) || 40;
  el.rotation = Number(el.rotation) || 0;
  if (el.type === 'text' && !el.paragraphs) el.paragraphs = [{ runs: [{ text: el.text || '' }] }];
  if (el.type === 'group') (el.children || []).forEach(normalizeElement);
}

export const store = new Store();
