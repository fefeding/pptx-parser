// PPTX editor webview main logic.
//
// Preview mode: renders the original file bytes with pptxToHtml (identical to the parser's
// preview page) — always the most faithful view, also used for presenting.
// Edit mode: parses the bytes into the parser's editor document (pptxToStandard -> docFromPptx)
// and renders straight from that model (no PPTX byte round-trip, so nothing is lost in
// serialization). All edits go through the parser's editor core (createStore / createActions),
// which owns the model, selection and undo/redo history.
//
// The extension host only stores PPTX bytes: after each edit the webview serializes the
// document and syncs the result, so the host can save without understanding the model.
import { CSS } from './style';
import { h, clamp } from './util';
import { t, data, setLocale, kbd } from './i18n';
import { openMenu, openMenuAt, popupKeydown, MenuItem } from './menu';
import { createRibbon, RibbonState, readStore, writeStore } from './ribbon';
import { icon } from './icons';
import {
  pptxToHtml,
  pptxToStandard,
  jsonToPptx,
  docFromPptx,
  docToPptx,
  createStore,
  createActions,
  updateDiagramTexts,
  FONT_LIST,
  THEMES,
  LAYOUTS,
  SLIDE_SIZES,
  TRANSITIONS,
  ANIM_CLASSES,
  ANIM_TYPES,
  ANIM_DIRECTIONS,
  CHART_TYPES
} from '@fefeding/ppt-parser';

// 几何 API：基础矩形类（elementRect/effectMargin）可由具名导出稳定获得，但线段/连接线相关函数
// （isLineElement/lineEndpoints/geomPoints/connectionPoints/findGlueTarget/resyncGlue/setLineEndpoints）
// 无法以具名重导出暴露（rollup 会修剪），统一从 createActions.geometryApi 聚合对象取用。
// 该静态属性在 pptx-parser 模块加载时即被赋值，故此处可安全解构。
const GEOM_API: any = (createActions as any).geometryApi || {};
const {
  elementRect,
  effectMargin,
  isLineElement,
  lineEndpoints,
  geomPoints,
  connectionPoints,
  findGlueTarget,
  resyncGlue,
  setLineEndpoints
} = GEOM_API;
import {
  renderSlideInto,
  renderThumbInto,
  disposeDetachedCharts,
  disposeAllCharts
} from './render';
import { exportSlidesToPng, base64Of } from './exporter';
import type { ExportedSlide } from './exporter';
import {
  hasOpenModal,
  closeModal,
  openTableSizeDialog,
  openTableDialog,
  openChartDialog,
  openShapePicker,
  openMediaDialog,
  openShortcuts,
  openConfirm,
  openDiagramTextDialog,
  shapeMenuItems
} from './dialogs';
// Same ECharts renderer the preview side uses (option building incl. 3D matches examples/index.html)
import { chartRenderer } from '@fefeding/ppt-parser/chart-renderer';
import type { EditorMode, HostToWebview, WebviewToHost } from '../protocol';

declare function acquireVsCodeApi(): {
  postMessage(msg: WebviewToHost): void;
  getState(): any;
  setState(state: any): void;
};
const vscode = acquireVsCodeApi();

const PARSE_OPTS = { mediaProcess: true, themeProcess: true } as const;

// OOXML preset pattern fills (subset of the common names the renderer understands).
const PATTERN_LIST: [string, string][] = [
  ['pct5', '5%'], ['pct10', '10%'], ['pct20', '20%'], ['pct25', '25%'], ['pct30', '30%'],
  ['pct40', '40%'], ['pct50', '50%'], ['pct60', '60%'], ['pct70', '70%'], ['pct75', '75%'],
  ['pct80', '80%'], ['pct90', '90%'],
  ['diagCross', 'Diagonal Cross'], ['divot', 'Divot'], ['dotGrid', 'Dot Grid'],
  ['horizontal', 'Horizontal'], ['vertical', 'Vertical'], ['wave', 'Wave'],
  ['zigzag', 'ZigZag'], ['plaid', 'Plaid'], ['shingle', 'Shingle'],
  ['trellis', 'Trellis'], ['solidDmnd', 'Solid Diamond']
];

// ---------- Document state ----------
// One editor store per webview (the parser core keeps the document, selection and history).
const store: any = createStore();
const actions: any = createActions(store);
// Re-render whenever the group-edit context changes (enter/exit group editing).
store.on('groupEdit', () => renderAll());

const state: {
  mode: EditorMode;
  title: string;
  /** Original file bytes — preview renders these. */
  original: Uint8Array | null;
  /** Latest bytes: original until the webview reports an edit. */
  latest: Uint8Array | null;
  /** Preview render result (pptxToHtml output). */
  preview: { slides: string[]; charts: any[]; metadata?: any; customProps?: Record<string, string> } | null;
  present: { slides: string[]; charts: any[]; version: number; global?: string } | null;
  /** True when the host document has unsaved changes (drives the title-bar dirty dot). */
  dirty: boolean;
  /** Bumped whenever the serialized document changes, to invalidate cached renders. */
  version: number;
  slideSize: { width: number; height: number };
  current: number;
  zoom: number;
  userZoom: number | null;
  grid: boolean;
  editingText: boolean;
  syncing: boolean;
} = {
  mode: 'preview',
  title: 'presentation.pptx',
  original: null,
  latest: null,
  preview: null,
  present: null,
  dirty: false,
  version: 0,
  slideSize: { width: 1280, height: 720 },
  current: 0,
  zoom: 1,
  userZoom: null,
  grid: false,
  editingText: false,
  syncing: false
};

// ---------- DOM references ----------
let root!: HTMLElement;
let ribbonHost!: HTMLElement;
let ribbon!: { root: HTMLElement; sync(s: RibbonState): void };
let slideListEl!: HTMLElement;
let stageEl!: HTMLElement;
let stageInnerEl!: HTMLElement;
let slideHostEl!: HTMLElement;
let overlayEl!: HTMLElement;
let gridEl!: HTMLElement;
let inspectorEl!: HTMLElement;
/** Element-list drawer state (persisted): collapsed by default so properties stay in view. */
const EL_LIST_KEY = 'pptx.inspector.elList';
let elListOpen = readStore(EL_LIST_KEY) === '1';
/** Slide panel view (persisted): thumbnails by default, outline keeps long decks scannable. */
const SLIDE_VIEW_KEY = 'pptx.panel.slideView';
let slideView: 'thumb' | 'outline' = readStore(SLIDE_VIEW_KEY) === 'outline' ? 'outline' : 'thumb';
let slideTabThumbs!: HTMLElement;
let slideTabOutline!: HTMLElement;
let elFilter = '';
let elFilterFocused = false;
let elFilterNode: HTMLInputElement | null = null;
/** Selection signature of the last inspector render, used to keep the scroll position across rebuilds. */
let lastInspSelKey = '';
let statusCountEl!: HTMLElement;
let zoomLabelEl!: HTMLElement;
let gridBtn!: HTMLElement;
let snapBtn!: HTMLElement;
let presentEl!: HTMLElement;
let presentCountEl!: HTMLElement;
let groupHintEl!: HTMLElement;
let presentHostEl!: HTMLElement;
/** Element entrance animations on the current slide that advance on click (sequenced). */
let presentClickAnims: { node: HTMLElement; kf: any; opts: KeyframeAnimationOptions; cls: string }[] = [];
let presentClickIdx = 0;
let docInfoEl!: HTMLElement;
let toastEl!: HTMLElement;

// ---------- Utilities ----------
function num(v: any, def = 0): number {
  const n = typeof v === 'number' ? v : parseFloat(v);
  return Number.isFinite(n) ? n : def;
}
function post(msg: WebviewToHost) {
  vscode.postMessage(msg);
}
let toastTimer: any;
function toast(message: string, kind: 'info' | 'warn' | 'error' = 'info') {
  toastEl.textContent = message;
  toastEl.className = `toast show ${kind === 'info' ? '' : kind}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (toastEl.className = 'toast'), 2600);
}
function debounce(fn: (...a: any[]) => void, ms: number) {
  let t: any;
  return (...a: any[]) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...a), ms);
  };
}
function toBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
function toBase64(bytes: Uint8Array): string {
  let bin = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)) as any);
  }
  return btoa(bin);
}

// ---------- Layout ----------
/**
 * Nodes whose text is baked in at build time get an updater registered here, so a late locale
 * change (the host's `init.locale` disagreeing with `<html lang>`) can refresh them in place
 * instead of rebuilding the whole layout and duplicating its listeners.
 */
const localeUpdaters: (() => void)[] = [];
function localize(fn: () => void): void {
  localeUpdaters.push(fn);
  fn();
}

/** Rebuild the ribbon (every label lives inside it) and refresh the static captions. */
function applyLocale(): void {
  for (const fn of localeUpdaters) fn();
  const next = createRibbon({ cmd: ribbonCommands });
  ribbonHost.replaceChild(next.root, ribbon.root);
  ribbon = next;
  ribbon.sync(ribbonState());
}

function buildLayout() {
  root = document.getElementById('app')!;
  root.className = 'app mode-preview';

  // Ribbon (title bar + tabs + groups). Rebuilt wholesale when the language changes.
  ribbonHost = h('div', { class: 'ribbon-host' });
  ribbon = createRibbon({ cmd: ribbonCommands });
  ribbonHost.append(ribbon.root);

  const newSlide = () => runAction(() => actions.addSlide());

  // Slide list — the panel is virtualized, so rows are drawn only when they scroll near the view.
  slideListEl = h('div', { class: 'sp-list' });
  slideListEl.addEventListener('scroll', () => scheduleThumbFlush(), { passive: true });
  slideListEl.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    // Outline rows get the same slide menu as thumbnails.
    const thumb = (e.target as HTMLElement).closest('.thumb, .ol-row') as HTMLElement | null;
    if (!thumb) return;
    const idx = Array.prototype.indexOf.call(slideListEl.children, thumb);
    if (idx < 0) return;
    showSlideCtxMenu(e.clientX, e.clientY, idx);
  });
  const spTitle = h('span', { class: 'sp-title' }, [t('panel.slides')]);
  // `edit-only`: the slide list is read-only in preview mode, so adding a slide must be hidden there.
  const spNewBtn = h('button', { class: 'btn sp-new edit-only', type: 'button', onclick: newSlide }, [icon('newSlide', 14)]);
  // View tabs: thumbnails vs a text outline. Both stay available in preview mode (read-only).
  slideTabThumbs = h('button', {
    class: 'sp-tab' + (slideView === 'thumb' ? ' on' : ''),
    type: 'button',
    title: t('panel.tabThumbs'),
    onclick: () => setSlideView('thumb')
  }, [t('panel.tabThumbs')]);
  slideTabOutline = h('button', {
    class: 'sp-tab' + (slideView === 'outline' ? ' on' : ''),
    type: 'button',
    title: t('panel.tabOutline'),
    onclick: () => setSlideView('outline')
  }, [t('panel.tabOutline')]);
  const spTabs = h('div', { class: 'sp-tabs' }, [slideTabThumbs, slideTabOutline]);
  localize(() => {
    spTitle.textContent = t('panel.slides');
    spNewBtn.title = t('panel.newSlide');
    slideTabThumbs.textContent = t('panel.tabThumbs');
    slideTabThumbs.title = t('panel.tabThumbs');
    slideTabOutline.textContent = t('panel.tabOutline');
    slideTabOutline.title = t('panel.tabOutline');
  });
  const spHead = h('div', { class: 'sp-head' }, [spTitle, spTabs, spNewBtn]);
  const slidesPanel = h('div', { class: 'slides-panel' }, [spHead, slideListEl]);

  // Canvas: frame + overlay share one scaled coordinate space
  slideHostEl = h('div', { class: 'slide-host slide-frame' });
  overlayEl = h('div', { class: 'overlay' });
  gridEl = h('div', { class: 'grid-overlay' });
  stageInnerEl = h('div', { class: 'stage-inner' }, [slideHostEl, gridEl, overlayEl]);
  stageEl = h('div', { class: 'stage' }, [stageInnerEl]);
  const scroll = h('div', { class: 'canvas-scroll', id: 'canvasScroll' }, [stageEl]);
  // Ctrl/Cmd + wheel zooms the canvas (plain wheel still scrolls)
  scroll.addEventListener(
    'wheel',
    (e) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      e.preventDefault();
      setZoom(state.userZoom ?? state.zoom, true, e.deltaY < 0 ? 1.1 : 0.9);
    },
    { passive: false }
  );
  // Canvas interaction: hit-test to select, drag handles live on the overlay boxes.
  slideHostEl.addEventListener('pointerdown', (e) => {
    if (state.mode !== 'edit') return;
    const node = (e.target as HTMLElement).closest('.el[data-id]') as HTMLElement | null;
    const id = node?.dataset.id || null;
    // Inside a group-edit session only elements belonging to the active group are selectable;
    // clicking anywhere else leaves group editing (click again to act on the outer element).
    if (store.groupEdit) {
      const gid = store.groupEdit;
      const belongs = id && (id === gid || findParentGroupId(id) === gid);
      if (!belongs) {
        store.setGroupEdit(null);
        return;
      }
    }
    if (node) {
      if (store.editingId === id) return; // let contenteditable place the caret itself
      // NOTE: do NOT call e.preventDefault() here — on Chromium it suppresses the synthetic
      // mouse/dblclick events, which would break double-click-to-edit. Dragging still prevents
      // text selection via CSS (user-select:none) and startMove() guards selection on move.
      if (e.shiftKey) store.toggleSel(id);
      else if (!store.sel.includes(id)) store.setSel([id]);
      renderSelection();
      renderInspector();
      if (!e.shiftKey) startMove(e);
      return;
    }
    if (store.editingId) commitEditing();
    store.setSel([]);
    renderSelection();
    renderInspector();
    startMarquee(e);
  });
  slideHostEl.addEventListener('dblclick', onCanvasDblClick);
  const canvasArea = h('div', { class: 'canvas-area' }, [scroll]);
  groupHintEl = h('div', { class: 'group-hint', onclick: () => store.setGroupEdit(null) }, [t('group.hint')]);
  localize(() => {
    groupHintEl.title = t('group.exit');
    groupHintEl.textContent = t('group.hint');
  });
  groupHintEl.style.display = 'none';
  canvasArea.append(groupHintEl);

  // Properties panel (edit mode only)
  inspectorEl = h('div', { class: 'inspector edit-only' });
  const main = h('div', { class: 'main' }, [slidesPanel, canvasArea, inspectorEl]);

  // Status bar (PowerPoint keeps the slide counter and the zoom cluster down here)
  statusCountEl = h('span', {});
  zoomLabelEl = h('span', { class: 'chip', onclick: () => setZoom(1) }, ['100%']);
  const zoomOut = h('span', { class: 'chip', onclick: () => setZoom(state.userZoom ?? state.zoom, true, 0.9) }, ['−']);
  const zoomIn = h('span', { class: 'chip', onclick: () => setZoom(state.userZoom ?? state.zoom, true, 1.1) }, ['+']);
  const zoomFit = h('span', { class: 'chip', onclick: () => setZoom(null) }, [t('status.fit')]);
  gridBtn = h('span', { class: 'chip edit-only', onclick: toggleGrid }, [icon('grid', 13)]);
  snapBtn = h('span', { class: 'chip edit-only', onclick: toggleSnap }, [icon('snap', 13)]);
  localize(() => {
    zoomLabelEl.title = t('status.zoomReset');
    zoomOut.title = t('cmd.zoomOut');
    zoomIn.title = t('cmd.zoomIn');
    zoomFit.title = t('cmd.fit');
    zoomFit.textContent = t('status.fit');
    gridBtn.title = t('cmd.grid');
    snapBtn.title = t('cmd.snap');
  });
  if (store.snap !== false) snapBtn.classList.add('on');
  const statusbar = h('div', { class: 'statusbar' }, [
    statusCountEl,
    h('span', { class: 'spacer' }),
    gridBtn,
    snapBtn,
    zoomOut,
    zoomFit,
    zoomLabelEl,
    zoomIn
  ]);

  // Presentation layer: full-screen slide + a discreet overlay (counter / hint / exit)
  presentHostEl = h('div', { class: 'slide-host' });
  presentCountEl = h('span', { class: 'present-count' });
  const presentHint = h('span', { class: 'present-hint' }, [t('present.hint')]);
  const presentExitBtn = h('button', {
    class: 'btn present-exit',
    type: 'button',
    onclick: (e: MouseEvent) => {
      e.stopPropagation();
      stopPresent();
    }
  }, [icon('close', 12), h('span', {}, [t('present.exit')])]);
  localize(() => {
    presentHint.textContent = t('present.hint');
    presentExitBtn.title = t('present.exit');
    const span = presentExitBtn.querySelector('span:last-child');
    if (span) span.textContent = t('present.exit');
  });
  const presentUi = h('div', { class: 'present-ui', onclick: (e: MouseEvent) => e.stopPropagation() }, [
    presentCountEl,
    presentHint,
    presentExitBtn
  ]);
  presentEl = h('div', { class: 'present', onclick: presentClick }, [presentHostEl, presentUi]);

  docInfoEl = h('div', { class: 'doc-info' });
  toastEl = h('div', { class: 'toast' });

  // Right-click context menu (on canvas)
  canvasArea.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    showCtxMenu(e.clientX, e.clientY);
  });

  root.append(ribbonHost, main, statusbar, docInfoEl, presentEl, toastEl);

  document.addEventListener('keydown', onKeyDown);

  // Re-fit on viewport changes so the slide never overflows a small window (the default zoom is
  // "fit to window", matching the parser's preview & editor). Presentation mode re-computes its
  // own scale; otherwise only re-fit when the user hasn't picked a manual zoom.
  window.addEventListener('resize', onWindowResize);
}

/** Window resize handler: re-fit the canvas, or recompute the presentation scale if presenting. */
function onWindowResize(): void {
  if (presentEl.classList.contains('on')) {
    renderPresent(1, false);
  } else if (state.userZoom == null) {
    setZoom(null);
  }
}

// ---------- Ribbon bridge ----------
/**
 * Capability flags the ribbon buttons declare through `needs`. Keeping the decision here (and
 * not inside the ribbon) means the rule "is this action available?" lives next to the model.
 */
function ribbonCaps(): Record<string, boolean | string> {
  const editing = state.mode === 'edit';
  const sel = editing ? store.selected() : [];
  const one = sel.length === 1 ? sel[0] : null;
  const texty = !!one && (one.type === 'text' || (one.paragraphs || []).length > 0);
  // 渐变填充没有单一 color（渲染端读 stops），色板取首个停点色才能反映当前外观
  const fill = one && one.fill ? (typeof one.fill === 'object' && (one.fill.color || one.fill.stops?.[0]?.color) ? (one.fill.color || one.fill.stops[0].color) : typeof one.fill === 'string' && one.fill !== 'none' ? one.fill : '') : '';
  const line = one && one.line && one.line !== 'none' ? one.line.color || '#000000' : '#000000';
  const bg = editing ? store.slide?.background : '';
  return {
    edit: editing,
    preview: !editing,
    canUndo: editing && store.canUndo(),
    canRedo: editing && store.canRedo(),
    selection: sel.length > 0,
    single: sel.length === 1,
    multi: sel.length > 1,
    text: texty,
    shape: sel.length > 0 && sel.every((el: any) => el.type === 'shape' || el.type === 'group'),
    ungroupable: sel.some((el: any) => el.type === 'group'),
    locked: sel.length > 0 && sel.every((el: any) => el.locked),
    elementHidden: sel.length > 0 && sel.every((el: any) => el.hidden),
    manySlides: editing && store.slideCount > 1,
    notFirstSlide: state.current > 0,
    notLastSlide: editing && state.current < store.slideCount - 1,
    fitZoom: state.userZoom == null,
    presenting: presentEl.classList.contains('on'),
    // values shown by selects / colour swatches
    fontFace: texty ? one.fontFace || '' : '',
    fontSize: texty ? String(one.fontSize || '') : '',
    textColor: texty ? one.color || '' : '',
    bold: texty ? !!one.bold : false,
    italic: texty ? !!one.italic : false,
    underline: texty ? !!one.underline : false,
    bullet: texty ? !!one.bullet : false,
    align: texty ? one.align || 'left' : '',
    valign: texty ? one.valign || 'top' : '',
    lineSpacing: texty ? String(one.lineSpacing || '') : '',
    fillColor: fill || '',
    lineColor: line,
    slideBg: typeof bg === 'string' ? bg : ''
  };
}

/** trigger 在编辑器模型里是字符串，从 PPTX 解析回来可能是 { type } 对象 */
function normTrigger(t: any): 'onClick' | 'withPrev' | 'afterPrev' {
  const v = typeof t === 'string' ? t : t?.type;
  return v === 'onClick' || v === 'withPrev' ? v : 'afterPrev';
}

/** 当前正在编辑的动画（slide.animations 下标）；-1 表示跟随最后一条 */
let animEditIndex = -1;

/** 参数控件应编辑哪条动画：显式选中优先，否则取选中元素的最后一条（通常就是刚添加的） */
function currentAnimIndex(): number {
  // preview 模式下 store.doc 为 null，而 store.slide 的 getter 会直接读 doc.slides
  if (state.mode !== 'edit' || !store.doc) return -1;
  const anims: any[] = Array.isArray(store.slide?.animations) ? store.slide.animations : [];
  if (animEditIndex >= 0 && animEditIndex < anims.length) return animEditIndex;
  const sel: string[] = store.sel || [];
  if (!sel.length) return anims.length - 1;
  for (let i = anims.length - 1; i >= 0; i--) {
    if (sel.includes(anims[i].target)) return i;
  }
  // 选中的元素没有动画时不要退回别元素的最后一条，否则参数区会误显示/误改它
  return -1;
}

/** 把类型/方向/触发/时长立即写回当前动画，而不是只改「下一次添加」的草稿 */
function patchCurrentAnim(patch: any): void {
  if (state.mode !== 'edit') return;
  const anims: any[] = Array.isArray(store.slide?.animations) ? store.slide.animations : [];
  const idx = currentAnimIndex();
  if (idx < 0 || idx >= anims.length) return;
  animEditIndex = idx;
  const next: any = { ...patch };
  // 类型/类别/方向变了，旧的 OOXML presetID/presetSubtype 必须失效：导出端会优先用它们，
  // 与 preset 不一致时 PowerPoint 会忽略该动画或仍播旧效果（一直「从左侧飞入」）。
  if ('type' in next || 'presetClass' in next) next.presetId = undefined;
  if ('direction' in next) next.presetSubtype = undefined;
  runAction(() => actions.updateAnimation(idx, next));
}

function animLabel(a: any): string {
  const trig = normTrigger(a.trigger);
  return [data('animClass', a.presetClass || 'entr', a.presetClass), data('animType', a.type || '', a.type), data('trigger', trig, trig)].join(' / ');
}

/**
 * 预览模式没有编辑器模型，动画随 pptxToHtml 的 data-timing 下发。
 * 重新打开文件默认是预览模式，若动画窗格空白，用户会以为「保存后动画没了」。
 */
function previewAnimations(): any[] {
  const html = state.preview?.slides?.[state.current] || '';
  const m = html.match(/data-timing='([^']*)'/);
  if (!m) return [];
  try {
    const payload = JSON.parse(m[1]);
    if (!Array.isArray(payload.animations)) return [];
    return payload.animations.map((a: any) => ({
      type: a.type,
      presetClass: a.presetClass || 'entr',
      trigger: a.trigger || 'onClick',
      direction: a.direction,
      duration: a.duration
    }));
  } catch { return []; }
}

function ribbonState(): RibbonState {
  // In preview mode the editor document isn't loaded (`store.doc === null`), so every store
  // lookup has to go through the mode-aware `slideCount()` instead of `store.slideCount`.
  const editing = state.mode === 'edit';
  const trans: any = (editing && store.slide?.transition) || {};
  const anims: any[] = editing
    ? (Array.isArray(store.slide?.animations) ? store.slide.animations : [])
    : previewAnimations();
  const themeId = editing ? (typeof store.doc?.theme === 'string' ? store.doc.theme : 'blue') : '';
  return {
    mode: state.mode,
    dirty: state.dirty,
    canUndo: editing && store.canUndo(),
    canRedo: editing && store.canRedo(),
    zoom: state.zoom,
    grid: state.grid,
    snap: store.snap !== false,
    presenting: presentEl.classList.contains('on'),
    slideCount: slideCount(),
    slideIndex: state.current,
    slideHidden: editing ? !!store.slide?.hidden : false,
    theme: themeId,
    sizeKey: editing ? store.doc?.sizeKey || '16:9' : '',
    transition: { type: trans.type || 'none', duration: trans.duration ?? 800, advanceOnClick: trans.advanceOnClick !== false },
    animations: anims.map((a, index) => ({ index, text: animLabel(a) })),
    animIndex: editing ? currentAnimIndex() : -1,
    animPick: (() => {
      if (!editing) return { cls: 'entr', type: 'fadeIn', trigger: 'afterPrev', direction: 'l', duration: '500' };
      const cur: any = anims[currentAnimIndex()] || null;
      return {
        cls: cur?.presetClass || 'entr',
        type: cur?.type || 'fadeIn',
        trigger: cur ? normTrigger(cur.trigger) : 'afterPrev',
        direction: cur?.direction || 'l',
        duration: String(cur?.duration ?? 500)
      };
    })(),
    caps: ribbonCaps()
  };
}

/** Commands the ribbon can invoke (see `ribbon-tabs.ts` for the call sites). */
function ribbonCommands(name: string, ...args: any[]): any {
  const edit = () => state.mode === 'edit';
  const sel = () => store.selected();
  const patch = (p: any) => {
    const ids = store.selected().map((el: any) => el.id);
    if (!ids.length) return;
    runAction(() => {
      for (const id of ids) actions.updateElement(id, p);
    });
  };
  switch (name) {
    // --- file / app ---
    // 走宿主的标准保存流程（含保存前的 flush 握手），不要在这里直接写盘
    case 'save': return post({ type: 'save' });
    case 'undo': return runAction(() => store.undo());
    case 'redo': return runAction(() => store.redo());
    case 'setMode': return setMode(args[0]);
    case 'shortcuts': return openShortcuts();
    case 'docProps': return toggleDocInfo();
    case 'present': return startPresent(false);
    case 'presentHere': return startPresent(true);
    case 'endShow': return stopPresent();
    case 'toggleInspector': return inspectorEl.classList.toggle('hidden');
    case 'zoom':
      if (args[0] === 'in') return setZoom(state.userZoom ?? state.zoom, true, 1.1);
      if (args[0] === 'out') return setZoom(state.userZoom ?? state.zoom, true, 0.9);
      if (args[0] === 'actual') return setZoom(1);
      return setZoom(null);

    // --- clipboard / element ---
    case 'paste': return runAction(() => actions.paste());
    case 'copy': actions.copySelected(false); toast(t('toast.copied')); renderAll(); return;
    case 'cut': actions.copySelected(true); toast(t('toast.cut')); renderAll(); return;
    case 'duplicate': return runAction(() => actions.duplicateSelected());
    case 'delete': return runAction(() => actions.deleteSelected());
    case 'selectAll': actions.selectAll(); renderAll(); return;
    case 'zOrder': return runAction(() => actions.zOrder(args[0]));
    case 'align': return runAction(() => actions.alignElements(args[0]));
    case 'distribute': return runAction(() => actions.distribute(args[0]));
    case 'groupSel': return runAction(() => actions.groupSelection());
    case 'ungroupSel': return runAction(() => actions.ungroupSelection());
    case 'toggleLock': return runAction(() => actions.toggleLock());
    case 'toggleHidden': return runAction(() => actions.toggleHidden());
    case 'flip': {
      const axis = args[0] === 'h' ? 'flipH' : 'flipV';
      const on = !sel()[0]?.[axis];
      return patch({ [axis]: on });
    }

    // --- text ---
    case 'toggleBold': return runAction(() => actions.applyTextStyleSel({ bold: !sel()[0]?.bold }));
    case 'toggleItalic': return runAction(() => actions.applyTextStyleSel({ italic: !sel()[0]?.italic }));
    case 'toggleUnderline': return runAction(() => actions.applyTextStyleSel({ underline: !sel()[0]?.underline }));
    case 'toggleBullet': return runAction(() => actions.applyTextStyleSel({ bullet: !sel()[0]?.bullet }));
    case 'setAlign': return applyTextStyle({ align: args[0] });
    case 'setValign': return patch({ valign: args[0] });
    case 'setLineSpacing': return applyTextStyle({ lineSpacing: num(args[0], 1.15) });
    case 'setFontFace': return applyTextStyle({ fontFace: args[0] });
    case 'setFontSize': return applyTextStyle({ fontSize: num(args[0], 24) });
    case 'setTextColor': return applyTextStyle({ color: args[0] === 'none' ? undefined : args[0] });

    // --- shape ---
    case 'setFill': {
      const v = args[0];
      for (const el of sel()) {
        const cur = el.fill && typeof el.fill === 'object' ? el.fill : null;
        runAction(() =>
          actions.updateElement(el.id, {
            fill:
              v === 'none'
                ? 'none'
                : cur && cur.type === 'gradient'
                  ? { ...cur, stops: [{ color: v, position: 0 }, cur.stops?.[1] || { color: '#ffffff', position: 1 }] }
                  : v
          })
        );
      }
      return;
    }
    case 'setLineColor': {
      const v = args[0];
      for (const el of sel()) {
        const w = el.line && el.line !== 'none' ? el.line.width || 1 : 1;
        runAction(() => actions.updateElement(el.id, { line: v === 'none' ? 'none' : { color: v, width: w } }));
      }
      return;
    }
    case 'shapes': return openShapePicker({ addShape: (t2: string, opt?: any) => addShapeElement(t2, opt) });
    case 'shapesMenu': return shapeMenuItems({ addShape: (t2: string, opt?: any) => addShapeElement(t2, opt) });

    // --- insert ---
    case 'textBox': return addTextElement();
    case 'picture': return pickImage();
    case 'table': return insertTable();
    case 'chart': return insertChart();
    case 'media': return openMediaDialog({ addMedia });
    case 'exportImage': return exportCurrentImage();
    case 'exportImages': return exportAllImages();
    case 'exportPdf': return exportPdf();
    // Format conversion runs in the extension host (it needs system fonts / Node APIs).
    // The host already holds the latest PPTX bytes, so unsaved edits are included.
    case 'convertFormat': post({ type: 'convert' }); return;

    // --- slides ---
    case 'newSlide': return runAction(() => actions.addSlide(args[0]));
    case 'duplicateSlide': return runAction(() => actions.duplicateSlide(state.current));
    case 'deleteSlide': return deleteCurrentSlide();
    case 'toggleSlideHidden': return runAction(() => actions.toggleSlideHidden(state.current));
    case 'slideUp': return runAction(() => actions.moveSlide(state.current, state.current - 1));
    case 'slideDown': return runAction(() => actions.moveSlide(state.current, state.current + 1));
    case 'applyLayout': return applyLayoutWithConfirm(args[0]);
    case 'openLayoutMenu': {
      const anchor =
        document.activeElement && ribbonHost.contains(document.activeElement as Node)
          ? (document.activeElement as HTMLElement)
          : (ribbonHost.querySelector('.rb-panel') as HTMLElement);
      return openMenuAt(anchor, layoutMenuItems((id) => applyLayoutWithConfirm(id)), { minWidth: 180 });
    }

    // --- design ---
    case 'setTheme': return runAction(() => actions.applyTheme(args[0], true));
    case 'setSizeKey': {
      const size = SLIDE_SIZES[args[0]];
      if (size) changeSlideSize(size);
      return;
    }
    case 'setSlideBg': return runAction(() => actions.setBackground(args[0] === 'none' ? '#ffffff' : args[0]));
    case 'bgImage': return pickBackgroundImage();
    case 'bgGradient': return runAction(() => actions.setBackground({ type: 'gradient', angle: 45, stops: [{ color: '#1a73e8', position: 0 }, { color: '#ffffff', position: 1 }] }));
    case 'bgNone': return runAction(() => actions.setBackground('none'));

    // --- transitions ---
    case 'setTransition': return setTransition({ type: args[0] });
    case 'setTransSpeed': return setTransition({ duration: num(args[0], 800) });
    case 'setTransAdvance': return setTransition({ advanceOnClick: !!args[0] });
    case 'applyTransitionAll': return applyTransitionToAll();

    // --- animations ---
    case 'addAnim': return addAnimationFromRibbon(args[0]);
    case 'animSelect':
      animEditIndex = Number(args[0]);
      renderInspector(true);
      ribbon.sync(ribbonState());
      return;
    case 'animPatch': return patchCurrentAnim(args[0]);
    case 'animMove': return runAction(() => actions.moveAnimation(args[0], args[1]));
    case 'animRemove': return runAction(() => actions.removeAnimation(args[0]));

    // --- view helpers ---
    case 'grid': return toggleGrid();
    case 'snap': return toggleSnap();
    default:
      if (!edit()) return;
      return undefined;
  }
}

// ---------- Helpers behind ribbon commands ----------
function layoutMenuItems(run: (id: string) => void): MenuItem[] {
  return [
    { type: 'label', label: t('cmd.layout') },
    ...LAYOUTS.map((l: any) => ({ label: data('layout', l.id, l.name), icon: 'layout', onClick: () => run(l.id) }))
  ];
}

/** Applying a layout replaces every element of the slide, so confirm first. */
async function applyLayoutWithConfirm(id?: string) {
  if (state.mode !== 'edit') return;
  const layoutId = id || (store.slide && store.slide.layout) || 'titleBody';
  const ok = await openConfirm({ message: t('confirm.applyLayout'), okLabel: t('common.ok') });
  if (!ok) return;
  runAction(() => actions.applyLayout(layoutId));
}

async function deleteCurrentSlide() {
  if (store.slideCount <= 1) {
    toast(t('toast.needSelection'), 'warn');
    return;
  }
  const ok = await openConfirm({
    message: t('confirm.deleteSlide'),
    detail: t('confirm.deleteSlideBody', state.current + 1),
    okLabel: t('common.delete'),
    danger: true
  });
  if (!ok) return;
  runAction(() => actions.deleteSlide(state.current));
}

/** Patch the current slide's transition, keeping the fields the caller doesn't mention. */
function setTransition(patch: any) {
  const cur: any = store.slide?.transition || {};
  runAction(() =>
    actions.setTransition({
      type: cur.type || 'none',
      duration: cur.duration ?? 800,
      advanceOnClick: cur.advanceOnClick !== false,
      ...patch
    })
  );
}

function applyTransitionToAll() {
  const cur: any = store.slide?.transition || {};
  const next = { type: cur.type || 'none', duration: cur.duration ?? 800, advanceOnClick: cur.advanceOnClick !== false };
  if (next.type === 'none') {
    toast(t('toast.needSelection'), 'warn');
    return;
  }
  runAction(() => {
    store.update((doc: any) => {
      for (const slide of doc.slides || []) slide.transition = { ...next };
    });
  });
}

function addAnimationFromRibbon(pick: any) {
  const el = store.sel.length === 1 ? store.findElement(store.sel[0]) : null;
  if (!el) {
    toast(t('toast.needSelection'), 'warn');
    return;
  }
  // 新添加的动画成为参数控件的编辑目标
  animEditIndex = -1;
  runAction(() =>
    actions.addAnimation({
      target: el.id,
      type: pick.type,
      presetClass: pick.cls,
      direction: pick.direction,
      duration: pick.duration,
      trigger: pick.trigger
    })
  );
}

// ---------- Export (PNG / PDF, rasterized offscreen from the editor model) ----------

/** Editor document for export: the live model when loaded, otherwise a throwaway parse. */
async function ensureExportDoc(): Promise<any> {
  if (store.doc?.slides?.length) return store.doc;
  const bytes = state.latest || state.original;
  if (!bytes) return null;
  const standard = await pptxToStandard(bytes, PARSE_OPTS);
  return docFromPptx(standard, { fileName: state.title });
}

function baseName(title: string): string {
  return String(title || 'slides').replace(/\.pptx$/i, '');
}

async function exportSlides(indices: number[], pixelRatio = 2): Promise<ExportedSlide[] | null> {
  const doc = await ensureExportDoc();
  if (!doc || !doc.slides?.length) return null;
  const safe = indices.filter((i) => i >= 0 && i < doc.slides.length);
  if (!safe.length) return null;
  toast(t('toast.exporting'));
  try {
    return await exportSlidesToPng({ doc, indices: safe, pixelRatio });
  } catch (e: any) {
    toast(t('toast.exportFailed', e?.message || String(e)), 'error');
    return null;
  }
}

async function exportCurrentImage(): Promise<void> {
  const out = await exportSlides([state.current]);
  if (!out?.length) return;
  post({
    type: 'exportImage',
    bytes: base64Of(out[0].png),
    suggestedName: `${baseName(state.title)}-${state.current + 1}.png`
  });
}

async function exportAllImages(): Promise<void> {
  const doc = await ensureExportDoc();
  if (!doc) return;
  const out = await exportSlides(doc.slides.map((_: any, i: number) => i));
  if (!out?.length) return;
  post({
    type: 'exportImages',
    images: out.map((s) => ({
      name: `${baseName(state.title)}-${String(s.index + 1).padStart(2, '0')}.png`,
      bytes: base64Of(s.png)
    }))
  });
}

async function exportPdf(): Promise<void> {
  const doc = await ensureExportDoc();
  if (!doc) return;
  const out = await exportSlides(doc.slides.map((_: any, i: number) => i));
  if (!out?.length) return;
  post({
    type: 'exportPdf',
    images: out.map((s) => ({ bytes: base64Of(s.png) })),
    widthPx: doc.slideSize.width,
    heightPx: doc.slideSize.height,
    suggestedName: `${baseName(state.title)}.pdf`
  });
}

// ---------- Preview pipeline (pptxToHtml: most faithful rendering) ----------
/**
 * Inject the parser's global stylesheet. pptxToHtml emits classed markup (_css_1 /
 * _tbl_cell_css_3, …) whose layout comes entirely from `styles.global`; without it the slides
 * render as unstyled HTML. Kept in its own <style> so the editor's own CSS stays separate.
 */
let globalStyleEl: HTMLStyleElement | null = null;
function ensureGlobalStyles(css: string | undefined): void {
  if (!css) return;
  if (!globalStyleEl) {
    globalStyleEl = document.createElement('style');
    globalStyleEl.id = 'pptx-global-styles';
    document.head.appendChild(globalStyleEl);
  }
  if (globalStyleEl.textContent !== css) globalStyleEl.textContent = css;
}

async function renderPreview(): Promise<void> {
  const bytes = state.latest;
  if (!bytes || !bytes.length) return;
  const result = await pptxToHtml(bytes, PARSE_OPTS);
  ensureGlobalStyles(result.styles?.global);
  state.preview = {
    slides: result.slides.map((s: any) => s.html),
    charts: result.charts || [],
    metadata: result.metadata,
    customProps: result.customProps
  };
  state.slideSize = result.slideSize || state.slideSize;
  if (state.present && state.present.version === state.version) state.present = null;
  if (state.mode === 'preview') {
    state.current = clamp(state.current, 0, state.preview.slides.length - 1);
    renderAll();
    fitIfAuto();
  }
}

// ---------- Edit pipeline (editor document -> DOM) ----------
async function enterEditMode(): Promise<void> {
  const bytes = state.latest;
  if (!bytes || !bytes.length) return;
  // 动画由 parser 精确保真（p:spTgt@spid ↔ p:cNvPr@id 映射），重新解析即可还原，
  // 不再需要按元素「位置」猜测并在找不到时退回第一个元素的备份逻辑——那正是动画
  // 错绑到未设置元素上的原因。
  const standard = await pptxToStandard(bytes, PARSE_OPTS);
  const doc = docFromPptx(standard, { fileName: state.title });
  store.setDoc(doc, { noHistory: true });
  store.setSlide(state.current);
  state.slideSize = doc.slideSize || state.slideSize;
  state.current = clamp(state.current, 0, store.slideCount - 1);
  renderAll();
  fitIfAuto();
}

/** Serialize the edited document back to PPTX bytes. */
async function serializeDoc(): Promise<Uint8Array> {
  const out = await jsonToPptx(docToPptx(store.doc), { outputType: 'uint8array' });
  return out as Uint8Array;
}

/**
 * Push the current document to the host so it can save. Called after every edit (debounced) and
 * before anything that needs the up-to-date bytes (presenting, leaving edit mode).
 */
let syncQueued = false;
async function syncBytes(dirty: boolean): Promise<void> {
  if (state.syncing) {
    // Queue one follow-up pass instead of dropping it: edits made while a serialization is in
    // flight would otherwise never reach the host, and would silently be missing from the save.
    syncQueued = true;
    return;
  }
  state.syncing = true;
  try {
    if (!dirty) {
      // Undo landed back on the imported state: fall back to the original bytes
      state.latest = state.original;
      state.dirty = false;
      post({ type: 'sync', bytes: null, dirty: false });
      syncChrome();
      return;
    }
    state.latest = await serializeDoc();
    state.dirty = true;
    post({ type: 'sync', bytes: toBase64(state.latest), dirty: true });
    // syncBytes 是防抖后异步执行的，此时 renderAll 早已跑完，必须单独刷新功能区上的保存按钮
    syncChrome();
  } catch (e: any) {
    post({ type: 'error', message: t('toast.serializeFailed', e?.message || String(e)) });
  } finally {
    state.syncing = false;
    if (syncQueued) {
      syncQueued = false;
      void syncBytes(dirty);
    }
  }
}
const syncSoon = debounce(() => syncBytes(store.canUndo()), 350);

/** Run a store action and refresh everything that depends on the model. */
function runAction(fn: () => void, opts: { skipSync?: boolean } = {}) {
  if (state.mode !== 'edit') return;
  try {
    fn();
  } catch (e: any) {
    toast(e?.message || String(e), 'error');
    return;
  }
  state.version++;
  renderAll();
  if (!opts.skipSync) syncSoon();
}

/** 可逐 run 覆盖的字符级样式键（与 parser 的 applyTextStyle 保持一致） */
const RUN_STYLE_KEYS = ['fontSize', 'color', 'bold', 'italic', 'underline', 'strike', 'highlight', 'fontFace'];
/**
 * 文本样式改写（字色 / 字号 / 字体 / 对齐 / 行距…）
 *
 * 渲染端按「run 值优先、元素值兜底」绘制：st.color = run.color || el.color || 主题色。
 * 而导入 PPTX 时解析端把每个 run 的显式 rPr（含从 lstStyle / 母版 defRPr 继承来的值）
 * 全部写进了 run，段落级 align 也一样被填成元素值。只改元素级属性会被 run 上的原值
 * 盖掉——表现就是面板/ribbon 里改了颜色、字号、字体、对齐，画布纹丝不动。
 *
 * 这里把改动同步写进每个 run（元素级也写，保证新建元素与导出一致），并按 parser 约定
 * 打 propsSet / alignSet / lineSpacingSet 标记，导出时才不会被当成面板默认值丢弃。
 * 值为 undefined 表示「清除」，让渲染回落到主题色。
 */
function applyTextStyle(patch: Record<string, unknown>): void {
  runAction(() => {
    store.update((doc: any) => {
      for (const id of store.sel.slice()) {
        const el = findInDoc(doc, id);
        if (!el || (el.type !== 'text' && !(el.paragraphs || []).length)) continue;
        for (const [k, v] of Object.entries(patch)) {
          if (v === undefined) {
            delete el[k];
            if (el.propsSet) delete el.propsSet[k];
          } else {
            el[k] = v;
            if (el.propsSet && k in el.propsSet) el.propsSet[k] = true;
          }
        }
        for (const p of el.paragraphs || []) {
          for (const r of p.runs || []) {
            for (const k of RUN_STYLE_KEYS) {
              if (!(k in patch)) continue;
              if (patch[k] === undefined) delete r[k];
              else r[k] = patch[k];
            }
          }
          if ('align' in patch) {
            if (patch.align === undefined) {
              delete p.align;
              delete p.alignSet;
            } else {
              p.align = patch.align;
              p.alignSet = true;
            }
          }
          if ('lineSpacing' in patch) {
            if (patch.lineSpacing === undefined) {
              delete p.lineSpacing;
              delete p.lineSpacingSet;
            } else {
              p.lineSpacing = patch.lineSpacing;
              p.lineSpacingSet = true;
            }
          }
          if ('bullet' in patch) p.bullet = patch.bullet === false ? undefined : patch.bullet;
        }
      }
    });
  });
}

/**
 * 段落样式改写（左/右缩进、悬挂缩进、段前段后、列表层级）
 *
 * 这些属性在 OOXML 里是段落级的（a:pPr@marL/marR/indent、a:spcBef/a:spcAft、a:pPr@lvl），
 * 编辑模型把它们放在每个段落对象上，故按段逐个写入；值为 undefined 表示清除（回到继承值）。
 * 单位：px（编辑模型内部单位，导出时由 parser 转成 pt）。
 */
function applyParagraphStyle(patch: Record<string, unknown>): void {
  runAction(() => {
    store.update((doc: any) => {
      for (const id of store.sel.slice()) {
        const el = findInDoc(doc, id);
        if (!el || el.type !== 'text') continue;
        for (const p of el.paragraphs || []) {
          for (const [k, v] of Object.entries(patch)) {
            if (v === undefined) delete p[k];
            else p[k] = v;
          }
        }
      }
    });
  });
}

// ---------- Mode switching ----------
async function toggleMode() {
  await setMode(state.mode === 'preview' ? 'edit' : 'preview');
}

/** Switch between preview and edit mode (no-op when already there). */
async function setMode(next: EditorMode) {
  if (next === state.mode) return;
  state.mode = next;
  post({ type: 'setMode', mode: next });
  await applyMode();
}

async function applyMode() {
  root.className = 'app mode-' + state.mode;
  const editing = state.mode === 'edit';
  overlayEl.innerHTML = '';

  if (editing) {
    await enterEditMode();
    return;
  }
  // Leaving edit mode: flush the edited bytes first, then render them faithfully
  if (store.canUndo()) await syncBytes(true);
  disposeAllCharts();
  try {
    await renderPreview();
  } catch (e: any) {
    toast(t('toast.renderFailed', e?.message || String(e)), 'error');
  }
}

// ---------- Zoom ----------
/**
 * Re-fit to the window only when the user hasn't set an explicit zoom. This keeps a manual zoom
 * sticky across byte updates / mode switches / window resizes — the default zoom is "fit to
 * window" (like the parser's preview & editor), not a fixed 100%.
 */
function fitIfAuto(): void {
  if (state.userZoom == null) setZoom(null);
}

/** Fit the slide into the canvas area on both axes (capped at 2x). */
function getFitZoom(): number {
  const sc = document.getElementById('canvasScroll');
  const availW = (sc?.clientWidth || 800) - 40;
  const availH = (sc?.clientHeight || 600) - 40;
  return clamp(Math.min(availW / state.slideSize.width, availH / state.slideSize.height), 0.1, 2);
}
function setZoom(value: number | null, isUser = true, factor?: number) {
  if (value === null) {
    state.userZoom = null;
    state.zoom = getFitZoom();
  } else {
    let z = value;
    if (factor) z = (state.userZoom ?? state.zoom) * factor;
    z = clamp(z, 0.1, 4);
    state.userZoom = z;
    state.zoom = z;
  }
  applyStageTransform();
  // 把手尺寸按 zoom 归一（屏幕像素恒定），缩放后需重建 overlay 才能生效。
  if (state.mode === 'edit') buildOverlay();
  zoomLabelEl.textContent = Math.round(state.zoom * 100) + '%';
}
function applyStageTransform() {
  stageInnerEl.style.width = state.slideSize.width + 'px';
  stageInnerEl.style.height = state.slideSize.height + 'px';
  stageInnerEl.style.transform = `scale(${state.zoom})`;
  // Outer placeholder = scaled size, to avoid clipping when zoomed in and blank scroll area when zoomed out
  stageEl.style.width = Math.round(state.slideSize.width * state.zoom) + 'px';
  stageEl.style.height = Math.round(state.slideSize.height * state.zoom) + 'px';
}

// ---------- Charts ----------
/**
 * Paint the ECharts placeholders emitted by pptxToHtml (preview + presentation). In edit mode
 * charts are rendered from the model by render.ts instead.
 */
function paintCharts(host: HTMLElement, charts?: any[]) {
  const list0 = charts ?? state.preview?.charts;
  if (!list0 || !list0.length) return;
  if (typeof (window as any).echarts === 'undefined') return; // skip silently if echarts not loaded
  try {
    // Container scope, so repeated copies of the same slide HTML (thumbnail / canvas / presentation)
    // never resolve to each other's chart containers.
    const list = list0.filter((c: any) => !!host.querySelector(`[id="${c.chartId}"]`));
    if (list.length) chartRenderer.renderCharts(list, host);
  } catch (e: any) {
    console.warn('[pptx-webview] Chart rendering failed:', e);
  }
}

/** Charts of the currently presented slide (rendered from their own render cache). */
function paintPresentCharts() {
  if (!state.present) return;
  paintCharts(presentHostEl, state.present.charts);
}

/**
 * Two ECharts instance registries exist: chartRenderer's singleton (used for pptxToHtml output)
 * and the renderer module's own map (used for model-rendered charts). Both hold instances whose
 * DOM has been replaced, so drop them to avoid leaks.
 */
function disposeStaleCharts() {
  disposeDetachedCharts();
  const insts = (chartRenderer as any).chartInstances as Map<string, any> | undefined;
  if (!insts) return;
  for (const [id, inst] of Array.from(insts.entries())) {
    const dom = typeof inst?.getDom === 'function' ? inst.getDom() : null;
    if (!dom || !document.contains(dom)) {
      try {
        inst.dispose();
      } catch {
        /* noop */
      }
      insts.delete(id);
    }
  }
}

// ---------- Rendering ----------
function renderAll() {
  applyStageTransform();
  gridEl.className = 'grid-overlay' + (state.grid ? ' on' : '');
  renderSlideList();
  renderCanvas();
  renderInspector();
  const count = slideCount();
  statusCountEl.textContent = t('status.counter', state.current + 1, count);
  statusCountEl.title = t('status.slides', count);
  gridBtn.classList.toggle('on', state.grid);
  snapBtn.classList.toggle('on', store.snap !== false);
  if (docInfoEl.classList.contains('on')) renderDocInfo();
  ribbon.sync(ribbonState());
}

function slideCount(): number {
  return state.mode === 'edit' ? store.slideCount : state.preview?.slides.length || 0;
}

/**
 * Slide panel. Thumbnails can be dragged to reorder (edit mode), show their number and a
 * hidden-slide marker, and double-click jumps straight into edit mode from preview.
 *
 * Virtualized: every slide keeps a row placeholder (so scroll height, ordering and the
 * `slideListEl.children[i]` mapping that drag-reorder / context menus rely on stay correct),
 * but only rows near the viewport are drawn; the rest fill in on scroll, and a row is redrawn
 * only when its slide object actually changed. Rebuilding all thumbnails on every edit
 * measured ~318ms per navigation on a 21-slide deck.
 */
const THUMB_OVERSCAN = 2;
const thumbRows: HTMLElement[] = [];
const thumbRev: any[] = [];
let thumbObserver: IntersectionObserver | null = null;
let thumbEpoch = '';
/** Rows currently near the viewport, maintained by the observer. */
const visibleThumbs = new Set<number>();

/** Uniform row height: every slide shares the deck's aspect ratio, so width fixes the height. */
function thumbRowHeight(): number {
  const W = state.slideSize.width || 1280;
  const H = state.slideSize.height || 720;
  const cs = getComputedStyle(slideListEl);
  const pad = (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.paddingRight) || 0);
  let boxW = (slideListEl.clientWidth || 168) - pad;
  if (boxW <= 0) boxW = 168; // panel not laid out yet
  return Math.max(24, Math.round(H * (boxW / W)));
}

/** Vertical gap between rows, read from CSS so it never drifts from the stylesheet. */
function thumbGap(): number {
  const cs = getComputedStyle(slideListEl);
  const g = parseFloat(cs.rowGap || cs.gap || '');
  return Number.isFinite(g) ? g : 6;
}

/** Cheap deck-structure key: a change here invalidates every cached row. */
function deckKey(): string {
  if (state.mode === 'edit') return 'e:' + store.doc.slides.map((s: any) => s.id || '').join(',');
  return 'p:' + state.version + ':' + (state.preview?.slides.length || 0);
}

/** Drop all cached rows (document reload, mode switch, or structural edit). */
function resetThumbRows(): void {
  slideListEl.innerHTML = '';
  thumbRows.length = 0;
  thumbRev.length = 0;
  visibleThumbs.clear();
}

function ensureThumbObserver(): void {
  if (thumbObserver) return;
  thumbObserver = new IntersectionObserver((entries) => {
    for (const e of entries) {
      const i = thumbRows.indexOf(e.target as HTMLElement);
      if (i < 0) continue;
      if (e.isIntersecting) {
        visibleThumbs.add(i);
        ensureThumbContent(i, state.mode === 'edit', false);
      } else visibleThumbs.delete(i);
    }
  }, { root: slideListEl, rootMargin: '300px 0px' });
}

/** One placeholder row per slide, in slide order, with the drag wiring attached once. */
function syncThumbRows(count: number, editing: boolean): void {
  while (thumbRows.length > count) {
    const row = thumbRows.pop();
    thumbRev.pop();
    if (row) { if (thumbObserver) thumbObserver.unobserve(row); row.remove(); }
  }
  while (thumbRows.length < count) {
    const i = thumbRows.length;
    const row = h('div', { class: 'thumb', draggable: 'false' });
    row.addEventListener('click', () => {
      if (dragMoved) return; // the gesture was a reorder, not a selection
      selectSlide(i);
    });
    row.addEventListener('dblclick', () => {
      if (state.mode !== 'edit' && store.doc.slides.length) toggleMode();
    });
    if (editing) enableThumbDrag(row, i);
    slideListEl.append(row);
    thumbRows.push(row);
    thumbRev.push(undefined);
    if (thumbObserver) thumbObserver.observe(row);
  }
}

/** Cheap per-row sync: selection ring + hidden marker. Runs for every row, draws nothing. */
function updateThumbRow(i: number, editing: boolean): void {
  const row = thumbRows[i];
  if (!row) return;
  const hidden = editing && !!store.doc.slides[i]?.hidden;
  row.classList.toggle('active', i === state.current);
  row.classList.toggle('hidden-slide', !!hidden);
  row.title = t('panel.slideThumb', i + 1) + (hidden ? ' · ' + t('panel.hiddenSlide') : '');
}

/** Redraw a row only when the slide it shows has changed. */
function ensureThumbContent(i: number, editing: boolean, force: boolean): void {
  const row = thumbRows[i];
  if (!row) return;
  const slide = editing ? store.doc.slides[i] : null;
  // `store.update` clones the whole document, so identity changes exactly when content does.
  const key: any = editing ? slide : `p:${state.version}:${i}`;
  updateThumbRow(i, editing);
  if (!force && thumbRev[i] === key) return;
  thumbRev[i] = key;
  row.innerHTML = '';
  if (editing) {
    if (!slide) return; // index out of range (transient during structural edits)
    renderThumbInto(row, slide, store.doc);
  } else {
    const htmls = state.preview?.slides || [];
    const W = state.slideSize.width, H = state.slideSize.height;
    const scale = (row.clientWidth || 160) / W;
    const inner = h('div', { class: 'inner' });
    inner.style.width = `${W}px`;
    inner.style.height = `${H}px`;
    inner.style.transform = `scale(${scale})`;
    inner.innerHTML = htmls[i] || '';
    row.style.height = `${Math.round(H * scale)}px`;
    row.appendChild(inner);
  }
  row.append(h('span', { class: 'num' }, [String(i + 1)]));
  if (editing && !!slide?.hidden) {
    row.append(h('span', { class: 'hide-badge', title: t('panel.hiddenSlide') }, ['🚫']));
  }
}

function flushThumbWindow(editing: boolean, force = false): void {
  const count = thumbRows.length;
  if (!count) return;
  // Created lazily: the observer's root is the panel, which only gets laid out once mounted.
  ensureThumbObserver();
  const key = deckKey();
  if (key !== thumbEpoch) { thumbEpoch = key; thumbRev.length = 0; }
  const rowH = thumbRowHeight();
  for (let i = 0; i < count; i++) {
    thumbRows[i].style.height = `${rowH}px`;
    // Selection / hidden state must stay correct even for rows scrolled out of view.
    updateThumbRow(i, editing);
  }

  // The observer drives lazy drawing; on the first flush (or while the panel has no layout yet)
  // there is no observer data, so fall back to a window around the current slide.
  const anchors = new Set<number>(visibleThumbs);
  if (!anchors.size) {
    const gap = thumbGap();
    const step = rowH + gap;
    const top = slideListEl.scrollTop;
    const view = slideListEl.clientHeight || rowH * 6;
    const first = Math.max(0, Math.floor((top + gap) / step) - THUMB_OVERSCAN);
    const last = Math.min(count - 1, Math.ceil((top + view) / step) + THUMB_OVERSCAN);
    for (let i = first; i <= last; i++) anchors.add(i);
  }
  for (const i of anchors) ensureThumbContent(i, editing, force);
  // The active slide can be scrolled out of view — keep its thumbnail current anyway.
  ensureThumbContent(state.current, editing, force);
}

/** Coalesce scroll bursts into a single flush per frame. */
let thumbFlushRaf = 0;
function scheduleThumbFlush(): void {
  if (thumbFlushRaf) return;
  thumbFlushRaf = requestAnimationFrame(() => {
    thumbFlushRaf = 0;
    if (slideView === 'thumb') flushThumbWindow(state.mode === 'edit', false);
  });
}

function renderSlideList() {
  const count = slideCount();
  const editing = state.mode === 'edit';
  // Outline mode replaces the thumbnails with one text row per slide; rows keep the same
  // click / right-click behaviour, only the presentation differs.
  slideListEl.classList.toggle('ol-mode', slideView === 'outline');
  if (slideView === 'outline') {
    resetThumbRows();
    renderOutlineList(count, editing);
    return;
  }
  // Leaving outline: drop the text rows before rebuilding the thumbnails,
  // otherwise the leftover `.ol-row` elements stay visible alongside them.
  if (slideListEl.querySelector('.ol-row')) resetThumbRows();
  syncThumbRows(count, editing);
  flushThumbWindow(editing);
}

/** Switch the slide panel between thumbnails and the text outline (the choice is remembered). */
function setSlideView(view: 'thumb' | 'outline'): void {
  if (slideView === view) return;
  slideView = view;
  writeStore(SLIDE_VIEW_KEY, view === 'outline' ? 'outline' : 'thumb');
  slideTabThumbs.classList.toggle('on', view === 'thumb');
  slideTabOutline.classList.toggle('on', view === 'outline');
  renderSlideList();
}

/** Flatten an element's paragraphs into text (soft breaks become spaces). */
function textOfElement(el: any): string {
  const parts: string[] = [];
  for (const p of el.paragraphs || []) {
    parts.push((p.runs || []).map((r: any) => (r.break ? ' ' : r.text || '')).join(''));
  }
  return parts.join('\n').replace(/[ \t]+/g, ' ').trim();
}

/**
 * Slide title heuristic: the largest-font text block wins. The editor model carries no
 * placeholder metadata, so this mirrors what the parser does elsewhere when it derives a
 * title from a slide (`pptxToBlocks` picks the max-font-size text box).
 */
function slideTitleFromDoc(slide: any): string {
  let best = '';
  let bestSize = -1;
  if (!slide) return best;
  for (const el of slide.elements || []) {
    if (el.hidden) continue;
    const txt = textOfElement(el);
    if (!txt) continue;
    let size = Number(el.fontSize) || 0;
    for (const p of el.paragraphs || []) {
      for (const r of p.runs || []) size = Math.max(size, Number(r.fontSize) || 0);
    }
    if (size > bestSize) {
      bestSize = size;
      best = txt;
    }
  }
  return best ? best.split('\n')[0].slice(0, 80) : '';
}

/** Preview mode has no editor model, so its titles come from the rendered HTML (cached per render). */
let previewTitleCache: { src: any; titles: string[] } | null = null;
function previewTitles(): string[] {
  const src = state.preview;
  if (!src) return [];
  if (previewTitleCache && previewTitleCache.src === src) return previewTitleCache.titles;
  const titles = (src.slides || []).map(titleFromHtml);
  previewTitleCache = { src, titles };
  return titles;
}

/** Largest inline `font-size` text in a slide's HTML, falling back to its first text run. */
function titleFromHtml(html: string): string {
  const cands: { size: number; text: string }[] = [];
  let fallback = '';
  try {
    const doc = new DOMParser().parseFromString(`<div>${html}</div>`, 'text/html');
    const host = doc.body.firstElementChild as Element | null;
    if (host) {
      const walk = (node: Element) => {
        for (const child of Array.from(node.children)) {
          const el = child as HTMLElement;
          const txt = (el.textContent || '').trim();
          const size = parseFloat(el.style?.fontSize || '') || 0;
          if (txt) {
            if (!fallback) fallback = txt;
            cands.push({ size, text: txt });
          }
          walk(el);
        }
      };
      walk(host);
    }
  } catch {
    return '';
  }
  if (!cands.length) return '';
  let max = 0;
  for (const c of cands) max = Math.max(max, c.size);
  let best = fallback;
  if (max > 0) {
    // A title line is often split into several runs of slightly differing size, so picking the
    // biggest single run truncates it ("What is" instead of "What is PPTXjs?"). Among the nodes
    // set at ~the top size, take the longest text — that is the whole line.
    const near = cands.filter((c) => c.size >= max * 0.85);
    best = near.reduce((a, b) => (b.text.length > a.text.length ? b : a)).text;
  }
  return best.replace(/\s+/g, ' ').slice(0, 80);
}

/**
 * Exact slide titles in preview mode come from the parser's document model rather than the
 * rendered HTML: `pptxToHtml` splits one title line into sibling runs ("What is" / "PPTXjs" / "?"),
 * so any DOM heuristic would read a fragment. The parse is lazy and cached per document bytes.
 */
let titleDocCache: { src: Uint8Array; titles: string[] } | null = null;
async function docTitles(): Promise<string[] | null> {
  const src = state.latest;
  if (!src) return null;
  if (titleDocCache && titleDocCache.src === src) return titleDocCache.titles;
  try {
    const standard = await pptxToStandard(src, PARSE_OPTS);
    const titles = docFromPptx(standard, { fileName: state.title }).slides.map((s: any) => slideTitleFromDoc(s));
    titleDocCache = { src, titles };
    return titles;
  } catch {
    return null;
  }
}

/**
 * Outline rows: slide number + title. Rows are far cheaper to scroll than rendered thumbnails,
 * which is what makes a deck with dozens of slides navigable in a narrow panel.
 */
function renderOutlineList(count: number, editing: boolean): void {
  if (editing) {
    paintOutlineRows(count, true, (store.doc.slides || []).map((s: any) => slideTitleFromDoc(s)));
    return;
  }
  // Paint straight away from the rendered HTML, then refine to the exact model titles.
  paintOutlineRows(count, false, previewTitles());
  void docTitles().then((list) => {
    // Drop a stale load: the view may have changed while the parse was running.
    if (slideView !== 'outline' || state.mode === 'edit' || !list) return;
    paintOutlineRows(slideCount(), false, list);
  });
}

function paintOutlineRows(count: number, editing: boolean, titles: string[]): void {
  slideListEl.innerHTML = '';
  for (let i = 0; i < count; i++) {
    const slide = editing ? store.doc.slides[i] : null;
    const hidden = editing && !!slide?.hidden;
    const raw = titles[i] || '';
    const row = h('div', {
      class: 'ol-row' + (i === state.current ? ' active' : '') + (hidden ? ' hidden-slide' : ''),
      title: raw || t('panel.slideThumb', i + 1)
    });
    row.append(h('span', { class: 'ol-num' }, [String(i + 1)]));
    row.append(h('span', { class: 'ol-text' }, [raw || t('panel.untitled')]));
    if (editing && slide?.notes) row.append(h('span', { class: 'ol-notes', title: t('panel.hasNotes') }, ['N']));
    if (hidden) row.append(h('span', { class: 'hide-badge', title: t('panel.hiddenSlide') }, ['🚫']));
    row.addEventListener('click', () => selectSlide(i));
    slideListEl.append(row);
  }
}

/** Drag a thumbnail to another position to reorder the slides. */
let dragMoved = false;
function enableThumbDrag(thumb: HTMLElement, index: number): void {
  thumb.addEventListener('pointerdown', (e) => {
    if (e.button !== 0 || store.slideCount < 2) return;
    const startY = e.clientY;
    dragMoved = false;
    let target = index;
    const clearMarks = () =>
      slideListEl.querySelectorAll('.thumb').forEach((n) => n.classList.remove('drop-before', 'drop-after', 'dragging'));
    const onMove = (ev: PointerEvent) => {
      if (!dragMoved && Math.abs(ev.clientY - startY) < 5) return;
      dragMoved = true;
      // Pick the insertion slot whose midpoint the pointer is closest to.
      let best = index;
      let bestDist = Infinity;
      let after = false;
      slideListEl.querySelectorAll<HTMLElement>('.thumb').forEach((n, i) => {
        const r = n.getBoundingClientRect();
        const mid = r.top + r.height / 2;
        const d = Math.abs(ev.clientY - mid);
        if (d < bestDist) {
          bestDist = d;
          best = i;
          after = ev.clientY > mid;
        }
      });
      target = best;
      clearMarks();
      thumb.classList.add('dragging');
      slideListEl.children[best]?.classList.add(after ? 'drop-after' : 'drop-before');
    };
    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      clearMarks();
      if (!dragMoved) return;
      // `target` is a slot in the current list; convert it to a moveSlide destination.
      // Removing the dragged slide first shifts everything after it up by one.
      const to = target > index ? target - 1 : target;
      if (to !== index && to >= 0 && to < store.slideCount - 1) runAction(() => actions.moveSlide(index, to));
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  });
}

function renderCanvas() {
  disposeStaleCharts();
  if (state.mode === 'edit') {
    const slide = store.slide;
    if (!slide) return;
    renderSlideInto(slideHostEl, slide, store.doc, { grid: state.grid, editingId: store.editingId });
    buildOverlay();
    applyGroupEditVisual();
  } else {
    slideHostEl.innerHTML = state.preview?.slides[state.current] || '';
    paintCharts(slideHostEl);
    overlayEl.innerHTML = '';
  }
}

function findParentGroupId(id: string): string | null {
  for (const slide of store.doc.slides || []) {
    for (const el of slide.elements || []) {
      if (el.children && el.children.some((c: any) => c.id === id)) return el.id;
    }
  }
  return null;
}

function applyGroupEditVisual() {
  const gid = store.groupEdit;
  slideHostEl.classList.toggle('group-editing', !!gid);
  const groupNode = gid ? (slideHostEl.querySelector(`.el[data-id="${gid}"]`) as HTMLElement | null) : null;
  slideHostEl.querySelectorAll('.el').forEach((n) => {
    const node = n as HTMLElement;
    const inside = groupNode ? groupNode.contains(node) : false;
    node.classList.toggle('dimmed', !!gid && !inside);
    node.classList.toggle('group-edit-active', !!gid && node === groupNode);
  });
  if (groupHintEl) groupHintEl.style.display = gid ? 'flex' : 'none';
}

// ---------- Selection overlay (geometry from the parser core) ----------
function buildOverlay() {
  overlayEl.innerHTML = '';
  const slide = store.slide;
  if (!slide) return;
  for (const el of slide.elements || []) {
    if (el.hidden) continue;
    if (!store.sel.includes(el.id)) continue; // only draw selection chrome for selected elements
    const r = elementRect(el);
    const m = effectMargin(el);
    const rect = h('div', {
      class: 'sel-rect selected' + (el.locked ? ' locked' : ''),
      'data-id': el.id,
      style:
        `left:${r.x - m.left}px;top:${r.y - m.top}px;` +
        `width:${r.width + m.left + m.right}px;height:${r.height + m.top + m.bottom}px;` +
        `transform:${el.rotation ? `rotate(${el.rotation}deg)` : ''}`
    });
    rect.addEventListener('pointerdown', (e) => onRectPointerDown(e, el));
    // 线条：端点/顶点把手，不用 resize/rotate handle
    const isLine = isLineElement(el);
    if (!el.locked && !isLine) {
      // 命中区（--hit）除以 zoom，低缩放下也容易抓住；视觉尺寸仍由 CSS 固定。
      const rk = 1 / (state.zoom || 1);
      for (const d of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) {
        rect.append(h('div', { class: 'handle ' + d, 'data-dir': d, style: `--hit:${18 * rk}px` }));
      }
      const rot = h('div', { class: 'handle rot', 'data-dir': 'rot', title: 'Rotate (hold Shift to snap 15°)', style: `--hit:${20 * rk}px` });
      rot.style.top = '-24px';
      rect.append(rot);
    }
    if (el.locked) rect.append(h('span', { class: 'lock-badge' }, ['🔒']));
    overlayEl.append(rect);
    // 线条把手在 sel-rect 之后追加：确保端点位于选择框之上，命中区不被选择框遮挡。
    if (!el.locked && isLine) drawLineHandles(el);
  }
}

/** Cheap redraw used while the selection changes (keeps the canvas DOM untouched). */
function renderSelection() {
  buildOverlay();
  // Reflect the selection on the rendered elements too (used to unlock media controls)
  slideHostEl.querySelectorAll('.el[data-id]').forEach((node) => {
    node.classList.toggle('is-sel', store.sel.includes((node as HTMLElement).dataset.id!));
  });
}

// ---------- Drag / Resize / Rotate ----------
function onRectPointerDown(e: PointerEvent, el: any) {
  e.preventDefault();
  e.stopPropagation();
  const dir = (e.target as HTMLElement).getAttribute('data-dir');
  const startX = e.clientX;
  const startY = e.clientY;
  const orig = { x: num(el.x), y: num(el.y), w: num(el.width), h: num(el.height), rot: num(el.rotation) };
  const center = { x: orig.x + orig.w / 2, y: orig.y + orig.h / 2 };
  const startAngle = Math.atan2(e.clientY - center.y, e.clientX - center.x);
  const box = (e.currentTarget as HTMLElement);
  box.setPointerCapture(e.pointerId);
  // One history entry per gesture: snapshot up-front, then apply without history
  store.snapshot();

  const apply = (patch: any) => {
    store.update((doc: any) => {
      const t = findInDoc(doc, el.id);
      if (!t) return;
      Object.assign(t, patch);
    }, { history: false });
    const target: any = store.findElement(el.id);
    if (!target) return;
    const r = elementRect(target);
    const m = effectMargin(target);
    box.style.left = r.x - m.left + 'px';
    box.style.top = r.y - m.top + 'px';
    box.style.width = r.width + m.left + m.right + 'px';
    box.style.height = r.height + m.top + m.bottom + 'px';
    box.style.transform = target.rotation ? `rotate(${target.rotation}deg)` : '';
    // Keep the rendered element itself in step with the frame, otherwise only the box moves
    const node = slideHostEl.querySelector(`.el[data-id="${el.id}"]`) as HTMLElement | null;
    if (node) {
      node.style.left = `${r.x}px`;
      node.style.top = `${r.y}px`;
      node.style.width = `${r.width}px`;
      node.style.height = `${r.height}px`;
      node.style.transform = target.rotation ? `rotate(${target.rotation}deg)` : '';
    }
  };

  const onMove = (ev: PointerEvent) => {
    const dx = (ev.clientX - startX) / state.zoom;
    const dy = (ev.clientY - startY) / state.zoom;
    if (dir === 'rot') {
      // Rotate around the element centre; hold Shift to snap to 15°
      const ang = Math.atan2(ev.clientY - center.y, ev.clientX - center.x);
      let next = orig.rot + ((ang - startAngle) * 180) / Math.PI;
      if (ev.shiftKey) next = Math.round(next / 15) * 15;
      next = ((next + 180) % 360) - 180;
      apply({ rotation: Math.round(next * 10) / 10 });
      return;
    }
    if (!dir) {
      apply({ x: Math.round(orig.x + dx), y: Math.round(orig.y + dy) });
      return;
    }
    // Resize: rotate the pointer delta into the element's local frame so that rotated
    // elements still resize along their own axes
    const rad = (-orig.rot * Math.PI) / 180;
    const ldx = dx * Math.cos(rad) - dy * Math.sin(rad);
    const ldy = dx * Math.sin(rad) + dy * Math.cos(rad);
    const sx = dir.includes('w') ? -1 : dir.includes('e') ? 1 : 0;
    const sy = dir.includes('n') ? -1 : dir.includes('s') ? 1 : 0;
    let nw = orig.w;
    let nh = orig.h;
    if (sx) nw = Math.max(8, orig.w + sx * ldx);
    if (sy) nh = Math.max(8, orig.h + sy * ldy);
    if (ev.shiftKey && sx && sy) {
      const s = Math.max(nw / orig.w, nh / orig.h);
      nw = orig.w * s;
      nh = orig.h * s;
    }
    // The grabbed handle stays under the cursor: shift the centre by half the size delta,
    // then rotate that offset back into page coordinates.
    const lox = (sx * (nw - orig.w)) / 2;
    const loy = (sy * (nh - orig.h)) / 2;
    const rad2 = (orig.rot * Math.PI) / 180;
    const cx = orig.x + orig.w / 2 + lox * Math.cos(rad2) - loy * Math.sin(rad2);
    const cy = orig.y + orig.h / 2 + lox * Math.sin(rad2) + loy * Math.cos(rad2);
    apply({
      width: Math.round(nw),
      height: Math.round(nh),
      x: Math.round(cx - nw / 2),
      y: Math.round(cy - nh / 2)
    });
  };

  const onUp = (ev: PointerEvent) => {
    box.releasePointerCapture(ev.pointerId);
    box.removeEventListener('pointermove', onMove);
    box.removeEventListener('pointerup', onUp);
    state.version++;
    renderAll();
    syncSoon();
  };
  box.addEventListener('pointermove', onMove);
  box.addEventListener('pointerup', onUp);
}

// ---------- Line: endpoints / vertices / glue ----------
/** 屏幕坐标 → 幻灯片坐标（overlay 与 slide 共享 state.zoom 缩放空间） */
function clientToSlide(cx: number, cy: number) {
  const rect = slideHostEl.getBoundingClientRect();
  return { x: (cx - rect.left) / state.zoom, y: (cy - rect.top) / state.zoom };
}

/** 为线条绘制端点把手；顶点编辑态下额外绘制顶点 / 贝塞尔控制点把手 */
function drawLineHandles(el: any) {
  const { a, b } = lineEndpoints(el);
  // overlay 位于被 scale(zoom) 的 stage 内：把手尺寸除以 zoom，使可见圆点与命中区
  // 在屏幕上保持恒定，低缩放下也容易点中。
  const k = 1 / (state.zoom || 1);
  const hs = 10 * k, hit = 24 * k;
  for (const [which, pt] of [['start', a], ['end', b]] as [string, any][]) {
    const handle = h('div', {
      class: 'handle endpoint',
      style: `left:${pt.x}px;top:${pt.y}px;width:${hs}px;height:${hs}px;margin-left:${-hs / 2}px;margin-top:${-hs / 2}px;--hit:${hit}px;`
    });
    handle.addEventListener('pointerdown', (e) => startLineEndpointDrag(e, which as 'start' | 'end', el.id));
    overlayEl.append(handle);
  }
  if (store.vertexEdit === el.id) {
    const pts = geomPoints(el);
    const vhs = 9 * k, vhit = 18 * k;
    for (const pt of pts) {
      const isSel = store.vertexSel && store.vertexSel.cmd === pt.cmdIndex && store.vertexSel.kind === pt.kind;
      const cls = (pt.kind === 'point' ? 'handle vertex' : 'handle vctrl') + (isSel ? ' sel' : '');
      const handle = h('div', {
        class: cls,
        style: `left:${pt.abs.x}px;top:${pt.abs.y}px;width:${vhs}px;height:${vhs}px;margin-left:${-vhs / 2}px;margin-top:${-vhs / 2}px;--hit:${vhit}px;`
      });
      handle.addEventListener('pointerdown', (e) => startLineVertexDrag(e, { id: el.id, cmd: pt.cmdIndex, kind: pt.kind }));
      overlayEl.append(handle);
    }
  }
}

/** 端点拖拽：移动线条一端，靠近形状连接点时吸附成 glue */
function startLineEndpointDrag(e: PointerEvent, which: 'start' | 'end', id: string) {
  e.preventDefault(); e.stopPropagation();
  const el = store.findElement(id);
  if (!el || el.locked) return;
  store.snapshot();
  const { a, b } = lineEndpoints(el);
  // 监听器挂在 window 上：renderAll() 每帧会重建 overlay 把手（销毁当前 box），
  // 若绑在被重建的 box 上，指针捕获会丢失、拖拽在第一帧后失效。
  const onMove = (ev: PointerEvent) => {
    const p = clientToSlide(ev.clientX, ev.clientY);
    const glue = findGlueTarget(p, id, store.slide.elements, 14 / state.zoom);
    store.update((doc: any) => {
      const t = findInDoc(doc, id); if (!t) return;
      const ep = glue ? glue.point : p;
      const na = which === 'start' ? ep : a;
      const nb = which === 'start' ? b : ep;
      setLineEndpoints(t, na, nb);
      if (glue) { if (which === 'start') t.begin = { shapeId: glue.shapeId, site: glue.site }; else t.end = { shapeId: glue.shapeId, site: glue.site }; }
      else { if (which === 'start') t.begin = null; else t.end = null; }
    }, { history: false });
    renderCanvas();
  };
  const onUp = () => {
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
    state.version++;
    renderAll();
    syncSoon();
  };
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
}

/** 顶点 / 贝塞尔控制点拖拽（顶点编辑态） */
function startLineVertexDrag(e: PointerEvent, ds: { id: string; cmd: number; kind: string }) {
  e.preventDefault(); e.stopPropagation();
  const el = store.findElement(ds.id);
  if (!el || el.locked) return;
  store.setVertexSel(ds.cmd, ds.kind);
  store.snapshot();
  // 同端点拖拽：监听器挂 window，避免渲染重建 overlay 时丢失指针捕获。
  const onMove = (ev: PointerEvent) => {
    const p = clientToSlide(ev.clientX, ev.clientY);
    store.update((doc: any) => {
      const t = findInDoc(doc, ds.id); if (!t || !t.custGeom) return;
      const c = t.custGeom.paths[0].commands[ds.cmd]; if (!c) return;
      const W = t.width || 100, H = t.height || 100, fx = t.flipH, fy = t.flipV;
      const lx = Math.max(0, Math.min(W, fx ? W - (p.x - t.x) : (p.x - t.x)));
      const ly = Math.max(0, Math.min(H, fy ? H - (p.y - t.y) : (p.y - t.y)));
      if (ds.kind === 'point') { c.x = lx; c.y = ly; }
      else if (ds.kind === 'c1') { c.x1 = lx; c.y1 = ly; }
      else if (ds.kind === 'c2') { c.x2 = lx; c.y2 = ly; }
    }, { history: false });
    renderCanvas();
  };
  const onUp = () => {
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
    state.version++;
    renderAll();
    syncSoon();
  };
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
}

// ---------- Canvas interaction: selection, marquee, smart guides, rich text ----------
/** Locate an element across the whole doc (top level or inside a group). */
function findInDoc(doc: any, id: string): any {
  for (const slide of doc.slides || []) {
    for (const el of slide.elements || []) {
      if (el.id === id) return el;
      for (const c of el.children || []) if (c.id === id) return c;
    }
  }
  return null;
}

/** Union bounding box of the given ids, in slide coordinates. */
function selUnion(ids: string[]) {
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  for (const id of ids) {
    const el = store.findElement(id);
    if (!el) continue;
    const r = elementRect(el);
    x0 = Math.min(x0, r.x);
    y0 = Math.min(y0, r.y);
    x1 = Math.max(x1, r.x + r.width);
    y1 = Math.max(y1, r.y + r.height);
  }
  return Number.isFinite(x0) ? { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } : null;
}

/** Move the rendered nodes with the model during a drag, instead of rebuilding the slide. */
function syncElementNodes(ids: string[]) {
  for (const id of ids) {
    const el = store.findElement(id);
    const node = slideHostEl.querySelector(`.el[data-id="${id}"]`) as HTMLElement | null;
    if (!el || !node) continue;
    const r = elementRect(el);
    node.style.left = `${r.x}px`;
    node.style.top = `${r.y}px`;
    node.style.width = `${r.width}px`;
    node.style.height = `${r.height}px`;
    node.style.transform = el.rotation ? `rotate(${el.rotation}deg)` : '';
  }
  buildOverlay();
}

/** Update only the selection highlight, leaving the canvas DOM untouched. */
function syncSelClasses() {
  slideHostEl.querySelectorAll('.el[data-id]').forEach((node) => {
    node.classList.toggle('is-sel', store.sel.includes((node as HTMLElement).dataset.id!));
  });
}

/**
 * Smart guides: snap the moving box to the edges/centres of sibling elements and of the slide.
 * Returns the correction to add to the pointer delta, plus the guide lines to draw.
 */
function snapBox(box: any, others: any[]) {
  const tol = 6 / (state.zoom || 1);
  const size = store.doc.slideSize || { width: 1280, height: 720 };
  const tx = [box.x, box.x + box.width / 2, box.x + box.width];
  const ty = [box.y, box.y + box.height / 2, box.y + box.height];
  const candX = [0, size.width / 2, size.width];
  const candY = [0, size.height / 2, size.height];
  for (const o of others) {
    candX.push(o.x, o.x + o.width / 2, o.x + o.width);
    candY.push(o.y, o.y + o.height / 2, o.y + o.height);
  }
  let bestX: any = null;
  let bestY: any = null;
  for (const t of tx)
    for (const c of candX) {
      const d = c - t;
      if (Math.abs(d) <= tol && (!bestX || Math.abs(d) < Math.abs(bestX.d))) bestX = { d, pos: c };
    }
  for (const t of ty)
    for (const c of candY) {
      const d = c - t;
      if (Math.abs(d) <= tol && (!bestY || Math.abs(d) < Math.abs(bestY.d))) bestY = { d, pos: c };
    }
  const guides: any[] = [];
  if (bestX) guides.push({ type: 'v', pos: bestX.pos });
  if (bestY) guides.push({ type: 'h', pos: bestY.pos });
  return { dx: bestX ? bestX.d : 0, dy: bestY ? bestY.d : 0, guides };
}

function showGuides(guides: any[]) {
  clearGuides();
  for (const g of guides) {
    overlayEl.append(
      h('div', { class: `guide ${g.type}`, style: g.type === 'v' ? `left:${g.pos}px` : `top:${g.pos}px` })
    );
  }
}
function clearGuides() {
  overlayEl.querySelectorAll('.guide').forEach((n) => n.remove());
}

/** Drag every selected element together. One undo entry per gesture. */
function startMove(e: PointerEvent) {
  const ids = store.sel.filter((id: string) => {
    const el = store.findElement(id);
    return el && !el.locked;
  });
  if (!ids.length) return;
  // Intentionally NOT preventing default on pointerdown: doing so on Chromium suppresses the
  // synthetic dblclick event, breaking double-click-to-edit. We only prevent default once the
  // gesture actually becomes a drag (below), which is after any click/dblclick has already fired.
  const startX = e.clientX;
  const startY = e.clientY;
  const orig = ids.map((id) => {
    const el = store.findElement(id);
    return {
      id,
      x: num(el.x),
      y: num(el.y),
      kids: (el.children || []).map((c: any) => ({ id: c.id, x: num(c.x), y: num(c.y) }))
    };
  });
  const box0 = selUnion(ids);
  if (!box0) return;
  const others = (store.slide?.elements || [])
    .filter((n: any) => n && !ids.includes(n.id) && !n.hidden)
    .map((n: any) => elementRect(n));
  // 被移动的形状集合（不含线条与组合；组合移动时其 children 也视为被移动），供 glue 重算
  const movedShapeIds = new Set<string>();
  const allEls = ids.map((id: string) => store.findElement(id)).filter(Boolean);
  for (const e of allEls) {
    if (isLineElement(e) || e.type === 'group') continue;
    movedShapeIds.add(e.id);
  }
  for (const e of allEls) {
    if (e.type === 'group') (e.children || []).forEach((c: any) => movedShapeIds.add(c.id));
  }
  // Don't snapshot until the gesture becomes a real drag: a plain click would otherwise push a
  // no-op undo entry on every selection.
  let moved = false;
  let snapped = false;

  const onMove = (ev: PointerEvent) => {
    let dx = (ev.clientX - startX) / state.zoom;
    let dy = (ev.clientY - startY) / state.zoom;
    if (!moved && Math.abs(dx) < 0.5 && Math.abs(dy) < 0.5) return;
    moved = true;
    if (!snapped) {
      store.snapshot();
      snapped = true;
    }
    // Now that this is a real drag (not a click/dblclick), suppress text selection.
    ev.preventDefault();
    let guides: any[] = [];
    if (store.snap !== false) {
      const s = snapBox(
        { x: box0.x + dx, y: box0.y + dy, width: box0.width, height: box0.height },
        others
      );
      dx += s.dx;
      dy += s.dy;
      guides = s.guides;
    }
    store.update((doc: any) => {
      for (const o of orig) {
        const el = findInDoc(doc, o.id);
        if (!el) continue;
        el.x = Math.round(o.x + dx);
        el.y = Math.round(o.y + dy);
        for (const k of o.kids) {
          const c = (el.children || []).find((x: any) => x.id === k.id);
          if (c) {
            c.x = Math.round(k.x + dx);
            c.y = Math.round(k.y + dy);
          }
        }
      }
      if (movedShapeIds.size) resyncGlue(doc, store.slideIndex, [...movedShapeIds]);
    }, { history: false });
    syncElementNodes(ids);
    showGuides(guides);
  };
  const onUp = () => {
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
    clearGuides();
    if (moved) {
      state.version++;
      renderAll();
      syncSoon();
    }
  };
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
}

/** Rubber-band selection over the canvas. */
function startMarquee(e: PointerEvent) {
  e.preventDefault();
  const startX = e.clientX;
  const startY = e.clientY;
  const box = h('div', { class: 'marquee' });
  overlayEl.append(box);
  const onMove = (ev: PointerEvent) => {
    const x0 = Math.min(startX, ev.clientX);
    const y0 = Math.min(startY, ev.clientY);
    const x1 = Math.max(startX, ev.clientX);
    const y1 = Math.max(startY, ev.clientY);
    if (x1 - x0 < 3 && y1 - y0 < 3) return;
    const r = slideHostEl.getBoundingClientRect();
    box.style.left = `${(x0 - r.left) / state.zoom}px`;
    box.style.top = `${(y0 - r.top) / state.zoom}px`;
    box.style.width = `${(x1 - x0) / state.zoom}px`;
    box.style.height = `${(y1 - y0) / state.zoom}px`;
    const bx0 = (x0 - r.left) / state.zoom;
    const by0 = (y0 - r.top) / state.zoom;
    const bx1 = (x1 - r.left) / state.zoom;
    const by1 = (y1 - r.top) / state.zoom;
    const hit = (store.slide?.elements || [])
      .filter((el: any) => el && !el.locked && !el.hidden)
      .map((el: any) => ({ id: el.id, r: elementRect(el) }))
      .filter((t: any) => t.r.x < bx1 && t.r.x + t.r.width > bx0 && t.r.y < by1 && t.r.y + t.r.height > by0)
      .map((t: any) => t.id);
    if (hit.length !== store.sel.length || hit.some((id: string) => !store.sel.includes(id))) {
      store.setSel(hit);
      syncSelClasses();
    }
  };
  const onUp = () => {
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
    box.remove();
    renderSelection();
    renderInspector();
  };
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
}

// ---------- Inline rich-text editing ----------
function richBodyNode(id: string): HTMLElement | null {
  return slideHostEl.querySelector(`.el[data-id="${id}"] [contenteditable="true"]`) as HTMLElement | null;
}

/** Double click: enter group editing, or enter rich-text editing, or hand over to the type's own affordance. */
function onCanvasDblClick(e: MouseEvent) {
  if (state.mode !== 'edit') return;
  const node = (e.target as HTMLElement).closest('.el[data-id]') as HTMLElement | null;
  if (!node) {
    // Double-click on empty canvas exits group editing (if active).
    if (store.groupEdit) store.setGroupEdit(null);
    return;
  }
  const el = store.findElement(node.dataset.id!);
  if (!el || el.locked) return;
  // 线条：双击进入顶点编辑；顶点编辑态下双击空白处插入一个顶点
  if (isLineElement(el)) {
    if (store.vertexEdit === el.id) {
      actions.addVertexAt(el.id, clientToSlide(e.clientX, e.clientY));
    } else {
      actions.enterVertexEdit(el.id);
    }
    renderAll();
    return;
  }
  // Double-clicking a group (or any element inside it) enters group editing.
  const gid = el.children && el.children.length ? el.id : findParentGroupId(el.id);
  if (gid && gid !== store.groupEdit) store.setGroupEdit(gid);
  store.setSel([el.id]);
  renderSelection();
  renderInspector();
  // Any element that carries text can be edited inline (text boxes, shapes with text, group children…)
  if (el.paragraphs && el.paragraphs.length) enterTextEditing(el);
  else if (el.type === 'image') pickImage();
  else if (el.type === 'table') editTable(el.id);
  else if (el.type === 'chart') editChart(el.id);
  else if (el.type === 'diagram') editDiagramText(el.id);
  else if (el.type === 'video' || el.type === 'audio') toast(t('toast.mediaNotPlayable'));
  // other element types (line, connector, picture, shape without text) have no inline text
}

/**
 * SmartArt 文字编辑。
 * 图示文字同时存在于数据部件（ppt/diagrams/dataN.xml）与缓存绘图部件（drawingN.xml），
 * 只改内存里的 texts 不会写进文件，故交给 parser 的 updateDiagramTexts 同步改写两处部件。
 */
function editDiagramText(id: string) {
  const el = store.findElement(id);
  if (!el || el.type !== 'diagram') return;
  const texts: string[] = Array.isArray(el.texts) ? el.texts.map((v: any) => String(v ?? '')) : [];
  if (!texts.length) {
    toast(t('insp.diagramHint'));
    return;
  }
  openDiagramTextDialog(texts, {
    save: (next) =>
      runAction(() => {
        store.update((doc: any) => {
          const elx = findInDoc(doc, id);
          if (!elx) return;
          updateDiagramTexts(elx, next);
        });
      })
  });
}

/** Open the table editor against the freshest model state. */
function editTable(id: string) {
  const el = store.findElement(id);
  if (!el) return;
  openTableDialog(el, {
    update: (patch: any) => runAction(() => actions.updateElement(id, patch)),
    resize: (rows: number, cols: number) => runAction(() => actions.resizeTable(el, rows, cols)),
    reopen: () => editTable(id)
  });
}

/** Open the chart data editor against the freshest model state. */
function editChart(id: string) {
  const el = store.findElement(id);
  if (!el) return;
  openChartDialog(el, { update: (patch: any) => runAction(() => actions.updateElement(id, patch)) });
}

/** Turn one text element into a contenteditable body and keep it focused. */
function enterTextEditing(el: any) {
  if (store.editingId === el.id) return;
  store.editingId = el.id;
  state.editingText = true;
  renderAll();
  // Defer focusing to the next frame: renderAll() rebuilds the DOM, and a synchronous focus()
  // right after can be a no-op in the VS Code webview (layout not yet flushed).
  requestAnimationFrame(() => {
    const body = richBodyNode(el.id);
    if (!body) return;
    body.focus();
    const range = document.createRange();
    range.selectNodeContents(body);
    range.collapse(false);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
    let timer: any = null;
    body.addEventListener('keydown', (ev: KeyboardEvent) => {
      ev.stopPropagation(); // keep the global shortcut bus out of the rich-text session
      if (ev.key === 'Escape') {
        ev.preventDefault();
        (ev.target as HTMLElement).blur();
      }
    });
    body.addEventListener('blur', () => {
      clearTimeout(timer);
      commitEditing();
    });
    body.addEventListener('input', () => {
      clearTimeout(timer);
      timer = setTimeout(() => commitEditing(true), 250);
    });
  });
}

/** Write the contenteditable content back into the model as paragraphs/runs. */
function commitEditing(keepEditing = false) {
  const id = store.editingId;
  const el = id ? store.findElement(id) : null;
  const body = id ? richBodyNode(id) : null;
  if (!id || !el || !body) {
    store.editingId = null;
    state.editingText = false;
    return;
  }
  const paras = parseRichBody(body, el);
  const prev = Array.isArray(el.paragraphs) ? el.paragraphs : [];
  store.update((doc: any) => {
    const t = findInDoc(doc, id);
    if (!t) return;
    t.paragraphs = paras.map((p: any, i: number) => ({
      runs: p.runs,
      align: (prev[i] && prev[i].align) || el.align || 'left',
      bullet: (prev[i] && prev[i].bullet) || undefined,
      lineSpacing: (prev[i] && prev[i].lineSpacing) || el.lineSpacing || 1.15
    }));
  }, { coalesce: 'text:' + id });
  if (!keepEditing) {
    store.editingId = null;
    state.editingText = false;
    renderAll();
    syncSoon();
  }
}

/** contenteditable DOM -> model paragraphs/runs. Computed px styles are converted back to pt. */
function parseRichBody(body: HTMLElement, el: any) {
  const blocks = Array.from(body.querySelectorAll('div, p')) as HTMLElement[];
  const roots = blocks.length ? blocks : [body];
  const out: any[] = [];
  for (const block of roots) {
    const runs: any[] = [];
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    let n: Node | null;
    while ((n = walker.nextNode())) {
      const text = n.nodeValue ?? '';
      if (!text) continue;
      const parent = (n.parentElement as HTMLElement) || block;
      // Bullet markers are decoration rendered by the view, not part of the text
      if (parent.classList?.contains('bullet-mark') || parent.classList?.contains('bullet-image')) continue;
      const cs = getComputedStyle(parent);
      const run: any = { text };
      if (/^(bold|bolder)$/i.test(cs.fontWeight) || parseInt(cs.fontWeight, 10) >= 600) run.bold = true;
      if (cs.fontStyle === 'italic' || cs.fontStyle === 'oblique') run.italic = true;
      if (cs.textDecorationLine.includes('underline') || cs.textDecoration.includes('underline')) run.underline = true;
      const size = parseFloat(cs.fontSize);
      if (Number.isFinite(size)) run.fontSize = Math.round((size * 72) / 96);
      if (cs.color) run.color = rgbToHex(cs.color);
      const ff = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (ff) run.fontFace = ff;
      const last = runs[runs.length - 1];
      if (last && sameRunStyle(last, run)) last.text += run.text;
      else runs.push(run);
    }
    out.push({ runs: runs.length ? runs : [{ text: '' }] });
  }
  return out;
}

function sameRunStyle(a: any, b: any) {
  const keys = ['bold', 'italic', 'underline', 'fontSize', 'color', 'fontFace'];
  return keys.every((k) => (a[k] ?? undefined) === (b[k] ?? undefined));
}

function rgbToHex(css: string): string {
  const m = css.match(/rgba?\(([^)]+)\)/);
  if (!m) return css;
  const parts = m[1].split(',').map((v) => parseInt(v, 10));
  return '#' + parts.slice(0, 3).map((v) => clamp(v, 0, 255).toString(16).padStart(2, '0')).join('');
}

// ---------- Inspector ----------
/**
 * Selection changes call `renderSelection()` + `renderInspector()` (never `renderAll()`), so the
 * ribbon is refreshed here — it lives in `renderInspector()` because both are always paired.
 */
function syncChrome(): void {
  ribbon.sync(ribbonState());
}

function renderInspector(force = false) {
  if (state.mode !== 'edit') {
    inspectorEl.innerHTML = '';
    lastInspSelKey = '';
    return;
  }
  const selKey = store.sel.join(',');
  const selectionChanged = selKey !== lastInspSelKey;
  // 换了元素就必须重建：此时旧面板里的动画参数、几何值全属于上一个元素。
  if (selectionChanged && !force) animEditIndex = -1;
  if (state.editingText && store.sel.length) return; // don't rebuild while typing (keeps focus)
  // 正在操作面板里的控件（取色器 / 下拉框 / 数字框）时不要重建：
  // 每次 edit 都 renderAll 一次，重建会把控件连同它身上的焦点一起换掉，原生取色器随之
  // 关闭——连续调色只能生效第一下，看着就像「改了颜色但没实时生效」。画布由 renderAll
  // 里的 renderCanvas 独立重绘，不受这里跳过的影响。
  // 但这个保护只在「同一个选择」时有效：刚在下拉框里设完动画再点画布上另一个元素，
  // 焦点仍在旧 select 上，若不强制重建，属性面板会一直停留在上一个元素。
  const active = document.activeElement as HTMLElement | null;
  if (!force && !selectionChanged && active && inspectorEl.contains(active) && /^(INPUT|SELECT|TEXTAREA)$/.test(active.tagName)) {
    syncChrome();
    return;
  }
  // Re-rendering wipes the DOM, so remember where the panel was scrolled to: property edits must not
  // throw the user back to the top. A *different* selection is a new context and starts at the top.
  const keepScroll = !selectionChanged;
  const prevScroll = inspectorEl.scrollTop;
  lastInspSelKey = selKey;
  inspectorEl.innerHTML = '';
  const slide = store.slide;
  const count = store.slideCount;
  const els: any[] = slide?.elements || [];

  // Sticky header: slide counter + the element-list toggle. The list itself is a drawer (collapsed by
  // default) instead of a flat list above the properties — otherwise a slide with dozens of elements
  // pushes every property field far below the fold.
  const toggle = h(
    'button',
    {
      class: 'insp-toggle' + (elListOpen ? ' on' : ''),
      type: 'button',
      'aria-expanded': elListOpen ? 'true' : 'false',
      title: elListOpen ? t('insp.elCollapse') : t('insp.elExpand')
    },
    [icon('chevron', 12), h('span', {}, [t('insp.elements', els.length)])]
  );
  toggle.onclick = () => {
    elListOpen = !elListOpen;
    elFilter = '';
    writeStore(EL_LIST_KEY, elListOpen ? '1' : '0');
    renderInspector();
  };
  inspectorEl.append(h('div', { class: 'insp-head' }, [h('h3', {}, [t('status.counter', state.current + 1, count)]), toggle]));

  if (elListOpen) inspectorEl.append(renderElDrawer(els));

  if (store.sel.length > 1) {
    const selected = store.sel.map((id: string) => store.findElement(id)).filter(Boolean);
    renderMultiInspector(selected);
  } else if (store.sel.length === 1) {
    const sel = store.findElement(store.sel[0]);
    if (sel) renderElementInspector(sel);
    else renderSlideInspector(slide);
  } else {
    renderSlideInspector(slide);
  }
  inspectorEl.scrollTop = keepScroll ? prevScroll : 0;
  syncChrome();
}

/**
 * Element list drawer: name filter (once a slide has more than a handful of elements) plus a scrollable
 * list that keeps the selected entry in view. Rebuilding only the list on filter input keeps the caret
 * in the filter box; a full inspector rebuild restores focus explicitly.
 */
function renderElDrawer(els: any[]): HTMLElement {
  const drawer = h('div', { class: 'el-drawer' });
  const scroller = h('ul', { class: 'el-list' });
  const paint = () => {
    scroller.innerHTML = '';
    const q = elFilter.trim().toLowerCase();
    const shown = q
      ? els.filter((el) => {
          const s = elSnippet(el);
          return (s.name + ' ' + s.type).toLowerCase().includes(q);
        })
      : els;
    if (!shown.length) {
      scroller.append(h('li', { class: 'empty' }, [t('insp.elNoMatch')]));
      return;
    }
    let active: HTMLElement | null = null;
    for (const el of shown) {
      const tag = elSnippet(el);
      const li = h(
        'li',
        {
          class: store.sel.includes(el.id) ? 'active' : '',
          title: tag.name,
          onclick: () => {
            store.setSel([el.id]);
            renderSelection();
            renderInspector();
          }
        },
        [h('span', { class: 'name' }, [tag.name]), h('span', { class: 'tag' }, [tag.type])]
      );
      if (store.sel.includes(el.id)) active = li;
      scroller.append(li);
    }
    if (active) scrollIntoBox(scroller, active);
  };

  if (els.length > 6) {
    const input = h('input', { class: 'el-filter', type: 'text', placeholder: t('insp.elFilter') }, []);
    input.value = elFilter;
    input.addEventListener('input', () => {
      elFilter = input.value;
      paint();
    });
    // A rebuild triggered from elsewhere (an undo, a canvas click…) would drop the caret, so give the
    // filter box its focus back when it was the active element.
    input.addEventListener('blur', () => {
      if (document.activeElement !== input) elFilterFocused = false;
    });
    input.addEventListener('focus', () => (elFilterFocused = true));
    elFilterNode = input;
    drawer.append(input);
  } else {
    elFilterNode = null;
  }
  paint();
  drawer.append(scroller);
  if (elFilterNode && elFilterFocused) {
    elFilterNode.focus();
    elFilterNode.setSelectionRange(elFilterNode.value.length, elFilterNode.value.length);
  }
  return drawer;
}

/** Scroll `item` into view inside `box` only — never touches the ancestor scroll chains. */
function scrollIntoBox(box: HTMLElement, item: HTMLElement) {
  const b = box.getBoundingClientRect();
  const i = item.getBoundingClientRect();
  if (i.top < b.top) box.scrollTop -= b.top - i.top + 4;
  else if (i.bottom > b.bottom) box.scrollTop += i.bottom - b.bottom + 4;
}

// Multi-selection panel: arrange / order / group actions shared across the selection.
function renderMultiInspector(els: any[]) {
  inspectorEl.append(h('h4', {}, [t('insp.selectedCount', els.length)]));
  const btn = (label: string, title: string, fn: () => void) =>
    h('button', { class: 'btn', type: 'button', title, onclick: () => runAction(fn) }, [label]);
  inspectorEl.append(h('h4', {}, [t('insp.arrange')]));
  inspectorEl.append(
    h('div', { class: 'row' }, [
      btn(t('cmd.alignLeft'), t('cmd.alignLeft'), () => actions.alignElements('left')),
      btn(t('cmd.alignCenter'), t('cmd.alignCenter'), () => actions.alignElements('hcenter')),
      btn(t('cmd.alignRight'), t('cmd.alignRight'), () => actions.alignElements('right')),
      btn(t('cmd.alignTop'), t('cmd.alignTop'), () => actions.alignElements('top')),
      btn(t('cmd.alignMiddle'), t('cmd.alignMiddle'), () => actions.alignElements('vcenter')),
      btn(t('cmd.alignBottom'), t('cmd.alignBottom'), () => actions.alignElements('bottom'))
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      btn(t('cmd.distH'), t('cmd.distH'), () => actions.distribute('h')),
      btn(t('cmd.distV'), t('cmd.distV'), () => actions.distribute('v'))
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      btn(t('cmd.bringFront'), t('cmd.bringFront'), () => actions.zOrder('front')),
      btn(t('cmd.bringForward'), t('cmd.bringForward'), () => actions.zOrder('forward')),
      btn(t('cmd.sendBackward'), t('cmd.sendBackward'), () => actions.zOrder('backward')),
      btn(t('cmd.sendBack'), t('cmd.sendBack'), () => actions.zOrder('back'))
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      btn(t('cmd.group'), t('cmd.group'), () => actions.groupSelection()),
      btn(t('cmd.ungroup'), t('cmd.ungroup'), () => actions.ungroupSelection()),
      btn(t('cmd.duplicate'), t('cmd.duplicate'), () => actions.duplicateSelected()),
      btn(t('cmd.delete'), t('cmd.delete'), () => actions.deleteSelected())
    ])
  );
}

function elSnippet(el: any): { name: string; type: string } {
  const label = (type: string) => t(`el.type.${type}`, type);
  switch (el.type) {
    case 'text':
      return { name: (el.text || '').toString().split('\n')[0].slice(0, 24) || t('el.type.text'), type: label('text') };
    case 'shape':
      return { name: (el.shapeType || 'shape').toString(), type: label('shape') };
    case 'image':
      return { name: label('image'), type: label('image') };
    case 'chart':
      return { name: (el.chartType || 'chart').toString(), type: label('chart') };
    case 'table':
      return { name: label('table'), type: label('table') };
    case 'group':
      return { name: label('group'), type: label('group') };
    case 'diagram':
      return { name: label('diagram'), type: label('diagram') };
    case 'video':
      return { name: label('video'), type: label('video') };
    case 'audio':
      return { name: label('audio'), type: label('audio') };
    default:
      return { name: (el.type || label('element')).toString(), type: el.type || label('element') };
  }
}

// Apply a new slide size both to the model and to the live canvas transform.
function changeSlideSize(size: { width: number; height: number }) {
  runAction(() => actions.setSlideSize(size));
  state.slideSize = { width: size.width, height: size.height };
  fitIfAuto();
  renderAll();
}

// ---- Slide-level properties ----
function renderSlideInspector(slide: any) {
  inspectorEl.append(h('h4', {}, [t('insp.slideProps')]));

  inspectorEl.append(
    selectField(
      t('insp.theme'),
      THEMES.map((x: any) => ({ value: x.id, label: data('theme', x.id, x.name) })),
      typeof store.doc.theme === 'string' ? store.doc.theme : 'blue',
      (v) => runAction(() => actions.applyTheme(v, true))
    )
  );
  inspectorEl.append(
    selectField(
      t('insp.size'),
      Object.keys(SLIDE_SIZES).map((k) => ({ value: k, label: data('size', k, k) })),
      store.doc.sizeKey || '16:9',
      (v) => changeSlideSize(SLIDE_SIZES[v])
    )
  );
  const ss = state.slideSize;
  inspectorEl.append(
    h('div', { class: 'row' }, [
      numField(t('insp.width'), ss.width, (v) => changeSlideSize({ width: Math.round(v), height: ss.height })),
      numField(t('insp.height'), ss.height, (v) => changeSlideSize({ width: ss.width, height: Math.round(v) }))
    ])
  );
  const curTrans: any = slide?.transition || {};
  inspectorEl.append(
    selectField(
      t('insp.transition'),
      TRANSITIONS.map((x: any) => ({ value: x.value, label: data('transition', x.value, x.name) })),
      curTrans.type || 'none',
      (v) => setTransition({ type: v })
    )
  );
  if (curTrans.type && curTrans.type !== 'none') {
    inspectorEl.append(
      selectField(
        t('insp.transSpeed'),
        [
          ['1500', data('speed', 'slow', 'Slow')],
          ['800', data('speed', 'normal', 'Normal')],
          ['400', data('speed', 'fast', 'Fast')]
        ],
        String(curTrans.duration ?? 800),
        (v) => setTransition({ duration: Number(v) })
      )
    );
    const adv = h('input', { type: 'checkbox', ...(curTrans.advanceOnClick !== false ? { checked: 'checked' } : {}) });
    adv.addEventListener('change', () => setTransition({ advanceOnClick: adv.checked }));
    inspectorEl.append(h('div', { class: 'field' }, [h('label', {}, [t('insp.clickToAdvance')]), adv]));
  }
  inspectorEl.append(
    selectField(
      t('insp.layout'),
      LAYOUTS.map((l: any) => ({ value: l.id, label: data('layout', l.id, l.name) })),
      slide?.layout || 'titleBody',
      (v) => void applyLayoutWithConfirm(v)
    )
  );

  const bgColor = slide?.background && typeof slide.background === 'string' ? slide.background : '#ffffff';
  inspectorEl.append(colorField(t('insp.background'), bgColor, (v) => runAction(() => actions.setBackground(v))));
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', title: t('insp.backgroundImage'), onclick: pickBackgroundImage }, [t('insp.backgroundImage')]),
      h('button', {
        class: 'btn',
        type: 'button',
        title: t('insp.backgroundGradient'),
        onclick: () => runAction(() => actions.setBackground({ type: 'gradient', angle: 45, stops: [{ color: '#1a73e8', position: 0 }, { color: '#ffffff', position: 1 }] }))
      }, [t('insp.backgroundGradient')])
    ])
  );

  const hidden = h('input', { type: 'checkbox', ...(slide?.hidden ? { checked: 'checked' } : {}) });
  hidden.addEventListener('change', () =>
    runAction(() => store.update((doc: any) => {
      doc.slides[store.slideIndex].hidden = hidden.checked;
    }))
  );
  inspectorEl.append(h('div', { class: 'field' }, [h('label', {}, [t('insp.hidden')]), hidden]));

  const notes = h('textarea', { placeholder: t('insp.notesPlaceholder') }, [slide?.notes || '']);
  notes.addEventListener('focus', () => (state.editingText = true));
  notes.addEventListener('blur', () => (state.editingText = false));
  notes.addEventListener(
    'input',
    debounce(() => runAction(() => actions.setNotes(notes.value), { skipSync: true }), 400)
  );
  inspectorEl.append(h('div', { class: 'field', style: 'align-items:flex-start' }, [h('label', {}, [t('insp.notes')]), notes]));

  // ---- Animations on this slide ----
  inspectorEl.append(h('h4', {}, [t('insp.animations')]));
  const anims: any[] = Array.isArray(slide?.animations) ? slide.animations : [];
  if (!anims.length) inspectorEl.append(h('div', { class: 'empty' }, [t('insp.noAnimations')]));
  for (let i = 0; i < anims.length; i++) {
    inspectorEl.append(
      h('div', { class: 'row' }, [
        h('span', { style: 'flex:1;font-size:11px' }, [animLabel(anims[i])]),
        h('button', { class: 'btn', type: 'button', title: t('insp.animMoveUp'), onclick: () => runAction(() => actions.moveAnimation(i, -1)) }, ['↑']),
        h('button', { class: 'btn', type: 'button', title: t('insp.animMoveDown'), onclick: () => runAction(() => actions.moveAnimation(i, 1)) }, ['↓']),
        h('button', { class: 'btn', type: 'button', title: t('insp.animRemove'), onclick: () => runAction(() => actions.removeAnimation(i)) }, ['✕'])
      ])
    );
  }
  // Adding an animation needs a target element, so that form lives in the element panel
  inspectorEl.append(h('div', { class: 'empty' }, [t('insp.animSelectEl')]));

  // Slide actions
  inspectorEl.append(h('h4', {}, [t('insp.slideActions')]));
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', onclick: () => addTextElement() }, ['+ ' + t('insp.newTextBox')]),
      h('button', { class: 'btn', type: 'button', onclick: () => addShapeElement('rect') }, ['+ ' + t('insp.newRectangle')]),
      h('button', { class: 'btn', type: 'button', onclick: pickImage }, ['+ ' + t('insp.newImage')])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.duplicateSlide(state.current)) }, [t('cmd.duplicateSlide')]),
      h('button', { class: 'btn', type: 'button', onclick: () => void deleteCurrentSlide() }, [t('cmd.deleteSlide')])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.moveSlide(state.current, Math.max(0, state.current - 1))) }, [t('insp.moveUp')]),
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.moveSlide(state.current, Math.min(store.slideCount - 1, state.current + 1))) }, [t('insp.moveDown')])
    ])
  );
}

/** The hyperlink shared by an element's runs ('' when none of them is linked). */
function firstHref(el: any): string {
  for (const p of el.paragraphs || []) for (const r of p.runs || []) if (r.href) return r.href;
  return '';
}

/**
 * 文本元素的展示文字统一取自 `paragraphs`（真实 PPTX 解析后只填充 paragraphs，
 * 不会填充顶层 el.text）。拼接每行段落内 run 的文字，段落之间用换行分隔。
 */
function elementText(el: any): string {
  if (Array.isArray(el.paragraphs) && el.paragraphs.length) {
    return el.paragraphs
      .map((p: any) => (p.runs || []).map((r: any) => (r.text == null ? '' : String(r.text))).join(''))
      .join('\n');
  }
  return el.text || '';
}

/**
 * 把纯文本写回 paragraphs：按换行拆成多段，保留每段原有的段落级属性（对齐/项目符号/
 * 行距）以及首 run 的字符级样式；渲染端只读 paragraphs，因此这样改写才会真正生效。
 */
function paragraphsFromText(value: string, el: any): any[] {
  const lines = (value ?? '').split('\n');
  const prev = Array.isArray(el.paragraphs) ? el.paragraphs : [];
  return lines.map((line, i) => {
    const src = prev[i]?.runs?.[0] || {};
    const style: any = {};
    for (const k of RUN_STYLE_KEYS) if (src[k] !== undefined) style[k] = src[k];
    return {
      runs: line ? [{ text: line, ...style }] : [{ text: '' }],
      align: prev[i]?.align || el.align || 'left',
      bullet: prev[i]?.bullet || undefined,
      lineSpacing: prev[i]?.lineSpacing || el.lineSpacing || 1.15
    };
  });
}

// ---- Element-level properties ----
function renderElementInspector(el: any) {
  inspectorEl.append(h('h4', {}, [t('insp.element', t(`el.type.${el.type}`, t('el.type.unknown')))]));
  inspectorEl.append(h('div', { class: 'muted el-kind' }, [t(`el.type.${el.type}`, el.type || t('el.type.element'))]));

  const update = (patch: any) => runAction(() => actions.updateElement(el.id, patch));
  const geom = h('div', {});
  geom.append(numField(t('insp.x'), el.x, (v) => update({ x: v })));
  geom.append(numField(t('insp.y'), el.y, (v) => update({ y: v })));
  geom.append(numField(t('insp.w'), el.width, (v) => update({ width: Math.max(1, v) })));
  // 线/连接符允许 0 尺寸（水平线/垂直线是合法几何）
  const isLineEl = isLineElement(el);
  geom.append(numField(t('insp.h'), el.height, (v) => update({ height: isLineEl ? Math.max(0, v) : Math.max(1, v) })));
  geom.append(numField(t('insp.rotation'), el.rotation, (v) => update({ rotation: v })));
  inspectorEl.append(geom);

  if (el.type === 'text') {
    const ta = h('textarea', { placeholder: t('insp.textPlaceholder') }, [elementText(el)]);
    ta.addEventListener('focus', () => (state.editingText = true));
    ta.addEventListener('blur', () => (state.editingText = false));
    ta.addEventListener(
      'input',
      debounce(() => runAction(() => actions.updateElement(el.id, { paragraphs: paragraphsFromText(ta.value, el) }, { history: false })), 400)
    );
    inspectorEl.append(h('div', { class: 'field', style: 'align-items:flex-start' }, [h('label', {}, [t('insp.text')]), ta]));

    inspectorEl.append(colorField(t('insp.color'), el.color || '#1e1e1e', (v) => applyTextStyle({ color: v === 'none' ? undefined : v })));
    inspectorEl.append(numField(t('insp.fontSize'), el.fontSize, (v) => applyTextStyle({ fontSize: v })));
    inspectorEl.append(
      selectField(
        t('insp.align'),
        ['left', 'center', 'right', 'justify'].map((v) => ({ value: v, label: data('align', v, v) })),
        el.align || 'left',
        (v) => applyTextStyle({ align: v })
      )
    );
    inspectorEl.append(
      selectField(
        t('insp.valign'),
        ['top', 'middle', 'bottom'].map((v) => ({ value: v, label: data('valign', v, v) })),
        el.valign || 'top',
        (v) => update({ valign: v })
      )
    );
    inspectorEl.append(checkField(t('cmd.bold'), !!el.bold, (v) => runAction(() => actions.applyTextStyleSel({ bold: v }))));
    inspectorEl.append(checkField(t('cmd.italic'), !!el.italic, (v) => runAction(() => actions.applyTextStyleSel({ italic: v }))));
    inspectorEl.append(checkField(t('cmd.underline'), !!el.underline, (v) => runAction(() => actions.applyTextStyleSel({ underline: v }))));
    inspectorEl.append(
      selectField(
        t('insp.font'),
        FONT_LIST.map((f: string) => ({ value: f, label: data('font', f, f) })),
        el.fontFace || FONT_LIST[0],
        (v) => applyTextStyle({ fontFace: v })
      )
    );
    inspectorEl.append(numField(t('insp.lineSpacing'), el.lineSpacing || 1.15, (v) => applyTextStyle({ lineSpacing: Math.max(0.5, v) })));
    inspectorEl.append(checkField(t('cmd.bullets'), !!el.bullet, (v) => runAction(() => actions.applyTextStyleSel({ bullet: !!v }))));
    // Hyperlinks live on runs, so writes are propagated to every run of the element
    inspectorEl.append(
      textField(t('insp.link'), firstHref(el), (v) =>
        runAction(() => {
          const elx = store.findElement(el.id);
          if (!elx) return;
          const paras = (elx.paragraphs || []).map((p: any) => ({
            ...p,
            runs: (p.runs || []).map((r: any) => ({ ...r, href: v || undefined }))
          }));
          actions.updateElement(el.id, { paragraphs: paras });
        })
      )
    );

    // ---- 段落属性（缩进 / 段距 / 列表层级）----
    // 取第一段的属性作为面板显示值（多段文本统一改写时以第一段为准）
    const p0: any = (el.paragraphs || [])[0] || {};
    inspectorEl.append(h('h4', {}, [t('insp.paragraph')]));
    inspectorEl.append(numField(t('insp.indentLeft'), p0.indentLeft ?? 0, (v) => applyParagraphStyle({ indentLeft: Math.max(0, v) })));
    inspectorEl.append(numField(t('insp.indentRight'), p0.indentRight ?? 0, (v) => applyParagraphStyle({ indentRight: Math.max(0, v) })));
    inspectorEl.append(numField(t('insp.hangingIndent'), p0.indent ?? 0, (v) => applyParagraphStyle({ indent: v })));
    inspectorEl.append(numField(t('insp.spaceBefore'), p0.spaceBefore ?? 0, (v) => applyParagraphStyle({ spaceBefore: Math.max(0, v) })));
    inspectorEl.append(numField(t('insp.spaceAfter'), p0.spaceAfter ?? 0, (v) => applyParagraphStyle({ spaceAfter: Math.max(0, v) })));
    // 列表层级（a:pPr@lvl）：0 为第一级，面板按 1 起显示
    const curLevel = Number(p0.level || 0);
    inspectorEl.append(
      selectField(
        t('insp.listLevel'),
        [1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => ({ value: String(n - 1), label: String(n) })),
        String(curLevel),
        (v) => applyParagraphStyle({ level: Number(v) })
      )
    );
    inspectorEl.append(
      h('div', { class: 'row' }, [
        h('button', { class: 'btn', type: 'button', onclick: () => applyParagraphStyle({ level: Math.max(0, curLevel - 1) }) }, [t('insp.promote')]),
        h('button', { class: 'btn', type: 'button', onclick: () => applyParagraphStyle({ level: Math.min(8, curLevel + 1) }) }, [t('insp.demote')])
      ])
    );

    // ---- 文本框内边距与自动适应（元素级，a:bodyPr）----
    const inset: any = el.inset && typeof el.inset === 'object' ? el.inset : {};
    inspectorEl.append(h('h4', {}, [t('insp.textPadding')]));
    const padRow = h('div', {});
    padRow.append(numField(t('insp.left'), inset.l ?? 0, (v) => update({ inset: { ...inset, l: Math.max(0, v) } })));
    padRow.append(numField(t('insp.right'), inset.r ?? 0, (v) => update({ inset: { ...inset, r: Math.max(0, v) } })));
    padRow.append(numField(t('insp.top'), inset.t ?? 0, (v) => update({ inset: { ...inset, t: Math.max(0, v) } })));
    padRow.append(numField(t('insp.bottom'), inset.b ?? 0, (v) => update({ inset: { ...inset, b: Math.max(0, v) } })));
    inspectorEl.append(padRow);
    inspectorEl.append(
      selectField(
        t('insp.autofit'),
        ['none', 'normal', 'shape'].map((v) => ({ value: v, label: data('autofit', v, v) })),
        el.autofit || 'none',
        (v) => update({ autofit: v })
      )
    );
  } else if (el.type === 'shape') {
    inspectorEl.append(textField(t('insp.shape'), el.shapeType || 'rect', (v) => update({ shapeType: v })));
    const fillObj: any = el.fill && typeof el.fill === 'object' ? el.fill : null;
    // 渐变没有单一 color：取首个停点色作为面板上的「填充色」，否则会退化成硬编码蓝
    const flatFill = fillObj && fillObj.color
      ? fillObj.color
      : fillObj?.stops?.[0]?.color
        || (typeof el.fill === 'string' && el.fill !== 'none' ? el.fill : '#4285f4');
    const gradStops = (over: (string | undefined)[] = []): { color: string; position: number }[] =>
      [0, 1].map((i) => ({ color: over[i] ?? fillObj?.stops?.[i]?.color ?? (i === 0 ? flatFill : '#ffffff'), position: i }));
    inspectorEl.append(
      selectField(
        t('insp.fill'),
        ['solid', 'gradient', 'image', 'pattern', 'none'].map((v) => ({ value: v, label: data('fill', v, v) })),
        fillObj?.type || 'solid',
        (v) => {
          if (v === 'none') update({ fill: 'none' });
          else if (v === 'gradient') update({ fill: { type: 'gradient', angle: fillObj?.angle ?? 90, stops: gradStops() } });
          else if (v === 'image') update({ fill: { type: 'image', data: '', extension: 'png', tile: { sx: 1, sy: 1 }, srcRect: { l: 0, t: 0, r: 0, b: 0 } } });
          else if (v === 'pattern') update({ fill: { type: 'pattern', prst: 'pct5', fg: '#000000', bg: '#ffffff' } });
          else update({ fill: flatFill });
        }
      )
    );
    if (!fillObj || fillObj.type !== 'none') {
      if (!fillObj || (fillObj.type !== 'image' && fillObj.type !== 'pattern')) {
        inspectorEl.append(
          colorField(t('insp.fillColor'), flatFill, (v) => {
            // 渐变渲染端读 stops（与预览端一致），写 colors 不会有任何视觉变化
            if (fillObj && fillObj.type === 'gradient') update({ fill: { ...fillObj, stops: gradStops([v]) } });
            else update({ fill: v });
          })
        );
      }
      if (fillObj && fillObj.type === 'gradient') {
        inspectorEl.append(
          selectField(
            t('insp.gradient'),
            [
              ['0', t('data.gradient.horizontal')],
              ['90', t('data.gradient.vertical')],
              ['45', t('data.gradient.diagonal')]
            ],
            String(fillObj.angle ?? 90),
            (v) => update({ fill: { ...fillObj, angle: Number(v) } })
          )
        );
        inspectorEl.append(
          colorField(t('insp.fillColor2'), fillObj.stops?.[1]?.color || '#ffffff', (v) =>
            update({ fill: { ...fillObj, stops: gradStops([undefined, v]) } })
          )
        );
      }
      if (fillObj && fillObj.type === 'image') {
        inspectorEl.append(
          h('div', { class: 'row' }, [
            h('button', { class: 'btn', type: 'button', title: t('insp.chooseImage'), onclick: () => pickFillImage(el.id) }, [t('insp.chooseImage')])
          ])
        );
        const tile: any = fillObj.tile || { sx: 1, sy: 1 };
        inspectorEl.append(numField(t('insp.tileX'), tile.sx ?? 1, (v) => update({ fill: { ...fillObj, tile: { ...tile, sx: v } } })));
        inspectorEl.append(numField(t('insp.tileY'), tile.sy ?? 1, (v) => update({ fill: { ...fillObj, tile: { ...tile, sy: v } } })));
        const rc: any = fillObj.srcRect || { l: 0, t: 0, r: 0, b: 0 };
        inspectorEl.append(h('h4', {}, [t('insp.crop')]));
        inspectorEl.append(numField(t('insp.left'), rc.l ?? 0, (v) => update({ fill: { ...fillObj, srcRect: { ...rc, l: clamp(v, 0, 100) } } })));
        inspectorEl.append(numField(t('insp.top'), rc.t ?? 0, (v) => update({ fill: { ...fillObj, srcRect: { ...rc, t: clamp(v, 0, 100) } } })));
        inspectorEl.append(numField(t('insp.right'), rc.r ?? 0, (v) => update({ fill: { ...fillObj, srcRect: { ...rc, r: clamp(v, 0, 100) } } })));
        inspectorEl.append(numField(t('insp.bottom'), rc.b ?? 0, (v) => update({ fill: { ...fillObj, srcRect: { ...rc, b: clamp(v, 0, 100) } } })));
      }
      if (fillObj && fillObj.type === 'pattern') {
        inspectorEl.append(
          selectField(
            t('insp.pattern'),
            PATTERN_LIST.map((p) => ({ value: p[0], label: data('pattern', p[0], p[1]) })),
            fillObj.prst || 'pct5',
            (v) => update({ fill: { ...fillObj, prst: v } })
          )
        );
        inspectorEl.append(colorField(t('insp.foreground'), fillObj.fg || '#000000', (v) => update({ fill: { ...fillObj, fg: v } })));
        inspectorEl.append(colorField(t('insp.backgroundColor'), fillObj.bg || '#ffffff', (v) => update({ fill: { ...fillObj, bg: v } })));
      }
      if (!fillObj || (fillObj.type !== 'image' && fillObj.type !== 'pattern')) {
        inspectorEl.append(
          numField(t('insp.transparency'), fillObj?.transparency ?? 0, (v) =>
            update({ fill: { ...(fillObj || { type: 'solid', color: flatFill }), transparency: clamp(v, 0, 100) } })
          )
        );
      }
    }
    const lineColor = el.line && el.line !== 'none' ? el.line.color || '#000' : '#000';
    const lineWidth = el.line && el.line !== 'none' ? el.line.width || 1 : 1;
    inspectorEl.append(colorField(t('insp.borderColor'), lineColor, (v) => update({ line: { color: v, width: lineWidth } })));
    inspectorEl.append(numField(t('insp.borderWidth'), lineWidth, (v) => update({ line: v > 0 ? { color: lineColor, width: v } : 'none' })));
  } else if (el.type === 'image') {
    const adj: any = el.imageAdjust || {};
    inspectorEl.append(h('div', { class: 'row' }, [h('button', { class: 'btn', type: 'button', onclick: pickImage }, [t('insp.replaceImage')])]));
    inspectorEl.append(numField(t('insp.brightness'), adj.brightness ?? 0, (v) => update({ imageAdjust: { ...adj, brightness: clamp(v, -100, 100) } })));
    inspectorEl.append(numField(t('insp.contrast'), adj.contrast ?? 0, (v) => update({ imageAdjust: { ...adj, contrast: clamp(v, -100, 100) } })));
    inspectorEl.append(numField(t('insp.transparency'), adj.transparency ?? 0, (v) => update({ imageAdjust: { ...adj, transparency: clamp(v, 0, 100) } })));
  } else if (el.type === 'table') {
    const rowCount = (el.rows || []).length;
    const colCount = Math.max(1, ...(el.rows || []).map((r: any) => (r.cells || []).length));
    inspectorEl.append(numField(t('insp.rows'), rowCount, (v) => runAction(() => actions.resizeTable(el, Math.max(1, v), colCount))));
    inspectorEl.append(numField(t('insp.columns'), colCount, (v) => runAction(() => actions.resizeTable(el, rowCount, Math.max(1, v)))));
    inspectorEl.append(checkField(t('insp.headerRow'), !!el.headerRow, (v) => update({ headerRow: v })));
    inspectorEl.append(colorField(t('insp.borderColor'), el.border?.color || '#cfd8e3', (v) => update({ border: { ...(el.border || { width: 1 }), color: v } })));
    inspectorEl.append(numField(t('insp.borderWidth'), el.border?.width ?? 1, (v) => update({ border: { ...(el.border || { color: '#cfd8e3' }), width: v } })));
    inspectorEl.append(colorField(t('insp.headerFill'), el.headerFill || '#1A73E8', (v) => update({ headerFill: v })));
    inspectorEl.append(colorField(t('insp.cellFill'), el.cellFill || '#ffffff', (v) => update({ cellFill: v })));
    inspectorEl.append(numField(t('insp.fontSize'), el.fontSize || 16, (v) => update({ fontSize: v })));
    const inset: any = el.inset || {};
    inspectorEl.append(h('h4', {}, [t('insp.cellPadding')]));
    inspectorEl.append(
      h('div', { class: 'row' }, [
        numField(t('insp.left'), inset.l ?? 6, (v) => update({ inset: { ...inset, l: v } })),
        numField(t('insp.right'), inset.r ?? 6, (v) => update({ inset: { ...inset, r: v } }))
      ])
    );
    inspectorEl.append(
      h('div', { class: 'row' }, [
        numField(t('insp.top'), inset.t ?? 4, (v) => update({ inset: { ...inset, t: v } })),
        numField(t('insp.bottom'), inset.b ?? 4, (v) => update({ inset: { ...inset, b: v } }))
      ])
    );
    inspectorEl.append(h('div', { class: 'row' }, [h('button', { class: 'btn', type: 'button', onclick: () => editTable(el.id) }, [t('insp.editCells')])]));
  } else if (el.type === 'chart') {
    inspectorEl.append(
      selectField(
        t('insp.chartType'),
        (CHART_TYPES || []).map((x: any) => ({ value: x.value, label: data('chartType', x.value, x.name) })),
        el.chartType || 'barChart',
        (v) => update({ chartType: v })
      )
    );
    inspectorEl.append(checkField(t('insp.legend'), !!el.legend, (v) => update({ legend: v })));
    inspectorEl.append(checkField(t('insp.dataLabels'), !!el.dataLabels, (v) => update({ dataLabels: v })));
    inspectorEl.append(h('div', { class: 'row' }, [h('button', { class: 'btn', type: 'button', onclick: () => editChart(el.id) }, [t('insp.editData')])]));
    inspectorEl.append(h('div', { class: 'empty' }, [t('insp.chartHint')]));
  } else if (el.type === 'diagram') {
    // SmartArt：文字在数据部件与缓存绘图部件里，改动需由 parser 的 updateDiagramTexts 同步写回
    inspectorEl.append(textField(t('insp.name'), el.name || '', (v) => update({ name: v })));
    inspectorEl.append(
      h('div', { class: 'row' }, [
        h('button', { class: 'btn', type: 'button', onclick: () => editDiagramText(el.id) }, [t('insp.editDiagramText')])
      ])
    );
    inspectorEl.append(h('div', { class: 'empty' }, [t('insp.diagramHint')]));
  } else if (el.type === 'video' || el.type === 'audio') {
    inspectorEl.append(textField(t('insp.name'), el.name || '', (v) => update({ name: v })));
    inspectorEl.append(
      h('div', { class: 'row' }, [
        h('button', { class: 'btn', type: 'button', title: t('insp.replaceMedia'), onclick: () => pickReplaceMedia(el.id) }, [t('insp.replaceMedia')]),
        h('button', { class: 'btn', type: 'button', title: t('insp.replacePoster'), onclick: () => pickPoster(el.id) }, [t('insp.replacePoster')])
      ])
    );
  } else {
    inspectorEl.append(h('div', { class: 'empty' }, [t('insp.typeUnsupported')]));
  }

  // Visual options shared by every element type
  inspectorEl.append(h('h4', {}, [t('insp.appearance')]));
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', onclick: () => update({ flipH: !el.flipH }) }, [el.flipH ? t('cmd.unflipH') : t('cmd.flipH')]),
      h('button', { class: 'btn', type: 'button', onclick: () => update({ flipV: !el.flipV }) }, [el.flipV ? t('cmd.unflipV') : t('cmd.flipV')])
    ])
  );
  inspectorEl.append(
    checkField(t('insp.shadow'), !!el.shadow, (v) =>
      update({ shadow: v ? { type: 'outer', angle: 45, distance: 4, blur: 8, transparency: 60, color: '#000000' } : null })
    )
  );

  // Element actions
  inspectorEl.append(h('h4', {}, [t('insp.elementActions')]));
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.duplicateSelected()) }, [t('cmd.duplicate')]),
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.deleteSelected()) }, [t('cmd.delete')])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.zOrder('front')) }, [t('cmd.bringFront')]),
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.zOrder('back')) }, [t('cmd.sendBack')])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', title: t('cmd.alignLeft'), onclick: () => runAction(() => actions.alignElements('left')) }, [t('cmd.alignLeft')]),
      h('button', { class: 'btn', type: 'button', title: t('cmd.alignCenter'), onclick: () => runAction(() => actions.alignElements('hcenter')) }, [t('cmd.alignCenter')]),
      h('button', { class: 'btn', type: 'button', title: t('cmd.alignTop'), onclick: () => runAction(() => actions.alignElements('top')) }, [t('cmd.alignTop')]),
      h('button', { class: 'btn', type: 'button', title: t('cmd.alignMiddle'), onclick: () => runAction(() => actions.alignElements('vcenter')) }, [t('cmd.alignMiddle')])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', title: t('cmd.alignRight'), onclick: () => runAction(() => actions.alignElements('right')) }, [t('cmd.alignRight')]),
      h('button', { class: 'btn', type: 'button', title: t('cmd.alignBottom'), onclick: () => runAction(() => actions.alignElements('bottom')) }, [t('cmd.alignBottom')]),
      h('button', { class: 'btn', type: 'button', title: t('cmd.distH'), onclick: () => runAction(() => actions.distribute('h')) }, [t('cmd.distH')]),
      h('button', { class: 'btn', type: 'button', title: t('cmd.distV'), onclick: () => runAction(() => actions.distribute('v')) }, [t('cmd.distV')])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.zOrder('forward')) }, [t('cmd.bringForward')]),
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.zOrder('backward')) }, [t('cmd.sendBackward')]),
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.groupSelection()) }, [t('cmd.group')]),
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.ungroupSelection()) }, [t('cmd.ungroup')])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.toggleLock()) }, [el.locked ? t('cmd.unlock') : t('cmd.lock')]),
      h('button', { class: 'btn', type: 'button', onclick: () => runAction(() => actions.toggleHidden()) }, [el.hidden ? t('cmd.show') : t('cmd.hide')])
    ])
  );

  // ---- Animations targeting this element ----
  inspectorEl.append(h('h4', {}, [t('insp.animations')]));
  const allAnims: any[] = Array.isArray(store.slide?.animations) ? store.slide.animations : [];
  const mine = allAnims.map((a: any, i: number) => ({ a, i })).filter((x: any) => x.a.target === el.id);
  // 配置控件编辑「当前选中的动画」，而不是「下一次添加」的草稿。
  // 旧实现里改类型/方向只改草稿、不动已有动画，于是无论选什么，播放出来永远是最初
  // 添加的那条默认效果（从左侧飞入）。
  let editIdx = currentAnimIndex();
  if (editIdx < 0 || !allAnims[editIdx] || allAnims[editIdx].target !== el.id) {
    editIdx = mine.length ? mine[mine.length - 1].i : -1;
  }
  const editAnim: any = allAnims[editIdx] || null;
  // 没有动画时仍给一个合理的默认值，供「添加动画」按钮使用
  const pick = {
    cls: editAnim?.presetClass || 'entr',
    type: editAnim?.type || (ANIM_TYPES.entr[0] || {}).value,
    trigger: editAnim ? normTrigger(editAnim.trigger) : 'afterPrev',
    direction: editAnim?.direction || (ANIM_DIRECTIONS[0] || {}).value,
    duration: editAnim?.duration ?? 500
  };
  const patchAnim = (p: any) => {
    Object.assign(pick, p);
    if (editIdx >= 0 && editIdx < allAnims.length) {
      animEditIndex = editIdx;
      // 走统一入口，改类型时同步清掉过期的 presetID/presetSubtype
      patchCurrentAnim(p);
    }
  };
  if (!mine.length) inspectorEl.append(h('div', { class: 'empty' }, [t('insp.noAnimOnEl')]));
  for (const x of mine) {
    inspectorEl.append(
      h('div', {
        class: 'row anim-row' + (x.i === editIdx ? ' on' : ''),
        style: 'cursor:pointer',
        onclick: () => {
          animEditIndex = x.i;
          renderInspector(true);
          ribbon.sync(ribbonState());
        }
      }, [
        h('span', { style: 'flex:1;font-size:11px' }, [animLabel(x.a)]),
        h('button', {
          class: 'btn', type: 'button', title: t('insp.animRemove'),
          onclick: (e: any) => {
            e.stopPropagation();
            animEditIndex = -1;
            runAction(() => actions.removeAnimation(x.i));
          }
        }, ['✕'])
      ])
    );
  }
  const addAnim = () => {
    animEditIndex = -1;
    runAction(() =>
      actions.addAnimation({
        target: el.id,
        type: pick.type,
        presetClass: pick.cls,
        direction: pick.direction,
        duration: pick.duration,
        trigger: pick.trigger
      })
    );
  };
  // Single-click add: the button always adds an animation immediately (no hidden two-step form).
  const addButton = () =>
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', type: 'button', title: t('insp.animAdd'), onclick: addAnim }, ['＋ ' + t('insp.animAdd')])
    ]);
  if (!mine.length) {
    inspectorEl.append(addButton());
  } else {
    const typeWrap = h('div', {});
    const renderTypes = () => {
      typeWrap.innerHTML = '';
      typeWrap.append(
        selectField(
          t('insp.animType'),
          (ANIM_TYPES[pick.cls] || []).map((x: any) => ({ value: x.value, label: data('animType', x.value, x.name) })),
          pick.type,
          (v) => patchAnim({ type: v })
        )
      );
    };
    renderTypes();
    inspectorEl.append(
      selectField(
        t('insp.animClass'),
        ANIM_CLASSES.map((c: any) => ({ value: c.value, label: data('animClass', c.value, c.name) })),
        pick.cls,
        (v) => {
          const first = ((ANIM_TYPES as any)[v] || [])[0]?.value || pick.type;
          patchAnim({ presetClass: v, type: first });
          renderTypes();
        }
      )
    );
    inspectorEl.append(typeWrap);
    inspectorEl.append(
      selectField(
        t('insp.animTrigger'),
        ['onClick', 'withPrev', 'afterPrev'].map((v) => ({ value: v, label: data('trigger', v, v) })),
        pick.trigger,
        (v) => patchAnim({ trigger: v })
      )
    );
    inspectorEl.append(
      selectField(
        t('insp.animDirection'),
        ANIM_DIRECTIONS.map((d: any) => ({ value: d.value, label: data('animDir', d.value, d.name) })),
        pick.direction,
        (v) => patchAnim({ direction: v })
      )
    );
    inspectorEl.append(numField(t('insp.animDuration'), pick.duration, (v) => patchAnim({ duration: Math.max(0, Math.round(v)) })));
    inspectorEl.append(addButton());
  }
}

// ---------- Field builders ----------
function colorField(label: string, value: string, onInput: (v: string) => void) {
  const input = h('input', { type: 'color', value: toHex(value) });
  input.addEventListener('input', () => onInput(input.value));
  const clear = h('button', { class: 'btn', type: 'button', style: 'padding:2px 6px', title: t('insp.setNone'), onclick: () => onInput('none') }, [t('common.none')]);
  return h('div', { class: 'field' }, [h('label', {}, [label]), input, clear]);
}
function textField(label: string, value: string, onInput: (v: string) => void) {
  const input = h('input', { type: 'text', value: value ?? '' });
  input.addEventListener('change', () => onInput(input.value));
  return h('div', { class: 'field' }, [h('label', {}, [label]), input]);
}
function numField(label: string, value: any, onInput: (v: number) => void) {
  const input = h('input', { type: 'number', value: num(value) });
  input.addEventListener('change', () => onInput(num(input.value)));
  return h('div', { class: 'field' }, [h('label', {}, [label]), input]);
}
type FieldOption = string | [string, string] | { value: string; label: string };
function selectField(label: string, options: FieldOption[], value: string, onInput: (v: string) => void) {
  const sel = h('select', {});
  for (const o of options) {
    const val = typeof o === 'object' && !Array.isArray(o) ? o.value : Array.isArray(o) ? o[0] : o;
    const txt = typeof o === 'object' && !Array.isArray(o) ? o.label : Array.isArray(o) ? o[1] : o;
    sel.append(h('option', { value: val, ...(val === value ? { selected: 'selected' } : {}) }, [txt]));
  }
  sel.addEventListener('change', () => onInput(sel.value));
  return h('div', { class: 'field' }, [h('label', {}, [label]), sel]);
}
function checkField(label: string, value: boolean, onInput: (v: boolean) => void) {
  const input = h('input', { type: 'checkbox', ...(value ? { checked: 'checked' } : {}) });
  input.addEventListener('change', () => onInput(input.checked));
  return h('div', { class: 'field' }, [h('label', {}, [label]), input]);
}
function toHex(v: any): string {
  if (typeof v === 'string' && /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(v)) return v;
  return '#ffffff';
}

// ---------- Slide selection ----------
function selectSlide(i: number) {
  state.current = clamp(i, 0, slideCount() - 1);
  if (state.mode === 'edit') store.setSlide(state.current);
  renderAll();
}

// ---------- Element creation ----------
/** New element id, matching the shape of the ids created while importing (`e_xxxx`). */
function newId(): string {
  return 'e_' + Math.random().toString(36).slice(2, 10);
}

function addTextElement() {
  runAction(() =>
    actions.addElement({
      id: newId(),
      type: 'text',
      x: 120,
      y: 120,
      width: 480,
      height: 90,
      text: 'Double-click to edit text',
      fontSize: 24,
      color: '#1e1e1e',
      align: 'left',
      valign: 'top'
    })
  );
}
function addShapeElement(shapeType: string, opt: any = {}) {
  // 线/连接符：没有填充，且初始为水平线（高度 0）
  // —— 竖直线/水平线用 0 尺寸表达是合法几何，若给个默认高度会变成斜线
  const isLine = /^(curvedConnector|bentConnector|straightConnector|line)/.test(shapeType);
  // 连接符（p:cxnSp）才能写 a:stCxn/a:endCxn，从而在 PowerPoint 里跟随形状移动
  const isConnector = /^(curvedConnector|bentConnector|straightConnector)/.test(shapeType);
  runAction(() =>
    actions.addElement({
      id: newId(), type: 'shape', shapeType,
      x: 200, y: 200,
      width: isLine ? 240 : 200,
      height: isLine ? 0 : 120,
      fill: opt.fill ?? (isLine ? 'none' : '#4285f4'),
      line: opt.line ?? { color: '#1f2937', width: 1.5 },
      ...(isConnector ? { connector: true } : {}),
      ...(opt.name ? { name: opt.name } : {})
    })
  );
}
function pickImage() {
  const input = h('input', { type: 'file', accept: 'image/*', style: 'display:none' });
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result as string;
      const img = new Image();
      img.onload = () => {
        const w = 360;
        const hgt = Math.round((img.height / img.width) * w) || 240;
        const replaceId = store.sel.length === 1 && store.findElement(store.sel[0])?.type === 'image' ? store.sel[0] : null;
        runAction(() =>
          replaceId
            ? actions.updateElement(replaceId, { data: dataUrl })
            : actions.addElement({ id: newId(), type: 'image', data: dataUrl, x: 160, y: 160, width: w, height: hgt })
        );
      };
      img.src = dataUrl;
    };
    reader.readAsDataURL(file);
  });
  document.body.append(input);
  input.click();
}

function pickBackgroundImage() {
  const input = h('input', { type: 'file', accept: 'image/*', style: 'display:none' });
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => runAction(() => actions.setBackgroundImage(reader.result as string));
    reader.readAsDataURL(file);
  });
  document.body.append(input);
  input.click();
}

function pickReplaceMedia(id: string) {
  const input = h('input', { type: 'file', accept: 'video/*,audio/*', style: 'display:none' });
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (!file) return;
    const ext = (file.name.match(/\.(\w+)$/)?.[1] || '').toLowerCase();
    const reader = new FileReader();
    reader.onload = () => runAction(() => actions.updateElement(id, { data: reader.result as string, extension: ext, name: file.name }));
    reader.readAsDataURL(file);
  });
  document.body.append(input);
  input.click();
}

function pickPoster(id: string) {
  const input = h('input', { type: 'file', accept: 'image/*', style: 'display:none' });
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (!file) return;
    const ext = (file.name.match(/\.(\w+)$/)?.[1] || 'png').toLowerCase();
    const reader = new FileReader();
    reader.onload = () => runAction(() => actions.updateElement(id, { poster: { data: reader.result as string, extension: ext } }));
    reader.readAsDataURL(file);
  });
  document.body.append(input);
  input.click();
}

function pickFillImage(id: string) {
  const input = h('input', { type: 'file', accept: 'image/*', style: 'display:none' });
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (!file) return;
    const ext = (file.name.match(/\.(\w+)$/)?.[1] || 'png').toLowerCase();
    const reader = new FileReader();
    reader.onload = () => {
      const elx = store.findElement(id);
      const cur = elx && typeof elx.fill === 'object' && elx.fill.type === 'image' ? elx.fill : {};
      runAction(() =>
        actions.updateElement(id, {
          fill: { ...cur, type: 'image', data: reader.result as string, extension: ext, tile: cur.tile || { sx: 1, sy: 1 }, srcRect: cur.srcRect || { l: 0, t: 0, r: 0, b: 0 } }
        })
      );
    };
    reader.readAsDataURL(file);
  });
  document.body.append(input);
  input.click();
}

function insertTable() {
  openTableSizeDialog({
    create: (rows: number, cols: number) => {
      const tblRows = Array.from({ length: rows }, () => ({
        cells: Array.from({ length: cols }, () => ({ text: '' }))
      }));
      runAction(() =>
        actions.addElement(
          {
            id: newId(),
            type: 'table',
            x: 160,
            y: 160,
            width: Math.max(240, cols * 90),
            height: Math.max(80, rows * 40),
            rows: tblRows,
            colWidths: new Array(cols).fill(1),
            headerRow: true,
            border: { color: '#cfd8e3', width: 1 },
            headerFill: '#1A73E8',
            fontSize: 16
          },
          { center: true }
        )
      );
    }
  });
}

function insertChart() {
  runAction(() =>
    actions.addElement(
      {
        id: newId(),
        type: 'chart',
        x: 160,
        y: 140,
        width: 520,
        height: 320,
        chartType: 'barChart',
        title: 'Chart title',
        legend: true,
        dataLabels: false,
        categories: ['A', 'B', 'C', 'D'],
        series: [{ name: 'Series 1', values: [5, 3, 8, 4] }]
      },
      { center: true }
    )
  );
}

function addMedia(type: 'video' | 'audio', data: string, ext: string) {
  runAction(() =>
    actions.addElement(
      {
        id: newId(),
        type,
        data,
        extension: ext,
        name: `${type}.${ext}`,
        x: 200,
        y: 160,
        width: type === 'video' ? 420 : 240,
        height: type === 'video' ? 260 : 90
      },
      { center: true }
    )
  );
}

// ---------- Grid / Presentation / Document info ----------
function toggleGrid() {
  state.grid = !state.grid;
  gridEl.className = 'grid-overlay' + (state.grid ? ' on' : '');
  gridBtn.classList.toggle('on', state.grid);
  ribbon.sync(ribbonState());
}

function toggleSnap() {
  store.setView({ snap: store.snap === false });
  snapBtn.classList.toggle('on', store.snap !== false);
  ribbon.sync(ribbonState());
}
function toggleDocInfo() {
  const on = docInfoEl.classList.toggle('on');
  if (on) renderDocInfo();
}

// Document-property labels come from the dictionary (core.xml field names → meta.* keys).
const META_KEYS = [
  'title', 'subject', 'author', 'keywords', 'description', 'lastModifiedBy',
  'created', 'modified', 'category', 'status', 'contentType', 'language'
];

function kvTable(obj: any, labels?: Record<string, string>) {
  const keys = Object.keys(obj || {}).filter((k) => obj[k] != null && obj[k] !== '');
  if (!keys.length) return h('div', { class: 'empty' }, [t('doc.none')]);
  const table = h('table', { class: 'kv' });
  for (const k of keys) {
    table.append(
      h('tr', {}, [h('td', { class: 'k' }, [labels?.[k] || k]), h('td', { class: 'v' }, [String(obj[k])])])
    );
  }
  return table;
}

/** Render document properties: metadata (core.xml) + custom properties (custom.xml) */
function renderDocInfo() {
  docInfoEl.innerHTML = '';
  const labels: Record<string, string> = {};
  for (const k of META_KEYS) labels[k] = t(`meta.${k}`, k);
  docInfoEl.append(
    h('div', { class: 'doc-info-head' }, [
      h('span', {}, [t('doc.title')]),
      h('button', {
        class: 'btn',
        type: 'button',
        style: 'padding:2px 8px',
        onclick: () => docInfoEl.classList.remove('on')
      }, [t('common.close')])
    ])
  );
  const source = state.mode === 'edit' ? store.doc : state.preview;
  docInfoEl.append(h('h4', {}, [t('doc.metadata')]), kvTable(source?.metadata, labels));
  docInfoEl.append(h('h4', {}, [t('doc.custom')]), kvTable(source?.customProps));
}

/** `fromCurrent` shows the show from the current slide (Shift+F5) instead of the first one. */
async function startPresent(fromCurrent = false) {
  if (state.mode === 'edit' && store.editingId) commitEditing();
  if (!fromCurrent) state.current = 0;
  if (state.mode === 'edit') store.setSlide(state.current);
  presentEl.className = 'present on';
  await renderPresent();
  renderAll();
}

/** Leave the show (Esc, the overlay button, or stepping past the last slide). */
function stopPresent() {
  if (!presentEl.classList.contains('on')) return;
  presentEl.className = 'present';
  disposeStaleCharts();
  renderAll();
}

/**
 * Render the current slide into the presentation host.
 *
 * Edit mode renders straight from the editor model via `renderSlideInto` (so every element keeps
 * its `data-id`, which is what element animations target). Preview mode falls back to the faithful
 * `pptxToHtml` bytes rendering.
 *
 * Two stacked layers (`.present-layer`) are used so the outgoing and incoming slides can animate
 * simultaneously for slide transitions.
 *
 * @param dir   navigation direction (1 = forward, -1 = backward) — drives directional transitions.
 * @param withTransition  false on a pure resize redraw, where re-playing the transition is unwanted.
 */
async function renderPresent(dir = 1, withTransition = true): Promise<void> {
  disposeStaleCharts();
  const editing = state.mode === 'edit' && !!store.doc && !!store.doc.slides;
  const sz = editing ? store.doc.slideSize || state.slideSize : state.slideSize;
  const prev = presentHostEl.querySelector('.present-layer') as HTMLElement | null;

  const layer = h('div', { class: 'present-layer' });
  if (editing) {
    const slide: any = store.doc.slides[state.current];
    renderSlideInto(layer, slide, store.doc, {});
    applyEntranceAnimations(layer, slide);
  } else {
    if (!state.present || state.present.version !== state.version) {
      try {
        const bytes = state.latest;
        if (!bytes?.length) return;
        const result = await pptxToHtml(bytes, PARSE_OPTS);
        state.present = {
          slides: result.slides.map((s: any) => s.html),
          charts: result.charts || [],
          version: state.version,
          global: result.styles?.global
        };
      } catch (e: any) {
        toast(t('toast.presentFailed', e?.message || String(e)), 'error');
        return;
      }
    }
    if (state.present.global) ensureGlobalStyles(state.present.global);
    layer.innerHTML = state.present.slides[state.current] || '';
    // 预览模式没有编辑器模型，动画随 pptxToHtml 的 data-timing 下发（每项带 OOXML spid）。
    // 把 spid 映射成 target 后复用同一套播放逻辑：保存后重新打开（默认预览模式）仍有动画。
    let previewAnims: any[] = [];
    try {
      const slideEl = layer.querySelector('.slide[data-timing]') as HTMLElement | null;
      const payload = JSON.parse(slideEl?.dataset.timing || '{}');
      if (Array.isArray(payload.animations)) {
        previewAnims = payload.animations.map((a: any) => ({
          ...a,
          target: String(a.spid ?? a.target ?? ''),
          trigger: a.trigger || 'onClick'
        }));
      }
    } catch { previewAnims = []; }
    applyEntranceAnimations(layer, { animations: previewAnims });
  }

  presentHostEl.style.width = sz.width + 'px';
  presentHostEl.style.height = sz.height + 'px';
  presentCountEl.textContent = t('present.counter', state.current + 1, slideCount());

  const trans: any = editing ? store.doc.slides[state.current].transition || { type: 'none' } : { type: 'none' };
  const playTrans = withTransition && !!prev && trans.type && trans.type !== 'none';
  if (playTrans) {
    presentHostEl.appendChild(layer);
    const bandType = trans.type === 'blinds' || trans.type === 'split' || trans.type === 'randomBar';
    if (bandType) playBandTransition(prev!, layer, trans, dir >= 0 ? 1 : -1);
    else playSlideTransition(prev!, layer, trans, dir);
    const dur = (trans.duration || 800) + 80;
    setTimeout(() => {
      if (prev && prev.parentNode === presentHostEl) prev.remove();
    }, dur);
  } else {
    if (prev) prev.remove();
    presentHostEl.appendChild(layer);
  }

  const z = Math.min(window.innerWidth / sz.width, window.innerHeight / sz.height) * 0.96;
  presentHostEl.style.transform = `scale(${z})`;

  requestAnimationFrame(() => {
    if (!editing) paintPresentCharts();
    layer.querySelectorAll('video').forEach((v) => {
      (v as HTMLVideoElement).play().catch(() => {
        /* autoplay may still be blocked; controls remain available */
      });
    });
  });
}

/** The current slide's transition (edit mode only; preview has no model). */
function currentTransition(): any {
  if (state.mode === 'edit' && store.doc?.slides?.[state.current]?.transition) {
    return store.doc.slides[state.current].transition;
  }
  return null;
}

/**
 * A click (or Space / →) in the show first advances any pending `onClick` entrance animations on
 * the current slide; only once those are exhausted does it move to the next slide (matching
 * PowerPoint's "click steps through animations, then advances"). A transition with
 * `advanceOnClick: false` never advances on click at all.
 */
function presentClick() {
  if (presentClickIdx < presentClickAnims.length) {
    const item = presentClickAnims[presentClickIdx++];
    playEntrance(item.node, item.kf, { ...item.opts, delay: 0 });
    return;
  }
  const trans = currentTransition();
  if (trans && trans.advanceOnClick === false) return;
  presentGo(1);
}

function presentGo(delta: number) {
  const total = slideCount();
  const next = state.current + delta;
  if (next < 0) return;
  if (next > total - 1) {
    stopPresent(); // stepping past the last slide ends the show, like PowerPoint
    return;
  }
  state.current = next;
  renderPresent(delta);
  if (state.mode === 'edit') store.setSlide(state.current);
  renderAll();
}
function presentNext() {
  presentGo(1);
}
function presentPrev() {
  presentGo(-1);
}

// ---------- Slide transitions ----------

/** Build the [from, to] keyframe pairs for the outgoing and incoming layers of a slide transition. */
function transitionConfig(type: string, sign: number): { prev: [any, any]; next: [any, any] } {
  const off = 100 * sign; // % offset for directional moves
  switch (type) {
    case 'fade':
      return {
        prev: [{ opacity: 1 }, { opacity: 0 }],
        next: [{ opacity: 0 }, { opacity: 1 }]
      };
    case 'zoom':
      return {
        prev: [{ opacity: 1, transform: 'scale(1)' }, { opacity: 0, transform: 'scale(1.1)' }],
        next: [{ opacity: 0, transform: 'scale(0.9)' }, { opacity: 1, transform: 'scale(1)' }]
      };
    case 'push':
      return {
        prev: [{ transform: 'translateX(0)' }, { transform: `translateX(${-off}%)` }],
        next: [{ transform: `translateX(${off}%)` }, { transform: 'translateX(0)' }]
      };
    case 'cover':
      return {
        prev: [{ transform: 'translateX(0)' }, { transform: 'translateX(0)' }],
        next: [{ transform: `translateX(${off}%)` }, { transform: 'translateX(0)' }]
      };
    case 'fly':
      return {
        prev: [{ opacity: 1, transform: 'translateX(0)' }, { opacity: 1, transform: 'translateX(0)' }],
        next: [{ opacity: 0, transform: `translateX(${off}%)` }, { opacity: 1, transform: 'translateX(0)' }]
      };
    // wipe / reveal reduce to a directional clip reveal. blinds / split / randomBar are
    // handled separately by playBandTransition (see renderPresent) for a distinct effect.
    case 'wipe':
    case 'reveal':
    default: {
      const clipFrom = sign > 0 ? 'inset(0 100% 0 0)' : 'inset(0 0 0 100%)';
      return {
        prev: [{ transform: 'translateX(0)' }, { transform: 'translateX(0)' }],
        next: [{ opacity: 1, clipPath: clipFrom }, { opacity: 1, clipPath: 'inset(0 0 0 0)' }]
      };
    }
  }
}

/** Animate the outgoing + incoming layers for a slide transition. */
function playSlideTransition(prev: HTMLElement, next: HTMLElement, trans: any, dir: number): void {
  const cfg = transitionConfig(trans.type, dir >= 0 ? 1 : -1);
  const d = trans.duration || 800;
  const apply = (el: HTMLElement, frame: any) => {
    if (frame.transform != null) el.style.transform = frame.transform;
    if (frame.opacity != null) el.style.opacity = String(frame.opacity);
    if (frame.clipPath != null) el.style.clipPath = frame.clipPath;
  };
  apply(prev, cfg.prev[0]);
  apply(next, cfg.next[0]);
  const a1 = prev.animate(cfg.prev, { duration: d, easing: 'ease', fill: 'both' });
  const a2 = next.animate(cfg.next, { duration: d, easing: 'ease', fill: 'both' });
  const clear = (el: HTMLElement) => {
    el.style.transform = '';
    el.style.opacity = '';
    el.style.clipPath = '';
  };
  a1.onfinish = () => clear(prev);
  a2.onfinish = () => clear(next);
}

/**
 * Distinct slide transitions built from horizontal slats of the outgoing slide:
 *  - `blinds`    (百叶窗): N slats collapsing to reveal the incoming slide — the classic venetian-blind.
 *  - `split`     (分割):    two halves sliding apart (up/down) to reveal the incoming slide.
 *  - `randomBar` (随机条):  N slats collapsing with random per-slat delays.
 * The incoming slide (`next`) sits underneath; an overlay of clipped copies of `prev` is stacked on
 * top and animated away. `prev` is hidden (display:none) for the duration and removed by the caller.
 * The overlay removes itself once its longest animation ends.
 */
function playBandTransition(prev: HTMLElement, next: HTMLElement, trans: any, dir: number): HTMLElement {
  const type: string = trans.type;
  const d: number = trans.duration || 800;
  const N = type === 'split' ? 2 : type === 'randomBar' ? 12 : 8;
  const sz: any = (store.doc && store.doc.slideSize) || state.slideSize || { width: 960, height: 540 };
  const W: number = sz.width;
  const H: number = sz.height;

  prev.style.display = 'none';
  const overlay = h('div', { class: 'present-blinds' });
  overlay.style.position = 'absolute';
  overlay.style.inset = '0';
  overlay.style.zIndex = '3';
  overlay.style.overflow = 'hidden';
  overlay.style.pointerEvents = 'none';

  const bands: HTMLElement[] = [];
  for (let i = 0; i < N; i++) {
    const band = document.createElement('div');
    band.style.position = 'absolute';
    band.style.overflow = 'hidden';
    band.style.left = '0';
    band.style.width = '100%';
    band.style.height = H / N + 'px';
    band.style.top = (i * H) / N + 'px';

    const clone = prev.cloneNode(true) as HTMLElement;
    clone.className = '';
    // `prev` was already set to `display:none` (line above) before cloning, so the clone inherits
    // it and would stay invisible — clearing it here is what actually makes each slat show the
    // outgoing slide. Without this, blinds / split / randomBar transitions render nothing and the
    // incoming slide just pops in with no effect.
    clone.style.display = '';
    clone.style.position = 'absolute';
    clone.style.margin = '0';
    clone.style.left = '0';
    clone.style.top = (-(i * H) / N) + 'px';
    clone.style.width = W + 'px';
    clone.style.height = H + 'px';
    clone.style.transform = '';
    clone.style.opacity = '';
    clone.style.clipPath = '';

    band.appendChild(clone);
    overlay.appendChild(band);
    bands.push(band);
  }
  presentHostEl.appendChild(overlay);

  let maxEnd = d;
  for (let i = 0; i < N; i++) {
    const band = bands[i];
    let anim: Keyframe[];
    if (type === 'split') {
      anim = i === 0
        ? [{ transform: 'translateY(0)' }, { transform: 'translateY(-100%)' }]
        : [{ transform: 'translateY(0)' }, { transform: 'translateY(100%)' }];
    } else {
      band.style.transformOrigin = type === 'randomBar' && Math.random() < 0.5 ? 'bottom' : 'top';
      anim = [{ transform: 'scaleY(1)' }, { transform: 'scaleY(0)' }];
    }
    const delay = type === 'randomBar' ? Math.random() * d * 0.6 : 0;
    const end = d + delay;
    if (end > maxEnd) maxEnd = end;
    band.animate(anim, { duration: d, delay, easing: 'ease-in-out', fill: 'both' });
  }
  setTimeout(() => {
    if (overlay.parentNode === presentHostEl) overlay.remove();
  }, maxEnd + 60);
  return overlay;
}

// ---------- Element entrance animations ----------

interface EntranceFrame {
  opacity?: number;
  transform?: string;
  clipPath?: string;
  offset?: number;
}

/**
 * Full keyframe list for one element animation. `frames` is what actually plays; its first entry
 * doubles as the pre-animation "hidden" state primed by applyEntranceAnimations, so there is no
 * flash between slide paint and animation start. Multi-step effects (bounce / shake / flash) need
 * more than a [from, to] pair, hence a list with explicit offsets.
 */
interface EntranceFrames {
  frames: EntranceFrame[];
}

/**
 * Normalize an animation duration to milliseconds. The editor stores milliseconds (default 500),
 * but animations round-tripped through PPTX come back in seconds (e.g. 0.5), which would otherwise
 * produce an imperceptible 0.5ms animation. Anything below 20 is treated as seconds.
 */
function animDurationMs(d: any): number {
  const v = typeof d === 'number' && d > 0 ? d : 500;
  return Math.round(v < 20 ? v * 1000 : v);
}

/** Map an animation record to its keyframes (for the `entr`/`exit`/`emph` classes). */
function entranceKeyframes(a: any): EntranceFrames {
  const dir = a.direction || 'l';
  const dist = 80; // px
  // direction → offset / clip edge
  const move: Record<string, string> = {
    l: `translateX(${-dist}px)`,
    r: `translateX(${dist}px)`,
    t: `translateY(${-dist}px)`,
    b: `translateY(${dist}px)`
  };
  const clipFrom: Record<string, string> = {
    l: 'inset(0 0 0 100%)', // reveal expands rightward (comes from left)
    r: 'inset(0 100% 0 0)', // reveal expands leftward
    t: 'inset(100% 0 0 0)',
    b: 'inset(0 0 100% 0)'
  };
  const type = a.type || 'fadeIn';
  if (a.presetClass === 'emph') {
    // Emphasis: the element is already on stage; every effect returns to the natural state at the
    // end (PowerPoint emphasis effects never permanently change the object).
    switch (type) {
      case 'grow':
        return {
          frames: [
            { opacity: 1, transform: 'scale(1)', offset: 0 },
            { opacity: 1, transform: 'scale(1.3)', offset: 0.7 },
            { opacity: 1, transform: 'scale(1)', offset: 1 }
          ]
        };
      case 'shake':
        return {
          frames: [
            { opacity: 1, transform: 'translateX(0)', offset: 0 },
            { opacity: 1, transform: 'translateX(-11px)', offset: 0.16 },
            { opacity: 1, transform: 'translateX(11px)', offset: 0.32 },
            { opacity: 1, transform: 'translateX(-9px)', offset: 0.48 },
            { opacity: 1, transform: 'translateX(9px)', offset: 0.62 },
            { opacity: 1, transform: 'translateX(-5px)', offset: 0.76 },
            { opacity: 1, transform: 'translateX(5px)', offset: 0.9 },
            { opacity: 1, transform: 'translateX(0)', offset: 1 }
          ]
        };
      case 'flash':
        return {
          frames: [
            { opacity: 1, offset: 0 },
            { opacity: 0, offset: 0.25 },
            { opacity: 1, offset: 0.5 },
            { opacity: 0, offset: 0.75 },
            { opacity: 1, offset: 1 }
          ]
        };
      case 'pulse':
      default:
        // pulse: normal → enlarged → normal (smooth, no snap-back)
        return {
          frames: [
            { opacity: 1, transform: 'scale(1)', offset: 0 },
            { opacity: 1, transform: 'scale(1.12)', offset: 0.5 },
            { opacity: 1, transform: 'scale(1)', offset: 1 }
          ]
        };
    }
  }
  if (a.presetClass === 'exit') {
    // reverse: visible → hidden
    switch (type) {
      case 'fadeOut':
        return { frames: [{ opacity: 1 }, { opacity: 0 }] };
      case 'wipeOut':
      case 'wipe':
        return { frames: [{ opacity: 1, clipPath: 'none' }, { opacity: 1, clipPath: clipFrom[dir] }] };
      case 'zoomOut':
      case 'zoom':
        return { frames: [{ opacity: 1, transform: 'scale(1)' }, { opacity: 0, transform: 'scale(0.3)' }] };
      case 'flyOut':
      case 'fly':
      default:
        return { frames: [{ opacity: 1, transform: 'translateX(0)' }, { opacity: 0, transform: move[dir] }] };
    }
  }
  // entrance
  switch (type) {
    case 'flyIn':
    case 'fly':
      return { frames: [{ opacity: 0, transform: move[dir] }, { opacity: 1 }] };
    case 'wipeIn':
    case 'wipe':
      return { frames: [{ opacity: 1, clipPath: clipFrom[dir] }, { opacity: 1, clipPath: 'inset(0 0 0 0)' }] };
    case 'zoomIn':
    case 'zoom':
      return { frames: [{ opacity: 0, transform: 'scale(0.3)' }, { opacity: 1 }] };
    case 'riseUp':
      return { frames: [{ opacity: 0, transform: 'translateY(60px)' }, { opacity: 1 }] };
    case 'bounceIn':
    case 'bounce': {
      // Drop in from one side, then two diminishing rebounds before settling.
      // ('b' reverses the drop direction; all others fall from above, like PowerPoint's Bounce In.)
      // The drop distance is intentionally larger than the other effects so the bounce reads clearly.
      const fromBelow = dir === 'b';
      const top = fromBelow ? 'translateY(120px)' : 'translateY(-120px)';
      const rebound = (px: number) => `translateY(${fromBelow ? px : -px}px)`;
      return {
        frames: [
          { opacity: 0, transform: top, offset: 0 },
          { opacity: 1, transform: 'translateY(0)', offset: 0.4 },
          { opacity: 1, transform: rebound(28), offset: 0.58 },
          { opacity: 1, transform: 'translateY(0)', offset: 0.73 },
          { opacity: 1, transform: rebound(12), offset: 0.87 },
          { opacity: 1, transform: 'translateY(0)', offset: 1 }
        ]
      };
    }
    case 'fadeIn':
    default:
      return { frames: [{ opacity: 0 }, { opacity: 1 }] };
  }
}

/** Play one element entrance animation and clear inline styles when done. */
function playEntrance(node: HTMLElement, kf: EntranceFrames, opts: KeyframeAnimationOptions): void {
  const list = kf.frames.length ? kf.frames : [{ opacity: 0 }, { opacity: 1 }];
  const n = list.length;
  const frames: Keyframe[] = list.map((f, i) => ({
    opacity: f.opacity ?? 1,
    transform: f.transform ?? 'none',
    clipPath: f.clipPath ?? 'none',
    offset: f.offset ?? (n > 1 ? i / (n - 1) : 1)
  }));
  const anim = node.animate(frames, opts);
  anim.onfinish = () => {
    node.style.opacity = '';
    node.style.transform = '';
    node.style.clipPath = '';
  };
}

/**
 * Set up a slide's element entrance animations. Auto-triggered animations (`afterPrev`, `withPrev`)
 * play in sequence (with a shared start delay for `withPrev`); `onClick` animations stay hidden and
 * are stepped by `presentClick`.
 */
function applyEntranceAnimations(layer: HTMLElement, slide: any): void {
  presentClickAnims = [];
  presentClickIdx = 0;
  const anims: any[] = Array.isArray(slide.animations) ? slide.animations : [];
  if (!anims.length) return;
  let cum = 0;
  for (const a of anims) {
    const targetId = String(a.target ?? '').replace(/"/g, '\\"');
    // 编辑模式用 data-id（编辑器元素 id），预览模式用 _id（OOXML spid）
    const node = targetId
      ? (layer.querySelector(`[data-id="${targetId}"], [_id="${targetId}"]`) as HTMLElement | null)
      : null;
    if (!node) continue;
    const kf = entranceKeyframes(a);
    // prime the hidden/start state before the layer becomes visible (no flash)
    const start = kf.frames[0] || {};
    node.style.opacity = String(start.opacity ?? 1);
    if (start.transform) node.style.transform = start.transform;
    if (start.clipPath) node.style.clipPath = start.clipPath;
    const dur = animDurationMs(a.duration);
    const opts: KeyframeAnimationOptions = { duration: dur, easing: 'ease', fill: 'both' };
    const trigger = normTrigger(a.trigger);
    if (trigger === 'onClick') {
      presentClickAnims.push({ node, kf, opts, cls: a.presetClass || 'entr' });
    } else {
      playEntrance(node, kf, { ...opts, delay: cum });
      if (trigger !== 'withPrev') cum += dur;
    }
  }
}

// ---------- Context menu ----------
function showCtxMenu(x: number, y: number) {
  const items: MenuItem[] = [];
  if (state.mode === 'edit') {
    const hasSel = store.sel.length > 0;
    items.push(
      { label: t('menu.previewMode'), icon: 'showSlide', onClick: () => setMode('preview') },
      { type: 'sep' },
      { label: t('menu.bringFront'), icon: 'zFront', disabled: !hasSel, onClick: () => runAction(() => actions.zOrder('front')) },
      { label: t('menu.sendBack'), icon: 'zBack', disabled: !hasSel, onClick: () => runAction(() => actions.zOrder('back')) },
      { type: 'sep' },
      { label: t('menu.cut'), icon: 'cut', accel: kbd('mod+x'), disabled: !hasSel, onClick: () => { actions.copySelected(true); toast(t('toast.cut')); } },
      { label: t('menu.copy'), icon: 'copy', accel: kbd('mod+c'), disabled: !hasSel, onClick: () => { actions.copySelected(false); toast(t('toast.copied')); } },
      { label: t('menu.paste'), icon: 'paste', accel: kbd('mod+v'), onClick: () => runAction(() => actions.paste()) },
      { type: 'sep' },
      { label: t('menu.duplicate'), icon: 'duplicate', accel: kbd('mod+d'), disabled: !hasSel, onClick: () => runAction(() => actions.duplicateSelected()) },
      { label: t('menu.delete'), icon: 'deleteSlide', accel: kbd('del'), disabled: !hasSel, onClick: () => runAction(() => actions.deleteSelected()) },
      { type: 'sep' },
      { label: store.selected().some((el: any) => el.locked) ? t('menu.unlock') : t('menu.lock'), icon: 'lock', disabled: !hasSel, onClick: () => runAction(() => actions.toggleLock()) },
      { label: t('cmd.selectAll'), icon: 'selectAll', accel: kbd('mod+a'), onClick: () => { actions.selectAll(); renderAll(); } },
      { type: 'sep' },
      { label: t('menu.group'), icon: 'group', accel: kbd('mod+g'), disabled: store.sel.length < 2, onClick: () => runAction(() => actions.groupSelection()) },
      { label: t('menu.ungroup'), icon: 'ungroup', accel: kbd('mod+shift+g'), onClick: () => runAction(() => actions.ungroupSelection()) }
    );
  } else {
    items.push(
      { label: t('menu.editMode'), icon: 'edit', onClick: () => setMode('edit') },
      { type: 'sep' },
      { label: t('cmd.present'), icon: 'present', accel: kbd('f5'), onClick: () => startPresent(false) },
      { label: t('menu.docProps'), icon: 'props', onClick: () => toggleDocInfo() }
    );
  }
  openMenu(items, x, y);
}

function showSlideCtxMenu(x: number, y: number, idx: number) {
  const count = slideCount();
  const editing = state.mode === 'edit';
  const slide = editing ? store.doc.slides[idx] : null;
  const items: MenuItem[] = [
    { label: t('menu.slide.newAfter'), icon: 'newSlide', onClick: () => runAction(() => { const n = store.slideCount; actions.addSlide(); actions.moveSlide(n, idx + 1); }) },
    { label: t('menu.slide.duplicate'), icon: 'duplicateSlide', onClick: () => runAction(() => actions.duplicateSlide(idx)) },
    { type: 'sep' },
    { label: t('menu.slide.moveUp'), icon: 'up', disabled: idx <= 0, onClick: () => runAction(() => actions.moveSlide(idx, idx - 1)) },
    { label: t('menu.slide.moveDown'), icon: 'down', disabled: idx >= count - 1, onClick: () => runAction(() => actions.moveSlide(idx, idx + 1)) },
    { label: slide?.hidden ? t('menu.slide.show') : t('menu.slide.hide'), icon: slide?.hidden ? 'showSlide' : 'hideSlide', onClick: () => runAction(() => actions.toggleSlideHidden(idx)) },
    { type: 'sep' },
    { label: t('menu.slide.delete'), icon: 'deleteSlide', disabled: count <= 1, onClick: () => { state.current = idx; void deleteCurrentSlide(); } }
  ];
  openMenu(items, x, y);
}

// ---------- Keyboard ----------
function onKeyDown(e: KeyboardEvent) {
  const target = e.target as HTMLElement;
  // SELECT matters now that the ribbon lives above the canvas: arrow keys must still nudge a
  // selection, not scroll an open dropdown.
  const typing =
    target &&
    (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.tagName === 'SELECT' || target.isContentEditable);
  // An open popup menu owns Esc / ↑ ↓ / ↵ — checked first so Esc closes the menu rather than
  // falling through to "cancel the current selection".
  if (popupKeydown(e)) return;
  if (e.key === 'Escape' && hasOpenModal()) {
    closeModal();
    return;
  }
  if (e.key === 'Escape') {
    if (presentEl.classList.contains('on')) {
      stopPresent();
      return;
    }
    if (store.groupEdit) {
      store.setGroupEdit(null);
      return;
    }
    if (store.vertexEdit) {
      actions.exitVertexEdit();
      renderAll();
      return;
    }
    if (state.mode === 'edit' && store.sel.length) {
      store.setSel([]);
      renderSelection();
      renderInspector();
      ribbon.sync(ribbonState());
    }
    return;
  }
  if (typing) return;
  if (e.key === 'F5') {
    e.preventDefault();
    startPresent(e.shiftKey);
    return;
  }
  if (presentEl.classList.contains('on')) {
    if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') presentClick();
    else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp' || e.key === 'PageUp') presentPrev();
    return;
  }
  const mod = e.ctrlKey || e.metaKey;
  if (mod && !e.altKey) {
    const k = e.key.toLowerCase();
    if (k === 'b' || k === 'i' || k === 'u') {
      if (state.mode !== 'edit') return;
      e.preventDefault();
      const one = store.sel.length === 1 ? store.findElement(store.sel[0]) : null;
      if (!one) return;
      const patch = k === 'b' ? { bold: !one.bold } : k === 'i' ? { italic: !one.italic } : { underline: !one.underline };
      runAction(() => actions.applyTextStyleSel(patch));
      return;
    }
  }
  if (e.key === '?') {
    e.preventDefault();
    openShortcuts();
    return;
  }
  if (state.mode !== 'edit') return;
  // 顶点编辑态：Delete/Backspace 删除当前选中顶点
  if (store.vertexEdit && (e.key === 'Delete' || e.key === 'Backspace')) {
    e.preventDefault();
    actions.deleteSelectedVertex();
    renderAll();
    syncSoon();
    return;
  }
  if (mod && e.key.toLowerCase() === 'z') {
    e.preventDefault();
    runAction(() => (e.shiftKey ? store.redo() : store.undo()));
    return;
  }
  if (mod && e.key.toLowerCase() === 'y') {
    e.preventDefault();
    runAction(() => store.redo());
    return;
  }
  if (mod && e.key.toLowerCase() === 'a') {
    e.preventDefault();
    actions.selectAll();
    renderAll();
    return;
  }
  if (mod && e.key.toLowerCase() === 'c') {
    e.preventDefault();
    actions.copySelected(false);
    renderAll();
    syncSoon();
    return;
  }
  if (mod && e.key.toLowerCase() === 'x') {
    e.preventDefault();
    actions.copySelected(true);
    renderAll();
    syncSoon();
    return;
  }
  if (mod && e.key.toLowerCase() === 'v') {
    e.preventDefault();
    actions.paste();
    renderAll();
    syncSoon();
    return;
  }
  if (mod && e.key.toLowerCase() === 'd') {
    e.preventDefault();
    runAction(() => actions.duplicateSelected());
    return;
  }
  if (mod && e.key.toLowerCase() === 'g') {
    e.preventDefault();
    runAction(() => (e.shiftKey ? actions.ungroupSelection() : actions.groupSelection()));
    return;
  }
  if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault();
    runAction(() => actions.deleteSelected());
    return;
  }
  if (e.key === 'Enter') {
    const one = store.sel.length === 1 ? store.findElement(store.sel[0]) : null;
    if (one && one.type === 'text') {
      e.preventDefault();
      enterTextEditing(one);
    }
    return;
  }
  if (e.key === 'Tab') {
    e.preventDefault();
    // Cycle through the elements of the current slide
    const els = (store.slide?.elements || []).filter((n: any) => n && !n.hidden);
    if (!els.length) return;
    const cur = store.sel.length ? els.findIndex((n: any) => n.id === store.sel[0]) : -1;
    const next = e.shiftKey ? (cur <= 0 ? els.length - 1 : cur - 1) : (cur + 1) % els.length;
    store.setSel([els[next].id]);
    renderSelection();
    renderInspector();
    return;
  }
  const nudges: Record<string, [number, number]> = {
    ArrowLeft: [-1, 0],
    ArrowRight: [1, 0],
    ArrowUp: [0, -1],
    ArrowDown: [0, 1]
  };
  const nudge = nudges[e.key];
  if (nudge && store.sel.length) {
    e.preventDefault();
    const step = e.shiftKey ? 10 : 1;
    runAction(() => actions.nudge(nudge[0] * step, nudge[1] * step));
  }
}

// ---------- Host messages ----------
window.addEventListener('message', async (event: MessageEvent<HostToWebview>) => {
  const msg = event.data;
  try {
    switch (msg.type) {
      case 'init': {
        state.original = toBytes(msg.originalBytes);
        state.latest = toBytes(msg.bytes);
        state.title = msg.title || state.title;
        state.mode = msg.mode === 'edit' ? 'edit' : 'preview';
        state.dirty = msg.dirty;
        // `<html lang>` already set the locale at startup; honour the host's value as
        // authoritative and refresh every baked-in label if it differs.
        if (msg.locale && setLocale(msg.locale)) applyLocale();
        state.version++;
        await renderPreview();
        await applyMode();
        syncChrome();
        break;
      }
      case 'bytes': {
        state.latest = toBytes(msg.bytes);
        state.dirty = msg.dirty;
        state.version++;
        if (state.mode === 'preview') await renderPreview();
        syncChrome();
        break;
      }
      case 'flush':
        // 保存前立即把挂起的编辑序列化，避免 Ctrl+S 抢在 350ms 防抖 / 异步序列化之前。
        // syncBytes 会发出 `sync`，宿主的保存等待由那条消息释放。
        if (state.mode === 'edit') void syncBytes(store.canUndo());
        else post({ type: 'flushDone' });
        break;
      case 'saved':
        state.dirty = false;
        syncChrome();
        toast(t('toast.saved'));
        break;
      case 'reverted':
        state.dirty = false;
        syncChrome();
        break;
      case 'setMode':
        // 来自 VS Code 标题栏的 pptx-parser.viewMode / pptx-parser.editMode 命令（host 经 setMode 消息下发）
        await setMode(msg.mode === 'edit' ? 'edit' : 'preview');
        break;
      case 'info':
        toast(msg.message, msg.kind || 'info');
        break;
    }
  } catch (e: any) {
    console.error('[pptx-webview] Message failed:', e);
    toast(t('toast.renderFailed', e?.message || String(e)), 'error');
  }
});

// ---------- Bootstrap ----------
async function main() {
  const style = document.createElement('style');
  style.textContent = CSS;
  document.head.append(style);
  // Follow VS Code's display language. The host writes it into <html lang> before this runs
  // (so there is no flash of English); `init` carries it again as the authoritative value.
  setLocale(document.documentElement.lang);
  buildLayout();
  post({ type: 'ready' });
}
main().catch((e) => {
  console.error('[pptx-webview] Init failed:', e);
  document.body.textContent = 'Editor failed to start: ' + (e?.message || String(e));
});