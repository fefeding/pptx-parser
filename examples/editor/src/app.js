/**
 * 应用入口：装配各模块、绑定全局事件
 */
import { store } from './store.js';
import { renderCanvas, updateSelection, fitToScreen, zoomBy, setZoomValue, initCanvas } from './interact.js';
import { initToolbar } from './toolbar.js';
import { initInspector } from './inspector.js';
import { initSlidePanel } from './slides.js';
import { exportPptx, saveJson } from './io.js';
import { setNotes } from './actions.js';
import { startPresent } from './present.js';

const $ = (id) => document.getElementById(id);

function boot() {
  store.init();

  const dom = {
    stage: $('stage'), frame: $('frame'), overlay: $('overlay'), scroll: $('canvasScroll'),
    menubar: $('menubar'), quickbar: $('quickbar'),
    inspector: $('inspector'), slideList: $('slideList'), addBtn: $('addSlideBtn')
  };

  initCanvas({
    stage: dom.stage, frame: dom.frame, overlay: dom.overlay, scroll: dom.scroll
  });
  initToolbar({ menubar: dom.menubar, quickbar: dom.quickbar });
  initInspector({ inspector: dom.inspector });
  initSlidePanel({ list: dom.slideList, addBtn: dom.addBtn });

  store.on('doc', (payload) => {
    if (payload && payload.mute) return;       // 文本编辑中的实时输入：不重绘画布（保留光标）
    renderCanvas();
    syncTopbar();
  });
  store.on('slide', () => { renderCanvas(); syncTopbar(); });
  store.on('zoom', () => { renderCanvas(); updateZoom(); });

  bindTopbar();
  bindZoomBar();
  bindGlobalUI();

  renderCanvas();
  updateZoom();
  fitToScreen();
  syncTopbar();
  $('saveState') && ($('saveState').textContent = '已就绪');
}

function updateZoom() {
  const v = $('zoomVal');
  if (v) v.textContent = Math.round(store.zoom * 100) + '%';
}

function syncTopbar() {
  const t = $('docTitle'); if (t && t.value !== store.doc.title) t.value = store.doc.title;
  const n = $('notesInput'); if (n) n.value = store.slide.notes || '';
  const c = $('sbSlideCount');
  if (c) c.textContent = `${store.slideIndex + 1} / ${store.doc.slides.length} 页`;
  const info = $('statusInfo'); if (info) info.textContent = '就绪';
}

function bindTopbar() {
  $('btnExportTop')?.addEventListener('click', () => exportPptx());
  $('btnPresentTop')?.addEventListener('click', () => startPresent(store.slideIndex));
  const title = $('docTitle');
  title?.addEventListener('input', () => store.update((d) => { d.title = title.value; }, { coalesce: 'title' }));
  title?.addEventListener('keydown', (e) => e.stopPropagation());
  $('notesInput')?.addEventListener('input', () => setNotes($('notesInput').value));
  $('notesInput')?.addEventListener('keydown', (e) => e.stopPropagation());
}

function bindZoomBar() {
  $('zoomOut')?.addEventListener('click', () => zoomBy(0.9));
  $('zoomIn')?.addEventListener('click', () => zoomBy(1.1));
  $('zoomFit')?.addEventListener('click', () => fitToScreen());
  $('zoomVal')?.addEventListener('click', () => setZoomValue(1));
  $('zoomVal')?.addEventListener('dblclick', () => {
    const z = prompt('输入缩放百分比（如 100）：', Math.round(store.zoom * 100));
    if (z) setZoomValue(Math.max(10, Math.min(400, parseFloat(z) || 100)) / 100);
  });
}

function bindGlobalUI() {
  const gridBtn = $('gridToggle');
  const syncGrid = () => { gridBtn?.classList.toggle('on', store.showGrid); };
  gridBtn?.addEventListener('click', () => { store.setView({ showGrid: !store.showGrid }); renderCanvas(); syncGrid(); });
  syncGrid();

  const notesBtn = $('notesToggle');
  const notesPanel = $('notesPanel');
  const syncNotes = () => {
    notesBtn?.classList.toggle('on', store.showNotes);
    notesPanel?.classList.toggle('show', store.showNotes);
  };
  notesBtn?.addEventListener('click', () => { store.setView({ showNotes: !store.showNotes }); syncNotes(); });
  $('notesClose')?.addEventListener('click', () => { store.setView({ showNotes: false }); syncNotes(); });
  syncNotes();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', boot);
} else {
  boot();
}
