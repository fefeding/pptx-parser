/**
 * 顶部菜单栏 + 快捷插入工具栏
 */
import { store } from './store.js';
import { h, pickFile, readFileAsDataURL } from './util.js';
import {
  createTextElement, createShapeElement, createImageElement, createTableElement, createChartElement,
  createAudioElement, createVideoElement, CHART_TYPES, THEMES, createDoc
} from './model.js';
import {
  addElement, duplicateSelected, deleteSelected, copySelected, paste, selectAll,
  groupSelection, ungroupSelection, alignElements, zOrder, applyTheme as applyThemeAction
} from './actions.js';
import { openMenu, openShapePicker, openShortcuts, openPageSetup } from './dialogs.js';
import { exportPptx, saveJson, importPptxFile, loadJsonFile } from './io.js';
import { fitToScreen, setZoomValue, renderCanvas } from './interact.js';
import { startPresent } from './present.js';

let DOM = {};
let undoBtn, redoBtn;

export function initToolbar(dom) {
  DOM = dom;
  buildMenubar(dom.menubar);
  buildQuickbar(dom.quickbar);
  store.on('doc', updateState);
  store.on('slide', updateState);
  updateState();
}

function updateState() {
  if (undoBtn) undoBtn.disabled = !store.canUndo();
  if (redoBtn) redoBtn.disabled = !store.canRedo();
}

function menuBtn(label, onClick) {
  const b = h('button', { class: 'menu-btn', text: label });
  b.addEventListener('click', (e) => { e.stopPropagation(); onClick(b); });
  return b;
}

function buildMenubar(host) {
  if (!host) return;
  host.appendChild(menuBtn('文件', (b) => openMenu(b, [
    { label: '新建演示文稿', action: () => { if (confirm('清空当前内容并新建？')) store.setDoc(createDoc(store.doc.theme, store.doc.sizeKey || '16:9')); } },
    { label: '打开 PPTX…', action: openPptx },
    { label: '打开 JSON…', action: openJson },
    'sep',
    { label: '导出 PPTX', shortcut: 'Ctrl+E', action: exportPptx },
    { label: '保存 JSON', shortcut: 'Ctrl+S', action: saveJson }
  ])));

  host.appendChild(menuBtn('编辑', (b) => openMenu(b, [
    { label: '撤销', shortcut: 'Ctrl+Z', action: () => store.undo() },
    { label: '重做', shortcut: 'Ctrl+Shift+Z', action: () => store.redo() },
    'sep',
    { label: '复制', shortcut: 'Ctrl+C', action: () => copySelected(false) },
    { label: '剪切', shortcut: 'Ctrl+X', action: () => copySelected(true) },
    { label: '粘贴', shortcut: 'Ctrl+V', action: paste },
    { label: '再制', shortcut: 'Ctrl+D', action: duplicateSelected },
    'sep',
    { label: '全选', shortcut: 'Ctrl+A', action: selectAll },
    { label: '删除', shortcut: 'Del', action: deleteSelected }
  ])));

  host.appendChild(menuBtn('插入', (b) => openMenu(b, [
    { label: '文本框', action: () => addElement(createTextElement({ text: '双击编辑文字', x: 220, y: 220 })) },
    { label: '图片…', action: insertImage },
    { label: '音频…', action: () => insertMedia('audio') },
    { label: '视频…', action: () => insertMedia('video') },
    { label: '形状', action: (e) => openShapePickerSub(b) },
    { label: '表格', action: (e) => openTableSub(b) },
    { label: '图表', action: (e) => openChartSub(b) }
  ])));

  host.appendChild(menuBtn('排列', (b) => openMenu(b, [
    { title: '对齐到幻灯片' },
    { label: '左对齐', action: () => alignElements('left') },
    { label: '水平居中', action: () => alignElements('hcenter') },
    { label: '右对齐', action: () => alignElements('right') },
    { label: '顶对齐', action: () => alignElements('top') },
    { label: '垂直居中', action: () => alignElements('vcenter') },
    { label: '底对齐', action: () => alignElements('bottom') },
    'sep',
    { label: '置于顶层', action: () => zOrder('front') },
    { label: '置于底层', action: () => zOrder('back') },
    'sep',
    { label: '组合', shortcut: 'Ctrl+G', action: groupSelection },
    { label: '取消组合', shortcut: 'Ctrl+Shift+G', action: ungroupSelection }
  ])));

  host.appendChild(menuBtn('视图', (b) => openMenu(b, [
    { label: store.showGrid ? '✓ 网格线' : '网格线', action: () => { store.setView({ showGrid: !store.showGrid }); renderCanvas(); } },
    { label: store.snap ? '✓ 对齐吸附' : '对齐吸附', action: () => { store.setView({ snap: !store.snap }); } },
    'sep',
    { label: '适应屏幕', action: () => fitToScreen() },
    { label: '实际大小 (100%)', action: () => setZoomValue(1) },
    { label: '页面设置…', action: () => openPageSetup() },
    'sep',
    { label: '快捷键帮助', action: openShortcuts }
  ])));

  // 右侧主题与放映
  const right = h('div', { class: 'menubar-right' });
  const themeBtn = menuBtn('主题', (b) => openMenu(b, THEMES.map((t) => ({
    label: (store.doc.theme === t.id ? '✓ ' : '') + t.name,
    action: () => applyThemeAction(t.id, true)
  }))));
  const presentBtn = h('button', { class: 'btn-primary', text: '▶ 放映', onclick: (e) => { e.stopPropagation(); startPresent(store.slideIndex); } });
  undoBtn = h('button', { class: 'icon-btn', title: '撤销', text: '↶', onclick: () => doUndo() });
  redoBtn = h('button', { class: 'icon-btn', title: '重做', text: '↷', onclick: () => doRedo() });
  right.append(undoBtn, redoBtn, themeBtn, presentBtn);
  host.appendChild(right);
}

function buildQuickbar(host) {
  if (!host) return;
  const add = (label, fn, opts = {}) => {
    const b = h('button', { class: 'qbtn' + (opts.primary ? ' primary' : ''), html: label, title: opts.title });
    b.addEventListener('click', (e) => { e.stopPropagation(); fn(b); });
    host.appendChild(b);
    return b;
  };
  add('T', () => addElement(createTextElement({ text: '双击编辑文字', x: 220, y: 220 })), { title: '文本框' });
  add('▭', openShapePickerSub, { title: '形状' });
  add('🖼', insertImage, { title: '图片' });
  add('⊞', openTableSub, { title: '表格' });
  add('📊', openChartSub, { title: '图表' });
  add('♪', () => insertMedia('audio'), { title: '音频' });
  add('▶', () => insertMedia('video'), { title: '视频' });
}

function openShapePickerSub(anchor) {
  openShapePicker(anchor, (type, opt) => {
    addElement(createShapeElement(type, { center: true, ...(opt || {}) }), { center: true });
  });
}
function openTableSub(anchor) {
  const sizes = [[2, 2], [3, 3], [4, 4], [3, 5], [1, 5]];
  openMenu(anchor, sizes.map(([r, c]) => ({ label: `${r} 行 × ${c} 列`, action: () => addElement(createTableElement(r, c, { center: true }), { center: true }) })), { class: 'ctx-menu' });
}
function openChartSub(anchor) {
  openMenu(anchor, CHART_TYPES.map((t) => ({ label: t.name, action: () => addElement(createChartElement(t.value, { center: true }), { center: true }) })), { class: 'ctx-menu' });
}

async function insertImage() {
  const file = await pickFile('image/*');
  if (!file) return;
  const data = await readFileAsDataURL(file);
  addElement(createImageElement(data, { center: true }), { center: true });
}

/** 插入音频/视频：选文件 → 建元素（带占位封面） */
async function insertMedia(kind) {
  const accept = kind === 'video' ? 'video/*,.mp4,.m4v,.mov,.webm,.avi' : 'audio/*,.mp3,.m4a,.wav,.aac,.ogg,.wma';
  const file = await pickFile(accept);
  if (!file) return;
  const data = await readFileAsDataURL(file);
  const name = file.name.replace(/\.[^.]+$/, '');
  const ext = file.name.split('.').pop().toLowerCase();
  const el = kind === 'video'
    ? createVideoElement(data, { name, extension: ext, center: true })
    : createAudioElement(data, { name, extension: ext, center: true });
  addElement(el, { center: true });
}

function openPptx() {
  pickFile('.pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation').then((f) => { if (f) importPptxFile(f); });
}
function openJson() {
  pickFile('.json,application/json').then((f) => { if (f) loadJsonFile(f); });
}
