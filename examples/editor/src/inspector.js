/**
 * 右侧属性面板：根据选中元素类型渲染对应属性控件
 */
import { store } from './store.js';
import { h, normalizeColor } from './util.js';
import { FONT_LIST, FONT_SIZES, getTheme, SHAPES, CHART_TYPES, THEMES, syncTableStyle, TRANSITIONS, TRANSITION_SPEEDS, ANIM_CLASSES, ANIM_TYPES, ANIM_DIRECTIONS, ANIM_TRIGGERS } from './model.js';
import {
  updateElement, applyTextStyleSel, setElementGeo, setBackground, setBackgroundImage,
  setNotes, alignElements, distribute, zOrder, groupSelection, ungroupSelection,
  toggleLock, toggleHidden, deleteSelected, duplicateSelected, resizeTable, applyTheme,
  setTransition, addAnimation, updateAnimation, removeAnimation, moveAnimation
} from './actions.js';
import {
  openPalette, openShapePicker, openChartDialog, openTableDialog, openImagePicker
} from './dialogs.js';
import { readFileAsDataURL } from './util.js';

let DOM = {};
export function initInspector(dom) {
  DOM = dom;
  store.on('sel', renderInspector);
  store.on('doc', renderInspector);
  store.on('slide', renderInspector);
  store.on('view', () => {});
  renderInspector();
}

function renderInspector() {
  const host = DOM.inspector;
  if (!host) return;
  const els = store.selected();
  host.innerHTML = '';
  if (!els.length) { renderDocPanel(host); return; }
  const single = els[0];
  // 头部
  const head = h('div', { class: 'insp-head' });
  head.appendChild(h('span', { class: 'insp-type', text: typeLabel(single.type) + (els.length > 1 ? ` ×${els.length}` : '') }));
  const tools = h('div', { class: 'insp-tools' });
  tools.appendChild(iconBtn(els.every((e) => !e.locked) ? '🔓' : '🔒', '锁定', (e) => { e.stopPropagation(); toggleLock(); }));
  tools.appendChild(iconBtn('👁', '隐藏', (e) => { e.stopPropagation(); toggleHidden(); }));
  tools.appendChild(iconBtn('⧉', '复制', (e) => { e.stopPropagation(); duplicateSelected(); }));
  tools.appendChild(iconBtn('🗑', '删除', (e) => { e.stopPropagation(); deleteSelected(); }));
  head.appendChild(tools);
  host.appendChild(head);

  // 位置与尺寸（仅单选时完整显示）
  host.appendChild(section('位置与尺寸', positionFields(single, els)));

  // 排列
  host.appendChild(section('排列', arrangeFields(els)));

  // 类型专属
  if (single.type === 'text') host.appendChild(section('文本', textFields(single, els)));
  else if (single.type === 'shape') host.appendChild(section('形状', shapeFields(single, els)));
  else if (single.type === 'image') host.appendChild(section('图片', imageFields(single, els)));
  else if (single.type === 'table') host.appendChild(section('表格', tableFields(single, els)));
  else if (single.type === 'chart') host.appendChild(section('图表', chartFields(single, els)));
  else if (single.type === 'group') host.appendChild(h('div', { class: 'hint', text: '已组合：可整体移动 / 缩放 / 取消组合。' }));

  // 选中元素的动画编辑
  if (els.length === 1) {
    const animSec = buildAnimEditPanel(single);
    if (animSec) host.appendChild(animSec);
  }
}

/* ---------- 通用控件 ---------- */
function section(title, controls) {
  const s = h('div', { class: 'insp-section' });
  s.appendChild(h('div', { class: 'insp-title', text: title }));
  const c = h('div', { class: 'insp-content' });
  for (const node of controls || []) if (node) c.appendChild(node);
  s.appendChild(c);
  return s;
}
function field(label, control) {
  const f = h('div', { class: 'field' });
  f.appendChild(h('label', { text: label }));
  f.appendChild(control);
  return f;
}
function numInput(value, onInput, opts = {}) {
  const i = h('input', { type: 'number', value: String(Math.round(value * 10) / 10) });
  if (opts.min != null) i.min = String(opts.min);
  if (opts.max != null) i.max = String(opts.max);
  if (opts.step) i.step = String(opts.step);
  i.style.width = opts.w || '72px';
  i.addEventListener('input', () => { const v = parseFloat(i.value); if (!isNaN(v)) onInput(v); });
  i.addEventListener('keydown', (e) => e.stopPropagation());
  return i;
}
function swatchBtn(value, onChange, allowNone = false) {
  const b = h('button', { class: 'swatch' });
  if (value && value !== 'none') b.style.background = normalizeColor(value) || value;
  else b.classList.add('none');
  b.dataset.popanchor = '1';
  b.onclick = (e) => { e.stopPropagation(); openPalette(b, { value: value === 'none' ? null : value, allowNone, onChange }); };
  return b;
}
function segBtn(options, current, onPick) {
  const wrap = h('div', { class: 'seg' });
  for (const o of options) {
    const b = h('button', { class: 'seg-b' + (o.value === current ? ' on' : ''), html: o.icon || o.label, title: o.label });
    b.onclick = (e) => { e.stopPropagation(); onPick(o.value); };
    wrap.appendChild(b);
  }
  return wrap;
}
function iconBtn(icon, title, on) {
  const b = h('button', { class: 'icon-sm', text: icon, title });
  b.onclick = on;
  return b;
}

/* ---------- 位置 ---------- */
function positionFields(single, els) {
  if (els.length > 1) return [h('div', { class: 'hint', text: '多选时仅可统一调整宽高。' })];
  const ids = els.map((e) => e.id);
  const mk = (label, key, step) => field(label, numInput(single[key] || 0, (v) => ids.forEach((id) => setElementGeo(id, { [key]: v })), { step: step || 1, min: 0 }));
  return [
    row2(mk('X', 'x'), mk('Y', 'y')),
    row2(mk('宽', 'width', 1), mk('高', 'height', 1)),
    row2(mk('旋转', 'rotation'), field('', h('span')))
  ];
}
function row2(a, b) { const r = h('div', { class: 'frow' }); r.appendChild(a); r.appendChild(b); return r; }

/* ---------- 排列 ---------- */
function arrangeFields(els) {
  const alignOpts = [
    { label: '左对齐', icon: '⇤', value: 'left' },
    { label: '水平居中', icon: '↔', value: 'hcenter' },
    { label: '右对齐', icon: '⇥', value: 'right' },
    { label: '顶对齐', icon: '⇞', value: 'top' },
    { label: '垂直居中', icon: '↕', value: 'vcenter' },
    { label: '底对齐', icon: '⇟', value: 'bottom' }
  ];
  const align = segBtn(alignOpts, '', (v) => alignElements(v));
  const z = h('div', { class: 'btn-row' });
  ['置于顶层', '上移一层', '下移一层', '置于底层'].forEach((t, i) => {
    const m = ['front', 'forward', 'backward', 'back'];
    const b = h('button', { class: 'mini-btn', text: t, onclick: (e) => { e.stopPropagation(); zOrder(m[i]); } });
    z.appendChild(b);
  });
  const g = h('div', { class: 'btn-row' });
  const hasGroup = els.some((e) => e.type === 'group');
  g.appendChild(h('button', { class: 'mini-btn', text: hasGroup ? '取消组合' : '组合', onclick: (e) => { e.stopPropagation(); hasGroup ? ungroupSelection() : groupSelection(); } }));
  g.appendChild(h('button', { class: 'mini-btn', text: '水平分布', onclick: (e) => { e.stopPropagation(); distribute('h'); } }));
  g.appendChild(h('button', { class: 'mini-btn', text: '垂直分布', onclick: (e) => { e.stopPropagation(); distribute('v'); } }));
  return [field('对齐', align), field('层级', z), field('组合', g)];
}

/* ---------- 文本 ---------- */
function textFields(single, els) {
  const ids = els.map((e) => e.id);
  const fontSel = h('select', { onchange: (e) => { applyTextStyleSel({ fontFace: e.target.value }); } });
  for (const f of FONT_LIST) fontSel.appendChild(h('option', { value: f, text: f }));
  fontSel.value = single.fontFace || '微软雅黑';

  const sizeSel = h('select', { onchange: (e) => { applyTextStyleSel({ fontSize: parseFloat(e.target.value) }); } });
  for (const s of FONT_SIZES) sizeSel.appendChild(h('option', { value: String(s), text: String(s) }));
  if (!FONT_SIZES.includes(single.fontSize)) sizeSel.appendChild(h('option', { value: String(single.fontSize), text: String(single.fontSize) }));
  sizeSel.value = String(single.fontSize);

  const biu = h('div', { class: 'seg' });
  ['B', 'I', 'U'].forEach((ch, i) => {
    const key = ['bold', 'italic', 'underline'][i];
    const cur = single[key];
    const b = h('button', { class: 'seg-b' + (cur ? ' on' : ''), text: ch, style: i === 0 ? 'font-weight:700' : i === 1 ? 'font-style:italic' : 'text-decoration:underline' });
    b.onclick = (e) => { e.stopPropagation(); applyTextStyleSel({ [key]: !store.findElement(ids[0])[key] }); };
    biu.appendChild(b);
  });

  const align = segBtn([
    { label: '左对齐', icon: '⯇', value: 'left' },
    { label: '居中', icon: '≡', value: 'center' },
    { label: '右对齐', icon: '⯈', value: 'right' },
    { label: '两端对齐', icon: '☰', value: 'justify' }
  ], single.align, (v) => applyTextStyleSel({ align: v }));

  const valign = segBtn([
    { label: '顶', icon: '⇞', value: 'top' },
    { label: '中', icon: '↕', value: 'middle' },
    { label: '底', icon: '⇟', value: 'bottom' }
  ], single.valign, (v) => applyTextStyleSel({ valign: v }));

  const ls = numInput(single.lineSpacing || 1.15, (v) => applyTextStyleSel({ lineSpacing: v }), { step: 0.05, min: 0.5, w: '64px' });
  const bullet = h('input', { type: 'checkbox', onchange: (e) => applyTextStyleSel({ bullet: e.target.checked }) });
  bullet.checked = !!single.bullet;

  return [
    row2(field('字体', fontSel), field('字号', sizeSel)),
    field('颜色', swatchBtn(single.color, (c) => applyTextStyleSel({ color: c }), false)),
    field('样式', biu),
    field('对齐', align),
    field('垂直', valign),
    row2(field('行距', ls), field('项目符号', wrapCheck(bullet)))
  ];
}
function wrapCheck(input) { const w = h('label', { class: 'chk' }); w.appendChild(input); return w; }

/* ---------- 形状 ---------- */
function shapeFields(single, els) {
  const ids = els.map((e) => e.id);
  const fill = single.fill && single.fill !== 'none'
    ? (typeof single.fill === 'string' ? single.fill : single.fill.color)
    : null;
  const fillRow = field('填充', swatchBtn(fill || 'none', (c) => {
    ids.forEach((id) => updateElement(id, { fill: c ? { type: 'solid', color: c, transparency: 0 } : 'none' }));
  }, true));
  const trans = slider('透明度', (single.fill && single.fill.transparency) || 0, (v) => {
    ids.forEach((id) => {
      const el = store.findElement(id);
      const base = el.fill && el.fill !== 'none' ? el.fill : { type: 'solid', color: '#4285F4' };
      updateElement(id, { fill: { ...base, transparency: Math.round(v) } });
    });
  }, 100);

  const lineVal = single.line && single.line !== 'none' ? single.line.color : null;
  const lineRow = field('描边', swatchBtn(lineVal || 'none', (c) => {
    ids.forEach((id) => updateElement(id, { line: c ? { color: c, width: 1, dashType: 'solid' } : 'none' }));
  }, true));
  const lw = h('select', { onchange: (e) => ids.forEach((id) => {
    const el = store.findElement(id);
    const cur = (el.line && el.line !== 'none') ? el.line : { color: '#000000', width: 1 };
    updateElement(id, { line: { ...cur, width: parseFloat(e.target.value) } });
  }) });
  [0.75, 1, 1.5, 2.25, 3, 4.5].forEach((w) => lw.appendChild(h('option', { value: String(w), text: w + 'pt' })));
  lw.value = String((single.line && single.line.width) || 1);
  const lwRow = field('描边粗细', lw);

  const shadow = h('input', { type: 'checkbox', onchange: (e) => ids.forEach((id) => updateElement(id, { shadow: e.target.checked ? { type: 'outer', angle: 45, distance: 4, blur: 8, transparency: 60, color: '#000000' } : null })) });
  shadow.checked = !!single.shadow;

  const shapeBtn = h('button', { class: 'mini-btn', text: '更改形状', onclick: (e) => {
    e.stopPropagation();
    openShapePicker(e.currentTarget, (type) => ids.forEach((id) => updateElement(id, { shapeType: type })));
  } });

  return [fillRow, trans, lineRow, lwRow, field('阴影', wrapCheck(shadow)), field('', shapeBtn)];
}

/* ---------- 图片 ---------- */
function imageFields(single, els) {
  const ids = els.map((e) => e.id);
  const adj = (k) => slider(k === 'brightness' ? '亮度' : k === 'contrast' ? '对比度' : '透明度',
    (single.imageAdjust && single.imageAdjust[k]) || 0,
    (v) => ids.forEach((id) => {
      const el = store.findElement(id);
      const a = { ...(el.imageAdjust || {}), [k]: Math.round(v) };
      updateElement(id, { imageAdjust: a });
    }), 100, -100);
  const replace = h('button', { class: 'mini-btn', text: '替换图片', onclick: (e) => { e.stopPropagation(); openImagePicker(single); } });
  return [field('', replace), adj('brightness'), adj('contrast'), adj('transparency')];
}

/* ---------- 表格 ---------- */
function tableFields(single, els) {
  const id = single.id;
  const rows = h('input', { type: 'number', value: String(single.rows.length), min: '1', max: '30', style: { width: '64px' } });
  const cols = h('input', { type: 'number', value: String(single.rows[0]?.cells.length || 1), min: '1', max: '20', style: { width: '64px' } });
  const apply = () => resizeTable(single, Math.max(1, Math.min(30, Number(rows.value) || 1)), Math.max(1, Math.min(20, Number(cols.value) || 1)));
  rows.onchange = apply; cols.onchange = apply;

  const header = h('input', { type: 'checkbox', onchange: (e) => { updateElement(id, { headerRow: e.target.checked }); const el = store.findElement(id); syncTableStyleOn(el); } });
  header.checked = !!single.headerRow;

  const border = field('边框颜色', swatchBtn(single.border?.color, (c) => updateElement(id, { border: { ...single.border, color: c || '#CBD5E1' } }), true));
  const bw = h('select', { onchange: (e) => updateElement(id, { border: { ...single.border, width: parseFloat(e.target.value) } }) });
  [0.5, 1, 1.5, 2, 3].forEach((w) => bw.appendChild(h('option', { value: String(w), text: w + 'pt' })));
  bw.value = String(single.border?.width || 1);
  const headerFill = field('表头底色', swatchBtn(single.headerFill, (c) => { updateElement(id, { headerFill: c }); syncTableStyleOn(store.findElement(id)); }, false));
  const cellFill = field('单元格底色', swatchBtn(single.cellFill, (c) => { updateElement(id, { cellFill: c }); syncTableStyleOn(store.findElement(id)); }, false));
  const fs = h('select', { onchange: (e) => { updateElement(id, { fontSize: parseFloat(e.target.value) }); syncTableStyleOn(store.findElement(id)); } });
  FONT_SIZES.forEach((s) => fs.appendChild(h('option', { value: String(s), text: String(s) })));
  fs.value = String(single.fontSize || 14);
  const color = field('文字颜色', swatchBtn(single.color, (c) => { updateElement(id, { color: c }); syncTableStyleOn(store.findElement(id)); }, false));

  const edit = h('button', { class: 'mini-btn', text: '编辑内容', onclick: (e) => { e.stopPropagation(); openTableDialog(single); } });
  return [
    row2(field('行数', rows), field('列数', cols)),
    field('表头行', wrapCheck(header)),
    row2(border, field('边框', bw)),
    row2(headerFill, cellFill),
    row2(field('字号', fs), color),
    field('', edit)
  ];
}
function syncTableStyleOn(el) {
  if (!el || el.type !== 'table') return;
  syncTableStyle(el);
}

/* ---------- 图表 ---------- */
function chartFields(single, els) {
  const id = single.id;
  const typeSel = h('select', { onchange: (e) => updateElement(id, { chartType: e.target.value }) });
  for (const t of CHART_TYPES) typeSel.appendChild(h('option', { value: t.value, text: t.name }));
  typeSel.value = single.chartType;
  const edit = h('button', { class: 'mini-btn', text: '编辑数据', onclick: (e) => { e.stopPropagation(); openChartDialog(single); } });
  const legend = h('input', { type: 'checkbox', onchange: (e) => updateElement(id, { legend: e.target.checked }) });
  legend.checked = !!single.legend;
  const dl = h('input', { type: 'checkbox', onchange: (e) => updateElement(id, { dataLabels: e.target.checked }) });
  dl.checked = !!single.dataLabels;
  return [
    field('类型', typeSel),
    field('', edit),
    field('图例', wrapCheck(legend)),
    field('数据标签', wrapCheck(dl))
  ];
}

/* ---------- 文档/幻灯片属性 ---------- */
function renderDocPanel(host) {
  const doc = store.doc, slide = store.slide;
  const theme = getTheme(doc.theme);

  host.appendChild(section('文档', [
    field('标题', (() => {
      const i = h('input', { type: 'text', value: doc.title || '', style: { width: '100%' } });
      i.addEventListener('input', () => store.update((d) => { d.title = i.value; }, { coalesce: 'title' }));
      i.addEventListener('keydown', (e) => e.stopPropagation());
      return i;
    })()),
    row2(buildThemeSelect(doc.theme), h('span'))
  ]));

  // 背景
  const bg = slide.background;
  const bgColor = bg && bg.type === 'solid' ? bg.color : bg && bg.type === 'gradient' ? null : (bg && bg.type === 'image' ? null : theme.bg);
  const bgSec = section('幻灯片背景', [
    field('颜色', swatchBtn(bgColor, (c) => setBackground(c ? { type: 'solid', color: c } : { type: 'solid', color: theme.bg }), true)),
    (() => {
      const b = h('button', { class: 'mini-btn', text: '使用图片背景', onclick: async (e) => {
        e.stopPropagation();
        const file = await pickBgImage();
        if (file) setBackgroundImage(await readFileAsDataURL(file));
      } });
      return field('', b);
    })(),
    (() => {
      const b = h('button', { class: 'mini-btn', text: '渐变背景', onclick: (e) => {
        e.stopPropagation();
        openPalette(e.currentTarget, { value: theme.accent, onChange: async (c1) => {
          const c2 = theme.accents[2] || '#34A853';
          setBackground({ type: 'gradient', direction: 'diagonal', stops: [{ color: c1, position: 0 }, { color: c2, position: 1 }] });
        } });
      } });
      return field('', b);
    })()
  ]);
  host.appendChild(bgSec);

  // 切换效果
  host.appendChild(buildTransitionSection(slide));

  // 动画
  host.appendChild(buildAnimationSection(slide));

  // 备注
  const notes = h('textarea', { class: 'notes-area', placeholder: '在此输入演讲者备注…' });
  notes.value = slide.notes || '';
  notes.addEventListener('input', () => setNotes(notes.value));
  notes.addEventListener('keydown', (e) => e.stopPropagation());
  host.appendChild(section('演讲者备注', [notes]));
  host.appendChild(h('div', { class: 'hint', text: '提示：双击文本 / 图表 / 表格可编辑内容，双击图片可替换。' }));
}

function buildThemeSelect(current) {
  const sel = h('select', { onchange: (e) => applyTheme(e.target.value, true) });
  for (const t of THEMES) sel.appendChild(h('option', { value: t.id, text: t.name }));
  sel.value = current;
  return field('主题', sel);
}

/* ---------- 切换效果 ---------- */
function buildTransitionSection(slide) {
  const trans = slide.transition || {};
  const typeSel = h('select', { onchange: (e) => {
    const type = e.target.value;
    if (type === 'none') { setTransition(null); return; }
    setTransition({ type, duration: trans.duration || 800, advanceOnClick: true });
  }});
  for (const t of TRANSITIONS) typeSel.appendChild(h('option', { value: t.value, text: t.name }));
  typeSel.value = trans.type || 'none';

  const speedSel = h('select', { onchange: (e) => {
    setTransition({ type: trans.type || 'fade', duration: parseInt(e.target.value), advanceOnClick: trans.advanceOnClick !== false });
  }});
  for (const s of TRANSITION_SPEEDS) speedSel.appendChild(h('option', { value: s.value, text: s.name }));
  speedSel.value = trans.duration || 800;

  const advCheck = h('input', { type: 'checkbox', onchange: (e) => {
    setTransition({ type: trans.type || 'fade', duration: trans.duration || 800, advanceOnClick: e.target.checked });
  }});
  advCheck.checked = trans.advanceOnClick !== false;

  return section('切换效果', [
    field('类型', typeSel),
    field('速度', speedSel),
    field('点击切换', wrapCheck(advCheck))
  ]);
}

/* ---------- 动画 ---------- */
function buildAnimationSection(slide) {
  const anims = slide.animations || [];
  const controls = [];

  // 动画列表
  for (let i = 0; i < anims.length; i++) {
    const a = anims[i];
    const el = store.findElement(a.target);
    const label = `${i + 1}. ${el ? typeLabel(el.type) : '元素'} · ${animTypeName(a)}`;
    const row = h('div', { class: 'anim-row' });
    row.appendChild(h('span', { class: 'anim-label', text: label }));
    const btns = h('div', { class: 'anim-btns' });
    btns.appendChild(h('button', { class: 'icon-sm', text: '↑', title: '上移', onclick: (e) => { e.stopPropagation(); moveAnimation(i, -1); } }));
    btns.appendChild(h('button', { class: 'icon-sm', text: '↓', title: '下移', onclick: (e) => { e.stopPropagation(); moveAnimation(i, 1); } }));
    btns.appendChild(h('button', { class: 'icon-sm', text: '🗑', title: '删除', onclick: (e) => { e.stopPropagation(); removeAnimation(i); } }));
    row.appendChild(btns);
    controls.push(row);
  }

  // 添加动画按钮
  const addBtn = h('button', { class: 'mini-btn', text: '＋ 为选中元素添加动画', onclick: (e) => {
    e.stopPropagation();
    const sel = store.selected();
    if (!sel.length) return;
    const el = sel[0];
    addAnimation({
      target: el.id, type: 'flyIn', presetClass: 'entr', duration: 0.5,
      direction: 'l', trigger: { type: 'onClick' }
    });
  }});
  controls.push(addBtn);

  if (!anims.length) {
    controls.push(h('div', { class: 'hint', text: '选中一个元素后点击上方按钮添加动画。' }));
  }

  return section('动画', controls);
}

function animTypeName(a) {
  const types = ANIM_TYPES[a.presetClass] || [];
  const found = types.find((t) => t.value === a.type);
  return found ? found.name : a.type;
}

/** 选中元素的动画属性编辑面板 */
function buildAnimEditPanel(el) {
  const slide = store.slide;
  const anims = slide.animations || [];
  const idx = anims.findIndex((a) => a.target === el.id);
  if (idx < 0) return null;
  const a = anims[idx];

  const classSel = h('select', { onchange: (e) => {
    const cls = e.target.value;
    const types = ANIM_TYPES[cls] || [];
    updateAnimation(idx, { presetClass: cls, type: types[0] ? types[0].value : a.type });
  }});
  for (const c of ANIM_CLASSES) classSel.appendChild(h('option', { value: c.value, text: c.name }));
  classSel.value = a.presetClass || 'entr';

  const typeSel = h('select', { onchange: (e) => updateAnimation(idx, { type: e.target.value }) });
  const types = ANIM_TYPES[a.presetClass] || [];
  for (const t of types) typeSel.appendChild(h('option', { value: t.value, text: t.name }));
  typeSel.value = a.type;

  const dirSel = h('select', { onchange: (e) => updateAnimation(idx, { direction: e.target.value }) });
  for (const d of ANIM_DIRECTIONS) dirSel.appendChild(h('option', { value: d.value, text: d.name }));
  dirSel.value = a.direction || 'l';

  const trigSel = h('select', { onchange: (e) => updateAnimation(idx, { trigger: { type: e.target.value } }) });
  for (const t of ANIM_TRIGGERS) trigSel.appendChild(h('option', { value: t.value, text: t.name }));
  trigSel.value = (a.trigger && a.trigger.type) || 'onClick';

  const durInp = numInput((a.duration || 0.5) * 1000, (v) => updateAnimation(idx, { duration: v / 1000 }), { step: 100, min: 100 });

  return section('动画属性', [
    field('类别', classSel),
    field('效果', typeSel),
    field('方向', dirSel),
    field('触发', trigSel),
    field('时长(ms)', durInp)
  ]);
}

async function pickBgImage() {
  return new Promise((resolve) => {
    const inp = h('input', { type: 'file', accept: 'image/*', style: { display: 'none' } });
    inp.onchange = () => resolve(inp.files[0] || null);
    document.body.appendChild(inp);
    inp.click();
    setTimeout(() => inp.remove(), 2000);
  });
}

/* ---------- 滑块 ---------- */
function slider(label, value, onInput, max = 100, min = 0) {
  const wrap = h('div', { class: 'slider-field' });
  const lab = h('label', { text: label });
  const val = h('span', { class: 'sv', text: Math.round(value) });
  const input = h('input', { type: 'range', min: String(min), max: String(max), value: String(Math.round(value)) });
  input.addEventListener('input', () => { val.textContent = Math.round(input.value); onInput(parseFloat(input.value)); });
  input.addEventListener('keydown', (e) => e.stopPropagation());
  wrap.appendChild(lab); wrap.appendChild(val); wrap.appendChild(input);
  return wrap;
}

function typeLabel(t) {
  return ({ text: '文本框', shape: '形状', image: '图片', table: '表格', chart: '图表', group: '组合', raw: '其他元素' })[t] || t;
}
