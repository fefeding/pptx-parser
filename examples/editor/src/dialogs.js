/**
 * 弹层组件：菜单 / 颜色面板 / 形状面板 / 模态框（图表数据、表格内容、页面设置等）
 */
import { h, $, clone, normalizeColor, pickFile, readFileAsDataURL, toast } from './util.js';
import { store } from './store.js';
import { SHAPES, CHART_TYPES, getTheme, SLIDE_SIZES, THEMES, syncTableStyle, defaultMediaPoster } from './model.js';
import { renderChartSVG } from './charts.js';
import { resizeTable, updateElement, setSlideSize, applyTheme } from './actions.js';

const popRoot = () => document.getElementById('popRoot');
const modalRoot = () => document.getElementById('modalRoot');

let closeCurrentPop = null;

function closePop() {
  if (closeCurrentPop) closeCurrentPop();
}

document.addEventListener('pointerdown', (e) => {
  if (e.target.closest && (e.target.closest('.pop') || e.target.closest('.ctx-menu'))) return;
  if (e.target.closest && e.target.closest('[data-popanchor]')) return;
  closePop();
}, true);

function place(node, anchor, opts = {}) {
  document.body.appendChild(node);
  const r = anchor.getBoundingClientRect();
  const w = node.offsetWidth, hh = node.offsetHeight;
  let left = opts.x != null ? opts.x : r.left;
  let top = opts.y != null ? opts.y : r.bottom + 4;
  if (left + w > window.innerWidth - 8) left = window.innerWidth - w - 8;
  if (top + hh > window.innerHeight - 8) top = Math.max(8, r.top - hh - 4);
  node.style.position = 'fixed';
  node.style.left = `${Math.max(4, left)}px`;
  node.style.top = `${top}px`;
}

/* ======================= 通用菜单 ======================= */
/**
 * @param {HTMLElement} anchor
 * @param {Array<{label,action,shortcut,disabled,danger}|'sep'|{title:string}>} items
 */
export function openMenu(anchor, items, opts = {}) {
  closePop();
  const node = h('div', { class: opts.class || 'pop' });
  anchor.dataset.popanchor = '1';
  for (const item of items) {
    if (item === 'sep') { node.appendChild(h('div', { class: 'pop-sep' })); continue; }
    if (item.title) { node.appendChild(h('div', { class: 'pop-title', text: item.title })); continue; }
    const btn = h('button', {
      class: 'pop-item' + (item.danger ? ' danger' : ''),
      onclick: () => { closePop(); if (!item.disabled && item.action) item.action(); }
    }, h('span', { text: item.label }), item.shortcut ? h('span', { class: 'k', text: item.shortcut }) : null);
    if (item.disabled) btn.disabled = true;
    node.appendChild(btn);
  }
  place(node, anchor, opts);
  const close = () => { node.remove(); delete anchor.dataset.popanchor; closeCurrentPop = null; };
  closeCurrentPop = close;
  return close;
}

/* ======================= 颜色面板 ======================= */
const SWATCHES = [
  '#000000', '#3C4043', '#5F6368', '#9AA0A6', '#DADCE0', '#FFFFFF',
  '#D93025', '#EA4335', '#F29900', '#F9AB00', '#FBBC04',
  '#188038', '#1E8E3E', '#34A853', '#1A73E8', '#4285F4',
  '#8430CE', '#A142F4', '#12B5CB', '#E8710A'
];

/**
 * @param {HTMLElement} anchor
 * @param {{value:string, onChange:Function, allowNone?:boolean}} opts
 */
export function openPalette(anchor, opts) {
  closePop();
  const theme = getTheme(store.doc.theme);
  const node = h('div', { class: 'pop', style: { minWidth: '236px' } });
  const grid = h('div', { class: 'palette', style: { gridTemplateColumns: 'repeat(10, 20px)' } });
  const list = theme.accents.concat(SWATCHES);
  for (const c of list) {
    grid.appendChild(h('button', {
      style: { background: c },
      title: c,
      onclick: () => { opts.onChange(c); closePop(); }
    }));
  }
  node.appendChild(grid);
  const row = h('div', { style: { display: 'flex', gap: '6px', alignItems: 'center', marginTop: '8px', padding: '0 4px' } });
  if (opts.allowNone) {
    row.appendChild(h('button', {
      class: 'mini-btn', text: '无',
      onclick: () => { opts.onChange(null); closePop(); }
    }));
  }
  const input = h('input', {
    type: 'color', value: normalizeColor(opts.value) || '#ffffff',
    style: { width: '40px', height: '28px', border: '1px solid #dadce0', borderRadius: '6px', padding: '0', background: 'none' },
    onchange: (e) => { opts.onChange(e.target.value); closePop(); }
  });
  row.appendChild(input);
  row.appendChild(h('button', {
    class: 'mini-btn', text: '自定义…',
    onclick: () => input.click()
  }));
  node.appendChild(row);
  place(node, anchor);
  const close = () => { node.remove(); closeCurrentPop = null; };
  closeCurrentPop = close;
}

/* ======================= 形状面板 ======================= */
const SHAPE_GROUPS = [
  { key: 'basic', label: '基本形状' },
  { key: 'line', label: '线条与连接符' },
  { key: 'arrow', label: '箭头连接符' },
  { key: 'flowchart', label: '流程图' }
];

export function openShapePicker(anchor, onPick) {
  closePop();
  const node = h('div', { class: 'pop shape-picker', style: { minWidth: '286px', maxHeight: '70vh', overflowY: 'auto' } });
  for (const g of SHAPE_GROUPS) {
    const items = SHAPES.filter((s) => s.group === g.key);
    if (!items.length) continue;
    node.appendChild(h('div', { class: 'pop-title', text: g.label }));
    const grid = h('div', { class: 'shape-grid' });
    for (const s of items) {
      const isLine = s.stroke || (s.type && /connector/.test(s.type));
      const inner = isLine
        ? `<svg viewBox="0 0 24 24"><path d="${s.d}" fill="none" stroke="#5f6368" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`
        : `<svg viewBox="0 0 24 24"><path d="${s.d}" fill="#5f6368" stroke="none"/></svg>`;
      grid.appendChild(h('button', {
        title: s.name,
        html: inner,
        onclick: () => { onPick(s.type, s.opt); closePop(); }
      }));
    }
    node.appendChild(grid);
  }
  place(node, anchor);
  closeCurrentPop = () => { node.remove(); closeCurrentPop = null; };
}

/* ======================= 模态框 ======================= */
export function openModal({ title, body, width, okText = '确定', cancelText = '取消', onOk, onClose, hideCancel }) {
  const mask = h('div', { class: 'modal-mask' });
  const modal = h('div', { class: 'modal', style: width ? { width } : {} });
  const close = () => { mask.remove(); document.removeEventListener('keydown', onKey); if (onClose) onClose(); };
  const onKey = (e) => {
    if (e.key === 'Escape') { e.stopPropagation(); close(); }
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); ok && ok(); }
  };
  const ok = () => {
    if (onOk && onOk() === false) return;
    close();
  };
  modal.appendChild(h('div', { class: 'modal-head' },
    h('h3', { text: title }),
    h('button', { class: 'icon-btn', text: '✕', onclick: close })
  ));
  const bodyEl = h('div', { class: 'modal-body' });
  if (typeof body === 'string') bodyEl.innerHTML = body;
  else if (body) bodyEl.appendChild(body);
  modal.appendChild(bodyEl);
  const foot = h('div', { class: 'modal-foot' });
  if (!hideCancel) foot.appendChild(h('button', { class: 'btn', text: cancelText, onclick: close }));
  if (okText) foot.appendChild(h('button', { class: 'btn btn-primary', text: okText, onclick: ok }));
  modal.appendChild(foot);
  mask.appendChild(modal);
  mask.addEventListener('pointerdown', (e) => { if (e.target === mask) close(); });
  modalRoot().appendChild(mask);
  document.addEventListener('keydown', onKey);
  const first = modal.querySelector('input,select,textarea');
  if (first) setTimeout(() => first.focus(), 30);
  return { close, modal, body: bodyEl };
}

export function openConfirm(message, onYes) {
  return openModal({
    title: '确认',
    body: h('div', { class: 'hint', text: message }),
    okText: '确定',
    onOk: () => { onYes && onYes(); }
  });
}

/* ======================= 图表数据编辑 ======================= */
export function openChartDialog(el) {
  const draft = clone(el);
  const theme = getTheme(store.doc.theme);
  const wrap = h('div', {});

  const typeSel = h('select', {
    onchange: (e) => { draft.chartType = e.target.value; refresh(); }
  });
  for (const t of CHART_TYPES) typeSel.appendChild(h('option', { value: t.value, text: t.name }));
  typeSel.value = draft.chartType;
  wrap.appendChild(h('div', { class: 'mfield' }, h('label', { class: 'f', text: '图表类型' }), typeSel));

  const grid = h('div', {});
  const preview = h('div', { class: 'chart-preview' });

  const buildGrid = () => {
    grid.innerHTML = '';
    const table = h('table', { class: 'grid-table' });
    const head = h('tr', {}, h('th', {}, h('input', { value: '类别', disabled: true })));
    draft.series.forEach((s, i) => {
      const inp = h('input', {
        value: s.name || `系列 ${i + 1}`,
        oninput: (e) => { s.name = e.target.value; refresh(); }
      });
      head.appendChild(h('th', {}, inp, h('button', {
        class: 'mini-btn', text: '✕', title: '删除该系列',
        style: { marginLeft: '2px', padding: '0 4px' },
        onclick: () => { draft.series.splice(i, 1); buildGrid(); refresh(); }
      })));
    });
    table.appendChild(head);
    draft.categories.forEach((c, ri) => {
      const tr = h('tr', {});
      tr.appendChild(h('td', {}, h('input', {
        value: c,
        oninput: (e) => { draft.categories[ri] = e.target.value; refresh(); }
      })));
      draft.series.forEach((s, si) => {
        tr.appendChild(h('td', {}, h('input', {
          type: 'number', value: String(s.values?.[ri] ?? 0),
          oninput: (e) => { s.values[ri] = Number(e.target.value) || 0; refresh(); }
        })));
      });
      tr.appendChild(h('td', { style: { width: '28px', border: 'none' } }, h('button', {
        class: 'mini-btn', text: '✕', title: '删除该类别',
        onclick: () => { draft.categories.splice(ri, 1); draft.series.forEach((s) => s.values.splice(ri, 1)); buildGrid(); refresh(); }
      })));
      table.appendChild(tr);
    });
    grid.appendChild(table);
    const ops = h('div', { style: { display: 'flex', gap: '8px', marginTop: '8px' } });
    ops.appendChild(h('button', {
      class: 'mini-btn', text: '＋ 添加系列',
      onclick: () => {
        draft.series.push({ name: `系列 ${draft.series.length + 1}`, values: draft.categories.map(() => Math.round(Math.random() * 50 + 10)) });
        buildGrid(); refresh();
      }
    }));
    ops.appendChild(h('button', {
      class: 'mini-btn', text: '＋ 添加类别',
      onclick: () => {
        draft.categories.push(`类别 ${draft.categories.length + 1}`);
        draft.series.forEach((s) => s.values.push(Math.round(Math.random() * 50 + 10)));
        buildGrid(); refresh();
      }
    }));
    grid.appendChild(ops);
  };

  const refresh = () => {
    preview.innerHTML = renderChartSVG(draft, { palette: theme.accents, textColor: theme.text, titleColor: theme.title });
  };
  buildGrid();
  refresh();

  const opts = h('div', { style: { display: 'flex', gap: '12px', marginTop: '10px', flexWrap: 'wrap' } });
  const mkCheck = (label, key) => h('label', { class: 'chk' }, (() => {
    const c = h('input', { type: 'checkbox', onchange: (e) => { draft[key] = e.target.checked; refresh(); } });
    c.checked = !!draft[key];
    return c;
  })(), label);
  opts.appendChild(mkCheck('显示图例', 'legend'));
  opts.appendChild(mkCheck('显示数据标签', 'dataLabels'));
  opts.appendChild(mkCheck('平滑曲线', 'smooth'));
  opts.appendChild(mkCheck('显示标记', 'marker'));

  wrap.appendChild(grid);
  wrap.appendChild(opts);
  wrap.appendChild(h('div', { class: 'mfield', style: { marginTop: '10px' } },
    h('label', { class: 'f', text: '图表标题' }),
    h('input', { type: 'text', value: draft.title || '', oninput: (e) => { draft.title = e.target.value; refresh(); } })
  ));
  wrap.appendChild(preview);

  openModal({
    title: '图表数据',
    width: '620px',
    body: wrap,
    onOk: () => {
      updateElement(el.id, {
        chartType: draft.chartType,
        categories: draft.categories,
        series: draft.series,
        legend: draft.legend,
        dataLabels: draft.dataLabels,
        smooth: draft.smooth,
        marker: draft.marker,
        title: draft.title
      });
    }
  });
}

/* ======================= 表格内容编辑 ======================= */
export function openTableDialog(el) {
  const draft = clone(el);
  const wrap = h('div', {});
  const grid = h('div', {});

  const build = () => {
    grid.innerHTML = '';
    const table = h('table', { class: 'grid-table' });
    draft.rows.forEach((row, ri) => {
      const tr = h('tr', {});
      row.cells.forEach((cell, ci) => {
        tr.appendChild(h('td', {}, h('input', {
          value: cell.text || '',
          oninput: (e) => { draft.rows[ri].cells[ci].text = e.target.value; }
        })));
      });
      table.appendChild(tr);
    });
    grid.appendChild(table);
  };
  build();

  const sizeRow = h('div', { style: { display: 'flex', gap: '10px', alignItems: 'center', marginBottom: '10px' } });
  const rowsInput = h('input', { type: 'number', value: String(draft.rows.length), min: '1', max: '30', style: { width: '80px' } });
  const colsInput = h('input', { type: 'number', value: String(draft.rows[0]?.cells.length || 1), min: '1', max: '20', style: { width: '80px' } });
  const applySize = () => {
    const rows = Math.max(1, Math.min(30, Number(rowsInput.value) || 1));
    const cols = Math.max(1, Math.min(20, Number(colsInput.value) || 1));
    const curCols = draft.rows[0]?.cells.length || 1;
    for (const r of draft.rows) {
      while (r.cells.length < cols) r.cells.push({ text: '', fill: null });
      if (r.cells.length > cols) r.cells.length = cols;
    }
    while (draft.rows.length < rows) {
      draft.rows.push({ height: draft.rows[0]?.height || 40, cells: new Array(cols).fill(0).map(() => ({ text: '', fill: null })) });
    }
    if (draft.rows.length > rows) draft.rows.length = rows;
    if (cols !== curCols) draft.colWidths = new Array(cols).fill(Math.round((el.width || 400) / cols));
    build();
  };
  rowsInput.onchange = applySize;
  colsInput.onchange = applySize;
  sizeRow.append(h('label', { class: 'chk' }, '行数', rowsInput), h('label', { class: 'chk' }, '列数', colsInput));

  wrap.appendChild(sizeRow);
  wrap.appendChild(grid);

  openModal({
    title: '表格内容',
    width: '640px',
    body: wrap,
    onOk: () => {
      updateElement(el.id, { rows: draft.rows, colWidths: draft.colWidths });
      resizeTable(el, draft.rows.length, draft.rows[0]?.cells.length || 1);
    }
  });
}

/* ======================= 图片选择 ======================= */
export async function openImagePicker(el) {
  const file = await pickFile('image/*');
  if (!file) return;
  const data = await readFileAsDataURL(file);
  if (el && el.id) updateElement(el.id, { data });
  else toast('请先选中一个图片元素');
}

/* ======================= 音视频 ======================= */

/** 媒体类型 → accept 过滤串（与 ppt-parser 的 MEDIA_MIME 对齐） */
const MEDIA_ACCEPT = {
  video: 'video/*,.mp4,.m4v,.mov,.webm,.avi',
  audio: 'audio/*,.mp3,.m4a,.wav,.aac,.ogg,.wma'
};

const mediaExt = (name) => String(name || '').split('.').pop().toLowerCase();

/** 替换媒体源（视频/音频文件） */
export async function openMediaPicker(el) {
  if (!el || !el.id) { toast('请先选中一个音频或视频元素'); return; }
  const file = await pickFile(MEDIA_ACCEPT[el.type] || '*/*');
  if (!file) return;
  const data = await readFileAsDataURL(file);
  updateElement(el.id, {
    data,
    extension: mediaExt(file.name),
    name: el.name || file.name.replace(/\.[^.]+$/, '')
  });
}

/** 替换海报/占位图 */
export async function openPosterPicker(el) {
  if (!el || !el.id) { toast('请先选中一个音频或视频元素'); return; }
  const file = await pickFile('image/*');
  if (!file) return;
  updateElement(el.id, {
    poster: { data: await readFileAsDataURL(file), extension: mediaExt(file.name) }
  });
}

/**
 * 音视频编辑对话框（双击打开）：预览、名称、替换媒体源、替换/恢复海报图。
 * 与图片的 openImagePicker 分开：媒体元素除媒体源外还有海报图与播放行为。
 */
export function openMediaDialog(el) {
  if (!el || (el.type !== 'video' && el.type !== 'audio')) return;
  const isVideo = el.type === 'video';
  const wrap = h('div', {});

  // 预览：视频可播；音频显示封面（原生控件在极小尺寸下不可用）
  const preview = h('div', {
    style: {
      background: '#0F172A', borderRadius: '8px', padding: '8px',
      display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: '120px', marginBottom: '10px'
    }
  });
  const buildPreview = () => {
    preview.innerHTML = '';
    const posterSrc = (el.poster && (el.poster.data || el.poster.src)) || defaultMediaPoster(el.type);
    if (isVideo) {
      const v = document.createElement('video');
      if (el.data || el.src) v.src = el.data || el.src;
      v.poster = posterSrc;
      v.controls = true;
      v.preload = 'metadata';
      v.style.cssText = 'width:100%;max-width:420px;border-radius:6px;background:#000';
      preview.appendChild(v);
    } else {
      const img = h('img', { src: posterSrc, style: { width: '96px', height: '96px', objectFit: 'contain' } });
      const a = document.createElement('audio');
      if (el.data || el.src) a.src = el.data || el.src;
      a.controls = true;
      a.style.cssText = 'margin-left:12px;max-width:280px';
      preview.append(img, a);
    }
  };
  buildPreview();

  const nameInput = h('input', { value: el.name || '', placeholder: isVideo ? '视频名称' : '音频名称' });

  const row = h('div', { style: { display: 'flex', gap: '8px', flexWrap: 'wrap' } });
  const replaceBtn = h('button', {
    class: 'btn', text: '替换媒体…',
    onclick: async (e) => { e.stopPropagation(); await openMediaPicker(el); buildPreview(); }
  });
  const posterBtn = h('button', {
    class: 'btn', text: '替换海报图…',
    onclick: async (e) => { e.stopPropagation(); await openPosterPicker(el); buildPreview(); }
  });
  const resetBtn = h('button', {
    class: 'btn', text: '恢复默认海报',
    onclick: (e) => {
      e.stopPropagation();
      updateElement(el.id, { poster: { data: defaultMediaPoster(el.type), extension: 'svg' } });
      buildPreview();
    }
  });
  row.append(replaceBtn, posterBtn, resetBtn);

  wrap.append(
    preview,
    h('div', { class: 'mfield' }, h('label', { class: 'f', text: '名称' }), nameInput),
    h('div', { class: 'mfield' }, h('label', { class: 'f', text: '媒体' }), row)
  );

  openModal({
    title: isVideo ? '编辑视频' : '编辑音频',
    body: wrap,
    width: '480px',
    onOk: () => {
      const v = nameInput.value.trim();
      if (v && v !== el.name) updateElement(el.id, { name: v });
    }
  });
}

/* ======================= 页面设置 ======================= */
export function openPageSetup() {
  const doc = store.doc;
  const wrap = h('div', {});
  const sel = h('select', {});
  for (const [k, v] of Object.entries(SLIDE_SIZES)) sel.appendChild(h('option', { value: k, text: `${k}（${v.width} × ${v.height}）` }));
  const curKey = Object.keys(SLIDE_SIZES).find((k) => SLIDE_SIZES[k].width === doc.slideSize.width && SLIDE_SIZES[k].height === doc.slideSize.height);
  sel.appendChild(h('option', { value: 'custom', text: '自定义' }));
  sel.value = curKey || 'custom';
  const wInput = h('input', { type: 'number', value: String(doc.slideSize.width), min: '200' });
  const hInput = h('input', { type: 'number', value: String(doc.slideSize.height), min: '200' });
  const themeSel = h('select', {});
  for (const t of THEMES) themeSel.appendChild(h('option', { value: t.id, text: t.name }));
  themeSel.value = doc.theme;
  sel.onchange = () => {
    const v = SLIDE_SIZES[sel.value];
    if (v) { wInput.value = String(v.width); hInput.value = String(v.height); }
  };
  wrap.append(
    h('div', { class: 'mfield' }, h('label', { class: 'f', text: '幻灯片尺寸' }), sel),
    h('div', { class: 'mfield', style: { display: 'flex', gap: '8px' } },
      h('div', { style: { flex: '1' } }, h('label', { class: 'f', text: '宽(px)' }), wInput),
      h('div', { style: { flex: '1' } }, h('label', { class: 'f', text: '高(px)' }), hInput)
    ),
    h('div', { class: 'mfield' }, h('label', { class: 'f', text: '主题配色' }), themeSel)
  );
  openModal({
    title: '页面设置',
    body: wrap,
    onOk: () => {
      setSlideSize({ width: Number(wInput.value) || 1280, height: Number(hInput.value) || 720 });
      applyTheme(themeSel.value, true);
    }
  });
}

/* ======================= 快捷键帮助 ======================= */
export function openShortcuts() {
  const list = [
    ['撤销', 'Ctrl / ⌘ + Z'], ['重做', 'Ctrl / ⌘ + Shift + Z'],
    ['复制 / 粘贴', 'Ctrl + C / V'], ['剪切', 'Ctrl + X'],
    ['再制', 'Ctrl + D'], ['全选', 'Ctrl + A'], ['删除', 'Delete'],
    ['组合', 'Ctrl + G'], ['取消组合', 'Ctrl + Shift + G'],
    ['保存 JSON', 'Ctrl + S'], ['导出 PPTX', 'Ctrl + E'],
    ['移动元素', '方向键（Shift 加速）'], ['进入文本编辑', '双击 / Enter'],
    ['退出编辑 / 取消选择', 'Esc'], ['播放演示', 'F5'],
    ['切换元素', 'Tab / Shift + Tab'], ['按比例缩放', '拖拽时按 Shift'],
    ['15° 步进旋转', '旋转时按 Shift']
  ];
  const grid = h('div', { class: 'shortcut-grid' });
  for (const [k, v] of list) {
    grid.appendChild(h('div', {}, h('span', { text: k }), h('span', { class: 'kbd', text: v })));
  }
  openModal({ title: '键盘快捷键', body: grid, hideCancel: true, okText: '知道了' });
}
