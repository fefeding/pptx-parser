/**
 * 画布交互：渲染、选择、拖拽/缩放/旋转、框选、双击编辑、键盘快捷键
 */
import { store } from './store.js';
import { renderSlideInto, elementRect, rotatedRect, effectMargin, disposeAllCharts, disposeDetachedCharts, absoluteElementRect } from './render.js';
import { h, clamp, unionBBox, rotatePoint, debounce } from './util.js';
import { parseBody, focusBody, blurBody } from './richtext.js';
import { cloneElement, nudge, deleteSelected, paste, copySelected, duplicateSelected, selectAll, groupSelection, ungroupSelection, findInDoc } from './actions.js';
import { openChartDialog, openTableDialog, openImagePicker, openMediaDialog } from './dialogs.js';
import { startPresent } from './present.js';
import { exportPptx, saveJson } from './io.js';

let DOM = {};
let drag = null;
let selLayer = null;

export function initCanvas(dom) {
  DOM = dom;
  DOM.stage.addEventListener('pointerdown', onPointerDown);
  DOM.stage.addEventListener('dblclick', onDblClick);
  DOM.stage.addEventListener('contextmenu', onContextMenu);
  window.addEventListener('keydown', onKeyDown);
  window.addEventListener('resize', () => { if (store.fitted) fitToScreen(); });
  store.on('groupEdit', () => updateSelection());
}

/* ======================= 渲染 ======================= */
export function renderCanvas() {
  const doc = store.doc;
  if (!doc) return;
  const slide = store.slide;
  const W = doc.slideSize.width, H = doc.slideSize.height;
  // 不再每次重绘都销毁全部图表：复用 ECharts 实例（见 renderChartEl），
  // 仅清理已脱离文档的实例（图表被删 / 切页）
  renderSlideInto(DOM.frame, slide, doc, {
    grid: store.showGrid,
    scale: store.zoom,
    editingId: store.editingId
  });
  disposeDetachedCharts();
  DOM.stage.style.width = `${W * store.zoom}px`;
  DOM.stage.style.height = `${H * store.zoom}px`;
  DOM.frame.style.width = `${W}px`;
  DOM.frame.style.height = `${H}px`;
  DOM.overlay.style.width = '100%';
  DOM.overlay.style.height = '100%';
  updateSelection();
  if (store.editingId) {
    const body = DOM.frame.querySelector(`[data-id="${store.editingId}"] .tb-body`);
    if (body) bindEditing(body);
  }
}

export function updateSelection() {
  const doc = store.doc;
  if (!doc) return;
  const W = doc.slideSize.width, H = doc.slideSize.height;
  const zoom = store.zoom;
  DOM.overlay.innerHTML = '';
  const els = store.selected();
  // 选中态打在元素节点上：内嵌原生控件（音视频/图表）据此放开 pointer-events
  // （见 styles.css 的 .el.is-sel .el-media video/audio），未选中时让画布拖拽优先
  if (DOM.frame) {
    const selSet = new Set(store.sel);
    for (const n of DOM.frame.querySelectorAll('.el.is-sel')) {
      if (!selSet.has(n.dataset.id)) n.classList.remove('is-sel');
    }
    for (const id of store.sel) {
      const n = DOM.frame.querySelector(`.el[data-id="${id}"]`);
      if (n) n.classList.add('is-sel');
    }
  }
  applyGroupEditVisual();
  if (!els.length) { selLayer = null; return; }

  const layer = h('div', {
    style: {
      position: 'absolute', left: '0', top: '0', width: `${W}px`, height: `${H}px`,
      transform: `scale(${zoom})`, transformOrigin: 'top left'
    }
  });
  selLayer = layer;
  const hs = 9 / zoom;
  const rotOffset = 20 / zoom;

  const drawBox = (rect, rotation, opts) => {
    const box = h('div', {
      class: 'sel-box' + (opts.group ? ' group-box' : ''),
      style: {
        left: `${rect.x}px`, top: `${rect.y}px`, width: `${rect.width}px`, height: `${rect.height}px`,
        transform: rotation ? `rotate(${rotation}deg)` : '',
        transformOrigin: 'center center',
        borderColor: opts.group ? '#1a73e8' : '#1a73e8',
        borderWidth: `${1 / zoom}px`
      }
    });
    if (opts.handles) {
      const dirs = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
      for (const d of dirs) {
        const handle = h('div', {
          class: `handle ${d}`,
          dataset: { handle: d, id: opts.id || '' },
          style: handleStyle(d, hs, rect)
        });
        box.appendChild(handle);
      }
      if (opts.rotatable) {
        const rot = h('div', {
          class: 'handle rot',
          dataset: { handle: 'rot', id: opts.id || '' },
          style: {
            left: '50%', top: `${-rotOffset}px`,
            width: `${hs + 2}px`, height: `${hs + 2}px`, marginLeft: `${-(hs + 2) / 2}px`, marginTop: `${-(hs + 2) / 2}px`
          }
        });
        box.appendChild(rot);
      }
    }
    if (opts.locked) {
      box.appendChild(h('div', { class: 'lock-badge', text: '🔒', style: { fontSize: `${11 / zoom}px` } }));
    }
    layer.appendChild(box);
    return box;
  };

  if (els.length === 1) {
    const el = els[0];
    const rect = rectOf(el);
    const m = effectMargin(el);
    rect.x -= m.left; rect.y -= m.top;
    rect.width += m.left + m.right;
    rect.height += m.top + m.bottom;
    drawBox(rect, el.rotation, { handles: !el.locked, rotatable: !el.locked, id: el.id, locked: el.locked, group: el.type === 'group' });
  } else {
    const expandedRects = [];
    for (const el of els) {
      const rect = rectOf(el);
      const m = effectMargin(el);
      rect.x -= m.left; rect.y -= m.top;
      rect.width += m.left + m.right;
      rect.height += m.top + m.bottom;
      expandedRects.push(rect);
      drawBox(rect, el.rotation, { group: false, handles: false, locked: el.locked });
    }
    const box = unionBBox(expandedRects);
    drawBox(box, 0, { group: true, handles: true, id: 'multi' });
  }
  DOM.overlay.appendChild(layer);
}

/* ======================= 组合编辑态视觉 ======================= */
function applyGroupEditVisual() {
  if (!DOM.frame) return;
  const gid = store.groupEdit;
  DOM.frame.classList.toggle('group-editing', !!gid);
  const nodes = DOM.frame.querySelectorAll('.el');
  for (const n of nodes) n.classList.remove('group-edit-active', 'dimmed');
  let hint = DOM.scroll ? DOM.scroll.querySelector('.group-edit-hint') : null;
  if (!gid) { if (hint) hint.remove(); return; }
  const g = DOM.frame.querySelector(`.el-group[data-id="${gid}"]`);
  if (g) {
    g.classList.add('group-edit-active');
    // 不淡化当前组合的祖先组合（嵌套组合时其顶层祖先会被一并淡化，需豁免）
    let anc = g.parentElement;
    const keep = new Set();
    while (anc && anc !== DOM.frame) { keep.add(anc); anc = anc.parentElement; }
    for (const n of DOM.frame.querySelectorAll('.el')) {
      if (n.parentElement === DOM.frame && n.dataset.id !== gid && !keep.has(n)) n.classList.add('dimmed');
    }
  }
  if (DOM.scroll && !hint) {
    hint = h('div', { class: 'group-edit-hint', text: '组合编辑中 · 双击空白处或按 Esc 退出' });
    DOM.scroll.appendChild(hint);
  }
}

function handleStyle(dir, hs, rect) {
  const pos = { left: '0', top: '0' };
  if (dir.includes('w')) pos.left = '0';
  else if (dir.includes('e')) pos.left = '100%';
  else pos.left = '50%';
  if (dir.includes('n')) pos.top = '0';
  else if (dir.includes('s')) pos.top = '100%';
  else pos.top = '50%';
  return {
    left: pos.left, top: pos.top,
    width: `${hs}px`, height: `${hs}px`,
    marginLeft: `${-hs / 2}px`, marginTop: `${-hs / 2}px`
  };
}

/* ======================= 坐标换算 ======================= */
function toSlide(e) {
  const rect = DOM.frame.getBoundingClientRect();
  const z = store.zoom || 1;
  return { x: (e.clientX - rect.left) / z, y: (e.clientY - rect.top) / z };
}

/** 元素外接矩形（自动换算到幻灯片绝对坐标，含祖先组合偏移；顶层元素等价于 elementRect） */
function rectOf(el) {
  const top = store.elements();
  return top.some((t) => t.id === el.id) ? elementRect(el) : absoluteElementRect(el, top);
}

/* ======================= 指针交互 ======================= */
function onPointerDown(e) {
  if (e.button !== 0) return;
  const editing = store.editingId;
  if (editing) {
    const node = DOM.frame.querySelector(`[data-id="${editing}"]`);
    if (node && node.contains(e.target)) return;   // 编辑中，交给浏览器
    exitEditing();
    return;
  }
  const handleNode = e.target.closest && e.target.closest('[data-handle]');
  if (handleNode) {
    e.preventDefault();
    startTransform(e, handleNode.dataset.handle, handleNode.dataset.id);
    return;
  }
  const elNode = e.target.closest && e.target.closest('.el');
  if (elNode && elNode.dataset.id) {
    const gNode = elNode.closest('.el-group');
    const gid = gNode && gNode.dataset.id;
    let id = elNode.dataset.id;
    if (gid) {
      if (store.groupEdit === gid) {
        // 组合编辑态：选中实际点击的子元素（可单独移动/删除/复制）
        id = elNode.dataset.id;
      } else {
        // 普通态：点击组合内任何位置都选中整个组合
        id = gid;
      }
    }
    // 点击了其它顶层元素或不同组合：退出当前组合编辑态
    if (store.groupEdit && store.groupEdit !== gid) store.setGroupEdit(null);
    if (e.shiftKey) store.toggleSel(id);
    else if (!store.sel.includes(id)) store.setSel([id]);
    const el = store.findElement(id);
    if (!el || el.locked) return;
    startMove(e);
    return;
  }
  // 空白处：退出组合编辑态并框选
  if (store.groupEdit) store.setGroupEdit(null);
  if (!e.shiftKey) store.clearSel();
  startMarquee(e);
}

function onDblClick(e) {
  const elNode = e.target.closest && e.target.closest('.el');
  if (!elNode) {
    // 空白处双击：退出组合编辑态
    if (store.groupEdit) store.setGroupEdit(null);
    return;
  }
  const el = store.findElement(elNode.dataset.id);
  if (!el || el.locked) return;
  const gNode = elNode.closest('.el-group');
  const gid = gNode && gNode.dataset.id;
  if (gid) {
    // 双击组合（或其内部子元素）：进入组合编辑态，之后可单独选中/编辑子元素
    store.setGroupEdit(gid);
    // 双击文本子元素：进入组合编辑并直接编辑文本
    if (el.type === 'text') { enterEditing(el, { x: e.clientX, y: e.clientY }); return; }
    // 其余子元素：选中该子元素
    if (!store.sel.includes(el.id)) store.setSel([el.id]);
    return;
  }
  // 顶层元素（不在任何组合内）：原有行为
  if (el.type === 'text') {
    enterEditing(el, { x: e.clientX, y: e.clientY });
  } else if (el.type === 'table') {
    openTableDialog(el);
  } else if (el.type === 'chart') {
    openChartDialog(el);
  } else if (el.type === 'image') {
    openImagePicker(el);
  } else if (el.type === 'video' || el.type === 'audio') {
    openMediaDialog(el);
  }
}

function onContextMenu(e) {
  const elNode = e.target.closest && e.target.closest('.el');
  if (elNode && elNode.dataset.id) {
    let id = elNode.dataset.id;
    const gNode = elNode.closest('.el-group');
    if (gNode && gNode.dataset.id) id = gNode.dataset.id;
    if (!store.sel.includes(id)) store.setSel([id]);
  }
}

/* ---------- 移动 ---------- */
function startMove(e) {
  const p = toSlide(e);
  const ids = store.sel.slice();
  const els = store.selected().filter((el) => !el.locked);
  if (!els.length) return;
  store.snapshot();
  const orig = els.map((el) => ({ id: el.id, x: el.x, y: el.y, kids: (el.children || []).map((c) => ({ id: c.id, x: c.x, y: c.y })) }));
  const box = unionBBox(els.map(rectOf));
  const others = store.elements().filter((el) => !ids.includes(el.id)).map(elementRect);
  drag = { type: 'move', start: p, orig, box, others, moved: false };
  bindDrag();
}

/* ---------- 缩放 / 旋转 ---------- */
function startTransform(e, dir, id) {
  const p = toSlide(e);
  const els = id === 'multi' ? store.selected() : [store.findElement(id)].filter(Boolean);
  if (!els.length) return;
  store.snapshot();
  const single = els.length === 1 ? els[0] : null;
  const rect = single ? rectOf(single) : unionBBox(els.map(rectOf));
  const center = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
  const startAngle = Math.atan2(p.y - center.y, p.x - center.x) * 180 / Math.PI;
  drag = {
    type: dir === 'rot' ? 'rotate' : 'resize',
    dir, start: p, center, startAngle,
    base: single
      ? { id: single.id, x: single.x, y: single.y, w: single.width, h: single.height, rot: single.rotation || 0 }
      : null,
    items: els.map((el) => ({ id: el.id, ...rectOf(el) })),
    box: rect,
    rotation0: single ? (single.rotation || 0) : 0
  };
  bindDrag();
}

/* ---------- 框选 ---------- */
function startMarquee(e) {
  const p = toSlide(e);
  const rect = DOM.frame.getBoundingClientRect();
  drag = { type: 'marquee', start: p, rect, cur: p };
  bindDrag();
}

function bindDrag() {
  window.addEventListener('pointermove', onDragMove);
  window.addEventListener('pointerup', onDragUp, { once: true });
}

function onDragMove(e) {
  if (!drag) return;
  const p = toSlide(e);
  const shift = e.shiftKey;
  if (drag.type === 'marquee') {
    drag.cur = p;
    drawMarquee();
    return;
  }
  if (drag.type === 'move') {
    let dx = p.x - drag.start.x, dy = p.y - drag.start.y;
    if (store.snap) {
      const box = { x: drag.box.x + dx, y: drag.box.y + dy, width: drag.box.width, height: drag.box.height };
      const snap = snapBox(box, drag.others, store.doc.slideSize);
      dx += snap.dx; dy += snap.dy;
      showGuides(snap.guides);
    }
    drag.moved = true;
    store.update((doc) => {
      for (const o of drag.orig) {
        const el = findInDoc(doc, o.id);
        if (!el) continue;
        el.x = Math.round(o.x + dx);
        el.y = Math.round(o.y + dy);
        if (el.children) {
          o.kids.forEach((k, i) => {
            if (el.children[i]) { el.children[i].x = Math.round(k.x + dx); el.children[i].y = Math.round(k.y + dy); }
          });
        }
      }
    }, { history: false });
    return;
  }
  if (drag.type === 'resize') {
    applyResize(p, shift);
    return;
  }
  if (drag.type === 'rotate') {
    const ang = Math.atan2(p.y - drag.center.y, p.x - drag.center.x) * 180 / Math.PI;
    let delta = ang - drag.startAngle;
    let next = drag.rotation0 + delta;
    if (shift) next = Math.round(next / 15) * 15;
    next = ((next + 180) % 360) - 180;
    store.update((doc) => {
      const el = findInDoc(doc, drag.base.id);
      if (el) el.rotation = Math.round(next * 10) / 10;
    }, { history: false });
    return;
  }
}

function onDragUp() {
  window.removeEventListener('pointermove', onDragMove);
  clearGuides();
  if (drag && drag.type === 'marquee' && selLayer) {
    const a = drag.start, b = drag.cur;
    const box = {
      x: Math.min(a.x, b.x), y: Math.min(a.y, b.y),
      width: Math.abs(b.x - a.x), height: Math.abs(b.y - a.y)
    };
    if (box.width > 3 && box.height > 3) {
      const ids = store.elements().filter((el) => {
        if (el.locked || el.hidden) return false;
        const r = elementRect(el);
        return r.x < box.x + box.width && r.x + r.width > box.x && r.y < box.y + box.height && r.y + r.height > box.y;
      }).map((el) => el.id);
      store.setSel(ids);
    }
    const m = DOM.overlay.querySelector('.marquee');
    if (m) m.remove();
  }
  drag = null;
}

function drawMarquee() {
  if (!selLayer) {
    selLayer = h('div', {
      style: {
        position: 'absolute', left: '0', top: '0',
        width: `${store.doc.slideSize.width}px`, height: `${store.doc.slideSize.height}px`,
        transform: `scale(${store.zoom})`, transformOrigin: 'top left'
      }
    });
    DOM.overlay.appendChild(selLayer);
  }
  let m = selLayer.querySelector('.marquee');
  if (!m) { m = h('div', { class: 'marquee' }); selLayer.appendChild(m); }
  const a = drag.start, b = drag.cur;
  m.style.left = `${Math.min(a.x, b.x)}px`;
  m.style.top = `${Math.min(a.y, b.y)}px`;
  m.style.width = `${Math.abs(b.x - a.x)}px`;
  m.style.height = `${Math.abs(b.y - a.y)}px`;
}

function applyResize(p, shift) {
  const dx = p.x - drag.start.x, dy = p.y - drag.start.y;
  if (drag.base) {
    const b = drag.base;
    const rad = -(b.rot * Math.PI) / 180;
    const ldx = dx * Math.cos(rad) - dy * Math.sin(rad);
    const ldy = dx * Math.sin(rad) + dy * Math.cos(rad);
    const sx = drag.dir.includes('w') ? -1 : drag.dir.includes('e') ? 1 : 0;
    const sy = drag.dir.includes('n') ? -1 : drag.dir.includes('s') ? 1 : 0;
    let nw = b.w, nh = b.h;
    if (sx) nw = Math.max(8, b.w + sx * ldx);
    if (sy) nh = Math.max(8, b.h + sy * ldy);
    if (shift && sx && sy) {
      const s = Math.max(nw / b.w, nh / b.h);
      nw = b.w * s; nh = b.h * s;
    }
    const lox = sx * (nw - b.w) / 2, loy = sy * (nh - b.h) / 2;
    const rad2 = (b.rot * Math.PI) / 180;
    const pdx = lox * Math.cos(rad2) - loy * Math.sin(rad2);
    const pdy = lox * Math.sin(rad2) + loy * Math.cos(rad2);
    const cx = b.x + b.w / 2 + pdx, cy = b.y + b.h / 2 + pdy;
    store.update((doc) => {
      const el = findInDoc(doc, b.id);
      if (!el) return;
      if (el.type === 'group') {
        const fx = nw / b.w, fy = nh / b.h;
        for (const c of el.children || []) {
          c.x = b.x + (c.x - b.x) * fx;
          c.y = b.y + (c.y - b.y) * fy;
          c.width = Math.max(4, c.width * fx);
          c.height = Math.max(4, c.height * fy);
        }
      }
      el.width = Math.round(nw); el.height = Math.round(nh);
      el.x = Math.round(cx - nw / 2); el.y = Math.round(cy - nh / 2);
    }, { history: false });
    return;
  }
  // 多选：整体等比/自由缩放
  const box = drag.box;
  const sx = drag.dir.includes('w') ? -1 : drag.dir.includes('e') ? 1 : 0;
  const sy = drag.dir.includes('n') ? -1 : drag.dir.includes('s') ? 1 : 0;
  let nw = box.width + (sx ? sx * dx : 0);
  let nh = box.height + (sy ? sy * dy : 0);
  nw = Math.max(12, nw); nh = Math.max(12, nh);
  if (shift) {
    const s = Math.max(nw / box.width, nh / box.height);
    nw = box.width * s; nh = box.height * s;
  }
  const fx = nw / box.width, fy = nh / box.height;
  const ox = sx === -1 ? box.x + box.width - nw : box.x;
  const oy = sy === -1 ? box.y + box.height - nh : box.y;
  store.update((doc) => {
    for (const item of drag.items) {
      const el = findInDoc(doc, item.id);
      if (!el || el.locked) continue;
      if (el.type === 'group') {
        for (const c of el.children || []) {
          c.x = ox + (c.x - box.x) * fx;
          c.y = oy + (c.y - box.y) * fy;
          c.width = Math.max(4, c.width * fx);
          c.height = Math.max(4, c.height * fy);
        }
      }
      el.x = Math.round(ox + (item.x - box.x) * fx);
      el.y = Math.round(oy + (item.y - box.y) * fy);
      el.width = Math.round(Math.max(6, item.width * fx));
      el.height = Math.round(Math.max(6, item.height * fy));
    }
  }, { history: false });
}

/* ---------- 吸附参考线 ---------- */
function snapBox(box, others, size) {
  const tol = 6 / store.zoom;
  const candsX = [], candsY = [];
  for (const o of others) {
    candsX.push(o.x, o.x + o.width / 2, o.x + o.width);
    candsY.push(o.y, o.y + o.height / 2, o.y + o.height);
  }
  candsX.push(0, size.width / 2, size.width);
  candsY.push(0, size.height / 2, size.height);
  const selfX = [box.x, box.x + box.width / 2, box.x + box.width];
  const selfY = [box.y, box.y + box.height / 2, box.y + box.height];
  let bestX = null, bestY = null;
  for (const c of candsX) for (const s of selfX) {
    const d = c - s;
    if (Math.abs(d) <= tol && (!bestX || Math.abs(d) < Math.abs(bestX.d))) bestX = { d, pos: c };
  }
  for (const c of candsY) for (const s of selfY) {
    const d = c - s;
    if (Math.abs(d) <= tol && (!bestY || Math.abs(d) < Math.abs(bestY.d))) bestY = { d, pos: c };
  }
  const guides = [];
  if (bestX) guides.push({ type: 'v', pos: bestX.pos });
  if (bestY) guides.push({ type: 'h', pos: bestY.pos });
  return { dx: bestX ? bestX.d : 0, dy: bestY ? bestY.d : 0, guides };
}

function showGuides(guides) {
  clearGuides();
  if (!selLayer) return;
  for (const g of guides) {
    const node = h('div', { class: `guide ${g.type}` });
    if (g.type === 'v') node.style.left = `${g.pos}px`;
    else node.style.top = `${g.pos}px`;
    selLayer.appendChild(node);
  }
}
function clearGuides() {
  if (DOM.overlay) DOM.overlay.querySelectorAll('.guide').forEach((n) => n.remove());
}

/* ======================= 文本内联编辑 ======================= */
export function enterEditing(el, point) {
  if (store.editingId === el.id) return;
  store.editingId = el.id;
  renderCanvas();
  const body = DOM.frame.querySelector(`[data-id="${el.id}"] .tb-body`);
  if (!body) { store.editingId = null; return; }
  focusBody(body, point);
  bindEditing(body);
}

function bindEditing(body) {
  if (body.__bound) return;
  body.__bound = true;
  body.addEventListener('input', () => scheduleCommit(false));
  body.addEventListener('blur', () => scheduleCommit(true));
  body.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { e.preventDefault(); body.blur(); }
    e.stopPropagation();
  });
  body.addEventListener('pointerdown', (e) => e.stopPropagation());
}

const scheduleCommit = debounce((final) => commitEditing(final), 200);

export function exitEditing() {
  commitEditing(true);
}

function commitEditing(final) {
  const id = store.editingId;
  if (!id) return;
  const body = DOM.frame.querySelector(`[data-id="${id}"] .tb-body`);
  const el = store.findElement(id);
  if (body && el) {
    const defaults = {
      fontSize: el.fontSize || 18,
      color: el.color || '#202124',
      fontFace: el.fontFace || '微软雅黑'
    };
    const prev = el.paragraphs || [];
    const paras = parseBody(body, defaults).map((p, i) => ({
      runs: p.runs,
      align: (prev[i] && prev[i].align) || el.align || 'left',
      bullet: (prev[i] && prev[i].bullet) || undefined,
      lineSpacing: (prev[i] && prev[i].lineSpacing) || el.lineSpacing || 1.15
    }));
    store.update((doc) => {
      const t = doc.slides[store.slideIndex].elements.find((x) => x.id === id);
      if (t) t.paragraphs = paras;
    }, { mute: true, coalesce: 'text:' + id });
  }
  if (final) {
    store.editingId = null;
    blurBody(body);
    renderCanvas();
    store.emit('sel');
  }
}

/* ======================= 键盘 ======================= */
function isTypingTarget(t) {
  if (!t) return false;
  const tag = t.tagName;
  return tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || t.isContentEditable;
}

function onKeyDown(e) {
  if (isTypingTarget(e.target)) {
    if (e.key === 'Escape') e.target.blur();
    return;
  }
  const meta = e.ctrlKey || e.metaKey;
  const key = e.key;

  if (meta && key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? store.redo() : store.undo(); return; }
  if (meta && key.toLowerCase() === 'y') { e.preventDefault(); store.redo(); return; }
  if (meta && key.toLowerCase() === 'a') { e.preventDefault(); selectAll(); return; }
  if (meta && key.toLowerCase() === 'c') { e.preventDefault(); copySelected(false); return; }
  if (meta && key.toLowerCase() === 'x') { e.preventDefault(); copySelected(true); return; }
  if (meta && key.toLowerCase() === 'v') { e.preventDefault(); paste(); return; }
  if (meta && key.toLowerCase() === 'd') { e.preventDefault(); duplicateSelected(); return; }
  if (meta && key.toLowerCase() === 'g') {
    e.preventDefault();
    e.shiftKey ? ungroupSelection() : groupSelection();
    return;
  }
  if (meta && key.toLowerCase() === 's') { e.preventDefault(); saveJson(); return; }
  if (meta && key.toLowerCase() === 'e') { e.preventDefault(); exportPptx(); return; }

  if (key === 'F5') { e.preventDefault(); startPresent(0); return; }
  if (key === 'Delete' || key === 'Backspace') { e.preventDefault(); deleteSelected(); return; }
  if (key === 'Escape') {
    if (store.editingId) { exitEditing(); return; }
    if (store.groupEdit) { store.setGroupEdit(null); return; }
    store.clearSel();
    return;
  }
  if (key === 'Enter') {
    const els = store.selected();
    if (els.length === 1 && els[0].type === 'text') { e.preventDefault(); enterEditing(els[0]); }
    return;
  }
  if (key.startsWith('Arrow')) {
    if (!store.sel.length) return;
    e.preventDefault();
    const step = e.shiftKey ? 10 : 1;
    const dx = key === 'ArrowLeft' ? -step : key === 'ArrowRight' ? step : 0;
    const dy = key === 'ArrowUp' ? -step : key === 'ArrowDown' ? step : 0;
    nudge(dx, dy);
    return;
  }
  if (key === 'Tab') {
    e.preventDefault();
    if (store.groupEdit) store.setGroupEdit(null);
    const els = store.elements();
    if (!els.length) return;
    const idx = els.findIndex((el) => el.id === store.sel[0]);
    const next = els[(idx + (e.shiftKey ? -1 : 1) + els.length) % els.length];
    store.setSel([next.id]);
  }
}

/* ======================= 缩放 ======================= */
export function fitToScreen() {
  const doc = store.doc;
  if (!doc) return;
  const pad = 48;
  const availW = DOM.scroll.clientWidth - pad;
  const availH = DOM.scroll.clientHeight - pad;
  const z = Math.min(availW / doc.slideSize.width, availH / doc.slideSize.height);
  store.setZoom(clamp(z, 0.1, 4));
  store.fitted = true;
  renderCanvas();
}

export function zoomBy(factor) {
  store.setZoom(store.zoom * factor);
  store.fitted = false;
  renderCanvas();
}

export function setZoomValue(z) {
  store.setZoom(z);
  store.fitted = false;
  renderCanvas();
}
