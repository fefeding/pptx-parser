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

/* ======================= 线条：端点 / 顶点 / 连接点 ======================= */
/** 是否可端点/顶点编辑的线条元素（连接符预设或开放自定义曲线） */
export function isLineElement(el) {
  if (!el || el.type !== 'shape') return false;
  if (/^(curvedConnector|bentConnector|straightConnector|line)/.test(el.shapeType || '')) return true;
  if (el.custGeom && Array.isArray(el.custGeom.paths)) {
    return el.custGeom.paths.every((p) => !p.closed);
  }
  return false;
}

/** 线条两个端点的幻灯片绝对坐标（考虑 flip；rotation 为 0 时即对角点） */
export function lineEndpoints(el) {
  const x = el.x || 0, y = el.y || 0, w = el.width || 0, h = el.height || 0;
  let a = { x: el.flipH ? x + w : x, y: el.flipV ? y + h : y };
  let b = { x: el.flipH ? x : x + w, y: el.flipV ? y : y + h };
  if (el.rotation) {
    const cx = x + w / 2, cy = y + h / 2, deg = -el.rotation;
    a = rotatePoint(a.x, a.y, cx, cy, deg);
    b = rotatePoint(b.x, b.y, cx, cy, deg);
  }
  return { a, b };
}

/** 标准连接点（site）：8 方位 + 中心，返回绝对坐标 */
export function connectionPoints(el) {
  const r = elementRect(el);
  const cx = r.x + r.width / 2, cy = r.y + r.height / 2;
  return [
    { site: 'top', x: cx, y: r.y },
    { site: 'bottom', x: cx, y: r.y + r.height },
    { site: 'left', x: r.x, y: cy },
    { site: 'right', x: r.x + r.width, y: cy },
    { site: 'topLeft', x: r.x, y: r.y },
    { site: 'topRight', x: r.x + r.width, y: r.y },
    { site: 'bottomLeft', x: r.x, y: r.y + r.height },
    { site: 'bottomRight', x: r.x + r.width, y: r.y + r.height },
    { site: 'center', x: cx, y: cy }
  ];
}

/** 根据 site 取某形状当前连接点绝对坐标 */
function connectionPointOf(el, site) {
  const c = connectionPoints(el).find((c) => c.site === site) || connectionPoints(el).find((c) => c.site === 'center');
  return { x: c.x, y: c.y };
}

/** 寻找离 p 最近、且在吸附阈值内的形状连接点（排除线条自身与被排除元素） */
function findGlueTarget(p, excludeId) {
  const tol = 14 / (store.zoom || 1);
  let best = null, bestD = tol;
  for (const el of store.elements()) {
    if (el.id === excludeId) continue;
    if (isLineElement(el) || el.type === 'group') continue; // 只吸附到形状/图片等实体
    for (const cp of connectionPoints(el)) {
      const d = Math.hypot(cp.x - p.x, cp.y - p.y);
      if (d <= bestD) { bestD = d; best = { shapeId: el.id, site: cp.site, point: { x: cp.x, y: cp.y } }; }
    }
  }
  return best;
}

/** 用两个绝对坐标端点重设线条包围盒与 flip（rotation 为 0 情形） */
export function setLineEndpoints(el, a, b) {
  const minX = Math.min(a.x, b.x), minY = Math.min(a.y, b.y);
  el.x = Math.round(minX); el.y = Math.round(minY);
  el.width = Math.round(Math.abs(b.x - a.x)); el.height = Math.round(Math.abs(b.y - a.y));
  el.flipH = b.x < a.x; el.flipV = b.y < a.y;
}

/** 被移动的形状集合变化后，重算所有 glue 到这些形状的线条端点（在同一 draft 内完成） */
function resyncGlue(doc, shapeIds) {
  const idx = store.slideIndex;
  for (const s of doc.slides[idx].elements) {
    if (!isLineElement(s)) continue;
    const gb = s.begin && shapeIds.includes(s.begin.shapeId);
    const ge = s.end && shapeIds.includes(s.end.shapeId);
    if (!gb && !ge) continue;
    const a = gb ? connectionPointOf(findInDoc(doc, s.begin.shapeId), s.begin.site) : lineEndpoints(s).a;
    const b = ge ? connectionPointOf(findInDoc(doc, s.end.shapeId), s.end.site) : lineEndpoints(s).b;
    setLineEndpoints(s, a, b);
  }
}

/** 选中线条时绘制端点手柄（顶点编辑模式额外绘制路径顶点，见 drawVertices） */
function drawLineHandles(layer, el, hs) {
  const { a, b } = lineEndpoints(el);
  const mk = (which, pt) => h('div', {
    class: `handle endpoint ${which}`,
    dataset: { handle: 'endpoint', which, id: el.id },
    style: {
      left: `${pt.x}px`, top: `${pt.y}px`,
      width: `${hs + 2}px`, height: `${hs + 2}px`,
      marginLeft: `${-(hs + 2) / 2}px`, marginTop: `${-(hs + 2) / 2}px`
    }
  });
  layer.appendChild(mk('start', a));
  layer.appendChild(mk('end', b));
}

/* ---------- 顶点编辑（自定义几何 custGeom） ---------- */
/** 将 line/connector 预设几何展开为可编辑的 custGeom（局部坐标 0..W,0..H，烘焙 flip） */
export function toVertexEditable(el) {
  const W = el.width || 100, H = el.height || 100;
  const { a, b } = lineEndpoints(el);
  // 直接用绝对端点减元素位置得到局部坐标：进入顶点编辑时会清零 rotation，
  // custGeom 渲染（rotation=0）时局部坐标 + 元素位置即还原旋转后的真实端点。
  const ax = a.x - el.x, ay = a.y - el.y;
  const bx = b.x - el.x, by = b.y - el.y;
  const st = el.shapeType || 'line';
  let commands;
  if (st.startsWith('bentConnector')) {
    const mx = (ax + bx) / 2;
    commands = [
      { type: 'moveTo', x: ax, y: ay },
      { type: 'lnTo', x: mx, y: ay },
      { type: 'lnTo', x: mx, y: by },
      { type: 'lnTo', x: bx, y: by }
    ];
  } else if (st.startsWith('curvedConnector')) {
    const mx = (ax + bx) / 2, my = (ay + by) / 2;
    const c1x = (ax + mx) / 2, c1y = ay;
    const c2x = (mx + bx) / 2, c2y = by;
    commands = [
      { type: 'moveTo', x: ax, y: ay },
      { type: 'quadBezTo', x1: c1x, y1: c1y, x: mx, y: my },
      { type: 'quadBezTo', x1: c2x, y1: c2y, x: bx, y: by }
    ];
  } else {
    commands = [
      { type: 'moveTo', x: ax, y: ay },
      { type: 'lnTo', x: bx, y: by }
    ];
  }
  return { w: W, h: H, closed: false, commands };
}

/** 进入顶点编辑：预设线条转为 custGeom，便于自由编辑节点 */
export function enterVertexEdit(el) {
  store.snapshot();
  store.update((doc) => {
    const t = findInDoc(doc, el.id);
    if (!t) return;
    const wasRotated = !!t.rotation;
    if (!t.custGeom) {
      t.custGeom = { paths: [toVertexEditable(t)] };   // toVertexEditable 已把 rotation 反算进局部坐标，端点位置保留
      t.shapeType = null;
      t.flipH = false; t.flipV = false;
    }
    if (wasRotated) t.rotation = 0;   // 几何已烘焙进 custGeom，清零 rotation（不影响视觉端点）
  });
  store.vertexEdit = el.id;
  store.emit('sel');
}

function exitVertexEdit() {
  if (!store.vertexEdit) return;
  store.vertexEdit = null;
  store.emit('sel');
}

/** 取 custGeom 第一路径的全部可拖点（顶点 + 贝塞尔控制点），返回绝对坐标 */
export function geomPoints(el) {
  const p = (el.custGeom && el.custGeom.paths && el.custGeom.paths[0]) || null;
  if (!p) return [];
  const W = el.width || 100, H = el.height || 100;
  const fx = el.flipH, fy = el.flipV;
  const toAbs = (lx, ly) => ({ x: el.x + (fx ? W - lx : lx), y: el.y + (fy ? H - ly : ly) });
  const pts = [];
  p.commands.forEach((c, i) => {
    if (c.type === 'moveTo' || c.type === 'lnTo') {
      pts.push({ cmdIndex: i, kind: 'point', abs: toAbs(c.x, c.y) });
    } else if (c.type === 'cubicBezTo') {
      pts.push({ cmdIndex: i, kind: 'c1', abs: toAbs(c.x1, c.y1) });
      pts.push({ cmdIndex: i, kind: 'c2', abs: toAbs(c.x2, c.y2) });
      pts.push({ cmdIndex: i, kind: 'point', abs: toAbs(c.x, c.y) });
    } else if (c.type === 'quadBezTo') {
      pts.push({ cmdIndex: i, kind: 'c1', abs: toAbs(c.x1, c.y1) });
      pts.push({ cmdIndex: i, kind: 'point', abs: toAbs(c.x, c.y) });
    }
  });
  return pts;
}

/** 顶点编辑模式：绘制顶点（方块）与贝塞尔控制点（圆点 + 虚线手柄连线） */
function drawVertices(layer, el, hs) {
  const pts = geomPoints(el);
  if (pts.length) {
    const svgNS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(svgNS, 'svg');
    svg.setAttribute('style', 'position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;pointer-events:none');
    for (const pt of pts) {
      if (pt.kind === 'c1' || pt.kind === 'c2') {
        const vp = pts.find((q) => q.cmdIndex === pt.cmdIndex && q.kind === 'point');
        if (vp) {
          const ln = document.createElementNS(svgNS, 'line');
          ln.setAttribute('x1', String(pt.abs.x)); ln.setAttribute('y1', String(pt.abs.y));
          ln.setAttribute('x2', String(vp.abs.x)); ln.setAttribute('y2', String(vp.abs.y));
          ln.setAttribute('stroke', '#1a73e8'); ln.setAttribute('stroke-width', String(1 / (store.zoom || 1)));
          ln.setAttribute('stroke-dasharray', '3 3');
          svg.appendChild(ln);
        }
      }
    }
    layer.appendChild(svg);
  }
  for (const pt of pts) {
    const isSel = store.vertexSel && store.vertexSel.cmd === pt.cmdIndex && store.vertexSel.kind === pt.kind;
    const cls = (pt.kind === 'point' ? 'handle vertex' : 'handle vctrl') + (isSel ? ' sel' : '');
    layer.appendChild(h('div', {
      class: cls,
      dataset: { handle: pt.kind === 'point' ? 'vertex' : 'control', cmd: String(pt.cmdIndex), kind: pt.kind, id: el.id },
      style: { left: `${pt.abs.x}px`, top: `${pt.abs.y}px`, width: `${hs}px`, height: `${hs}px`, marginLeft: `${-hs / 2}px`, marginTop: `${-hs / 2}px` }
    }));
  }
}

/** 顶点 / 控制点拖拽启动 */
function startVertexDrag(e, ds) {
  const el = store.findElement(ds.id);
  if (!el || el.locked) return;
  store.vertexSel = { cmd: Number(ds.cmd), kind: ds.kind };
  store.snapshot();
  drag = { type: 'vertex', id: ds.id, cmd: Number(ds.cmd), kind: ds.kind };
  bindDrag();
}

/* ---------- 顶点增删 ---------- */
function pointSegDist(px, py, ax, ay, bx, by) {
  const dx = bx - ax, dy = by - ay;
  const len2 = dx * dx + dy * dy;
  let t = len2 ? ((px - ax) * dx + (py - ay) * dy) / len2 : 0;
  t = Math.max(0, Math.min(1, t));
  const cx = ax + t * dx, cy = ay + t * dy;
  return Math.hypot(px - cx, py - cy);
}

/** 顶点编辑：在离 p 最近的直线段中点插入一个 lnTo 顶点 */
export function addVertexAt(el, p) {
  const cg = el.custGeom && el.custGeom.paths && el.custGeom.paths[0];
  if (!cg) return;
  const W = el.width || 100, H = el.height || 100, fx = el.flipH, fy = el.flipV;
  const lx = clamp(fx ? W - (p.x - el.x) : (p.x - el.x), 0, W);
  const ly = clamp(fy ? H - (p.y - el.y) : (p.y - el.y), 0, H);
  const cmds = cg.commands;
  let best = null, bestD = 1e9;
  for (let i = 0; i < cmds.length - 1; i++) {
    const a = cmds[i], b = cmds[i + 1];
    const aPt = (a.type === 'moveTo' || a.type === 'lnTo') ? { x: a.x, y: a.y } : null;
    const bPt = (b.type === 'moveTo' || b.type === 'lnTo') ? { x: b.x, y: b.y } : null;
    if (!aPt || !bPt) continue;
    const d = pointSegDist(lx, ly, aPt.x, aPt.y, bPt.x, bPt.y);
    if (d < bestD) { bestD = d; best = { i, aPt, bPt }; }
  }
  if (!best) return;
  store.update((doc) => {
    const t = findInDoc(doc, el.id); if (!t || !t.custGeom) return;
    const cmds2 = t.custGeom.paths[0].commands;
    const mx = (best.aPt.x + best.bPt.x) / 2, my = (best.aPt.y + best.bPt.y) / 2;
    cmds2.splice(best.i + 1, 0, { type: 'lnTo', x: mx, y: my });
  }, { history: false });
}

/** 顶点编辑：删除当前选中顶点（保留至少 2 个直线点，禁止删起点 moveTo） */
export function deleteSelectedVertex() {
  if (!store.vertexEdit) return;
  const sel = store.vertexSel; if (!sel) return;
  store.update((doc) => {
    const t = findInDoc(doc, store.vertexEdit); if (!t || !t.custGeom) return;
    const cmds = t.custGeom.paths[0].commands;
    // 统计所有顶点（moveTo + 各线条段终点），曲线终点也算，统一约束至少保留 2 个
    const pts = cmds.filter((c) => c.type === 'moveTo' || c.type === 'lnTo' || c.type === 'quadBezTo' || c.type === 'cubicBezTo').length;
    if (pts <= 2) return;
    const c = cmds[sel.cmd];
    if (!c || c.type === 'moveTo') return;
    cmds.splice(sel.cmd, 1);
  }, { history: false });
  store.vertexSel = null;
}

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
    const inVertex = store.vertexEdit === el.id && !!el.custGeom;
    drawBox(rect, el.rotation, { handles: !el.locked && !inVertex, rotatable: false, id: el.id, locked: el.locked, group: el.type === 'group' });
    if (!el.locked && isLineElement(el)) {
      if (inVertex) drawVertices(layer, el, hs);
      else drawLineHandles(layer, el, hs);
    }
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
    const hd = handleNode.dataset.handle;
    if (hd === 'endpoint') { startEndpointDrag(e, handleNode.dataset.which, handleNode.dataset.id); return; }
    if (hd === 'vertex' || hd === 'control') { startVertexDrag(e, handleNode.dataset); return; }
    startTransform(e, hd, handleNode.dataset.id);
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
    if (store.vertexEdit) { exitVertexEdit(); return; }
    // 空白处双击：退出组合编辑态
    if (store.groupEdit) store.setGroupEdit(null);
    return;
  }
  const el = store.findElement(elNode.dataset.id);
  if (!el || el.locked) return;
  // 线条：双击进入/保持顶点编辑（与 Office “编辑顶点”一致）
  if (isLineElement(el)) {
    if (store.vertexEdit !== el.id) { enterVertexEdit(el); return; }
    addVertexAt(el, toSlide(e));   // 已在顶点编辑：双击线段中点插入顶点
    return;
  }
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

/* ---------- 端点拖拽（改变起点/终点） ---------- */
function startEndpointDrag(e, which, id) {
  const el = store.findElement(id);
  if (!el || el.locked) return;
  store.snapshot();
  const { a, b } = lineEndpoints(el);
  const fixed = which === 'start' ? b : a;   // 另一端保持不动
  drag = { type: 'endpoint', which, id, fixed };
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
      // 连接线跟随：被移动的形状若被线条 glue，则重算线条端点
      const movedShapeIds = new Set();
      const all = drag.orig.map((o) => findInDoc(doc, o.id)).filter(Boolean);
      for (const e of all) {
        if (isLineElement(e) || e.type === 'group') continue;
        movedShapeIds.add(e.id);
      }
      // 移动组合时，其内部形状也视为被移动，需触发 glue 重算
      for (const e of all) {
        if (e.type === 'group') (e.children || []).forEach((c) => movedShapeIds.add(c.id));
      }
      if (movedShapeIds.size) resyncGlue(doc, [...movedShapeIds]);
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
  if (drag.type === 'endpoint') {
    let p = toSlide(e);
    const glue = findGlueTarget(p, drag.id);
    if (glue) p = glue.point;
    const other = drag.fixed;
    store.update((doc) => {
      const t = findInDoc(doc, drag.id);
      if (!t) return;
      setLineEndpoints(t, other, p);
      if (glue) {
        if (drag.which === 'start') t.begin = { shapeId: glue.shapeId, site: glue.site };
        else t.end = { shapeId: glue.shapeId, site: glue.site };
      } else {
        if (drag.which === 'start') t.begin = null; else t.end = null;
      }
    }, { history: false });
    return;
  }
  if (drag.type === 'vertex' || drag.type === 'control') {
    const p = toSlide(e);
    const el = store.findElement(drag.id);
    if (!el) return;
    const W = el.width || 100, H = el.height || 100;
    const fx = el.flipH, fy = el.flipV;
    const lx = clamp(fx ? W - (p.x - el.x) : (p.x - el.x), 0, W);
    const ly = clamp(fy ? H - (p.y - el.y) : (p.y - el.y), 0, H);
    store.update((doc) => {
      const t = findInDoc(doc, drag.id);
      if (!t || !t.custGeom) return;
      const c = t.custGeom.paths[0].commands[drag.cmd];
      if (!c) return;
      if (drag.kind === 'point') { c.x = lx; c.y = ly; }
      else if (drag.kind === 'c1') { c.x1 = lx; c.y1 = ly; }
      else if (drag.kind === 'c2') { c.x2 = lx; c.y2 = ly; }
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
  if (key === 'Delete' || key === 'Backspace') {
    e.preventDefault();
    if (store.vertexEdit) { deleteSelectedVertex(); return; }
    deleteSelected();
    return;
  }
  if (key === 'Escape') {
    if (store.editingId) { exitEditing(); return; }
    if (store.vertexEdit) { exitVertexEdit(); return; }
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
