/**
 * 文档操作集合（增删改 / 层级 / 对齐 / 组合 / 幻灯片管理）
 */
import { store } from './store.js';
import { clone, uid, unionBBox } from './util.js';
import { elementRect } from './render.js';
import { createSlide, buildSlideFromLayout, getTheme, syncTableStyle, createGroupElement, applyTextStyle } from './model.js';

/* ======================= 元素增删改 ======================= */
export function cloneElement(el) {
  const c = clone(el);
  c.id = uid();
  if (c.children) c.children = c.children.map(cloneElement);
  return c;
}

export function addElement(el, opts = {}) {
  store.update((doc) => {
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

export function deleteSelected() {
  const ids = store.sel.slice();
  if (!ids.length) return;
  store.update((doc) => {
    const slide = doc.slides[store.slideIndex];
    slide.elements = slide.elements.filter((e) => !ids.includes(e.id));
  });
  store.setSel([]);
}

export function duplicateSelected() {
  const els = store.selected();
  if (!els.length) return;
  const copies = els.map((el) => {
    const c = cloneElement(el);
    c.x += 16; c.y += 16;
    return c;
  });
  store.update((doc) => {
    doc.slides[store.slideIndex].elements.push(...copies);
  });
  store.setSel(copies.map((c) => c.id));
}

export function copySelected(cut = false) {
  const els = store.selected();
  if (!els.length) return;
  store.clipboard = els.map(clone);
  if (cut) deleteSelected();
}

export function paste() {
  const items = store.clipboard || [];
  if (!items.length) return;
  const copies = items.map((el) => {
    const c = cloneElement(el);
    c.x += 20; c.y += 20;
    return c;
  });
  store.update((doc) => {
    doc.slides[store.slideIndex].elements.push(...copies);
  });
  store.setSel(copies.map((c) => c.id));
}

export function selectAll() {
  store.setSel(store.elements().filter((e) => !e.locked).map((e) => e.id));
}

export function nudge(dx, dy) {
  const els = store.selected();
  if (!els.length) return;
  store.update((doc) => {
    for (const el of doc.slides[store.slideIndex].elements) {
      if (!store.sel.includes(el.id) || el.locked) continue;
      el.x += dx; el.y += dy;
      if (el.children) for (const c of el.children) { c.x += dx; c.y += dy; }
    }
  }, { coalesce: 'nudge' });
}

/* ======================= 层级 ======================= */
export function zOrder(op) {
  const ids = store.sel.slice();
  if (!ids.length) return;
  store.update((doc) => {
    const list = doc.slides[store.slideIndex].elements;
    const picked = list.filter((e) => ids.includes(e.id));
    const rest = list.filter((e) => !ids.includes(e.id));
    if (op === 'front') doc.slides[store.slideIndex].elements = rest.concat(picked);
    else if (op === 'back') doc.slides[store.slideIndex].elements = picked.concat(rest);
    else {
      const arr = list.slice();
      const idxs = ids.map((id) => arr.findIndex((e) => e.id === id)).sort((a, b) => a - b);
      if (op === 'forward') {
        for (let i = idxs.length - 1; i >= 0; i--) {
          const idx = idxs[i];
          if (idx < arr.length - 1 && !ids.includes(arr[idx + 1].id)) {
            [arr[idx], arr[idx + 1]] = [arr[idx + 1], arr[idx]];
          }
        }
      } else {
        for (let i = 0; i < idxs.length; i++) {
          const idx = idxs[i];
          if (idx > 0 && !ids.includes(arr[idx - 1].id)) {
            [arr[idx], arr[idx - 1]] = [arr[idx - 1], arr[idx]];
          }
        }
      }
      doc.slides[store.slideIndex].elements = arr;
    }
  });
}

/* ======================= 对齐 / 分布 ======================= */
export function alignElements(op) {
  const els = store.selected();
  if (!els.length) return;
  const slide = store.slide;
  const box = els.length > 1
    ? unionBBox(els.map(elementRect))
    : { x: 0, y: 0, width: store.doc.slideSize.width, height: store.doc.slideSize.height };
  store.update((doc) => {
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
function moveTo(el, x, y) {
  const dx = x == null ? 0 : x - el.x;
  const dy = y == null ? 0 : y - el.y;
  el.x += dx; el.y += dy;
  if (el.children) for (const c of el.children) { c.x += dx; c.y += dy; }
}

export function distribute(op) {
  const els = store.selected();
  if (els.length < 3) return;
  store.update((doc) => {
    const work = doc.slides[store.slideIndex].elements.filter((e) => store.sel.includes(e.id) && !e.locked);
    const rects = work.map((e) => ({ el: e, r: elementRect(e) }));
    if (op === 'h') {
      rects.sort((a, b) => a.r.x - b.r.x);
      const first = rects[0], last = rects[rects.length - 1];
      const span = (last.r.x + last.r.width) - first.r.x;
      const total = rects.reduce((s, it) => s + it.r.width, 0);
      const gap = (span - total) / (rects.length - 1);
      let cursor = first.r.x;
      for (const it of rects) {
        moveTo(it.el, cursor - (it.r.x - it.el.x), null);
        cursor += it.r.width + gap;
      }
    } else {
      rects.sort((a, b) => a.r.y - b.r.y);
      const first = rects[0], last = rects[rects.length - 1];
      const span = (last.r.y + last.r.height) - first.r.y;
      const total = rects.reduce((s, it) => s + it.r.height, 0);
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
export function groupSelection() {
  const els = store.selected();
  if (els.length < 2) return;
  const group = createGroupElement(els.map(clone));
  const rect = unionBBox(els.map(elementRect));
  group.x = rect.x; group.y = rect.y; group.width = rect.width; group.height = rect.height;
  const ids = els.map((e) => e.id);
  store.update((doc) => {
    const slide = doc.slides[store.slideIndex];
    const minIdx = Math.min(...ids.map((id) => slide.elements.findIndex((e) => e.id === id)));
    slide.elements = slide.elements.filter((e) => !ids.includes(e.id));
    slide.elements.splice(Math.min(minIdx, slide.elements.length), 0, group);
  });
  store.setSel([group.id]);
}

export function ungroupSelection() {
  const els = store.selected().filter((e) => e.type === 'group');
  if (!els.length) return;
  const newIds = [];
  store.update((doc) => {
    const slide = doc.slides[store.slideIndex];
    const out = [];
    for (const el of slide.elements) {
      if (els.some((g) => g.id === el.id)) {
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

export function toggleLock(force) {
  const els = store.selected();
  if (!els.length) return;
  const val = force === undefined ? !els[0].locked : force;
  store.update((doc, prevDoc) => {
    for (const el of doc.slides[store.slideIndex].elements) {
      if (store.sel.includes(el.id)) el.locked = val;
    }
  });
  if (val) store.setSel([]);
}

export function toggleHidden(force) {
  const els = store.selected();
  if (!els.length) return;
  const val = force === undefined ? !els[0].hidden : force;
  store.update((doc) => {
    for (const el of doc.slides[store.slideIndex].elements) {
      if (store.sel.includes(el.id)) el.hidden = val;
    }
  });
}

/* ======================= 幻灯片 ======================= */
export function addSlide(layoutId, opts = {}) {
  const doc = store.doc;
  const slide = buildSlideFromLayout(layoutId || 'titleBody', doc.theme, doc.slideSize);
  store.update((d) => {
    d.slides.splice(store.slideIndex + 1, 0, slide);
  });
  store.setSlide(store.slideIndex + 1);
  store.setSel([]);
  return slide;
}

export function duplicateSlide(index = store.slideIndex) {
  const copy = clone(store.doc.slides[index]);
  copy.id = uid('s');
  copy.elements = copy.elements.map(cloneElement);
  store.update((d) => { d.slides.splice(index + 1, 0, copy); });
  store.setSlide(index + 1);
}

export function deleteSlide(index = store.slideIndex) {
  if (store.doc.slides.length <= 1) return;
  store.update((d) => { d.slides.splice(index, 1); });
  store.setSlide(Math.min(index, store.doc.slides.length - 1), { force: true });
}

export function moveSlide(from, to) {
  if (from === to || from < 0 || to < 0) return;
  store.update((d) => {
    const [s] = d.slides.splice(from, 1);
    d.slides.splice(to, 0, s);
  });
  store.setSlide(to, { force: true });
}

/** 隐藏 / 显示幻灯片（隐藏页在放映时跳过，导入的 PPTX show="0" 也会带此标记） */
export function toggleSlideHidden(index = store.slideIndex) {
  const idx = Math.max(0, Math.min(index, store.doc.slides.length - 1));
  store.update((d) => { d.slides[idx].hidden = !d.slides[idx].hidden; });
}

export function applyLayout(layoutId) {
  const theme = getTheme(store.doc.theme);
  const elements = (store.doc, buildSlideFromLayout(layoutId, store.doc.theme, store.doc.slideSize).elements);
  store.update((doc) => {
    const slide = doc.slides[store.slideIndex];
    slide.elements = elements;
  }, { history: true });
  store.setSel([]);
}

export function setBackground(bg) {
  store.update((doc) => { doc.slides[store.slideIndex].background = bg; }, { coalesce: 'bg' });
}

export function setSlideSize(size) {
  store.update((doc) => { doc.slideSize = { ...size }; });
}

export function applyTheme(themeId, all = false) {
  const theme = getTheme(themeId);
  store.update((doc) => {
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

export function setNotes(text) {
  store.update((doc) => { doc.slides[store.slideIndex].notes = text; }, { coalesce: 'notes' });
}

/** 对选中文本元素应用样式（同步元素默认 + 所有 run），并同步到组内子元素 */
export function applyTextStyleSel(patch, opts = {}) {
  store.update((doc) => {
    for (const el of doc.slides[store.slideIndex].elements) {
      if (!store.sel.includes(el.id)) continue;
      if (el.type !== 'text') continue;
      applyTextStyle(el, patch);
      if (el.children) for (const c of el.children) if (c.type === 'text') applyTextStyle(c, patch);
    }
  }, opts);
}

export function setBackgroundImage(data) {
  store.update((doc) => { doc.slides[store.slideIndex].background = { type: 'image', data }; }, { coalesce: 'bg' });
}

/** 设置元素几何（支持组合：子元素同步移动/缩放） */
export function setElementGeo(id, geo) {
  store.update((doc) => {
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


export function resizeTable(el, rows, cols) {
  store.update((doc) => {
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
    target.rows.forEach((r) => { r.height = Math.round((el.height || 200) / Math.max(1, rows)); });
    syncTableStyle(target);
  });
}

function findInDoc(doc, id) {
  for (const el of doc.slides[store.slideIndex].elements) {
    if (el.id === id) return el;
    if (el.children) {
      const f = el.children.find((c) => c.id === id);
      if (f) return f;
    }
  }
  return null;
}

export function updateElement(id, patch, opts = {}) {
  store.update((doc) => {
    const el = findInDoc(doc, id);
    if (el) Object.assign(el, patch);
  }, opts);
}

export { findInDoc };

/* ======================= 过渡 / 动画 ======================= */
export function setTransition(trans) {
  store.update((doc) => {
    doc.slides[store.slideIndex].transition = trans;
  }, { coalesce: 'trans' });
}

export function addAnimation(anim) {
  store.update((doc) => {
    const slide = doc.slides[store.slideIndex];
    if (!slide.animations) slide.animations = [];
    slide.animations.push({ ...anim });
  });
}

export function updateAnimation(index, patch) {
  store.update((doc) => {
    const slide = doc.slides[store.slideIndex];
    if (slide.animations && slide.animations[index]) {
      Object.assign(slide.animations[index], patch);
    }
  });
}

export function removeAnimation(index) {
  store.update((doc) => {
    const slide = doc.slides[store.slideIndex];
    if (slide.animations) slide.animations.splice(index, 1);
  });
}

export function moveAnimation(index, dir) {
  store.update((doc) => {
    const slide = doc.slides[store.slideIndex];
    if (!slide.animations) return;
    const ni = index + dir;
    if (ni < 0 || ni >= slide.animations.length) return;
    const [a] = slide.animations.splice(index, 1);
    slide.animations.splice(ni, 0, a);
  });
}
