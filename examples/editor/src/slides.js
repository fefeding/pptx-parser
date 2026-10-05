/**
 * 左侧幻灯片缩略图面板
 */
import { store } from './store.js';
import { h } from './util.js';
import { renderThumbInto } from './render.js';
import { addSlide, duplicateSlide, deleteSlide, moveSlide, toggleSlideHidden } from './actions.js';
import { openMenu } from './dialogs.js';
import { applyLayout } from './actions.js';

let DOM = {};
export function initSlidePanel(dom) {
  DOM = dom;
  DOM.addBtn.addEventListener('click', (e) => openLayoutMenu(e.currentTarget));
  store.on('doc', renderSlideList);
  store.on('slide', renderSlideList);
  renderSlideList();
}

function openLayoutMenu(anchor) {
  const items = [
    { title: '新建幻灯片（选择版式）' }
  ];
  for (const l of LAYOUTS()) items.push({ label: l.name, action: () => addSlide(l.id) });
  openMenu(anchor, items);
}

function LAYOUTS() {
  return [
    { id: 'title', name: '标题页' },
    { id: 'titleBody', name: '标题 + 正文' },
    { id: 'titleOnly', name: '仅标题' },
    { id: 'section', name: '章节标题' },
    { id: 'twoCol', name: '两栏内容' },
    { id: 'comparison', name: '对比' },
    { id: 'quote', name: '引言' },
    { id: 'imageText', name: '图文混排' },
    { id: 'blank', name: '空白页' }
  ];
}

function renderSlideList() {
  const host = DOM.list;
  if (!host) return;
  host.innerHTML = '';
  store.doc.slides.forEach((slide, i) => {
    const item = h('div', { class: 'slide-item' + (i === store.slideIndex ? ' active' : '') + (slide.hidden ? ' hidden' : '') });
    item.addEventListener('click', () => store.setSlide(i));
    item.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      openMenu(e.currentTarget, [
        { label: '在此后新建', action: () => addSlide('titleBody') },
        { label: '复制幻灯片', action: () => duplicateSlide(i) },
        { label: '应用版式', action: () => openLayoutMenuFor(i) },
        'sep',
        { label: '上移', disabled: i === 0, action: () => moveSlide(i, i - 1) },
        { label: '下移', disabled: i === store.doc.slides.length - 1, action: () => moveSlide(i, i + 1) },
        'sep',
        { label: slide.hidden ? '显示幻灯片' : '隐藏幻灯片', action: () => toggleSlideHidden(i) },
        { label: '删除', danger: true, action: () => deleteSlide(i) }
      ], { class: 'ctx-menu' });
    });
    const num = h('div', { class: 'slide-num', text: String(i + 1) });
    const thumb = h('div', { class: 'slide-thumb' });
    renderThumbInto(thumb, slide, store.doc);
    if (slide.hidden) {
      const badge = h('div', {
        class: 'slide-hidden-badge',
        title: '已隐藏（放映时跳过）— 点击可显示',
        text: '⊘ 已隐藏',
        onclick: (e) => { e.stopPropagation(); toggleSlideHidden(i); }
      });
      item.appendChild(badge);
    }
    const bar = h('div', { class: 'slide-bar' });
    bar.appendChild(h('button', { class: 'mini-btn', text: '＋', title: '在此后新建', onclick: (e) => { e.stopPropagation(); addSlide('titleBody'); } }));
    bar.appendChild(h('button', { class: 'mini-btn', text: '⧉', title: '复制', onclick: (e) => { e.stopPropagation(); duplicateSlide(i); } }));
    bar.appendChild(h('button', {
      class: 'mini-btn', text: slide.hidden ? '👁' : '⊘', title: slide.hidden ? '显示幻灯片' : '隐藏幻灯片（放映时跳过）',
      onclick: (e) => { e.stopPropagation(); toggleSlideHidden(i); }
    }));
    bar.appendChild(h('button', { class: 'mini-btn', text: '🗑', title: '删除', onclick: (e) => { e.stopPropagation(); deleteSlide(i); } }));
    item.append(num, thumb, bar);
    host.appendChild(item);
  });
}

function openLayoutMenuFor(index) {
  const items = LAYOUTS().map((l) => ({ label: l.name, action: () => { store.setSlide(index, { force: true }); applyLayout(l.id); } }));
  openMenu(DOM.list.children[index] || DOM.addBtn, [{ title: '应用版式' }].concat(items), { class: 'ctx-menu' });
}
