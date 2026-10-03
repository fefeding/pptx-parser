/**
 * 演示模式：全屏放映，支持键盘 / 点击翻页
 */
import { store } from './store.js';
import { h, clamp } from './util.js';
import { renderSlideInto } from './render.js';

let state = null;

export function startPresent(startIndex = 0) {
  if (state) return;
  const doc = store.doc;
  const overlay = document.getElementById('presentRoot');
  if (!overlay) return;
  overlay.innerHTML = '';
  overlay.classList.add('on');
  const idx = clamp(startIndex, 0, doc.slides.length - 1);
  state = { index: idx, overlay, doc };
  document.addEventListener('keydown', onKey);
  overlay.addEventListener('click', onClick);
  show(idx);
  store.setSlide(idx);
}

function show(i) {
  const { overlay, doc } = state;
  state.index = clamp(i, 0, doc.slides.length - 1);
  const slide = doc.slides[state.index];
  const size = doc.slideSize;
  overlay.innerHTML = '';
  const frame = h('div', { class: 'present-frame' });
  renderSlideInto(frame, slide, doc, {});
  const z = Math.min((window.innerWidth - 24) / size.width, (window.innerHeight - 24) / size.height);
  frame.style.position = 'absolute';
  frame.style.left = `${Math.round((window.innerWidth - size.width * z) / 2)}px`;
  frame.style.top = `${Math.round((window.innerHeight - size.height * z) / 2)}px`;
  frame.style.transform = `scale(${z})`;
  frame.style.transformOrigin = 'top left';
  overlay.appendChild(frame);

  const bar = h('div', { class: 'present-bar' });
  bar.appendChild(h('button', { class: 'present-nav', text: '‹', onclick: (e) => { e.stopPropagation(); prev(); } }));
  bar.appendChild(h('span', { class: 'present-idx', text: `${state.index + 1} / ${doc.slides.length}` }));
  bar.appendChild(h('button', { class: 'present-nav', text: '›', onclick: (e) => { e.stopPropagation(); next(); } }));
  bar.appendChild(h('button', { class: 'present-exit', text: '✕ 退出', onclick: (e) => { e.stopPropagation(); exit(); } }));
  overlay.appendChild(bar);
}

function onClick(e) {
  if (e.target.closest('.present-exit') || e.target.closest('.present-nav')) return;
  next();
}
function next() { if (state.index < state.doc.slides.length - 1) show(state.index + 1); }
function prev() { if (state.index > 0) show(state.index - 1); }

function onKey(e) {
  if (!state) return;
  if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') { e.preventDefault(); next(); }
  else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { e.preventDefault(); prev(); }
  else if (e.key === 'Home') show(0);
  else if (e.key === 'End') show(state.doc.slides.length - 1);
  else if (e.key === 'Escape') exit();
}

export function exit() {
  if (!state) return;
  document.removeEventListener('keydown', onKey);
  const overlay = state.overlay;
  overlay.classList.remove('on');
  overlay.innerHTML = '';
  state = null;
}
