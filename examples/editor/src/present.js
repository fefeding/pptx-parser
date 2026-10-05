/**
 * 演示模式：全屏放映，支持键盘 / 点击翻页 + 过渡 + 动画播放
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
  // 隐藏页不参与放映：从 startIndex 起找第一个可见页；若之前无可见页则取其后最近的可见页
  const vis = doc.slides.map((s, i) => s.hidden ? -1 : i).filter((i) => i >= 0);
  if (!vis.length) return;
  let idx = vis.find((i) => i >= startIndex);
  if (idx === undefined) idx = vis[0];
  state = { index: idx, overlay, doc, animQueue: [], animStep: 0 };
  document.addEventListener('keydown', onKey);
  overlay.addEventListener('click', onClick);
  show(idx, null);
  store.setSlide(idx);
}

function show(i, prevDir) {
  const { overlay, doc } = state;
  state.index = clamp(i, 0, doc.slides.length - 1);
  const slide = doc.slides[state.index];
  const size = doc.slideSize;
  const trans = slide.transition;
  const transType = trans && trans.type !== 'none' ? trans.type : null;
  const transDur = trans ? (trans.duration || 800) : 0;

  // 旧帧（如果有过渡）
  const oldFrame = overlay.querySelector('.present-frame');

  const z = Math.min((window.innerWidth - 24) / size.width, (window.innerHeight - 24) / size.height);
  const left = Math.round((window.innerWidth - size.width * z) / 2);
  const top = Math.round((window.innerHeight - size.height * z) / 2);

  const frame = h('div', { class: 'present-frame' });
  renderSlideInto(frame, slide, doc, {});
  frame.style.position = 'absolute';
  frame.style.left = `${left}px`;
  frame.style.top = `${top}px`;
  frame.style.transform = `scale(${z})`;
  frame.style.transformOrigin = 'top left';

  if (transType && oldFrame) {
    // 两帧叠加做过渡
    applyTransition(oldFrame, frame, transType, transDur, prevDir);
  } else {
    if (oldFrame) oldFrame.remove();
    overlay.appendChild(frame);
  }
  // 记录当前帧引用，供动画系统使用（避免 querySelector 查到旧帧）
  state.frame = frame;

  // 导航栏
  const bar = overlay.querySelector('.present-bar');
  if (bar) bar.remove();
  const newBar = h('div', { class: 'present-bar' });
  newBar.appendChild(h('button', { class: 'present-nav', text: '‹', onclick: (e) => { e.stopPropagation(); prev(); } }));
  newBar.appendChild(h('span', { class: 'present-idx', text: `${state.index + 1} / ${doc.slides.length}` }));
  newBar.appendChild(h('button', { class: 'present-nav', text: '›', onclick: (e) => { e.stopPropagation(); next(); } }));
  newBar.appendChild(h('button', { class: 'present-exit', text: '✕ 退出', onclick: (e) => { e.stopPropagation(); exit(); } }));
  overlay.appendChild(newBar);

  // 准备动画队列
  setupAnimations(slide);
}

/** 过渡效果：旧帧退出 + 新帧进入 */
function applyTransition(oldFrame, newFrame, type, dur, dir) {
  const overlay = state.overlay;
  const ms = Math.min(dur, 2000);
  newFrame.style.opacity = '0';
  overlay.appendChild(newFrame);

  // 新帧进入
  const enterAnim = transitionEnter(type, ms, dir || 'r');
  if (enterAnim) {
    newFrame.style.animation = enterAnim;
    newFrame.style.animationFillMode = 'forwards';
    // 过渡结束后清除 animation，避免残留影响元素动画
    setTimeout(() => {
      newFrame.style.animation = '';
      newFrame.style.opacity = '1';
    }, ms + 60);
  } else {
    newFrame.style.transition = `opacity ${ms}ms ease`;
    requestAnimationFrame(() => { newFrame.style.opacity = '1'; });
  }

  // 旧帧退出
  const exitAnim = transitionExit(type, ms, dir || 'r');
  if (exitAnim) {
    oldFrame.style.animation = exitAnim;
    oldFrame.style.animationFillMode = 'forwards';
    setTimeout(() => oldFrame.remove(), ms + 50);
  } else {
    oldFrame.style.transition = `opacity ${ms}ms ease`;
    oldFrame.style.opacity = '0';
    setTimeout(() => oldFrame.remove(), ms + 50);
  }
}

function transitionEnter(type, ms, dir) {
  const map = {
    fade: `presFadeIn ${ms}ms ease`,
    wipe: `presWipeIn ${ms}ms ease`,
    push: dir === 'r' ? `presPushInR ${ms}ms ease` : dir === 'l' ? `presPushInL ${ms}ms ease` : `presPushInR ${ms}ms ease`,
    cover: dir === 'r' ? `presCoverR ${ms}ms ease` : `presCoverL ${ms}ms ease`,
    blinds: `presBlindsIn ${ms}ms ease`,
    split: `presSplitIn ${ms}ms ease`,
    zoom: `presZoomIn ${ms}ms ease`,
    fly: `presFlyIn ${ms}ms ease`,
    reveal: `presRevealIn ${ms}ms ease`,
    randomBar: `presRandomBarIn ${ms}ms ease`,
  };
  return map[type] || null;
}
function transitionExit(type, ms, dir) {
  const map = {
    fade: `presFadeOut ${ms}ms ease`,
    wipe: `presWipeOut ${ms}ms ease`,
    push: dir === 'r' ? `presPushOutL ${ms}ms ease` : `presPushOutR ${ms}ms ease`,
    cover: null,
    blinds: `presBlindsOut ${ms}ms ease`,
    split: `presSplitOut ${ms}ms ease`,
    zoom: `presZoomOut ${ms}ms ease`,
    fly: null,
    reveal: `presRevealOut ${ms}ms ease`,
    randomBar: `presRandomBarOut ${ms}ms ease`,
  };
  return map[type] || null;
}

/** 动画队列：按 trigger 分组播放 */
function setupAnimations(slide) {
  state.animQueue = [];
  state.animStep = 0;
  const anims = slide.animations || [];
  if (!anims.length) return;

  // 隐藏所有有动画的元素，等待触发
  const frame = state.frame;
  if (!frame) return;
  for (const a of anims) {
    const el = frame.querySelector(`[data-id="${a.target}"]`);
    if (el) {
      el.style.visibility = 'hidden';
      el.dataset.animPending = '1';
    }
  }
  state.animQueue = anims.slice();
  // 自动播放 withPrev / afterPrev 的第一组
  playNextAnim();
}

function playNextAnim() {
  if (!state || !state.animQueue.length) return;
  const a = state.animQueue[0];
  const trigger = (a.trigger && a.trigger.type) || 'onClick';
  if (trigger === 'onClick') return; // 等待点击
  // withPrev / afterPrev 自动播放
  if (trigger === 'afterPrev' && state.animStep > 0) {
    // 等一小段延迟
    setTimeout(() => doPlayAnim(), 200);
  } else {
    doPlayAnim();
  }
}

function doPlayAnim() {
  if (!state || !state.animQueue.length) return;
  const a = state.animQueue.shift();
  state.animStep++;
  const frame = state.frame;
  if (!frame) return;
  const el = frame.querySelector(`[data-id="${a.target}"]`);
  if (!el) { playNextAnim(); return; }
  el.style.visibility = '';
  delete el.dataset.animPending;
  const ms = Math.round((a.duration || 0.5) * 1000);
  const anim = animCSS(a, ms);
  if (anim) {
    el.style.animation = anim;
    el.style.animationFillMode = 'both';
  }
  // 播放完后尝试下一个
  setTimeout(() => playNextAnim(), ms + 100);
}

function animCSS(a, ms) {
  const cls = a.presetClass || 'entr';
  const type = a.type;
  const dir = a.direction || 'l';
  const key = `${cls}_${type}_${dir}`;
  const map = {
    'entr_flyIn_l': `animFlyInL ${ms}ms ease`,
    'entr_flyIn_r': `animFlyInR ${ms}ms ease`,
    'entr_flyIn_t': `animFlyInT ${ms}ms ease`,
    'entr_flyIn_b': `animFlyInB ${ms}ms ease`,
    'entr_fadeIn_l': `animFadeIn ${ms}ms ease`,
    'entr_fadeIn_r': `animFadeIn ${ms}ms ease`,
    'entr_fadeIn_t': `animFadeIn ${ms}ms ease`,
    'entr_fadeIn_b': `animFadeIn ${ms}ms ease`,
    'entr_wipeIn_l': `animWipeIn ${ms}ms ease`,
    'entr_wipeIn_r': `animWipeIn ${ms}ms ease`,
    'entr_wipeIn_t': `animWipeIn ${ms}ms ease`,
    'entr_wipeIn_b': `animWipeIn ${ms}ms ease`,
    'entr_zoomIn_l': `animZoomIn ${ms}ms ease`,
    'entr_zoomIn_r': `animZoomIn ${ms}ms ease`,
    'entr_zoomIn_t': `animZoomIn ${ms}ms ease`,
    'entr_zoomIn_b': `animZoomIn ${ms}ms ease`,
    'entr_riseUp_l': `animRiseUp ${ms}ms ease`,
    'entr_riseUp_r': `animRiseUp ${ms}ms ease`,
    'entr_riseUp_t': `animRiseUp ${ms}ms ease`,
    'entr_riseUp_b': `animRiseUp ${ms}ms ease`,
    'entr_bounceIn_l': `animBounceIn ${ms}ms ease`,
    'entr_bounceIn_r': `animBounceIn ${ms}ms ease`,
    'entr_bounceIn_t': `animBounceIn ${ms}ms ease`,
    'entr_bounceIn_b': `animBounceIn ${ms}ms ease`,
    'exit_flyOut_l': `animFlyOutL ${ms}ms ease`,
    'exit_flyOut_r': `animFlyOutR ${ms}ms ease`,
    'exit_flyOut_t': `animFlyOutT ${ms}ms ease`,
    'exit_flyOut_b': `animFlyOutB ${ms}ms ease`,
    'exit_fadeOut_l': `animFadeOut ${ms}ms ease`,
    'exit_fadeOut_r': `animFadeOut ${ms}ms ease`,
    'exit_fadeOut_t': `animFadeOut ${ms}ms ease`,
    'exit_fadeOut_b': `animFadeOut ${ms}ms ease`,
    'exit_wipeOut_l': `animWipeOut ${ms}ms ease`,
    'exit_wipeOut_r': `animWipeOut ${ms}ms ease`,
    'exit_wipeOut_t': `animWipeOut ${ms}ms ease`,
    'exit_wipeOut_b': `animWipeOut ${ms}ms ease`,
    'exit_zoomOut_l': `animZoomOut ${ms}ms ease`,
    'exit_zoomOut_r': `animZoomOut ${ms}ms ease`,
    'exit_zoomOut_t': `animZoomOut ${ms}ms ease`,
    'exit_zoomOut_b': `animZoomOut ${ms}ms ease`,
    'emph_pulse_l': `animPulse ${ms}ms ease`,
    'emph_pulse_r': `animPulse ${ms}ms ease`,
    'emph_pulse_t': `animPulse ${ms}ms ease`,
    'emph_pulse_b': `animPulse ${ms}ms ease`,
    'emph_shake_l': `animShake ${ms}ms ease`,
    'emph_shake_r': `animShake ${ms}ms ease`,
    'emph_shake_t': `animShake ${ms}ms ease`,
    'emph_shake_b': `animShake ${ms}ms ease`,
    'emph_flash_l': `animFlash ${ms}ms ease`,
    'emph_flash_r': `animFlash ${ms}ms ease`,
    'emph_flash_t': `animFlash ${ms}ms ease`,
    'emph_flash_b': `animFlash ${ms}ms ease`,
    'emph_grow_l': `animGrow ${ms}ms ease`,
    'emph_grow_r': `animGrow ${ms}ms ease`,
    'emph_grow_t': `animGrow ${ms}ms ease`,
    'emph_grow_b': `animGrow ${ms}ms ease`,
  };
  return map[key] || `animFadeIn ${ms}ms ease`;
}

function onClick(e) {
  if (e.target.closest('.present-exit') || e.target.closest('.present-nav')) return;
  // 超链接：内部跳转（run.href='#N' → a[data-slide-jump]）翻到对应页；外部链接放行浏览器默认行为
  const link = e.target.closest && e.target.closest('a[data-slide-jump]');
  if (link) {
    e.preventDefault();
    e.stopPropagation();
    const n = parseInt(link.dataset.slideJump, 10);
    if (!isNaN(n)) show(n - 1);
    return;
  }
  if (e.target.closest && e.target.closest('a[href^="http"]')) return;
  // 先尝试播放下一个 onClick 动画
  if (state && state.animQueue.length) {
    const a = state.animQueue[0];
    const trigger = (a.trigger && a.trigger.type) || 'onClick';
    if (trigger === 'onClick') { doPlayAnim(); return; }
  }
  next();
}
function visibleIndexes() { return state.doc.slides.map((s, i) => s.hidden ? -1 : i).filter((i) => i >= 0); }
function next() {
  const vis = visibleIndexes();
  const pos = vis.indexOf(state.index);
  if (pos >= 0 && pos < vis.length - 1) show(vis[pos + 1], 'r');
}
function prev() {
  const vis = visibleIndexes();
  const pos = vis.indexOf(state.index);
  if (pos > 0) show(vis[pos - 1], 'l');
}

function onKey(e) {
  if (!state) return;
  if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') { e.preventDefault(); next(); }
  else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { e.preventDefault(); prev(); }
  else if (e.key === 'Home') { const v = visibleIndexes(); if (v.length) show(v[0], 'r'); }
  else if (e.key === 'End') { const v = visibleIndexes(); if (v.length) show(v[v.length - 1], 'r'); }
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
