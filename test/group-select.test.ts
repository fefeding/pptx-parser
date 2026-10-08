// @vitest-environment jsdom
// 回归：组合后点击子元素应选中整个组合；delete/duplicate 兼容组合内子元素
import { describe, it, expect, beforeAll } from 'vitest';
import { createStore, createActions } from '../dist/ppt-parser.browser.js';

const HTML = `
<header><input id="docTitle"/><span id="saveState"></span></header>
<nav class="menubar" id="menubar"></nav>
<div class="toolbar" id="quickbar"></div>
<main class="workspace">
  <aside class="slides-panel"><div class="sp-list" id="slideList"></div><button id="addSlideBtn"></button></aside>
  <section class="canvas-area">
    <div class="canvas-scroll" id="canvasScroll"><div class="stage" id="stage">
      <div class="slide-frame" id="frame"></div><div class="overlay" id="overlay"></div>
    </div></div>
    <div class="notes-bar" id="notesPanel"><textarea id="notesInput"></textarea><button id="notesClose"></button></div>
  </section>
  <aside class="inspector" id="inspector"></aside>
</main>
<footer><span id="statusInfo"></span><span id="sbSlideCount"></span>
  <button id="zoomOut"></button><button id="zoomFit"></button><button id="zoomVal"></button><button id="zoomIn"></button></footer>
<div id="popRoot"></div><div id="modalRoot"></div><div id="presentRoot"></div><div id="toastRoot"></div>
<input type="file" id="hiddenFile" hidden/>
<button id="btnExportTop"></button><button id="btnPresentTop"></button>
<button id="notesToggle"></button><button id="gridToggle"></button>
`;

function groupedDoc() {
  return {
    title: 't', theme: 'blue', slideSize: { width: 1280, height: 720 },
    slides: [{
      id: 's1', background: null, notes: '', hidden: false, transition: null, animations: [],
      elements: [{
        id: 'g1', type: 'group', x: 0, y: 0, width: 200, height: 200,
        children: [
          { id: 'a', type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 50, height: 50, fill: '#f00' },
          { id: 'b', type: 'shape', shapeType: 'rect', x: 50, y: 50, width: 50, height: 50, fill: '#0f0' },
        ]
      }, { id: 't', type: 'shape', shapeType: 'rect', x: 300, y: 300, width: 80, height: 80, fill: '#000' }]
    }]
  };
}

describe('点击组合内子元素应选中整个组合', () => {
  beforeAll(async () => {
    document.body.innerHTML = HTML;
    await import('../examples/editor/src/app.js');
  }, 20000);

  it('在子元素 a 上点击 → sel = [g1]，且绘制 group 选择框', async () => {
    const { store } = await import('../examples/editor/src/store.js');
    store.setDoc(groupedDoc());
    const childNode = document.querySelector('#frame .el[data-id="a"]') as any;
    expect(childNode).toBeTruthy();
    const ev = new (window as any).Event('pointerdown', { bubbles: true });
    ev.button = 0; ev.shiftKey = false;
    childNode.dispatchEvent(ev);
    expect(store.sel).toEqual(['g1']);
    const gBox = document.querySelector('#overlay .sel-box.group-box');
    expect(gBox).toBeTruthy();
  });

  it('点击组合的空白区域 → 仍选中 g1（行为不变）', async () => {
    const { store } = await import('../examples/editor/src/store.js');
    store.setDoc(groupedDoc());
    const groupNode = document.querySelector('#frame .el.el-group[data-id="g1"]') as any;
    const ev = new (window as any).Event('pointerdown', { bubbles: true });
    ev.button = 0; ev.shiftKey = false;
    groupNode.dispatchEvent(ev);
    expect(store.sel).toEqual(['g1']);
  });
});

describe('deleteSelected / duplicateSelected 兼容组合内子元素', () => {
  function mk() {
    const store = createStore({ storage: undefined });
    store.init();
    store.setDoc(groupedDoc());
    return { store, a: createActions(store) };
  }

  it('删除组合内子元素 a → 仅从 group.children 移除，组合保留', () => {
    const { store, a } = mk();
    store.setSel(['a']);
    a.deleteSelected();
    const g = store.doc.slides[0].elements.find((e: any) => e.id === 'g1');
    expect(g.children.map((c: any) => c.id)).toEqual(['b']);
    expect(store.doc.slides[0].elements.map((e: any) => e.id)).toEqual(['g1', 't']);
  });

  it('删除组合全部子元素 → 空组合一并移除', () => {
    const { store, a } = mk();
    store.setSel(['a', 'b']);
    a.deleteSelected();
    expect(store.doc.slides[0].elements.map((e: any) => e.id)).toEqual(['t']);
  });

  it('复制组合内子元素 a → 在 group.children 内新增副本', () => {
    const { store, a } = mk();
    store.setSel(['a']);
    a.duplicateSelected();
    const g = store.doc.slides[0].elements.find((e: any) => e.id === 'g1');
    expect(g.children.length).toBe(3);
    const clone = g.children.find((c: any) => c.id !== 'a' && c.id !== 'b');
    expect(clone).toBeTruthy();
    expect(clone.x).toBe(16);
  });

  it('复制整个组合 → 顶层新增一个完整组合（含子元素）', () => {
    const { store, a } = mk();
    store.setSel(['g1']);
    a.duplicateSelected();
    const groups = store.doc.slides[0].elements.filter((e: any) => e.type === 'group');
    expect(groups.length).toBe(2);
    expect(groups[1].children.length).toBe(2);
  });
});
