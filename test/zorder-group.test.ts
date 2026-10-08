// @vitest-environment jsdom
// 回归：组合(group)内的子元素被选中时，zOrder 应在其所属 group 内部重排，而非静默失效
import { describe, it, expect } from 'vitest';
import { createStore, createActions } from '../dist/ppt-parser.browser.js';

function mkStore() {
  const store = createStore({ storage: undefined });
  store.init();
  const group = {
    id: 'g1', type: 'group', x: 0, y: 0, width: 200, height: 200,
    children: [
      { id: 'a', type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 50, height: 50, fill: '#f00' },
      { id: 'b', type: 'shape', shapeType: 'rect', x: 50, y: 50, width: 50, height: 50, fill: '#0f0' },
      { id: 'c', type: 'shape', shapeType: 'rect', x: 100, y: 100, width: 50, height: 50, fill: '#00f' },
    ]
  };
  store.setDoc({
    title: 't', theme: 'blue', slideSize: { width: 1280, height: 720 },
    slides: [{
      id: 's1', background: null, notes: '', hidden: false, transition: null, animations: [],
      elements: [group, { id: 't', type: 'shape', shapeType: 'rect', x: 300, y: 300, width: 80, height: 80, fill: '#000' }]
    }]
  });
  return store;
}

describe('组合内子元素 zOrder（修复回归）', () => {
  it('选中 group 子元素 a，置于顶层 → a 在 group.children 末尾', () => {
    const store = mkStore();
    const a = createActions(store);
    store.setSel(['a']);
    a.zOrder('front');
    const g = store.doc.slides[0].elements.find((e: any) => e.id === 'g1');
    expect(g.children.map((c: any) => c.id)).toEqual(['b', 'c', 'a']);
    // 顶层 elements 顺序不应改变
    expect(store.doc.slides[0].elements.map((e: any) => e.id)).toEqual(['g1', 't']);
  });

  it('选中 group 子元素 c，置于底层 → c 在 group.children 开头', () => {
    const store = mkStore();
    const a = createActions(store);
    store.setSel(['c']);
    a.zOrder('back');
    const g = store.doc.slides[0].elements.find((e: any) => e.id === 'g1');
    expect(g.children.map((c: any) => c.id)).toEqual(['c', 'a', 'b']);
  });

  it('选中 group 子元素 b，上移一层 → b 上移一位（a,c,b）', () => {
    const store = mkStore();
    const a = createActions(store);
    store.setSel(['b']);
    a.zOrder('forward');
    const g = store.doc.slides[0].elements.find((e: any) => e.id === 'g1');
    expect(g.children.map((c: any) => c.id)).toEqual(['a', 'c', 'b']);
  });

  it('选中 group 子元素 b，下移一层 → b 下移一位（b,a,c）', () => {
    const store = mkStore();
    const a = createActions(store);
    store.setSel(['b']);
    a.zOrder('backward');
    const g = store.doc.slides[0].elements.find((e: any) => e.id === 'g1');
    expect(g.children.map((c: any) => c.id)).toEqual(['b', 'a', 'c']);
  });

  it('顶层元素仍正常重排', () => {
    const store = mkStore();
    const a = createActions(store);
    store.setSel(['t']);
    a.zOrder('front');
    expect(store.doc.slides[0].elements.map((e: any) => e.id)).toEqual(['g1', 't']); // t 本就在末尾
    a.zOrder('back');
    expect(store.doc.slides[0].elements.map((e: any) => e.id)).toEqual(['t', 'g1']);
  });
});
