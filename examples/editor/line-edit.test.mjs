// @vitest-environment jsdom
// 线条编辑：核心几何逻辑 + 编辑器 UI 手柄渲染。
import { describe, it, expect, beforeAll } from 'vitest';
import * as interact from './src/interact.js';

describe('line/connector editing helpers (geometry)', () => {
  it('lineEndpoints reflects flip (起点/终点对角点)', () => {
    const el = { type: 'shape', shapeType: 'line', x: 100, y: 100, width: 200, height: 100 };
    let ep = interact.lineEndpoints(el);
    expect(ep.a).toEqual({ x: 100, y: 100 });
    expect(ep.b).toEqual({ x: 300, y: 200 });
    el.flipH = true;
    ep = interact.lineEndpoints(el);
    expect(ep.a).toEqual({ x: 300, y: 100 });
    expect(ep.b).toEqual({ x: 100, y: 200 });
  });

  it('toVertexEditable expands presets into custGeom', () => {
    const line = { type: 'shape', shapeType: 'line', x: 0, y: 0, width: 200, height: 100 };
    expect(interact.toVertexEditable(line).commands.map((c) => c.type)).toEqual(['moveTo', 'lnTo']);
    const bent = { type: 'shape', shapeType: 'bentConnector3', x: 0, y: 0, width: 200, height: 100 };
    expect(interact.toVertexEditable(bent).commands.map((c) => c.type))
      .toEqual(['moveTo', 'lnTo', 'lnTo', 'lnTo']);
    const curved = { type: 'shape', shapeType: 'curvedConnector3', x: 0, y: 0, width: 200, height: 100 };
    expect(interact.toVertexEditable(curved).commands.map((c) => c.type))
      .toEqual(['moveTo', 'quadBezTo', 'quadBezTo']);
  });

  it('geomPoints maps local coords to absolute (含 flip)', () => {
    const el = {
      type: 'shape', shapeType: null, x: 50, y: 50, width: 200, height: 100,
      custGeom: { paths: [{ w: 200, h: 100, closed: false, commands: [
        { type: 'moveTo', x: 0, y: 0 }, { type: 'lnTo', x: 200, y: 100 }
      ] }] }
    };
    const pts = interact.geomPoints(el);
    expect(pts.length).toBe(2);
    expect(pts[0].abs).toEqual({ x: 50, y: 50 });
    expect(pts[1].abs).toEqual({ x: 250, y: 150 });
    el.flipH = true;
    expect(interact.geomPoints(el)[0].abs).toEqual({ x: 250, y: 50 });
  });

  it('setLineEndpoints derives bbox + flip (改变起点/终点)', () => {
    const el = {};
    interact.setLineEndpoints(el, { x: 300, y: 200 }, { x: 100, y: 100 });
    expect(el.x).toBe(100); expect(el.y).toBe(100);
    expect(el.width).toBe(200); expect(el.height).toBe(100);
    expect(el.flipH).toBe(true); expect(el.flipV).toBe(true);
  });
});

describe('editor UI handles (jsdom)', () => {
  beforeAll(async () => {
    document.body.innerHTML = `
      <header class="topbar"><input id="docTitle" /><span id="saveState"></span>
        <button id="notesToggle"></button><button id="gridToggle"></button>
        <button id="btnPresentTop"></button><button id="btnExportTop"></button></header>
      <nav class="menubar" id="menubar"></nav>
      <div class="toolbar" id="quickbar"></div>
      <main class="workspace">
        <aside class="slides-panel"><button id="addSlideBtn"></button><div id="slideList"></div></aside>
        <section class="canvas-area">
          <div class="canvas-scroll" id="canvasScroll"><div class="stage" id="stage"><div class="slide-frame" id="frame"></div><div class="overlay" id="overlay"></div></div></div>
          <div class="notes-bar" id="notesPanel"><textarea id="notesInput"></textarea></div>
        </section>
        <aside class="inspector" id="inspector"></aside>
      </main>
      <footer class="statusbar"><span id="statusInfo"></span><span id="sbSlideCount"></span>
        <button id="zoomOut"></button><button id="zoomFit"></button><button id="zoomVal"></button><button id="zoomIn"></button></footer>
      <div id="popRoot"></div><div id="modalRoot"></div><div id="presentRoot"></div><div id="toastRoot"></div>
      <input type="file" id="hiddenFile" />`;
    if (!HTMLElement.prototype.getBoundingClientRect) HTMLElement.prototype.getBoundingClientRect = () => ({ left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0 });
    window.requestAnimationFrame = window.requestAnimationFrame || ((cb) => setTimeout(cb, 0));
    await import('./src/app.js');
    await new Promise((r) => setTimeout(r, 50));
  });

  it('selected line shows two endpoint handles', async () => {
    const { store } = await import('./src/store.js');
    const { createShapeElement } = await import('./src/model.js');
    const { addElement } = await import('./src/actions.js');
    const line = createShapeElement('line', { x: 100, y: 100, width: 200, height: 100, line: { color: '#000000', width: 2 } });
    addElement(line);
    store.setSel([line.id]);
    interact.renderCanvas();
    const overlay = document.getElementById('overlay');
    expect(overlay.querySelectorAll('.handle.endpoint').length).toBe(2);
  });

  it('vertex edit converts to custGeom and shows vertex handles', async () => {
    const { store } = await import('./src/store.js');
    const line = store.slide.elements.find((e) => e.shapeType === 'line');
    expect(line).toBeTruthy();
    interact.enterVertexEdit(line);
    interact.renderCanvas();
    const updated = store.findElement(line.id);
    expect(updated.custGeom).toBeTruthy();
    expect(store.vertexEdit).toBe(line.id);
    const overlay = document.getElementById('overlay');
    expect(overlay.querySelectorAll('.handle.vertex').length).toBe(2);
    // 贝塞尔控制点（curvedConnector 才有；此处直线无控制点，验证至少顶点存在）
  });

  it('addVertexAt inserts a midpoint vertex on double-click position', async () => {
    const { store } = await import('./src/store.js');
    const line = store.slide.elements.find((e) => e.custGeom);
    const updated = store.findElement(line.id);
    const before = updated.custGeom.paths[0].commands.length;
    interact.addVertexAt(updated, { x: 200, y: 150 }); // 元素 x:100 y:100 w:200 h:100 → 局部 (100,50) 即直线段中点
    const after = store.findElement(line.id);
    expect(after.custGeom.paths[0].commands.length).toBe(before + 1);
  });

  it('deleteSelectedVertex removes a selected vertex', async () => {
    const { store } = await import('./src/store.js');
    const line = store.slide.elements.find((e) => e.custGeom);
    store.vertexEdit = line.id;
    store.vertexSel = { cmd: 2, kind: 'point' };
    const before = store.findElement(line.id).custGeom.paths[0].commands.length;
    interact.deleteSelectedVertex();
    const after = store.findElement(line.id);
    expect(after.custGeom.paths[0].commands.length).toBe(before - 1);
  });

  it('enterVertexEdit under rotation preserves endpoints (clears rotation, bakes geometry)', async () => {
    const { store } = await import('./src/store.js');
    const { createShapeElement } = await import('./src/model.js');
    const { addElement } = await import('./src/actions.js');
    const line = createShapeElement('line', { x: 100, y: 100, width: 100, height: 100, rotation: 45, line: { color: '#000', width: 2 } });
    addElement(line);
    const before = interact.lineEndpoints(store.findElement(line.id));
    interact.enterVertexEdit(store.findElement(line.id));
    const t = store.findElement(line.id);
    expect(t.rotation).toBe(0);
    expect(t.custGeom).toBeTruthy();
    const pts = interact.geomPoints(t);
    const aAbs = pts[0].abs, bAbs = pts[pts.length - 1].abs;
    expect(aAbs.x).toBeCloseTo(before.a.x, 1);
    expect(aAbs.y).toBeCloseTo(before.a.y, 1);
    expect(bAbs.x).toBeCloseTo(before.b.x, 1);
    expect(bAbs.y).toBeCloseTo(before.b.y, 1);
  });
});
