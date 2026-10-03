// @vitest-environment jsdom
import { describe, it, expect, beforeAll } from 'vitest';

const ERRORS = [];
beforeAll(() => {
  // 捕获运行期错误
  const orig = console.error;
  console.error = (...a) => { ERRORS.push(a.join(' ')); orig(...a); };
  if (typeof window !== 'undefined') {
    window.addEventListener('error', (e) => ERRORS.push('window.error: ' + e.message));
  }
});

describe('editor smoke', () => {
  it('boots and renders without throwing', async () => {
    // 注入必要 DOM
    document.body.innerHTML = `
      <header class="topbar">
        <input id="docTitle" /><span id="saveState"></span>
        <button id="notesToggle"></button><button id="gridToggle"></button>
        <button id="btnPresentTop"></button><button id="btnExportTop"></button>
      </header>
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
      <footer class="statusbar">
        <span id="statusInfo"></span><span id="sbSlideCount"></span>
        <button id="zoomOut"></button><button id="zoomFit"></button><button id="zoomVal"></button><button id="zoomIn"></button>
      </footer>
      <div id="popRoot"></div><div id="modalRoot"></div><div id="presentRoot"></div><div id="toastRoot"></div>
      <input type="file" id="hiddenFile" />`;

    // 提供 jsdom 缺失的 API
    if (!HTMLElement.prototype.getBoundingClientRect) HTMLElement.prototype.getBoundingClientRect = () => ({ left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0 });
    window.requestAnimationFrame = window.requestAnimationFrame || ((cb) => setTimeout(cb, 0));

    const app = await import('./src/app.js');
    // 等待动态 import 与渲染完成
    await new Promise((r) => setTimeout(r, 50));

    const { store } = await import('./src/store.js');
    expect(store.doc).toBeTruthy();
    expect(Array.isArray(store.doc.slides)).toBe(true);
    expect(store.doc.slides.length).toBeGreaterThan(0);

    const frame = document.getElementById('frame');
    expect(frame.children.length).toBeGreaterThan(0); // 至少渲染出元素

    const menubar = document.getElementById('menubar');
    expect(menubar.children.length).toBeGreaterThan(0); // 菜单已构建

    const slideList = document.getElementById('slideList');
    expect(slideList.children.length).toBe(store.doc.slides.length);

    // 选中一个元素，属性面板应渲染
    store.setSel([store.slide.elements[0].id]);
    await new Promise((r) => setTimeout(r, 10));
    const inspector = document.getElementById('inspector');
    expect(inspector.innerHTML.length).toBeGreaterThan(0);

    // 基本动作：新增文本元素
    const before = store.slide.elements.length;
    const { createTextElement } = await import('./src/model.js');
    const { addElement } = await import('./src/actions.js');
    addElement(createTextElement({ text: '测试', x: 100, y: 100 }));
    expect(store.slide.elements.length).toBe(before + 1);

    // 富文本编辑回写（注意要选画布内的 body，而非缩略图里的同名元素）
    const { enterEditing, exitEditing } = await import('./src/interact.js');
    const textEl = store.slide.elements.find((e) => e.type === 'text');
    store.setSel([textEl.id]);
    enterEditing(textEl, { x: 5, y: 5 });
    const frameEl = document.getElementById('frame');
    const body = frameEl.querySelector(`[data-id="${textEl.id}"] .tb-body`);
    body.innerHTML = '<div>新标题</div><div>第二行</div>';
    const { parseBody } = await import('./src/richtext.js');
    expect(parseBody(body, { fontSize: 24, color: '#000000', fontFace: '微软雅黑' }).map((p) => p.runs.map((r) => r.text)))
      .toEqual([['新标题'], ['第二行']]);
    exitEditing();
    await new Promise((r) => setTimeout(r, 50));
    const updated = store.slide.elements.find((e) => e.id === textEl.id);
    expect(updated.paragraphs.length).toBe(2);
    expect(updated.paragraphs[0].runs[0].text).toBe('新标题');

    // 导出 PPTX 不应抛错
    const { exportPptx } = await import('./src/io.js');
    // 拦截下载
    const origCreate = document.createElement.bind(document);
    let downloaded = null;
    document.createElement = (tag) => { const el = origCreate(tag); if (tag === 'a') { el.click = () => { downloaded = el.download; }; } return el; };
    const URLorig = globalThis.URL.createObjectURL;
    const blobs = [];
    globalThis.URL.createObjectURL = (b) => { blobs.push(b); return 'blob:x'; };
    try {
      // 追加各类元素以确保序列化路径都被覆盖
      const { createChartElement, createTableElement, createShapeElement, createImageElement } = await import('./src/model.js');
      addElement(createChartElement('barChart', { x: 50, y: 50 }));
      addElement(createTableElement(3, 4, { x: 50, y: 400 }));
      addElement(createShapeElement('roundRect', { x: 700, y: 50, fill: { type: 'gradient', direction: 'diagonal', stops: [{ color: '#1A73E8', position: 0 }, { color: '#34A853', position: 1 }] } }));
      addElement(createImageElement('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', { x: 700, y: 400 }));
      await exportPptx();
    } finally {
      globalThis.URL.createObjectURL = URLorig; document.createElement = origCreate;
    }
    expect(downloaded).toMatch(/\.pptx$/);

    // 整库序列化 / 解析 往返：直接拿字节
    const { jsonToPptx, pptxToStandard } = await import('../../dist/ppt-parser.browser.js');
    const docToPptx = (await import('./src/model.js')).docToPptx;
    const bytes = await jsonToPptx(docToPptx(store.doc), { outputType: 'uint8array' });
    expect(bytes && bytes.length).toBeTruthy();
    const round = await pptxToStandard(bytes);
    expect(round.slides.length).toBe(store.doc.slides.length);

    // 过渡 + 动画 序列化往返
    const { setTransition, addAnimation } = await import('./src/actions.js');
    const firstSlide = store.doc.slides[0];
    const animTarget = firstSlide.elements.find((e) => e.type === 'text');
    setTransition({ type: 'fade', duration: 800, advanceOnClick: true });
    addAnimation({ target: animTarget.id, type: 'flyIn', presetClass: 'entr', duration: 0.5, direction: 'l', trigger: { type: 'onClick' } });
    const pptxWithAnim = docToPptx(store.doc);
    expect(pptxWithAnim.slides[0].transition.type).toBe('fade');
    expect(pptxWithAnim.slides[0].animations.length).toBe(1);
    expect(pptxWithAnim.slides[0].animations[0].type).toBe('flyIn');
    const animBytes = await jsonToPptx(pptxWithAnim, { outputType: 'uint8array' });
    const animRound = await pptxToStandard(animBytes);
    expect(animRound.slides[0].transition).toBeTruthy();
    expect(animRound.slides[0].transition.type).toBe('fade');
    expect(animRound.slides[0].animations.length).toBe(1);

    if (ERRORS.length) console.log('CAPTURED ERRORS:\n' + ERRORS.join('\n'));
    expect(ERRORS.length).toBe(0);
  });
});
