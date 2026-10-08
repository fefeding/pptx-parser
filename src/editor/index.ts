/**
 * Headless 编辑器内核：与 UI 框架无关的编辑器业务逻辑。
 *
 * 组合方式（第二个编辑器可直接复用，自行实现渲染层与交互层）：
 *   const store = createStore({ storage: localStorage });
 *   const actions = createActions(store);
 *   store.setDoc(docFromPptx(pptxDoc));   // 导入
 *   actions.alignElements('hcenter');     // 编辑
 *   const pptxDoc = docToPptx(store.doc); // 导出
 *   const svg = renderChartSVG(chartEl);  // 图表（字符串，宿主自定渲染）
 */
export * from './model';
export { createStore, EditorStore, normalizeDoc, normalizeElement } from './store';
export type { StorageAdapter } from './store';
export { createActions } from './actions';
export type { EditorActions } from './actions';
export { renderChartSVG } from './charts';
export { elementRect, rotatedRect, effectMargin } from './geometry';
