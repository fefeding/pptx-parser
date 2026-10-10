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

// 基础矩形/边距类几何 API（仅被 geometry 内部使用，rollup 不会摇除其具名重导出）。
export { elementRect, rotatedRect, absoluteElementRect, effectMargin } from './geometry';

// 线段/连接线相关几何 API（isLineElement / lineEndpoints / geomPoints / connectionPoints /
// connectionPointOf / findGlueTarget / resyncGlue / setLineEndpoints / toVertexEditable）
// 无法以具名重导出稳定保留（被 actions 跨模块具名导入后，rollup 会将其同名 re-export 修剪掉），
// 故统一通过 createActions.geometryApi 暴露。消费方（vscode-pptx）从该聚合对象取用即可：
//   import { createActions } from '@fefeding/ppt-parser';
//   const GEOM = (createActions as any).geometryApi;
