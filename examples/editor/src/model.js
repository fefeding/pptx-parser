/**
 * 兼容层：文档模型与 PPTX 双向转换已下沉到库（src/editor/model.ts）。
 *
 * 此处从 dist 原样 re-export，使其它 editor 模块的 `import ... from './model.js'` 无需改动；
 * 新代码请直接从库导入，第二个编辑器可复用同一套模型与导入导出能力。
 */
export * from '../../../dist/ppt-parser.browser.js';
