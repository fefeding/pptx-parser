/**
 * 兼容层：状态仓库已下沉到库（src/editor/store.ts），此处仅创建本编辑器使用的实例。
 *
 * 库内 store 不再硬编码 localStorage / document：持久化通过注入 StorageAdapter 完成，
 * 因此同一套状态与撤销栈可在 Node、测试或第二个编辑器中复用。
 */
import { createStore } from '../../../dist/ppt-parser.browser.js';

// 浏览器环境注入 localStorage（保留 UI 偏好持久化）；非浏览器时为 undefined → 库回退内存实现
export const store = createStore({
  storage: typeof localStorage !== 'undefined' ? localStorage : undefined
});

export { createStore };
