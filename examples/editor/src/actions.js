/**
 * 兼容层：编辑操作已下沉到库（src/editor/actions.ts）。
 *
 * 库内为 createActions(store) 工厂（不再依赖单例），此处绑定到本编辑器的 store 实例后
 * 原样导出，使其它 editor 模块的 `import ... from './actions.js'` 无需改动。
 */
import { createActions } from '../../../dist/ppt-parser.browser.js';
import { store } from './store.js';

export const {
  cloneElement, addElement, deleteSelected, duplicateSelected, copySelected, paste, selectAll, nudge,
  zOrder, alignElements, distribute, groupSelection, ungroupSelection, toggleLock, toggleHidden,
  addSlide, duplicateSlide, deleteSlide, moveSlide, toggleSlideHidden, applyLayout,
  setBackground, setSlideSize, applyTheme, setNotes, applyTextStyleSel, setBackgroundImage,
  setElementGeo, resizeTable, updateElement, findInDoc,
  setTransition, addAnimation, updateAnimation, removeAnimation, moveAnimation
} = createActions(store);
