# examples/editor（已冻结 / Frozen）

> ⚠️ **本目录为历史演示副本，已冻结，不再随主逻辑更新。**

编辑器**可维护的内核源码在 [`src/editor/`](../src/editor/)**（UI 无关：模型、状态、编辑操作、图表与几何），
并通过 `docFromPptx` / `docToPptx` 与标准 `PptxDocument` 双向互通。本目录的 JS 实现是从 `src/editor`
早期版本「下沉」出的快照，与主逻辑已出现分叉，**请勿在此修改功能**，否则会与 `src/editor` 漂移。

- 运行在线编辑器演示：见根 README 的 Live Demo（`examples/index.html`）。
- 在自有项目中复用编辑器能力：直接 `import { createStore, createActions, docFromPptx, docToPptx, ... } from '@fefeding/pptx-parser'`，详见根 README「Headless Editor Core」。
- 如需改动编辑器行为，请改 `src/editor/` 并同步更新相应测试。
