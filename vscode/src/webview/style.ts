// Webview styles (injected as a string to avoid extra resource requests).
//
// Colours come from the VS Code theme variables (`--vscode-*`) with the old hard-coded dark
// palette as fallback, so the editor matches the user's colour theme. The ribbon/menu/dialog
// chrome follows PowerPoint's structure: title bar → tabs → groups with captions underneath.
export const CSS = `
:root {
  --bg: var(--vscode-editor-background, #1e1e1e);
  --panel: var(--vscode-sideBar-background, #252526);
  --panel-2: var(--vscode-editorWidget-background, #2d2d30);
  --panel-3: var(--vscode-toolbar-hoverBackground, #3a3a3d);
  --border: var(--vscode-panel-border, #3c3c3c);
  --text: var(--vscode-foreground, #e6e6e6);
  --muted: var(--vscode-descriptionForeground, #9d9d9d);
  --accent: var(--vscode-charts-yellow, #f9ab00);
  --accent-2: var(--vscode-focusBorder, #4285f4);
  --accent-bg: var(--vscode-toolbar-activeBackground, rgba(66,133,244,.28));
  --danger: var(--vscode-inputValidation-errorBackground, #f14c4c);
  --ok: var(--vscode-charts-green, #4ec9b0);
  --hover: var(--vscode-toolbar-hoverBackground, rgba(255,255,255,.08));
}
* { box-sizing: border-box; }
html, body { margin: 0; height: 100%; }
body {
  font-family: var(--vscode-font-family, -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif);
  font-size: var(--vscode-font-size, 13px);
  background: var(--bg);
  color: var(--text);
  overflow: hidden;
}
#app { display: flex; flex-direction: column; height: 100vh; }

/* Hide editing controls in preview mode */
.app.mode-preview .edit-only { display: none !important; }

button { font-family: inherit; }
.btn {
  background: var(--panel-2); color: var(--text); border: 1px solid var(--border);
  border-radius: 4px; padding: 4px 10px; cursor: pointer; font-size: 12px;
  display: inline-flex; align-items: center; gap: 4px;
}
.btn:hover { background: var(--panel-3); }
.btn.primary { background: var(--accent-2); border-color: var(--accent-2); color: #fff; }
.btn.primary:hover { filter: brightness(1.1); }
.btn.danger { background: var(--danger); border-color: var(--danger); color: #fff; }
.btn:disabled { opacity: .45; cursor: default; }

/* ================= Ribbon (PowerPoint-style) ================= */
.ribbon { display: flex; flex-direction: column; flex: 0 0 auto; background: var(--panel); border-bottom: 1px solid var(--border); }

.rb-titlebar {
  display: flex; align-items: stretch; gap: 6px; padding: 0 6px 0 10px;
  background: var(--panel); font-size: 12px; min-height: 36px;
}
.rb-doc { display: flex; align-items: center; gap: 6px; min-width: 0; color: var(--muted); }
.rb-doc-title { color: var(--text); font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 26vw; }
.rb-dirty { width: 7px; height: 7px; border-radius: 50%; background: transparent; flex: 0 0 auto; }
.rb-dirty.on { background: var(--accent); }
.rb-qat { display: flex; align-items: center; gap: 2px; }
.rb-spring { flex: 1 1 auto; }
.rb-collapse { align-self: center; }
.rb-collapse .ico { transform: rotate(180deg); }
.rb-collapse.flip .ico { transform: rotate(0deg); }

.rb-seg { display: inline-flex; border: 1px solid var(--border); border-radius: 4px; overflow: hidden; margin-right: 2px; }
.rb-seg-btn {
  background: transparent; color: var(--text); border: 0; padding: 3px 10px;
  font-size: 12px; cursor: pointer;
}
.rb-seg-btn:hover:not(.disabled) { background: var(--hover); }
.rb-seg-btn.on { background: var(--accent-bg); }
.rb-seg-btn.disabled { opacity: .5; cursor: default; }

/* Tab strip sits inline in the title bar; the active tab is marked by a bottom accent. */
.rb-tabs { display: flex; align-items: stretch; gap: 1px; flex: 0 1 auto; margin: 0 2px; overflow: hidden; }
.rb-tab {
  background: transparent; border: 0; border-bottom: 2px solid transparent; border-radius: 4px 4px 0 0;
  color: var(--text); padding: 0 10px; font-size: 12px; cursor: pointer; white-space: nowrap;
  display: inline-flex; align-items: center;
}
.rb-tab:hover { background: var(--hover); }
.rb-tab.active { border-bottom-color: var(--accent-2); background: var(--panel-2); }

.rb-body { background: var(--panel-2); overflow-x: auto; overflow-y: hidden; }
.rb-panel { display: flex; align-items: stretch; gap: 0; min-height: 92px; padding: 4px 2px 0; }
.rb-group {
  display: flex; flex-direction: column; justify-content: space-between;
  padding: 2px 8px 0; border-right: 1px solid var(--border); flex: 0 0 auto;
}
.rb-group:last-child { border-right: 0; }
.rb-group-body { display: flex; flex-direction: column; gap: 3px; justify-content: center; flex: 1 1 auto; }
.rb-group-title {
  font-size: 10px; color: var(--muted); text-align: center; padding: 2px 0 3px;
  white-space: nowrap;
}
.rb-row { display: flex; align-items: center; gap: 2px; }

.rb-btn {
  background: transparent; color: var(--text); border: 1px solid transparent; border-radius: 4px;
  cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 4px;
  padding: 2px 6px; font-size: 12px;
}
.rb-btn:hover:not(.disabled) { background: var(--hover); border-color: var(--border); }
.rb-btn.disabled { opacity: .38; cursor: default; }
.rb-btn.on { background: var(--accent-bg); border-color: var(--accent-2); }
.rb-btn-mini { min-width: 24px; height: 24px; padding: 2px 4px; }
/* 保存按钮：有未保存改动时高亮（.on）并在右上角显示小圆点 */
.rb-save { position: relative; }
.rb-save.on::after {
  content: ''; position: absolute; right: 0; top: 0;
  width: 6px; height: 6px; border-radius: 50%; background: var(--accent);
}
.rb-btn-big { flex-direction: column; gap: 2px; height: 58px; min-width: 52px; padding: 4px 6px; font-size: 11px; }
.rb-btn-big .rb-glyph { font-size: 14px; }
.rb-glyph { font-size: 13px; line-height: 1; }
.rb-split { display: flex; align-items: stretch; }
.rb-split-main { border-radius: 4px 0 0 4px; }
.rb-split-arrow { width: 14px; height: 58px; border-radius: 0 4px 4px 0; padding: 0; margin-left: 1px; }
.rb-ico { display: inline-flex; align-items: center; }
.rb-ico svg { display: block; }

.rb-btn-color { flex-direction: column; gap: 0; height: 50px; min-width: 30px; position: relative; }
.rb-color-stack { position: relative; display: inline-flex; }
.rb-color-swatch {
  position: absolute; left: 1px; right: 1px; bottom: -2px; height: 4px;
  border: 1px solid var(--border); border-radius: 1px;
}
.rb-caret { line-height: 0; }
.rb-color-input { position: absolute; opacity: 0; width: 0; height: 0; pointer-events: none; }

.rb-field { display: flex; flex-direction: column; gap: 1px; }
.rb-field label { font-size: 10px; color: var(--muted); }
.rb-select {
  background: var(--bg); color: var(--text); border: 1px solid var(--border);
  border-radius: 3px; font-size: 11px; padding: 2px 3px; height: 22px;
}
.rb-select.disabled { opacity: .4; }

.rb-check { display: inline-flex; align-items: center; gap: 4px; font-size: 11px; color: var(--text); padding: 2px 4px; }

.rb-stack { display: flex; flex-direction: column; gap: 2px; }
.rb-caption { font-size: 10px; color: var(--muted); }
.rb-gallery { display: flex; gap: 2px; overflow-x: auto; padding-bottom: 2px; max-width: 620px; }
.rb-chip {
  background: var(--bg); color: var(--text); border: 1px solid var(--border);
  border-radius: 3px; font-size: 11px; padding: 3px 8px; cursor: pointer; white-space: nowrap;
}
.rb-chip:hover { background: var(--hover); }
.rb-chip.on { background: var(--accent-bg); border-color: var(--accent-2); }
.rb-gallery.disabled .rb-chip { opacity: .45; cursor: default; }

.rb-theme {
  display: flex; flex-direction: column; align-items: center; gap: 3px; padding: 3px 5px;
  background: transparent; border: 1px solid transparent; border-radius: 4px; cursor: pointer; color: var(--text);
}
.rb-theme:hover { background: var(--hover); }
.rb-theme.on { border-color: var(--accent-2); background: var(--accent-bg); }
.rb-theme-sw { display: grid; grid-template-columns: 10px 10px; gap: 1px; }
.rb-theme-sw i { width: 10px; height: 10px; display: block; border: 1px solid var(--border); }
.rb-theme-name { font-size: 10px; white-space: nowrap; }

.rb-anim-list { width: 190px; max-height: 74px; overflow-y: auto; display: flex; flex-direction: column; gap: 2px; }
.rb-anim-row { display: flex; align-items: center; gap: 2px; padding: 1px 2px; border-radius: 3px; }
.rb-anim-row:hover { background: var(--panel-2); }
.rb-anim-row.on { outline: 1px solid var(--accent-2); background: var(--panel-2); }
.rb-anim-text { flex: 1 1 auto; font-size: 10px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.rb-anim-list .empty { padding: 4px; }

/* Collapsed ribbon: keep the title bar + tabs, hide the panel (like PowerPoint's auto-hide) */
.ribbon.rb-collapsed .rb-body { display: none; }

/* Preview mode: no expandable tab ribbon — just a compact shortcut menu in the header. */
.rb-preview-menu { display: none; align-items: center; align-self: center; }
.ribbon.rb-preview .rb-tabs { display: none; }
.ribbon.rb-preview .rb-collapse { display: none; }
.ribbon.rb-preview .rb-body { display: none; }
.ribbon.rb-preview .rb-preview-menu { display: inline-flex; }
/* Edit-only quick-access buttons (save / undo / redo) are always disabled here — drop them. */
.ribbon.rb-preview .rb-qat .rb-btn.disabled { display: none; }

/* ================= Popup menus ================= */
.menu {
  position: fixed; z-index: 90; min-width: 168px; padding: 4px;
  background: var(--vscode-menu-background, var(--panel-2));
  color: var(--vscode-menu-foreground, var(--text));
  border: 1px solid var(--vscode-menu-border, var(--border));
  border-radius: 6px; box-shadow: 0 6px 24px rgba(0,0,0,.45);
  max-height: 80vh; overflow-y: auto;
}
.menu-item {
  display: flex; align-items: center; gap: 6px; padding: 4px 8px; border-radius: 4px;
  font-size: 12px; cursor: pointer; white-space: nowrap; outline: none;
}
.menu-item.active, .menu-item:focus { background: var(--vscode-menu-selectionBackground, var(--accent-bg)); }
.menu-item.disabled { opacity: .4; cursor: default; }
.menu-ico { width: 16px; display: inline-flex; justify-content: center; flex: 0 0 auto; font-size: 11px; }
.menu-text { flex: 1 1 auto; }
.menu-accel { color: var(--muted); font-size: 11px; margin-left: 12px; }
.menu-sep { height: 1px; background: var(--border); margin: 4px 2px; }
.menu-caption { font-size: 10px; color: var(--muted); padding: 4px 8px 2px; text-transform: uppercase; letter-spacing: .04em; }

/* ================= Main area ================= */
.main { display: flex; flex: 1 1 auto; min-height: 0; }
.slides-panel { width: 180px; background: var(--panel); border-right: 1px solid var(--border); display: flex; flex-direction: column; flex: 0 0 auto; }
.sp-head { padding: 5px 6px 5px 8px; font-size: 12px; color: var(--muted); display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid var(--border); }
.sp-new { padding: 2px 6px; }
.sp-list { overflow-y: auto; padding: 6px; display: flex; flex-direction: column; gap: 6px; }
.thumb {
  position: relative; border: 2px solid transparent; border-radius: 4px; cursor: pointer; background: #fff;
  overflow: hidden; flex: 0 0 auto;
}
.thumb.active { border-color: var(--accent); }
.thumb .num { position: absolute; top: 2px; left: 4px; font-size: 10px; color: #333; background: rgba(255,255,255,.7); border-radius: 3px; padding: 0 3px; }
.thumb .inner { transform-origin: top left; pointer-events: none; }
.thumb.hidden-slide { opacity: .55; }
.thumb.hidden-slide .num { text-decoration: line-through; }
.hide-badge { position: absolute; right: 3px; top: 2px; font-size: 11px; }
/* Drag-to-reorder feedback */
.thumb.dragging { opacity: .35; }
.thumb.drop-before { box-shadow: -3px 0 0 var(--accent-2); }
.thumb.drop-after { box-shadow: 3px 0 0 var(--accent-2); }

/* Slide panel head: title takes the slack so the view tabs and the "new slide" button sit right. */
.sp-head { gap: 4px; }
.sp-head > .sp-title { flex: 1 1 auto; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sp-tabs { display: inline-flex; gap: 2px; flex: 0 0 auto; }
.sp-tab {
  background: transparent; border: 1px solid transparent; border-radius: 4px;
  color: var(--muted); cursor: pointer; font-size: 11px; font-family: inherit; padding: 2px 5px;
}
.sp-tab:hover { background: var(--panel-2); }
.sp-tab.on { background: var(--accent-bg); border-color: var(--accent-2); color: var(--text); }

/* Outline view: one compact row per slide (number + title) — cheap to scroll through long decks. */
.sp-list.ol-mode { gap: 1px; padding: 4px; }
.ol-row {
  display: flex; align-items: flex-start; gap: 6px; padding: 4px 6px; border-radius: 4px;
  cursor: pointer; border: 1px solid transparent; font-size: 12px; line-height: 1.35;
}
.ol-row:hover { background: var(--panel-2); }
.ol-row.active { background: var(--accent-bg); border-color: var(--accent-2); }
.ol-row.hidden-slide { opacity: .55; }
.ol-row .ol-num {
  flex: 0 0 auto; min-width: 16px; text-align: right;
  color: var(--muted); font-variant-numeric: tabular-nums;
}
.ol-row.hidden-slide .ol-num { text-decoration: line-through; }
.ol-row .ol-text { flex: 1 1 auto; min-width: 0; word-break: break-word; }
.ol-row .ol-notes { flex: 0 0 auto; color: var(--muted); font-size: 10px; }

.canvas-area { flex: 1 1 auto; min-width: 0; display: flex; flex-direction: column; background: var(--vscode-editor-background, #161616); }
/* Center via margin:auto on the child (.stage) instead of the flex container:
   a centered flex item that overflows the viewport makes the top/left unreachable by scrolling. */
.canvas-scroll { flex: 1 1 auto; overflow: auto; display: flex; padding: 20px; }
/* Slide content should not inherit the editor UI font settings (font size / family would change
  text metrics and line wrapping; align with examples/index.html defaults: 16px + sans-serif stack) */
.slide {
  font-size: 16px;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif;
}

/* Outer size = scaled size (ensures correct scroll range and centering); inner is 1:1 layout then scaled as a whole */
.stage { position: relative; background: #fff; box-shadow: 0 4px 20px rgba(0,0,0,.4); overflow: hidden; flex: 0 0 auto; margin: auto; }
.stage-inner { position: absolute; top: 0; left: 0; transform-origin: top left; }
.slide-host { position: absolute; top: 0; left: 0; user-select: none; -webkit-user-select: none; }
/* Inline-editing needs native text selection on the contenteditable body (parent disables it) */
.tb-body[contenteditable="true"] { user-select: text; -webkit-user-select: text; }
.slide-host .slide { margin: 0 !important; }
.overlay { position: absolute; top: 0; left: 0; pointer-events: none; }
.grid-overlay { position: absolute; inset: 0; pointer-events: none; display: none; }
.grid-overlay.on { display: block; background-image: linear-gradient(rgba(0,0,0,.08) 1px, transparent 1px), linear-gradient(90deg, rgba(0,0,0,.08) 1px, transparent 1px); background-size: 40px 40px; }

/* ===== Rendered slide elements =====
   Ported from the parser's editor renderer (examples/editor/styles.css) so elements
   built by render.ts get the same box model / text metrics as the reference editor. */
.slide-frame { position: absolute; top: 0; left: 0; }
.slide-frame.show-grid {
  background-image: linear-gradient(to right, rgba(0,0,0,.06) 1px, transparent 1px),
                    linear-gradient(to bottom, rgba(0,0,0,.06) 1px, transparent 1px);
}
.el { position: absolute; box-sizing: border-box; }
.el.locked { cursor: default; }
.el.hidden-el { display: none; }
.el-image img { width: 100%; height: 100%; display: block; }
.el-media video { pointer-events: none; }
.el.is-sel .el-media video { pointer-events: auto; }
.el-media .media-badge { display: flex; align-items: center; justify-content: center; }
.el.is-sel .el-media .media-badge { display: none; }
.el-text .tb-body {
  width: 100%; height: 100%; display: flex; flex-direction: column;
  justify-content: flex-start; outline: none; word-break: break-word;
  position: relative; z-index: 1;
  /* a:bodyPr lIns/tIns are inner padding and must sit inside the element box, otherwise a
     short text frame (e.g. 23.6px tall) loses its content area and the text drifts down. */
  box-sizing: border-box;
}
.el-text .tb-body > div { min-height: 1em; flex-shrink: 0; }
/* Hanging indent via negative first-line indent, so wrapped lines align with the bullet text */
.el-text .tb-body > div.bullet-para { padding-left: 1.3em; text-indent: -1.3em; }
.el-text .tb-body > div.num-para { padding-left: 1.8em; text-indent: -1.8em; }
.el-text .tb-body > div .bullet-mark {
  display: inline-block; min-width: 1.2em; text-align: left;
  white-space: pre; margin-right: 0.2em;
}
.el-text .tb-body > div .bullet-img img { display: inline-block; max-width: none; object-fit: contain; }
.el-table { display: grid; }
.el-table .cell {
  border: 1px solid #cbd5e1; padding: 4px 6px; overflow: hidden; display: flex;
  align-items: flex-start; word-break: break-word; white-space: pre-wrap;
}
.el-table .cell.cell-top { align-items: flex-start; }
.el-table .cell.cell-middle { align-items: center; }
.el-table .cell.cell-bottom { align-items: flex-end; }
.el-chart svg { display: block; width: 100%; height: 100%; }
.el-raw {
  border: 1px dashed #9aa0a6; background: repeating-linear-gradient(45deg,#f8f9fa,#f8f9fa 8px,#eceff1 8px,#eceff1 16px);
  display: flex; align-items: center; justify-content: center; color: var(--muted); font-size: 12px;
  text-align: center; padding: 6px;
}
/* Group edit state: dim other top-level elements and outline the active group */
.slide-host.group-editing .el.dimmed { opacity: 0.2; pointer-events: none; }
.slide-host.group-editing .el.group-edit-active { outline: 2px dashed var(--accent-2); outline-offset: -2px; }
.group-hint { position: absolute; left: 50%; top: 8px; transform: translateX(-50%); z-index: 30;
  background: var(--accent); color: #fff; padding: 4px 12px; border-radius: 14px; font-size: 12px;
  cursor: pointer; box-shadow: 0 2px 8px rgba(0,0,0,.25); }
.el[data-editing="1"] .tb-body { outline: 2px solid var(--accent-2); }

/* Comment cards drawn on the canvas (matches the preview side's .pptx-comment styling) */
.slide-comment {
  position: absolute; width: 200px; box-sizing: border-box; z-index: 60;
  background: #fff7cc; border: 1px solid #f1d27a; border-radius: 8px;
  padding: 8px 10px; box-shadow: 0 2px 6px rgba(0,0,0,.15);
  font: 12px/1.4 sans-serif; color: #3b2f00; pointer-events: none;
}
.slide-comment .sc-author { font-weight: 700; margin-bottom: 2px; }
.slide-comment .sc-text { white-space: pre-wrap; }
.slide-comment .sc-date { margin-top: 4px; font-size: 10px; color: #8a7a3a; }

.sel-rect { position: absolute; border: 1px solid var(--accent); background: rgba(249,171,0,.08); pointer-events: auto; cursor: move; }
.sel-rect.selected { border-color: var(--accent-2); background: rgba(66,133,244,.1); }
.sel-rect.locked { border-style: dashed; cursor: default; }
.handle { position: absolute; width: 8px; height: 8px; background: #fff; border: 1px solid var(--accent-2); border-radius: 1px; }
.handle.nw { left: -4px; top: -4px; cursor: nwse-resize; }
.handle.n  { left: 50%; top: -4px; margin-left: -4px; cursor: ns-resize; }
.handle.ne { right: -4px; top: -4px; cursor: nesw-resize; }
.handle.e  { right: -4px; top: 50%; margin-top: -4px; cursor: ew-resize; }
.handle.se { right: -4px; bottom: -4px; cursor: nwse-resize; }
.handle.s  { left: 50%; bottom: -4px; margin-left: -4px; cursor: ns-resize; }
.handle.sw { left: -4px; bottom: -4px; cursor: nesw-resize; }
.handle.w  { left: -4px; top: 50%; margin-top: -4px; cursor: ew-resize; }
/* Rotation grip, positioned by JS above the top edge */
.handle.rot {
  left: 50%; margin-left: -5px; cursor: grab;
  width: 10px; height: 10px; border-radius: 50%;
  background: var(--accent-2); border-color: #fff;
}
/* Line endpoint handles (drag to resize/move an endpoint, snap to shapes as glue) */
.handle.endpoint {
  border-radius: 50%; background: var(--accent-2); border: 2px solid #fff;
  cursor: grab; pointer-events: auto; box-sizing: border-box;
}
.handle.endpoint:active { cursor: grabbing; }
/* Vertex points / bezier control points while in vertex-edit mode */
.handle.vertex {
  background: #fff; border: 1.5px solid var(--accent-2);
  cursor: move; pointer-events: auto; box-sizing: border-box;
}
.handle.vertex.sel { background: var(--accent-2); }
.handle.vctrl {
  border-radius: 50%; background: #fff; border: 1.5px solid var(--accent-2);
  cursor: move; pointer-events: auto; box-sizing: border-box;
}
/* Hit-area expansion: keeps the visible dot unchanged while enlarging the clickable region.
   Size is injected by JS as --hit (already divided by zoom → constant screen pixels). */
.handle::after {
  content: ''; position: absolute; left: 50%; top: 50%;
  width: var(--hit, 20px); height: var(--hit, 20px);
  transform: translate(-50%, -50%);
}
.lock-badge {
  position: absolute; top: -9px; left: -9px; font-size: 11px;
  background: var(--panel); border: 1px solid var(--border); border-radius: 4px; padding: 0 3px;
}
/* Snap guides while dragging */
.guide { position: absolute; background: #ea4335; pointer-events: none; }
.guide.h { height: 1px; left: 0; right: 0; }
.guide.v { width: 1px; top: 0; bottom: 0; }
/* Rubber-band selection box */
.marquee {
  position: absolute; border: 1px dashed var(--accent-2);
  background: rgba(66,133,244,.12); pointer-events: none;
}

/* ===== Modal dialogs (tables, charts, shapes, media, confirm, shortcuts) ===== */
.modal {
  position: fixed; inset: 0; background: rgba(0,0,0,.45);
  display: flex; align-items: center; justify-content: center; z-index: 95;
}
.dlg {
  background: var(--panel); border: 1px solid var(--border); border-radius: 8px;
  box-shadow: 0 12px 32px rgba(0,0,0,.35); min-width: 380px; max-width: 92vw;
  max-height: 86vh; display: flex; flex-direction: column;
}
.dlg-head {
  display: flex; align-items: center; justify-content: space-between; gap: 12px;
  padding: 8px 10px; border-bottom: 1px solid var(--border); font-weight: 600; font-size: 13px;
}
.dlg-x { cursor: pointer; opacity: .6; padding: 0 4px; background: none; border: 0; color: inherit; }
.dlg-x:hover { opacity: 1; }
.dlg-body { padding: 12px; overflow: auto; display: flex; flex-direction: column; gap: 10px; }
.dlg-msg { margin: 0; font-size: 13px; line-height: 1.5; }
.dlg-detail { margin: 0; font-size: 12px; color: var(--muted); }
.dlg-line { display: flex; align-items: center; gap: 8px; }
.dlg-line > span { width: 84px; flex: 0 0 auto; font-size: 12px; opacity: .8; }
.dlg-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.dlg-col { display: flex; flex-direction: column; gap: 6px; }
.dlg-block { border: 1px solid var(--border); border-radius: 6px; padding: 6px; margin-bottom: 6px; }
.dlg-split { display: flex; gap: 16px; }
.dlg-split > div { flex: 1 1 0; min-width: 0; }
.dlg-in, .dlg-cell {
  border: 1px solid var(--border); border-radius: 4px; padding: 3px 6px;
  font-size: 12px; background: var(--vscode-input-background, var(--bg)); color: inherit; min-width: 0;
}
.dlg-cell { width: 100%; box-sizing: border-box; }
.dlg-table { display: grid; gap: 4px; }
.dlg-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 6px; }
.dlg-check-row { display: flex; align-items: center; gap: 4px; font-size: 12px; }
.dlg-shapes { display: grid; grid-template-columns: repeat(6, 1fr); gap: 6px; }
.shape-item {
  display: flex; flex-direction: column; align-items: center; gap: 2px;
  padding: 6px 2px; border: 1px solid var(--border); border-radius: 6px;
  background: transparent; color: inherit; cursor: pointer; font-size: 10px;
}
.shape-item:hover { border-color: var(--accent-2); background: var(--accent-bg); }
.dlg-group-title { font-size: 11px; color: var(--muted); font-weight: 600; margin-top: 2px; }
.shortcuts h4 { margin: 10px 0 4px; font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: .04em; }
.shortcuts table { width: 100%; border-collapse: collapse; font-size: 12px; }
.shortcuts td { padding: 2px 4px; border-bottom: 1px solid var(--border); vertical-align: top; }
.shortcuts td.key { width: 34%; color: var(--accent); font-family: ui-monospace, SFMono-Regular, Menlo, monospace; white-space: nowrap; }

/* ===== Right-side properties panel ===== */
/* The panel is one scroll container; only the header is sticky, so long property lists scroll
   under it and the "elements" toggle is always reachable. */
.inspector { width: 290px; background: var(--panel); border-left: 1px solid var(--border); overflow-y: auto; padding: 0 10px 10px; flex: 0 0 auto; }
.inspector.hidden { display: none; }
.inspector h3 { font-size: 12px; margin: 0 0 8px; color: var(--accent); text-transform: uppercase; letter-spacing: .04em; }
.inspector h4 { font-size: 11px; margin: 12px 0 6px; color: var(--muted); }
.insp-head {
  position: sticky; top: 0; z-index: 3; background: var(--panel);
  margin: 0 -10px; padding: 8px 10px 7px; border-bottom: 1px solid var(--border);
  display: flex; align-items: center; gap: 8px;
}
.insp-head h3 { flex: 1 1 auto; min-width: 0; margin: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.insp-toggle {
  flex: 0 0 auto; display: inline-flex; align-items: center; gap: 3px;
  background: var(--panel-2); color: var(--text); border: 1px solid var(--border);
  border-radius: 4px; padding: 2px 6px; cursor: pointer; font-size: 11px; font-family: inherit;
}
.insp-toggle:hover { background: var(--panel-3); }
.insp-toggle.on { background: var(--accent-bg); border-color: var(--accent-2); }
.insp-toggle .rb-ico svg { transition: transform .12s ease; }
.insp-toggle.on .rb-ico svg { transform: rotate(180deg); }
.el-kind { font-size: 10px; text-transform: uppercase; letter-spacing: .04em; margin: -4px 0 8px; }
.field { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; }
.field label { width: 72px; color: var(--muted); flex: 0 0 auto; font-size: 11px; }
.field input[type=text], .field input[type=number], .field select, .field textarea {
  flex: 1 1 auto; width: 100%; background: var(--vscode-input-background, var(--panel-2));
  border: 1px solid var(--border); color: var(--text);
  border-radius: 4px; padding: 4px 6px; font-size: 12px; font-family: inherit;
}
.field input[type=color] { width: 36px; height: 26px; padding: 0; border: 1px solid var(--border); background: none; }
.field textarea { resize: vertical; min-height: 60px; }
.row { display: flex; gap: 6px; flex-wrap: wrap; }
.row .btn { flex: 1 1 auto; justify-content: center; }
.anim-row { padding: 4px 6px; border: 1px solid transparent; border-radius: 4px; align-items: center; }
.anim-row:hover { background: var(--panel-2); }
.anim-row.on { border-color: var(--accent-2); background: var(--panel-2); }
.chk { display: flex; align-items: center; gap: 6px; }
.el-list { list-style: none; margin: 0; padding: 0; }
.el-list li { padding: 5px 8px; border: 1px solid var(--border); border-radius: 4px; margin-bottom: 4px; cursor: pointer; display: flex; justify-content: space-between; gap: 6px; }
.el-list li:hover { background: var(--panel-2); }
.el-list li.active { border-color: var(--accent-2); background: var(--panel-2); }
.el-list li .tag { color: var(--muted); font-size: 11px; }
.el-list li .name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.el-list li.empty { border-color: transparent; cursor: default; }
/* Drawer for the element list: its own scroll box, so a slide with dozens of elements can never
   push the property fields below the fold. */
.el-drawer { border: 1px solid var(--border); border-radius: 4px; padding: 6px; margin: 8px 0 2px; }
.el-drawer .el-list { max-height: 32vh; overflow-y: auto; }
.el-filter {
  width: 100%; box-sizing: border-box; margin-bottom: 6px; font-family: inherit; font-size: 12px;
  background: var(--vscode-input-background, var(--panel-2)); color: var(--text);
  border: 1px solid var(--border); border-radius: 4px; padding: 3px 6px;
}
.empty { color: var(--muted); font-size: 12px; padding: 8px 0; }

/* ===== Status bar ===== */
.statusbar { display: flex; align-items: center; gap: 8px; padding: 3px 10px; background: var(--panel); border-top: 1px solid var(--border); font-size: 12px; color: var(--muted); flex: 0 0 auto; }
.statusbar .spacer { flex: 1 1 auto; }
.chip {
  background: var(--panel-2); border: 1px solid var(--border); border-radius: 4px;
  padding: 2px 8px; cursor: pointer; display: inline-flex; align-items: center; gap: 3px;
}
.chip:hover { background: var(--panel-3); }
.chip.on { background: var(--accent-bg); border-color: var(--accent-2); color: var(--text); }

/* ===== Presentation mode ===== */
.present { position: fixed; inset: 0; background: #000; display: none; align-items: center; justify-content: center; z-index: 88; }
.present.on { display: flex; }
.present .stage { box-shadow: none; background: #000; }
.present .slide-host { position: relative; transform-origin: center center; overflow: hidden; flex: none; }
/* Stacked layers let the outgoing + incoming slides animate simultaneously for transitions. */
.present-layer { position: absolute; inset: 0; overflow: hidden; will-change: transform, opacity; }
.present .slide-host .slide { margin: 0 !important; }
.present-ui {
  position: absolute; left: 0; right: 0; bottom: 0; display: flex; align-items: center;
  justify-content: flex-end; gap: 12px; padding: 10px 16px; font-size: 12px; color: #cfcfcf;
  background: linear-gradient(transparent, rgba(0,0,0,.55)); opacity: 0; transition: opacity .2s;
}
.present:hover .present-ui { opacity: 1; }
.present-count { font-variant-numeric: tabular-nums; }
.present-hint { opacity: .7; }
.present-exit { background: rgba(255,255,255,.12); border-color: transparent; color: #fff; }

/* ===== Document properties panel (available in both preview and edit mode) ===== */
.doc-info {
  position: fixed; top: 40px; right: 12px; width: 330px; max-height: 70vh; overflow: auto;
  background: var(--panel); border: 1px solid var(--border); border-radius: 6px; padding: 10px;
  display: none; z-index: 85; box-shadow: 0 6px 24px rgba(0,0,0,.45);
}
.doc-info.on { display: block; }
.doc-info-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px; font-weight: 600; color: var(--accent); }
.doc-info h4 { font-size: 11px; margin: 10px 0 4px; color: var(--muted); }
.doc-info table.kv { width: 100%; border-collapse: collapse; font-size: 12px; }
.doc-info table.kv td { padding: 3px 4px; border-bottom: 1px solid var(--border); vertical-align: top; word-break: break-all; }
.doc-info table.kv td.k { width: 92px; color: var(--muted); }

/* ===== Toast ===== */
.toast { position: fixed; bottom: 44px; left: 50%; transform: translateX(-50%); background: #333; color: #fff; padding: 6px 14px; border-radius: 6px; font-size: 12px; opacity: 0; transition: opacity .2s; pointer-events: none; z-index: 96; }
.toast.show { opacity: 1; }
.toast.error { background: var(--danger); }
.toast.warn { background: #b89500; }

.muted { color: var(--muted); }
`;