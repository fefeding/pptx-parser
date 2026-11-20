// Message protocol between extension host (Node) and webview (browser).
//
// The webview owns the document model and all editing (via @fefeding/ppt-parser's
// editor core: createStore / createActions / docFromPptx / docToPptx), so the host only
// moves PPTX bytes and tracks the dirty state. Both payloads are base64: VS Code's webview
// typed-array transport ($$vscode_array_buffer_reference$$) does not restore Uint8Array in
// custom editors, passing binary directly makes parsing fail.

export type EditorMode = 'preview' | 'edit';

/** Webview → Host */
export type WebviewToHost =
  | { type: 'ready' }
  | { type: 'setMode'; mode: EditorMode }
  /**
   * Editor produced new PPTX bytes (edit mode). `bytes` is null when the edit mode was
   * left without any edit (e.g. undo back to the imported state), meaning "keep original".
   */
  | { type: 'sync'; bytes: string | null; dirty: boolean }
  /** Reply to `flush` when there is nothing to serialize (e.g. preview mode). */
  | { type: 'flushDone' }
  | { type: 'save' }
  | { type: 'saveAs'; bytes: string | null }
  | { type: 'revert' }
  /** Save one rendered slide (PNG, base64) — the host owns file dialogs and the filesystem. */
  | { type: 'exportImage'; bytes: string; suggestedName: string }
  /** Save every rendered slide into a picked folder. */
  | { type: 'exportImages'; images: { name: string; bytes: string }[] }
  /** Host assembles the rendered slide PNGs into a single PDF (pdf-lib lives host-side only). */
  | { type: 'exportPdf'; images: { bytes: string }[]; widthPx: number; heightPx: number; suggestedName: string }
  /** Ribbon "Convert Document": ask the host to run the convert command for the active document. */
  | { type: 'convert' }
  /** Surface a webview-side problem (parse/serialize failure) in the host UI. */
  | { type: 'error'; message: string };

/** Host → Webview */
export type HostToWebview =
  | {
    type: 'init';
    mode: EditorMode;
    /** Original file bytes (base64) — preview mode renders these faithfully. */
    originalBytes: string;
    /** Latest bytes (base64); equals originalBytes until the webview reports an edit. */
    bytes: string;
    title: string;
    /** True when the document has unsaved changes (drives the webview's dirty dot). */
    dirty: boolean;
    /**
     * VS Code display language (e.g. `zh-cn`, `en`). The webview also reads `<html lang>`,
     * which the host fills with the same value; this field is the authoritative fallback.
     */
    locale: string;
  }
  | { type: 'bytes'; bytes: string; dirty: boolean }
  /** Push a preview/edit mode change from the editor title bar (host command) into the webview. */
  | { type: 'setMode'; mode: EditorMode }
  | { type: 'saved' }
  /** Save is about to happen: serialize any pending edit immediately (beats the 350ms debounce). */
  | { type: 'flush' }
  | { type: 'reverted' }
  | { type: 'info'; message: string; kind?: 'info' | 'warn' | 'error' };