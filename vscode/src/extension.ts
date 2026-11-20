import * as vscode from 'vscode';
import { PDFDocument } from 'pdf-lib';
import { PptxCustomDocument } from './pptx-document';
import { CONVERT_TARGETS, convertPptxBytes, isPptxPath } from './convert';
import type { ConvertTargetMeta } from './convert';
import type { EditorMode, HostToWebview, WebviewToHost } from './protocol';

const VIEW_TYPE = 'pptx-parser.pptxEditor';

/**
 * Localized host-side string. `vscode.l10n.t` echoes the key back when no bundle is installed
 * (or the locale is untranslated), so the English fallback is kept as a second argument.
 */
function loc(key: string, fallback: string, ...args: string[]): string {
  const value = vscode.l10n.t(key, ...args);
  return !value || value === key ? fallback.replace('{0}', args[0] ?? '') : value;
}

function getNonce(): string {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}

class PptxEditorProvider implements vscode.CustomEditorProvider<PptxCustomDocument> {
  public static readonly viewType = VIEW_TYPE;

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<
    vscode.CustomDocumentEditEvent<PptxCustomDocument> | vscode.CustomDocumentContentChangeEvent<PptxCustomDocument>
  >();
  public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  /** The currently focused document (used by commands like "Export Copy"). */
  private _activeDoc: PptxCustomDocument | undefined;

  /** The webview panel of the currently focused document (so title-bar commands can message it). */
  private _activePanel: vscode.WebviewPanel | undefined;

  /** Resolver for a pending "serialize now, before saving" handshake. */
  private _pendingFlush: (() => void) | null = null;

  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(
    uri: vscode.Uri,
    _openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken
  ): Promise<PptxCustomDocument> {
    const document = await PptxCustomDocument.create(uri);

    // Notify the editor when document content changes (for dirty indicator / save prompt)
    document.onDidChange(() => this._onDidChangeCustomDocument.fire({ document }));
    document.onDidRevert(() => this._onDidChangeCustomDocument.fire({ document }));
    document.onDidSave(() => this._onDidChangeCustomDocument.fire({ document }));
    return document;
  }

  async resolveCustomEditor(
    document: PptxCustomDocument,
    webviewPanel: vscode.WebviewPanel,
    _token: vscode.CancellationToken
  ): Promise<void> {
    webviewPanel.webview.options = {
      enableScripts: true,
      enableCommandUris: false,
      localResourceRoots: [this.context.extensionUri]
    };
    webviewPanel.webview.html = this.getWebviewHtml(webviewPanel.webview);

    // Not every host fires onDidChangeViewState when a panel opens. Register right away so
    // title-bar / ribbon commands can resolve the active document on the first click.
    this._activeDoc = document;
    this._activePanel = webviewPanel;

    webviewPanel.onDidChangeViewState((e) => {
      if (e.webviewPanel.active) {
        this._activeDoc = document;
        this._activePanel = webviewPanel;
        this.syncContext();
      }
    });

    const send = (msg: HostToWebview) => webviewPanel.webview.postMessage(msg);
    const sendInit = async () => {
      const s = document.snapshot();
      send({
        type: 'init',
        mode: s.mode,
        originalBytes: s.originalBytes,
        bytes: s.bytes,
        title: document.uri.path.split('/').pop() || 'presentation.pptx',
        dirty: document.isDirty,
        locale: vscode.env.language
      });
    };

    // Surface the host's real dirty state so the webview's dot matches VS Code's own save/revert.
    // The webview no longer implements saving (VS Code handles Ctrl+S / File › Save), so the
    // authoritative dirty flag lives here and is pushed down on every change.
    document.onDidSave(() => send({ type: 'saved' }));
    document.onDidRevert(() => {
      void sendInit();
      send({ type: 'reverted' });
    });

    // Register listener before sending, so the webview never misses a message if the host throws
    webviewPanel.webview.onDidReceiveMessage(async (message: WebviewToHost) => {
      try {
        switch (message.type) {
          case 'ready':
            await sendInit();
            break;
          case 'setMode':
            document.mode = message.mode;
            // Keep the `pptx:editing` context key in sync with the active document's mode.
            if (this._activeDoc === document) this.syncContext();
            break;
          case 'sync':
            document.setBytes(message.bytes);
            document.setDirty(message.dirty);
            this._pendingFlush?.();
            break;
          case 'flushDone':
            this._pendingFlush?.();
            break;
          case 'save':
            // 走 VS Code 标准保存：它会调用 saveCustomDocument（含保存前的 flush 握手），
            // 并正确清除标题栏的未保存标记。直接 document.save() 会绕过 VS Code，导致
            // 标题栏仍显示未保存。写盘完成后 onDidSave 会把 `saved` 推回 webview。
            await vscode.commands.executeCommand('workbench.action.files.save');
            break;
          case 'saveAs':
            await this.saveAs(document, message.bytes);
            break;
          case 'exportImage':
            await this.saveImageFile(document, message);
            break;
          case 'exportImages':
            await this.saveImageFiles(document, message);
            break;
          case 'exportPdf':
            await this.savePdf(document, message);
            break;
          case 'convert':
            // Ribbon button: always convert the document this panel belongs to, never
            // guess from activeTextEditor (which may be a different file).
            await this.convertUri(document.uri);
            break;
          case 'revert':
            await document.revert();
            await sendInit();
            send({ type: 'info', kind: 'info', message: loc('extension.reloaded', 'Reloaded') });
            break;
          case 'error':
            send({ type: 'info', kind: 'error', message: message.message });
            break;
        }
      } catch (err: any) {
        send({ type: 'info', kind: 'error', message: err?.message || String(err) });
      }
    });

    // First send: failures only show a toast in the webview; don't let the entire editor stall on an empty view
    try {
      await sendInit();
    } catch (err: any) {
      send({ type: 'info', kind: 'error', message: loc('extension.initFailed', 'Initialization failed: {0}', String(err?.message || err)) });
    }
  }

  // ---- CustomEditorProvider required methods ----
  /**
   * Edits are pushed to the host with a 350ms debounce plus async PPTX serialization. Without
   * this handshake, Ctrl+S right after an edit can save the *previous* bytes (e.g. an animation
   * added a moment ago is absent after reopening). Ask the webview to flush and wait for its
   * `sync` reply (bounded by a timeout so saving never hangs).
   */
  private async flushPendingSync(document: PptxCustomDocument): Promise<void> {
    const panel = this._activeDoc === document ? this._activePanel : undefined;
    if (!panel) return;
    await new Promise<void>((resolve) => {
      const timer = setTimeout(() => {
        this._pendingFlush = null;
        resolve();
      }, 2000);
      this._pendingFlush = () => {
        clearTimeout(timer);
        this._pendingFlush = null;
        resolve();
      };
      void panel.webview.postMessage({ type: 'flush' } as HostToWebview);
    });
  }

  async saveCustomDocument(document: PptxCustomDocument, _token: vscode.CancellationToken): Promise<void> {
    await this.flushPendingSync(document);
    await document.save();
  }
  async saveCustomDocumentAs(
    document: PptxCustomDocument,
    target: vscode.Uri,
    _token: vscode.CancellationToken
  ): Promise<void> {
    await document.save(target);
  }
  async revertCustomDocument(document: PptxCustomDocument, _token: vscode.CancellationToken): Promise<void> {
    await document.revert();
  }
  async backupCustomDocument(
    document: PptxCustomDocument,
    context: vscode.CustomDocumentBackupContext,
    _token: vscode.CancellationToken
  ): Promise<vscode.CustomDocumentBackup> {
    await document.backup(context.destination);
    return {
      id: context.destination.toString(),
      delete: () => {
        vscode.workspace.fs.delete(context.destination).then(
          () => undefined,
          () => undefined
        );
      }
    };
  }

  private async saveAs(document: PptxCustomDocument, bytes?: string | null): Promise<void> {
    const uri = await vscode.window.showSaveDialog({
      defaultUri: document.uri,
      saveLabel: loc('extension.exportCopy', 'Export Copy'),
      filters: { 'PowerPoint': ['pptx'] }
    });
    if (!uri) return;
    // Adopt the webview's latest bytes when provided, then export (saving to another
    // targetUri intentionally leaves the dirty flag untouched)
    if (bytes) document.setBytes(bytes);
    await document.save(uri);
    vscode.window.showInformationMessage(loc('extension.exported', 'Exported copy: {0}', uri.fsPath));
  }

  // ---- Export targets (PNG / PDF): the webview renders slides offscreen, the host writes files ----

  private async saveImageFile(
    document: PptxCustomDocument,
    msg: { bytes: string; suggestedName: string }
  ): Promise<void> {
    const uri = await vscode.window.showSaveDialog({
      defaultUri: vscode.Uri.joinPath(document.uri, '..', msg.suggestedName || 'slide.png'),
      filters: { 'PNG': ['png'] }
    });
    if (!uri) return;
    await vscode.workspace.fs.writeFile(uri, Buffer.from(msg.bytes, 'base64'));
    vscode.window.showInformationMessage(loc('extension.imageSaved', 'Image saved: {0}', uri.fsPath));
  }

  private async saveImageFiles(
    document: PptxCustomDocument,
    msg: { images: { name: string; bytes: string }[] }
  ): Promise<void> {
    const folders = await vscode.window.showOpenDialog({
      canSelectFiles: false,
      canSelectFolders: true,
      canSelectMany: false,
      defaultUri: vscode.Uri.joinPath(document.uri, '..'),
      openLabel: loc('extension.chooseFolder', 'Choose Folder')
    });
    if (!folders?.length) return;
    const dir = folders[0];
    for (const img of msg.images) {
      await vscode.workspace.fs.writeFile(vscode.Uri.joinPath(dir, img.name), Buffer.from(img.bytes, 'base64'));
    }
    vscode.window.showInformationMessage(
      loc('extension.imagesSaved', 'Saved {0} images to {1}', String(msg.images.length), dir.fsPath)
    );
  }

  private async savePdf(
    document: PptxCustomDocument,
    msg: { images: { bytes: string }[]; widthPx: number; heightPx: number; suggestedName: string }
  ): Promise<void> {
    if (!msg.images?.length) return;
    const pdf = await PDFDocument.create();
    // px → pt (1px = 0.75pt, the same conversion the parser uses elsewhere)
    const w = (msg.widthPx || 1280) * 0.75;
    const h = (msg.heightPx || 720) * 0.75;
    for (const img of msg.images) {
      const embedded = await pdf.embedPng(Buffer.from(img.bytes, 'base64'));
      const page = pdf.addPage([w, h]);
      page.drawImage(embedded, { x: 0, y: 0, width: w, height: h });
    }
    const uri = await vscode.window.showSaveDialog({
      defaultUri: vscode.Uri.joinPath(document.uri, '..', msg.suggestedName || 'slides.pdf'),
      filters: { 'PDF': ['pdf'] }
    });
    if (!uri) return;
    await vscode.workspace.fs.writeFile(uri, Buffer.from(await pdf.save()));
    vscode.window.showInformationMessage(loc('extension.pdfSaved', 'PDF saved: {0}', uri.fsPath));
  }

  public async revertActive(): Promise<void> {
    if (this._activeDoc) await this._activeDoc.revert();
  }
  public async exportActive(): Promise<void> {
    if (this._activeDoc) await this.saveAs(this._activeDoc);
  }

  /** 刷新编辑器标题栏切换所需的上下文键（pptx:editing）。 */
  private syncContext(): void {
    void vscode.commands.executeCommand('setContext', 'pptx:editing', this._activeDoc?.mode === 'edit');
  }

  /** 标题栏「预览 / 编辑」命令：把模式推给当前活动的 webview。 */
  public async setActiveMode(mode: EditorMode): Promise<void> {
    const doc = this._activeDoc;
    if (!doc) return;
    doc.mode = mode;
    this.syncContext();
    await this._activePanel?.webview.postMessage({ type: 'setMode', mode } as HostToWebview);
  }

  /** 标题栏切换按钮：在预览 / 编辑之间翻转当前活动文档的模式。 */
  public async toggleActiveMode(): Promise<void> {
    const doc = this._activeDoc;
    if (!doc) return;
    await this.setActiveMode(doc.mode === 'edit' ? 'preview' : 'edit');
  }

  // ---- Document format conversion (PPTX -> docx / pdf / md / html) ----

  /**
   * Convert a PPTX file into another document format.
   *
   * `uri` comes from the Explorer / editor context menus. When invoked from the command
   * palette or the ribbon we fall back to the active custom editor first (a custom editor
   * has no activeTextEditor) and then to the active text editor. For the active custom
   * editor we use its in-memory bytes so unsaved edits are reflected in the output.
   */
  public async convertUri(uri?: vscode.Uri): Promise<void> {
    const active = this._activeDoc;
    const target = uri || active?.uri || vscode.window.activeTextEditor?.document.uri;
    if (!target) {
      vscode.window.showInformationMessage(
        loc('extension.pickConvert', 'Open a PPTX file or select one in the Explorer to convert it.')
      );
      return;
    }
    if (!isPptxPath(target.path)) {
      vscode.window.showWarningMessage(
        loc('extension.unsupported', 'Document conversion currently supports .pptx files only.')
      );
      return;
    }

    const sourceName = target.path.split('/').pop() || 'presentation.pptx';
    const base = sourceName.replace(/\.[^.]+$/, '');
    const items = CONVERT_TARGETS.map((meta) => ({ label: loc(meta.labelKey, meta.label), meta }));
    const picked = await vscode.window.showQuickPick(items, {
      placeHolder: loc('extension.convertTo', 'Convert "{0}" to which format?', sourceName),
      title: loc('extension.convertTitle', 'PPTX: Convert Document Format')
    });
    if (!picked) return;

    const meta: ConvertTargetMeta = picked.meta;
    await vscode.window.withProgress(
      { location: vscode.ProgressLocation.Notification, title: loc('extension.converting', 'Converting document…') },
      async () => {
        try {
          const fromMemory = !!active && target.toString() === active.uri.toString();
          const bytes = fromMemory
            ? Uint8Array.from(Buffer.from(active!.snapshot().bytes, 'base64'))
            : await vscode.workspace.fs.readFile(target);
          const result = await convertPptxBytes(bytes, meta.id, base);
          const dir = target.path.replace(/[^/]+$/, '');
          const outUri = await vscode.window.showSaveDialog({
            defaultUri: target.with({ path: dir + base + '.' + meta.ext }),
            saveLabel: loc('extension.saveResult', 'Save converted file'),
            filters: meta.filters
          });
          if (!outUri) return;
          const output = typeof result === 'string' ? new TextEncoder().encode(result) : result;
          await vscode.workspace.fs.writeFile(outUri, output);
          vscode.window.showInformationMessage(
            loc('extension.converted', 'Converted and saved: {0}', outUri.fsPath)
          );
        } catch (err: any) {
          vscode.window.showErrorMessage(
            loc('extension.convertFailed', 'Conversion failed: {0}', String(err?.message || err))
          );
        }
      }
    );
  }

  private getWebviewHtml(webview: vscode.Webview): string {
    const scriptUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview.js')
    );
    // Chart dependencies: echarts / echarts-gl loaded as global scripts (order-sensitive, must precede webview.js)
    const echartsUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'media', 'echarts.min.js')
    );
    const echartsGlUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'media', 'echarts-gl.min.js')
    );
    // Base layout skeleton (.block/.content etc.) now comes from the parser's styles.global,
    // injected by the webview via ensureGlobalStyles().
    const nonce = getNonce();
    const csp = [
      `default-src 'none';`,
      `img-src ${webview.cspSource} data: blob:;`,
      `media-src ${webview.cspSource} data: blob:;`,
      `style-src ${webview.cspSource} 'unsafe-inline';`,
      // 'unsafe-eval' is required by echarts-gl: it uses new Function() to parse expr() expressions,
      // otherwise 3D charts throw "Invalid expression." (CSP blocks it and the real cause is swallowed by catch)
      `script-src 'nonce-${nonce}' 'unsafe-eval' ${webview.cspSource};`
    ].join(' ');
    return `<!DOCTYPE html>
<html lang="${vscode.env.language}">
<head>
<meta charset="utf-8" />
<meta http-equiv="Content-Security-Policy" content="${csp}" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>PPTX Editor</title>
</head>
<body>
<div id="app"></div>
<script nonce="${nonce}" src="${echartsUri}"></script>
<script nonce="${nonce}" src="${echartsGlUri}"></script>
<script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
  }
}

export function activate(context: vscode.ExtensionContext): void {
  const provider = new PptxEditorProvider(context);
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(VIEW_TYPE, provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: false
    })
  );

  // Seed `pptx:editing` with a concrete boolean so the context key is never undefined.
  void vscode.commands.executeCommand('setContext', 'pptx:editing', false);

  context.subscriptions.push(
    vscode.commands.registerCommand('pptx-parser.revert', () => provider.revertActive()),
    vscode.commands.registerCommand('pptx-parser.exportCopy', () => provider.exportActive()),
    vscode.commands.registerCommand('pptx-parser.viewMode', () => provider.setActiveMode('preview')),
    vscode.commands.registerCommand('pptx-parser.editMode', () => provider.setActiveMode('edit')),
    vscode.commands.registerCommand('pptx-parser.toggleMode', () => provider.toggleActiveMode()),
    vscode.commands.registerCommand('pptx-parser.convert', (uri?: vscode.Uri) => provider.convertUri(uri))
  );
}

export function deactivate(): void {
  // no-op
}
