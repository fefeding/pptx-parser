import * as vscode from 'vscode';
import type { EditorMode } from './protocol';

/**
 * PPTX custom document: a thin byte holder for the custom editor.
 *
 * All parsing/model/editing lives in the webview (browser environment) using
 * @fefeding/ppt-parser's editor core, so the host never needs a DOM and never touches the
 * document model. Responsibilities are limited to: read the file, keep the latest bytes the
 * webview produced, and drive save / revert / backup / dirty bookkeeping.
 *
 * In newer @types/vscode, CustomDocument is an interface (only needs uri + dispose),
 * so we use implements rather than extends.
 */
export class PptxCustomDocument implements vscode.CustomDocument {
  private readonly _onDidChange = new vscode.EventEmitter<{ contentChanged: boolean }>();
  public readonly onDidChange = this._onDidChange.event;

  private readonly _onDidRevert = new vscode.EventEmitter<void>();
  public readonly onDidRevert = this._onDidRevert.event;

  private readonly _onDidSave = new vscode.EventEmitter<void>();
  public readonly onDidSave = this._onDidSave.event;

  public readonly uri: vscode.Uri;

  private _originalBytes!: Uint8Array; // file bytes as loaded from disk
  private _currentBytes!: Uint8Array; // latest bytes (webview output after edits, else original)
  private _dirty = false;
  private _mode: EditorMode = 'preview';

  private constructor(uri: vscode.Uri) {
    this.uri = uri;
  }

  static async create(uri: vscode.Uri): Promise<PptxCustomDocument> {
    const doc = new PptxCustomDocument(uri);
    const bytes = await vscode.workspace.fs.readFile(uri);
    doc._originalBytes = bytes;
    doc._currentBytes = bytes;
    return doc;
  }

  get mode(): EditorMode {
    return this._mode;
  }
  set mode(m: EditorMode) {
    this._mode = m;
  }

  get isDirty(): boolean {
    return this._dirty;
  }

  /**
   * Accept the PPTX bytes produced by the webview editor.
   * `bytes === null` means the editor discarded its edits (back to the imported state).
   */
  setBytes(bytes: string | null): void {
    this._currentBytes = bytes ? Buffer.from(bytes, 'base64') : this._originalBytes;
    this.setDirty(true);
  }

  setDirty(dirty: boolean): void {
    if (this._dirty === dirty) return;
    this._dirty = dirty;
    this._onDidChange.fire({ contentChanged: dirty });
  }

  /** Payload for the webview: original bytes (faithful preview) + latest bytes. */
  snapshot(): { mode: EditorMode; originalBytes: string; bytes: string } {
    return {
      mode: this._mode,
      originalBytes: Buffer.from(this._originalBytes).toString('base64'),
      bytes: Buffer.from(this._currentBytes).toString('base64')
    };
  }

  // ---- Save / Revert / Backup ----
  async save(targetUri?: vscode.Uri): Promise<void> {
    const dest = targetUri || this.uri;
    await vscode.workspace.fs.writeFile(dest, this._currentBytes);
    if (!targetUri) {
      this._originalBytes = this._currentBytes;
      this.setDirty(false);
      this._onDidSave.fire();
    }
  }

  async revert(): Promise<void> {
    const bytes = await vscode.workspace.fs.readFile(this.uri);
    this._originalBytes = bytes;
    this._currentBytes = bytes;
    this.setDirty(false);
    this._onDidRevert.fire();
  }

  async backup(destination: vscode.Uri): Promise<void> {
    await vscode.workspace.fs.writeFile(destination, this._currentBytes);
  }

  dispose(): void {
    this._onDidChange.dispose();
    this._onDidRevert.dispose();
    this._onDidSave.dispose();
  }
}