// Host-side glue for the PPTX format conversion command.
//
// The conversion engine lives in `src/jdoc/` (a trimmed copy of jdoc's PPTX conversion
// chain). It runs in the extension host rather than the webview because:
//   * generating a Chinese PDF needs system fonts (jdoc's autoFont probing uses node:fs);
//   * the host already holds the latest PPTX bytes, including unsaved webview edits;
//   * the webview bundle stays a single self-contained IIFE (no extra worker/CSP concerns).

import { convertPptx, type ConvertTarget } from './jdoc';

export type { ConvertTarget };

export interface ConvertTargetMeta {
  id: ConvertTarget;
  /** vscode.l10n key used for the QuickPick label. */
  labelKey: string;
  /** English fallback (vscode.l10n.t echoes the key when no bundle is installed). */
  label: string;
  /** Output file extension. */
  ext: string;
  /** showSaveDialog filters. */
  filters: Record<string, string[]>;
}

export const CONVERT_TARGETS: readonly ConvertTargetMeta[] = [
  { id: 'docx', labelKey: 'extension.target.docx', label: 'Word Document (.docx)', ext: 'docx', filters: { Word: ['docx'] } },
  { id: 'pdf', labelKey: 'extension.target.pdf', label: 'PDF Document (.pdf)', ext: 'pdf', filters: { PDF: ['pdf'] } },
  { id: 'md', labelKey: 'extension.target.md', label: 'Markdown (.md)', ext: 'md', filters: { Markdown: ['md', 'markdown'] } },
  { id: 'html', labelKey: 'extension.target.html', label: 'HTML Web Page (.html)', ext: 'html', filters: { HTML: ['html', 'htm'] } }
];

/** Only PPTX is a supported source for this port. */
export function isPptxPath(filePath: string): boolean {
  return /\.pptx$/i.test(filePath);
}

/** Guard against a stalled generator turning into a "clicked but nothing happened" state. */
function withTimeout<T>(promise: Promise<T>, ms: number): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<T>((_, reject) => {
    timer = setTimeout(() => reject(new Error(`Conversion timed out after ${Math.round(ms / 1000)}s`)), ms);
  });
  return Promise.race([promise, timeout]).finally(() => {
    if (timer) clearTimeout(timer);
  });
}

/**
 * Convert PPTX bytes to docx / pdf / md / html.
 * Binary targets return Uint8Array; text targets return a string.
 */
export function convertPptxBytes(
  bytes: Uint8Array,
  target: ConvertTarget,
  title: string
): Promise<Uint8Array | string> {
  return withTimeout(convertPptx(bytes, target, { title, autoFont: true }), 60_000);
}
