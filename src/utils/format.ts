/**
 * 字符串 / dataURL 工具。下沉自编辑器 util.js（纯逻辑、无 DOM 依赖）。
 */

/** HTML 转义 */
export function escapeHtml(s: unknown): string {
  return String(s == null ? '' : s).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string));
}

/** dataURL → 裸 base64 */
export function stripDataUrl(dataUrl: string): string {
  if (typeof dataUrl !== 'string') return '';
  const i = dataUrl.indexOf(',');
  return i >= 0 ? dataUrl.slice(i + 1) : dataUrl;
}

/** dataURL → 扩展名（jpg/png/svg/...），无法识别回退 fallback */
export function extOfDataUrl(dataUrl: string, fallback = 'png'): string {
  const m = /^data:image\/([a-zA-Z0-9.+-]+)/.exec(String(dataUrl || ''));
  if (!m) return fallback;
  const e = m[1].toLowerCase();
  return e === 'jpeg' ? 'jpg' : e === 'svg+xml' ? 'svg' : e;
}
