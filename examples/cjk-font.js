/**
 * 浏览器端中文 PDF 字体辅助。
 *
 * pdf-lib 的 14 种标准字体不含 CJK 字形，浏览器又不能像 Node 那样探测系统字体。
 * 这里按 PPTX 中实际出现的汉字向 Google Fonts 请求一个只含这些字形的子集字体
 * （通常几十 KB），避免下载完整中文字体，也避免把整份字体嵌入 PDF。
 *
 * 网络不可用时调用方应回退到用户手动上传 TTF/OTF。
 */
const CJK_RE = /[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF]/g;
const FONT_CSS_URL = 'https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400&text=';
// encodeURIComponent 会把每个汉字转成 9 个字符，太长会被浏览器/服务端拒绝
const MAX_CHARS = 800;

const fontCache = new Map();
const pending = new Map();

/** 从任意文本（JSON 字符串亦可）中提取去重后的汉字集合 */
export function collectCjkChars(text) {
  return Array.from(new Set(String(text || '').match(CJK_RE) || [])).join('');
}

/** 加载覆盖 chars 的字体字节；无汉字时返回 null */
export async function loadCjkFont(chars) {
  if (!chars) return null;
  if (fontCache.has(chars)) return fontCache.get(chars);
  if (chars.length > MAX_CHARS) {
    throw new Error(`文稿含 ${chars.length} 个不同汉字，超出自动字体上限（${MAX_CHARS}），请手动上传 TTF/OTF 字体`);
  }
  if (pending.has(chars)) return pending.get(chars);

  const task = (async () => {
    const cssResp = await fetch(FONT_CSS_URL + encodeURIComponent(chars));
    if (!cssResp.ok) throw new Error(`字体服务返回 HTTP ${cssResp.status}`);
    const css = await cssResp.text();
    const match = css.match(/url\((https:[^)]+)\)/);
    if (!match) throw new Error('未能从字体服务获取字体文件地址');

    const fontResp = await fetch(match[1]);
    if (!fontResp.ok) throw new Error(`字体文件下载失败（HTTP ${fontResp.status}）`);
    const bytes = new Uint8Array(await fontResp.arrayBuffer());
    fontCache.set(chars, bytes);
    return bytes;
  })();

  pending.set(chars, task);
  try {
    return await task;
  } finally {
    pending.delete(chars);
  }
}
