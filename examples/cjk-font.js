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

/**
 * 从任意文本（JSON 字符串亦可）中提取需要子集的字符集合。
 * 不只收集汉字：标题里的 "T1 · " 等 ASCII/标点同样要进子集，否则 PDF 里这些字符会丢失。
 * 排除空白与控制字符，避免 text= 参数过长/无意义。
 */
export function collectCjkChars(text) {
  const chars = Array.from(new Set(String(text || '').match(/[\u0020-\u007E\u00A0-\u024F\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF]/g) || []));
  return chars.filter((c) => c.trim()).join('');
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
    // 优先走 dev server 的字体代理（Node fetch 拿到真正的 TTF）。
    // 浏览器直连 Google Fonts 会得到 WOFF2，pdf-lib 子集化后文字不可见。
    let bytes;
    try {
      const proxyResp = await fetch('/api/cjk-font?chars=' + encodeURIComponent(chars));
      if (proxyResp.ok) bytes = new Uint8Array(await proxyResp.arrayBuffer());
    } catch { /* 静态部署无代理，回退直连 */ }

    if (!bytes) {
      const cssResp = await fetch(FONT_CSS_URL + encodeURIComponent(chars));
      if (!cssResp.ok) throw new Error(`字体服务返回 HTTP ${cssResp.status}`);
      const css = await cssResp.text();
      const match = css.match(/url\((https:[^)]+)\)/);
      if (!match) throw new Error('未能从字体服务获取字体文件地址');
      const fontResp = await fetch(match[1]);
      if (!fontResp.ok) throw new Error(`字体文件下载失败（HTTP ${fontResp.status}）`);
      bytes = new Uint8Array(await fontResp.arrayBuffer());
    }

    // 兜底防御：若拿到的仍是 WOFF2/EOT，明确报错引导用户上传 TTF，
    // 避免生成「文字能提取却不可见」的坏 PDF。
    const magic = String.fromCharCode(bytes[0] || 0, bytes[1] || 0, bytes[2] || 0, bytes[3] || 0);
    if (magic === 'wOF2' || (bytes[0] === 0x4c && bytes[1] === 0x0a)) {
      throw new Error('字体服务返回了 WOFF2/EOT，pdf-lib 无法正确处理；请手动上传 TTF/OTF 字体');
    }
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
