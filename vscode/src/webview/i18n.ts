// Webview 国际化。
//
// 语言来源（两者同源，HTML 上的先到先用，init 消息兜底）：
//   1. 宿主把 `vscode.env.language` 写进 webview HTML 的 `<html lang="…">`，启动即可用，无闪烁；
//   2. `init` 消息里的 `locale` 字段（宿主仍会下发，便于将来做运行时切换）。
//
// 新增语言：在 `locales/` 下新增一个与 en.ts 同构的文件，再到 BUNDLES 注册即可。
// 语言包类型受 `keyof typeof en` 约束，漏译会在 `tsc --noEmit` 时报错。
import { en } from './locales/en';
import { zhCN } from './locales/zh-cn';

type Dict = Record<string, string>;

const BUNDLES: Record<string, Dict> = {
  en,
  'zh-cn': zhCN
};

export const DEFAULT_LOCALE = 'en';

let current: string = DEFAULT_LOCALE;

/** 把 VS Code 的 locale（zh-cn / zh_CN / zh / en-US …）归一到已注册的语言包。 */
function normalizeLocale(locale?: string | null): string {
  const raw = String(locale || '').trim().toLowerCase().replace(/_/g, '-');
  if (!raw) return DEFAULT_LOCALE;
  if (BUNDLES[raw]) return raw;
  const base = raw.split('-')[0];
  if (BUNDLES[base]) return base;
  // zh-Hans / zh-Hant / zh-TW 暂统一回退到简中；注册 'zh-tw' 后会自动命中
  if (base === 'zh') return 'zh-cn';
  return DEFAULT_LOCALE;
}

/** 切换语言，返回是否真的发生了变化（变化后调用方需要重建 UI 文案）。 */
export function setLocale(locale?: string | null): boolean {
  const next = normalizeLocale(locale);
  if (next === current) return false;
  current = next;
  document.documentElement.lang = next;
  return true;
}

export function getLocale(): string {
  return current;
}

function lookup(key: string): string | undefined {
  const dict = BUNDLES[current] || BUNDLES[DEFAULT_LOCALE];
  return dict[key] ?? BUNDLES[DEFAULT_LOCALE][key];
}

/** `{0}` / `{1}` 位置占位符替换。 */
function interpolate(s: string, args: any[]): string {
  if (!args.length) return s;
  return s.replace(/\{(\d+)\}/g, (m, i) => {
    const v = args[Number(i)];
    return v == null ? m : String(v);
  });
}

/** 取词条；缺 key 时返回 key 本身（便于定位漏译）。 */
export function t(key: string, ...args: any[]): string {
  const s = lookup(key);
  if (s == null) return key;
  return interpolate(s, args);
}

/** 取词条；缺 key 时退回调用方给的文案（用于解析器提供的英文数据名）。 */
export function tOr(key: string, fallback: string, ...args: any[]): string {
  const s = lookup(key);
  return interpolate(s == null ? fallback : s, args);
}

/** 解析器数据（主题 / 版式 / 切换 / 形状 / 图表类型 …）的本地化标签。 */
export function data(ns: string, value: string, fallback = ''): string {
  return tOr(`data.${ns}.${value}`, fallback || value);
}

// ---------- 快捷键显示 ----------
const IS_MAC =
  typeof navigator !== 'undefined' && /Mac|iPhone|iPad|iPod/.test(navigator.userAgent || '');

const KEY_GLYPH: Record<string, string> = {
  arrowleft: '←',
  arrowright: '→',
  arrowup: '↑',
  arrowdown: '↓',
  esc: 'Esc',
  escape: 'Esc',
  enter: '↵',
  return: '↵',
  space: 'Space',
  backspace: 'Backspace',
  delete: 'Del',
  tab: 'Tab',
  '?': '?',
  plus: '+'
};

const MOD_GLYPH: Record<string, string> = IS_MAC
  ? { mod: '⌘', ctrl: '⌃', alt: '⌥', shift: '⇧' }
  : { mod: 'Ctrl', ctrl: 'Ctrl', alt: 'Alt', shift: 'Shift' };

/**
 * 把 `mod+shift+z` 这样的写法转成平台相关的显示文本（mac 用 ⌘⇧，其它平台用 Ctrl+Shift）。
 * `mod` = ⌘ / Ctrl。功能键（F5）与单键直接原样大写。
 */
export function kbd(spec: string): string {
  const parts = String(spec)
    .split('+')
    .map((p) => p.trim().toLowerCase())
    .filter(Boolean);
  if (!parts.length) return '';
  const key = parts.pop()!;
  const mods = parts.filter((p) => p in MOD_GLYPH);
  const keyText = KEY_GLYPH[key] || (/^f\d+$/.test(key) ? key.toUpperCase() : key.length === 1 ? key.toUpperCase() : key[0].toUpperCase() + key.slice(1));
  if (!mods.length) return keyText;
  // mac 的修饰键没有 “+”，其它平台用 “+” 连接
  return IS_MAC ? mods.map((m) => MOD_GLYPH[m]).join('') + keyText : [...mods, keyText].map((m) => (m in MOD_GLYPH ? MOD_GLYPH[m] : m)).join('+');
}

/** 词条 + 快捷键，用作 tooltip / 菜单右侧提示：`复制  ⌘C`。 */
export function tip(key: string, accel?: string, ...args: any[]): string {
  const label = t(key, ...args);
  return accel ? `${label}  ${kbd(accel)}` : label;
}

/** 给任意文案附加快捷键提示（词条已在外部解析时使用）。 */
export function withAccel(label: string, accel?: string): string {
  return accel ? `${label}  ${kbd(accel)}` : label;
}