/**
 * PPTXjs通用工具函数 - TypeScript转译版
 * 对齐PPTXjs.js中的各种工具函数
 */

import JSZip from 'jszip';

/**
 * ArrayBuffer转Base64 - 对齐PPTXjs的base64ArrayBuffer函数
 */
export function base64ArrayBuffer(arrayBuffer: ArrayBuffer): string {
  let base64 = '';
  const encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  const bytes = new Uint8Array(arrayBuffer);
  const byteLength = bytes.byteLength;
  const byteRemainder = byteLength % 3;
  const mainLength = byteLength - byteRemainder;
  let a: number, b: number, c: number, d: number, chunk: number;

  // 主循环，每次处理3个字节
  for (let i = 0; i < mainLength; i += 3) {
    chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];

    a = (chunk & 16515072) >> 18; // 0xFC0000
    b = (chunk & 258048) >> 12;   // 0x03F000
    c = (chunk & 4032) >> 6;     // 0x000FC0
    d = chunk & 63;               // 0x00003F

    base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d];
  }

  // 处理剩余的字节
  if (byteRemainder === 1) {
    chunk = bytes[mainLength];
    a = (chunk & 252) >> 2; // 0xFC
    b = (chunk & 3) << 4;   // 0x03

    base64 += encodings[a] + encodings[b] + '==';
  } else if (byteRemainder === 2) {
    chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];
    a = (chunk & 64512) >> 10; // 0xFC00
    b = (chunk & 1008) >> 4;   // 0x03F0
    c = (chunk & 15) << 2;     // 0x000F

    base64 += encodings[a] + encodings[b] + encodings[c] + '=';
  }

  return base64;
}

/**
 * 从zip中读取图片并转为Base64
 */
export function getImageBase64(zip: JSZip, imagePath: string): string | null {
  try {
    const imageFile = zip.file(imagePath);
    if (!imageFile) {
      return null;
    }

    const arrayBuffer = imageFile.asArrayBuffer();
    return base64ArrayBuffer(arrayBuffer);
  } catch (e) {
    console.error(`Error reading image ${imagePath}:`, e);
    return null;
  }
}

/**
 * 获取图片MIME类型
 */
export function getImageMimeType(fileName: string): string {
  const ext = fileName.split('.').pop()?.toLowerCase() || '';
  
  const mimeTypes: Record<string, string> = {
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'gif': 'image/gif',
    'bmp': 'image/bmp',
    'tiff': 'image/tiff',
    'webp': 'image/webp',
    'svg': 'image/svg+xml',
  };

  return mimeTypes[ext] || 'image/png';
}

/**
 * 生成Data URL
 */
export function generateDataUrl(base64Data: string, mimeType: string): string {
  return `data:${mimeType};base64,${base64Data}`;
}

/**
 * 处理数值（处理可能的null/undefined）
 */
export function safeParseInt(value: any, defaultValue = 0): number {
  if (value === undefined || value === null || value === '') {
    return defaultValue;
  }
  const parsed = parseInt(String(value), 10);
  return isNaN(parsed) ? defaultValue : parsed;
}

/**
 * 处理浮点数
 */
export function safeParseFloat(value: any, defaultValue = 0): number {
  if (value === undefined || value === null || value === '') {
    return defaultValue;
  }
  const parsed = parseFloat(String(value));
  return isNaN(parsed) ? defaultValue : parsed;
}

/**
 * 深度克隆对象
 */
export function deepClone<T>(obj: T): T {
  if (obj === null || typeof obj !== 'object') {
    return obj;
  }

  if (Array.isArray(obj)) {
    return obj.map(item => deepClone(item)) as unknown as T;
  }

  const cloned: any = {};
  for (const key in obj) {
    if (obj.hasOwnProperty(key)) {
      cloned[key] = deepClone(obj[key]);
    }
  }

  return cloned;
}

/**
 * 检查是否为RTL语言
 */
export function isRtlLanguage(lang: string): boolean {
  const rtlLangs = ['he-IL', 'ar-AE', 'ar-SA', 'dv-MV', 'fa-IR', 'ur-PK'];
  return rtlLangs.includes(lang);
}

/**
 * 规范化十六进制颜色值
 */
export function normalizeHexColor(color: string): string {
  if (!color) return '';

  // 移除可能的#前缀
  let hex = color.replace('#', '');

  // 如果是3位或4位十六进制，转换为6位或8位
  if (hex.length === 3) {
    hex = hex.split('').map(c => c + c).join('');
  } else if (hex.length === 4) {
    hex = hex.split('').map(c => c + c).join('');
  }

  // 确保是大写
  hex = hex.toUpperCase();

  return hex;
}

/**
 * 检查是否为有效的十六进制颜色
 */
export function isValidHexColor(color: string): boolean {
  if (!color) return false;
  const hex = color.replace('#', '');
  return /^[0-9A-Fa-f]{3}$|^[0-9A-Fa-f]{6}$|^[0-9A-Fa-f]{8}$/.test(hex);
}

/**
 * 生成唯一ID
 */
export function generateUniqueId(prefix = 'id'): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * 延迟执行
 */
export function delay(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * 重试函数
 */
export async function retry<T>(
  fn: () => Promise<T>,
  maxRetries = 3,
  delayMs = 1000
): Promise<T> {
  let lastError: Error | undefined;

  for (let i = 0; i < maxRetries; i++) {
    try {
      return await fn();
    } catch (error) {
      lastError = error as Error;
      if (i < maxRetries - 1) {
        await delay(delayMs * (i + 1));
      }
    }
  }

  throw lastError || new Error('Retry failed');
}

/**
 * 数组去重
 */
export function unique<T>(array: T[]): T[] {
  return Array.from(new Set(array));
}

/**
 * 合并对象（深度合并）
 */
export function deepMerge<T extends object>(target: T, ...sources: Partial<T>[]): T {
  if (!sources.length) return target;
  const source = sources.shift();

  if (isObject(target) && isObject(source)) {
    for (const key in source) {
      if (isObject(source[key])) {
        if (!target[key]) Object.assign(target, { [key]: {} });
        deepMerge(target[key], source[key]);
      } else {
        Object.assign(target, { [key]: source[key] });
      }
    }
  }

  return deepMerge(target, ...sources);
}

/**
 * 检查是否为对象
 */
function isObject(item: any): boolean {
  return item && typeof item === 'object' && !Array.isArray(item);
}

/**
 * 格式化文件大小
 */
export function formatFileSize(bytes: number): string {
  if (bytes === 0) return '0 Bytes';

  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));

  return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
}

/**
 * 截断字符串
 */
export function truncateString(str: string, maxLength: number, suffix = '...'): string {
  if (str.length <= maxLength) return str;
  return str.substring(0, maxLength - suffix.length) + suffix;
}

/**
 * 缓存装饰器
 */
export function memoize<T extends (...args: any[]) => any>(fn: T): T {
  const cache = new Map();

  return ((...args: any[]) => {
    const key = JSON.stringify(args);
    if (cache.has(key)) {
      return cache.get(key);
    }
    const result = fn(...args);
    cache.set(key, result);
    return result;
  }) as T;
}
