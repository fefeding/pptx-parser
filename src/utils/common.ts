/**
 * 通用工具函数
 * ID生成、日志等通用功能
 */

/**
 * 生成唯一ID
 * @param prefix 前缀
 * @returns 唯一ID
 */
export function generateId(prefix: string = 'ppt-node'): string {
  return `${prefix}-${Date.now()}-${Math.floor(Math.random() * 10000)}`;
}

/**
 * 日志工具（避免生产环境输出）
 * @param level 日志级别
 * @param message 消息
 * @param data 附加数据
 */
export function log(level: 'info' | 'warn' | 'error', message: string, data?: unknown): void {
  // @ts-ignore
  const showLog = true; // (typeof process !== 'undefined' && process?.env?.NODE_ENV === 'development') || (import.meta?.env?.NODE_ENV === 'development');
  if (showLog) {
    const prefix = `[pptx-parser ${level.toUpperCase()}]`;
    if (data !== undefined) {
      console[level](prefix, message, data);
    } else {
      console[level](prefix, message);
    }
  }
}
