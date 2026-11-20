/**
 * 统一错误类型模块
 *
 * 底层库（mammoth / exceljs / pdfjs-dist / ppt-parser）抛出的错误形态各异：字符串、Error、
 * 带 code 的自定义对象。这里统一成带 `format` / `code` 的 JdocError，便于调用方按格式分支处理。
 *
 * @module core/errors
 */

/** 支持的文件格式（四类二进制 OOXML/PDF + Markdown + 纯文本） */
export type DocFormat = 'pptx' | 'docx' | 'xlsx' | 'pdf' | 'md' | 'txt';

/** 统一错误类型 */
export class JdocError extends Error {
    /** 出错时所在的格式 */
    readonly format: DocFormat;
    /** 底层错误码（若底层库提供） */
    readonly code?: string;
    /** 底层原始错误，便于排查 */
    readonly cause?: unknown;

    constructor(
        message: string,
        format: DocFormat,
        options: { code?: string; cause?: unknown } = {}
    ) {
        super(message);
        this.name = 'JdocError';
        this.format = format;
        this.code = options.code;
        this.cause = options.cause;
        // 继承内置 Error 时修正原型链，保证 instanceof 在 ES5 target 下也成立
        Object.setPrototypeOf(this, JdocError.prototype);
    }
}

/** 从任意 catch 值中提取可读信息 */
export function describeError(error: unknown): string {
    if (error instanceof Error) return error.message;
    if (typeof error === 'string') return error;
    try {
        return JSON.stringify(error);
    } catch {
        return String(error);
    }
}

/**
 * 包装底层库抛出的错误为 JdocError，同时保留原始错误
 * @param format - 当前处理的格式
 * @param action - 正在执行的动作（解析/生成），用于拼装错误信息
 */
export function wrapError(format: DocFormat, action: string, error: unknown): JdocError {
    if (error instanceof JdocError) return error;
    const code = (error as { code?: string; name?: string })?.code;
    return new JdocError(`${format} ${action}失败：${describeError(error)}`, format, {
        code,
        cause: error
    });
}