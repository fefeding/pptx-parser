/**
 * 数字格式化（Excel 数字格式串）
 *
 * 抽到 core 是因为图表轴标签、单元格显示、TEXT() 函数三处都要用同一套规则，
 * 放在 xlsx/formula.ts 里会让「与格式无关」的图表模块反向依赖 xlsx。
 *
 * 支持的格式子集（够日常表格用）：
 *   - 数字段：`0` `#,##0` `0.00` `#,##0.00` 等（含千分位、小数位）
 *   - 百分号：`%`（按出现次数放大 100 倍）
 *   - 字面量前后缀：`¥` `$` `"文字"` 等非数字段字符原样保留
 *   - 日期段：`yyyy` `yy` `mm` `m` `dd` `d` `hh` `h` `ss` `s`（序列号按 UTC 取年月日）
 *
 * @module core/numfmt
 */

/** Excel 日期序列号 -> 按格式输出（序列号无时区，统一按 UTC 取年月日） */
export function formatDate(serial: number, fmt: string): string {
    const ms = Math.round((serial - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (Number.isNaN(d.getTime())) return String(serial);
    const pad = (n: number, w = 2): string => String(n).padStart(w, '0');
    const hasTime = /[hs]/i.test(fmt);
    return fmt
        .replace(/yyyy/gi, String(d.getUTCFullYear()))
        .replace(/yy/gi, pad(d.getUTCFullYear() % 100))
        .replace(/mm/g, hasTime ? pad(d.getUTCMinutes()) : pad(d.getUTCMonth() + 1))
        .replace(/m/g, hasTime ? String(d.getUTCMinutes()) : String(d.getUTCMonth() + 1))
        .replace(/dd/gi, pad(d.getUTCDate()))
        .replace(/\bd\b/g, String(d.getUTCDate()))
        .replace(/hh/gi, pad(d.getUTCHours()))
        .replace(/\bh\b/g, String(d.getUTCHours()))
        .replace(/ss/gi, pad(d.getUTCSeconds()))
        .replace(/\bs\b/g, String(d.getUTCSeconds()));
}

/** 数字格式（TEXT 函数、图表轴标签、单元格显示共用） */
export function formatByPattern(value: number | string, pattern: string): string {
    const fmt = String(pattern ?? '');
    if (!fmt) return String(value ?? '');

    // 日期格式
    if (/[ymdhs]/i.test(fmt) && !/[#0]/.test(fmt.replace(/0/g, ''))) {
        const serial = typeof value === 'number' ? value : parseFloat(String(value));
        if (!Number.isNaN(serial)) return formatDate(serial, fmt);
    }
    if (typeof value !== 'number') return String(value);

    // 数字格式：拆出「数字段」与前后字面量
    const match = /[#0?][#0?,.]*/.exec(fmt);
    const numericPart = match ? match[0] : '';
    const prefix = match ? fmt.slice(0, match.index) : '';
    const suffix = match ? fmt.slice(match.index + numericPart.length) : '';

    const percentCount = (fmt.match(/%/g) || []).length;
    const v = value * Math.pow(100, percentCount);

    const dotIdx = numericPart.indexOf('.');
    const decimals = dotIdx >= 0 ? numericPart.length - dotIdx - 1 : 0;
    const thousands = numericPart.includes(',');

    let text = Math.abs(v).toFixed(decimals);
    if (thousands) text = text.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    const negative = v < 0;
    // 数值段里没有 '#' 时补齐整数位
    if (!numericPart.includes('#') && numericPart.includes('0')) {
        const intLen = numericPart.split('.')[0].length;
        const intPart = text.split('.')[0];
        if (intPart.length < intLen) text = '0'.repeat(intLen - intPart.length) + intPart;
    }
    const negativeInPattern = numericPart.startsWith('-');
    const sign = negative && !negativeInPattern ? '-' : '';
    return `${sign}${prefix}${text}${suffix}`;
}
