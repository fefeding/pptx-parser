/**
 * XML 构建工具模块
 *
 * 提供 JSON→PPTX 序列化所需的底层 XML 生成能力：
 * - XML 转义
 * - 单位换算（px→EMU、pt→OOXML 百分值）
 * - 颜色规范化（借助 tinycolor2，支持 #hex / rgb / 常见颜色名）
 * - XML 节点构建与拼接
 *
 * @module serializer/xml-builder
 */

import TinyColor from 'tinycolor2';

/** XML 构建阶段使用的节点形态：{ tagName, attrs, children } */
export interface BuilderNode {
    tagName: string;
    attrs?: Record<string, string>;
    children?: (string | BuilderNode)[];
}

/** px → EMU 换算因子（96px = 1inch = 914400EMU） */
const PX_TO_EMU = 914400 / 96;

/** pt → EMU 换算因子（1pt = 12700EMU） */
const PT_TO_EMU = 12700;

/** OOXML 命名空间 */
export const NS = {
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    table: 'http://schemas.openxmlformats.org/drawingml/2006/table',
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
    cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    dc: 'http://purl.org/dc/elements/1.1/',
    dcterms: 'http://purl.org/dc/terms/',
    dcmitype: 'http://purl.org/dc/dcmitype/',
    xsi: 'http://www.w3.org/2001/XMLSchema-instance',
    ext: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
};

/**
 * XML 特殊字符转义
 * @param {unknown} str - 原始文本
 * @returns {string} 转义后的文本
 */
export function escapeXml(str: unknown): string {
    if (str === undefined || str === null) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * 像素转 EMU
 * @param {unknown} px - 像素值
 * @returns {number} EMU 值（取整）
 */
export function pxToEmu(px: unknown): number {
    return Math.round((Number(px) || 0) * PX_TO_EMU);
}

/**
 * 磅转 EMU（用于线条宽度）
 * @param {unknown} pt - 磅值
 * @returns {number} EMU 值（取整）
 */
export function ptToEmu(pt: unknown): number {
    return Math.round((Number(pt) || 0) * PT_TO_EMU);
}

/**
 * 磅转 OOXML 字号（sz 属性，百分之一磅）
 * @param {unknown} pt - 磅值
 * @returns {number} sz 值
 */
export function ptToSz(pt: unknown): number {
    return Math.round((Number(pt) || 0) * 100);
}

/**
 * 角度转 OOXML 旋转值（1/60000 度）
 * @param {unknown} deg - 角度
 * @returns {number} OOXML 旋转值
 */
export function degToRot(deg: unknown): number {
    return Math.round((Number(deg) || 0) * 60000);
}

/**
 * 颜色规范化为 6 位大写 HEX（不含 #），供 a:srgbClr val 使用
 * @param {unknown} color - 颜色值（#hex / rgb() / 颜色名）
 * @returns {string} 6 位 HEX 字符串，无法解析时返回 '000000'
 */
export function colorToHex(color: unknown): string {
    if (color === undefined || color === null || color === '') return '000000';
    const c = new TinyColor(String(color));
    if (!c.isValid()) return '000000';
    return c.toHexString().replace('#', '').toUpperCase();
}

/**
 * 构建单个 XML 标签
 * @param {string} tagName - 标签名（可含命名空间前缀，如 a:off）
 * @param {Object} [attrs] - 属性表（值为 null/undefined 的属性会被忽略）
 * @param {...(string|Object)} children - 子内容：字符串或其他节点对象
 * @returns {BuilderNode} 节点对象 { tagName, attrs, children }
 */
export function xmlNode(tagName: string, attrs?: Record<string, unknown> | BuilderNode | null, ...children: (string | BuilderNode | null | undefined)[]): BuilderNode {
    // 容错：第二参数误传节点对象时（如 xmlNode('a:solidFill', xmlNode(...))），
    // 自动将其视为子节点，避免节点属性被当成属性表序列化
    if (attrs && typeof attrs === 'object' && !Array.isArray(attrs) && typeof (attrs as BuilderNode).tagName === 'string') {
        children.unshift(attrs as unknown as BuilderNode);
        attrs = null;
    }
    const filteredAttrs: Record<string, string> = {};
    if (attrs) {
        const src = attrs as Record<string, unknown>;
        for (const key in src) {
            const val = src[key];
            if (val !== undefined && val !== null) {
                filteredAttrs[key] = String(val);
            }
        }
    }
    return {
        tagName,
        attrs: filteredAttrs,
        children: children.filter(c => c !== undefined && c !== null && c !== '') as (string | BuilderNode)[]
    };
}

/**
 * 节点转 XML 字符串
 * @param {BuilderNode|string} node - 节点对象或纯文本
 * @param {string} [indent=''] - 缩进（内部递归使用）
 * @returns {string} XML 字符串
 */
export function nodeToString(node: string | BuilderNode | null | undefined, indent = ''): string {
    if (node === null || node === undefined) return '';
    if (typeof node === 'string') {
        return escapeXml(node);
    }

    const attrs = node.attrs ?? {};
    const attrStr = Object.keys(attrs)
        .map(key => ` ${key}="${escapeXml(attrs[key])}"`)
        .join('');

    if (!node.children || node.children.length === 0) {
        return `<${node.tagName}${attrStr}/>`;
    }

    const inner = node.children
        .map((child) => nodeToString(child))
        .join('');

    // 纯文本子节点直接内联，避免多余空白
    if (node.children.every((c) => typeof c === 'string')) {
        return `<${node.tagName}${attrStr}>${inner}</${node.tagName}>`;
    }

    return `<${node.tagName}${attrStr}>${inner}</${node.tagName}>`;
}

/**
 * 生成带 XML 声明的完整文档
 * @param {BuilderNode} rootNode - 根节点对象
 * @returns {string} XML 文档字符串
 */
export function toXmlDocument(rootNode: BuilderNode): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${nodeToString(rootNode)}`;
}
