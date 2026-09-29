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
 * @param {string} str - 原始文本
 * @returns {string} 转义后的文本
 */
export function escapeXml(str: any) {
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
 * @param {number} px - 像素值
 * @returns {number} EMU 值（取整）
 */
export function pxToEmu(px: any) {
    return Math.round((Number(px) || 0) * PX_TO_EMU);
}

/**
 * 磅转 EMU（用于线条宽度）
 * @param {number} pt - 磅值
 * @returns {number} EMU 值（取整）
 */
export function ptToEmu(pt: any) {
    return Math.round((Number(pt) || 0) * PT_TO_EMU);
}

/**
 * 磅转 OOXML 字号（sz 属性，百分之一磅）
 * @param {number} pt - 磅值
 * @returns {number} sz 值
 */
export function ptToSz(pt: any) {
    return Math.round((Number(pt) || 0) * 100);
}

/**
 * 角度转 OOXML 旋转值（1/60000 度）
 * @param {number} deg - 角度
 * @returns {number} OOXML 旋转值
 */
export function degToRot(deg: any) {
    return Math.round((Number(deg) || 0) * 60000);
}

/**
 * 颜色规范化为 6 位大写 HEX（不含 #），供 a:srgbClr val 使用
 * @param {string} color - 颜色值（#hex / rgb() / 颜色名）
 * @returns {string} 6 位 HEX 字符串，无法解析时返回 '000000'
 */
export function colorToHex(color: any) {
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
 * @returns {Object} 节点对象 { tagName, attrs, children }
 */
export function xmlNode(tagName: any, attrs?: any, ...children: any[]) {
    // 容错：第二参数误传节点对象时（如 xmlNode('a:solidFill', xmlNode(...))），
    // 自动将其视为子节点，避免节点属性被当成属性表序列化
    if (attrs && typeof attrs === 'object' && !Array.isArray(attrs) && typeof attrs.tagName === 'string') {
        children.unshift(attrs);
        attrs = null;
    }
    const filteredAttrs: any = {};
    if (attrs) {
        for (const key in attrs) {
            const val = attrs[key];
            if (val !== undefined && val !== null) {
                filteredAttrs[key] = String(val);
            }
        }
    }
    return {
        tagName,
        attrs: filteredAttrs,
        children: children.filter(c => c !== undefined && c !== null && c !== '')
    };
}

/**
 * 节点转 XML 字符串
 * @param {Object|string} node - 节点对象或纯文本
 * @param {string} [indent=''] - 缩进（内部递归使用）
 * @returns {string} XML 字符串
 */
export function nodeToString(node: any, indent = '') {
    if (typeof node === 'string') {
        return escapeXml(node);
    }

    const attrs = Object.keys(node.attrs)
        .map(key => ` ${key}="${escapeXml(node.attrs[key])}"`)
        .join('');

    if (!node.children || node.children.length === 0) {
        return `<${node.tagName}${attrs}/>`;
    }

    const inner = node.children
        .map((child: any) => nodeToString(child))
        .join('');

    // 纯文本子节点直接内联，避免多余空白
    if (node.children.every((c: any) => typeof c === 'string')) {
        return `<${node.tagName}${attrs}>${inner}</${node.tagName}>`;
    }

    return `<${node.tagName}${attrs}>${inner}</${node.tagName}>`;
}

/**
 * 生成带 XML 声明的完整文档
 * @param {Object} rootNode - 根节点对象
 * @returns {string} XML 文档字符串
 */
export function toXmlDocument(rootNode: any) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${nodeToString(rootNode)}`;
}
