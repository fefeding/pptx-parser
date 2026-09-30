/**
 * tXml - Tiny XML Parser
 * 轻量级 XML 解析器，支持简化、过滤和流式解析
 * @module tXml
 */

import { RawXmlNode, RawXmlChildren, RawXmlParseOptions, XmlNode } from './types';

let order = 1;

/** XML 预定义实体 */
const XML_ENTITIES: Record<string, string> = {
    amp: '&',
    lt: '<',
    gt: '>',
    quot: '"',
    apos: "'"
};

/**
 * 解码 XML 实体：&amp; &lt; &gt; &quot; &apos; 及数字字符引用 &#nn; / &#xhh;
 * 未识别的实体保持原样，避免破坏非标准内容。
 * @param {string} str - 可能含实体的字符串
 * @returns {string} 解码后的字符串
 */
function decodeEntities(str: string): string {
    if (!str || str.indexOf('&') === -1) return str;
    return str.replace(/&(#x[0-9a-f]+|#[0-9]+|[a-z]+);/gi, (match, code: string) => {
        if (code.charAt(0) === '#') {
            const isHex = code.charAt(1) === 'x' || code.charAt(1) === 'X';
            const num = parseInt(isHex ? code.slice(2) : code.slice(1), isHex ? 16 : 10);
            if (isNaN(num) || num < 0 || num > 0x10ffff) return match;
            try {
                return String.fromCodePoint(num);
            } catch {
                return match;
            }
        }
        const named = XML_ENTITIES[code.toLowerCase()];
        return named === undefined ? match : named;
    });
}

/** tXml 的可调用形态：既是函数，又挂载 simplify/filter 等方法 */
interface TXmlFn {
    (xml: string, options?: RawXmlParseOptions): RawXmlNode[] | RawXmlNode;
    simplify: (nodes: RawXmlChildren) => XmlNode | string;
    filter: (nodes: RawXmlChildren, filterFn: (node: RawXmlNode) => boolean) => RawXmlNode[];
    stringify: (nodes: RawXmlChildren) => string;
    toContentString: (node: RawXmlNode | RawXmlChildren | string) => string;
    getElementById: (xml: string, id: string, simplify: boolean) => RawXmlNode | RawXmlNode[];
    getElementsByClassName: (xml: string, className: string, simplify: boolean) => RawXmlNode[];
    parseStream: (source: string | NodeJS.ReadableStream, chunkSize?: number | ((chunk: Buffer) => void)) => NodeJS.ReadableStream;
}

const tXml = (function (xml: string, options: RawXmlParseOptions = {}): RawXmlNode[] | RawXmlNode | RawXmlChildren {
    const POS = options.pos || 0;

    // 字符常量
    const CHAR_LT = '<';
    const CHAR_GT = '>';
    const CHAR_SLASH = '/';
    const CHAR_DASH = '-';
    const CHAR_EXCLAMATION = '!';
    const CHAR_SINGLE_QUOTE = "'";
    const CHAR_DOUBLE_QUOTE = '"';
    const STOP_CHARS = "\n\t>/= ";
    const VOID_ELEMENTS = ['img', 'br', 'input', 'meta', 'link'];

    // 字符码
    const CODE_LT = CHAR_LT.charCodeAt(0);
    const CODE_GT = CHAR_GT.charCodeAt(0);
    const CODE_DASH = CHAR_DASH.charCodeAt(0);
    const CODE_SLASH = CHAR_SLASH.charCodeAt(0);
    const CODE_EXCLAMATION = CHAR_EXCLAMATION.charCodeAt(0);
    const CODE_SINGLE_QUOTE = CHAR_SINGLE_QUOTE.charCodeAt(0);
    const CODE_DOUBLE_QUOTE = CHAR_DOUBLE_QUOTE.charCodeAt(0);

    let pos = POS;

    /**
     * 解析所有子节点
     * @returns {Array} 子节点数组
     */
    function parseChildren(): RawXmlChildren {
        const children: RawXmlChildren = [];

        while (xml[pos]) {
            const charCode = xml.charCodeAt(pos);

            if (charCode === CODE_LT) {
                const nextCharCode = xml.charCodeAt(pos + 1);

                // 结束标签 </tag>
                if (nextCharCode === CODE_SLASH) {
                    pos = xml.indexOf(CHAR_GT, pos);
                    if (pos + 1) pos += 1;
                    return children;
                }

                // 注释或 DOCTYPE
                if (nextCharCode === CODE_EXCLAMATION) {
                    // CDATA 或 注释
                    if (xml.charCodeAt(pos + 2) === CODE_DASH) {
                        // 跳过注释 <!-- -->
                        while (pos !== -1 &&
                            !(xml.charCodeAt(pos) === CODE_GT &&
                                xml.charCodeAt(pos - 1) === CODE_DASH &&
                                xml.charCodeAt(pos - 2) === CODE_DASH)) {
                            pos = xml.indexOf(CHAR_GT, pos + 1);
                        }
                        if (pos === -1) pos = xml.length;
                    } else {
                        // DOCTYPE
                        pos += 2;
                        while (xml.charCodeAt(pos) !== CODE_GT && xml[pos]) {
                            pos++;
                        }
                    }
                    pos++;
                    continue;
                }

                const node = parseNode();
                children.push(node);
            } else {
                const text = parseText();
                if (text.trim().length > 0) {
                    children.push(text);
                }
                pos++;
            }
        }

        return children;
    }

    /**
     * 解析文本节点（并解码 XML 实体）
     * @returns {string} 文本内容
     */
    function parseText(): string {
        const start = pos;
        pos = xml.indexOf(CHAR_LT, pos) - 1;
        if (pos === -2) {
            pos = xml.length;
        }
        return decodeEntities(xml.slice(start, pos + 1));
    }

    /**
     * 解析标签名
     * @returns {string} 标签名
     */
    function parseTagName(): string {
        const start = pos;
        while (STOP_CHARS.indexOf(xml[pos]) === -1 && xml[pos]) {
            pos++;
        }
        return xml.slice(start, pos);
    }

    /**
     * 解析属性值（并解码 XML 实体）
     * @returns {string|null} 属性值
     */
    function parseAttributeValue(): string {
        const quoteChar = xml[pos];
        const start = ++pos;
        pos = xml.indexOf(quoteChar, start);
        return decodeEntities(xml.slice(start, pos));
    }

    /**
     * 查找属性位置
     * @returns {number} 属性位置索引
     */
    function findAttributePosition(): number {
        const pattern = new RegExp(`\\s${options.attrName}\\s*=[\'"]${options.attrValue}[\'"]`);
        const match = pattern.exec(xml);
        return match ? match.index : -1;
    }

    /**
     * 解析单个 XML 节点
     * @returns {Object} 节点对象
     */
    function parseNode(): RawXmlNode {
        const node: RawXmlNode = { tagName: '' };
        pos++;
        node.tagName = parseTagName();

        let hasAttributes = false;

        while (xml.charCodeAt(pos) !== CODE_GT && xml[pos]) {
            const charCode = xml.charCodeAt(pos);

            // 检查是否是属性名开始
            if ((charCode > 64 && charCode < 91) || (charCode > 96 && charCode < 123)) {
                const attrName = parseTagName();
                let attrValue: string | null = null;

                // 跳过空白和等号
                let currentCharCode = xml.charCodeAt(pos);
                while (currentCharCode &&
                    currentCharCode !== CODE_SINGLE_QUOTE &&
                    currentCharCode !== CODE_DOUBLE_QUOTE &&
                    !((currentCharCode > 64 && currentCharCode < 91) ||
                        (currentCharCode > 96 && currentCharCode < 123)) &&
                    currentCharCode !== CODE_GT) {
                    pos++;
                    currentCharCode = xml.charCodeAt(pos);
                }

                // 解析属性值
                if (currentCharCode === CODE_SINGLE_QUOTE ||
                    currentCharCode === CODE_DOUBLE_QUOTE) {
                    attrValue = parseAttributeValue();
                    if (pos === -1) return node;
                } else {
                    attrValue = null;
                    pos--;
                }

                if (!hasAttributes) {
                    node.attributes = {};
                    hasAttributes = true;
                }
                node.attributes![attrName] = attrValue;
            }
            pos++;
        }

        // 处理自闭合标签或解析子元素
        if (xml.charCodeAt(pos - 1) !== CODE_SLASH) {
            if (node.tagName === 'script') {
                const contentStart = pos + 1;
                pos = xml.indexOf('</script>', pos);
                node.children = [xml.slice(contentStart, pos - 1)];
                pos += 8;
            } else if (node.tagName === 'style') {
                const contentStart = pos + 1;
                pos = xml.indexOf('</style>', pos);
                node.children = [xml.slice(contentStart, pos - 1)];
                pos += 7;
            } else if (VOID_ELEMENTS.indexOf(node.tagName) === -1) {
                pos++;
                node.children = parseChildren();
            } else {
                pos++;
            }
        } else {
            pos++;
        }

        return node;
    }

    let result: RawXmlNode[] | RawXmlNode | RawXmlChildren;

    if (options.attrValue !== undefined) {
        options.attrName = options.attrName || 'id';
        result = [];
        let attrPos;
        while ((attrPos = findAttributePosition()) !== -1) {
            pos = xml.lastIndexOf(CHAR_LT, attrPos);
            if (pos !== -1) {
                (result as RawXmlNode[]).push(parseNode());
            }
            xml = xml.substr(pos);
            pos = 0;
        }
    } else {
        result = options.parseNode ? parseNode() : parseChildren();
    }

    if (options.filter) {
        result = tXml.filter(result as RawXmlChildren, options.filter);
    }
    if (options.simplify) {
        result = tXml.simplify(result as RawXmlChildren) as RawXmlNode[] | RawXmlNode;
    }

    (result as RawXmlNode).pos = pos;
    return result;
}) as unknown as TXmlFn;

/**
 * 简化解析结果
 * @param {Array} nodes - 节点数组
 * @returns {Object|string} 简化后的对象
 */
tXml.simplify = (nodes: RawXmlChildren): XmlNode | string => {
    const result: XmlNode = {};
    if (nodes === undefined) {
        return {};
    }
    if (nodes.length === 1 && typeof nodes[0] === 'string') {
        return nodes[0];
    }
    nodes.forEach((node: RawXmlNode | string) => {
        if (typeof node !== 'object') {
            return;
        }
        if (!result[node.tagName]) {
            result[node.tagName] = [];
        }
        const simplified: XmlNode | string = tXml.simplify(node.children || []);
        (result[node.tagName] as (XmlNode | string)[]).push(simplified);
        // 只在对象是对象类型时设置属性
        if (typeof simplified === 'object' && simplified !== null) {
            if (node.attributes) {
                simplified.attrs = node.attributes;
            }
            if (simplified.attrs === undefined) {
                simplified.attrs = { order };
            }
            else {
                simplified.attrs.order = order;
            }
            order++;
        }
    });
    // 如果数组只有一个元素，直接返回该元素
    for (const key in result) {
        const arr = result[key] as unknown as ArrayLike<unknown>;
        if (arr.length === 1) {
            result[key] = result[key][0];
        }
    }
    return result;
};

/**
 * 过滤节点
 * @param {Array} nodes - 节点数组
 * @param {Function} filterFn - 过滤函数
 * @returns {Array} 过滤后的节点
 */
tXml.filter = (nodes: RawXmlChildren, filterFn: (node: RawXmlNode) => boolean): RawXmlNode[] => {
    const result: RawXmlNode[] = [];
    nodes.forEach((node: RawXmlNode | string) => {
        if (typeof node === 'object' && filterFn(node)) {
            result.push(node);
        }
        if (typeof node === 'object' && node.children) {
            const filtered = tXml.filter(node.children, filterFn);
            result.push(...filtered);
        }
    });
    return result;
};

/**
 * 将节点数组转换为 XML 字符串
 * @param {Array} nodes - 节点数组
 * @returns {string} XML 字符串
 */
tXml.stringify = (nodes: RawXmlChildren): string => {
    let xmlString = '';
    function processNodes(nodes: RawXmlChildren) {
        if (!nodes)
            return;
        for (const item of nodes) {
            if (typeof item === 'string') {
                xmlString += item.trim();
            }
            else {
                processNode(item);
            }
        }
    }
    function processNode(node: RawXmlNode) {
        xmlString += `<${node.tagName}`;
        if (node.attributes) {
            for (const attr in node.attributes) {
                const value = node.attributes[attr];
                if (value === null) {
                    xmlString += ` ${attr}`;
                }
                else if (value.indexOf('"') === -1) {
                    xmlString += ` ${attr}="${value.trim()}"`;
                }
                else {
                    xmlString += ` ${attr}='${value.trim()}'`;
                }
            }
        }
        xmlString += '>';
        processNodes(node.children ?? []);
        xmlString += `</${node.tagName}>`;
    }
    processNodes(nodes);
    return xmlString;
};

/**
 * 获取节点的文本内容
 * @param {Array|Object|string} node - 节点
 * @returns {string} 文本内容
 */
tXml.toContentString = (node: RawXmlNode | RawXmlChildren | string): string => {
    if (Array.isArray(node)) {
        let text = '';
        node.forEach((child) => {
            text += ` ${tXml.toContentString(child)}`;
            text = text.trim();
        });
        return text;
    }
    if (typeof node === 'object') {
        return tXml.toContentString(node.children ?? []);
    }
    return ` ${node}`;
};

/**
 * 通过 ID 获取元素
 * @param {string} xml - XML 字符串
 * @param {string} id - 元素 ID
 * @param {boolean} simplify - 是否简化结果
 * @returns {Object} 元素对象
 */
tXml.getElementById = (xml: string, id: string, simplify: boolean): RawXmlNode | RawXmlNode[] => {
    const result = tXml(xml, {
        attrValue: id,
        simplify: simplify
    });
    return simplify ? result : (result as RawXmlNode[])[0];
};

/**
 * 通过 class 名获取元素
 * @param {string} xml - XML 字符串
 * @param {string} className - 类名
 * @param {boolean} simplify - 是否简化结果
 * @returns {Array} 元素数组
 */
tXml.getElementsByClassName = (xml: string, className: string, simplify: boolean): RawXmlNode[] => {
    return tXml(xml, {
        attrName: 'class',
        attrValue: `[a-zA-Z0-9-s ]*${className}[a-zA-Z0-9-s ]*`,
        simplify: simplify
    }) as RawXmlNode[];
};

/**
 * 流式解析 XML
 * @param {string|Stream} source - XML 源
 * @param {number|Function} chunkSize - 块大小或回调函数
 * @returns {EventEmitter} 事件发射器
 */
tXml.parseStream = (source: string | NodeJS.ReadableStream, chunkSize: number | ((chunk: Buffer) => void) = 0): NodeJS.ReadableStream => {
    let callback: ((chunk: Buffer) => void) | undefined;
    if (typeof chunkSize === 'function') {
        callback = chunkSize;
        chunkSize = 0;
    }
    let stream: NodeJS.ReadableStream;
    if (typeof source === 'string') {
        const fs = require('fs');
        stream = fs.createReadStream(source, { start: chunkSize });
    } else {
        stream = source;
    }
    let pos = chunkSize;
    let buffer = '';
    let chunkIndex = 0;
    stream.on('data', (chunk: Buffer) => {
        chunkIndex++;
        buffer += chunk;
        let lastPos = 0;
        while (true) {
            pos = buffer.indexOf('<', pos) + 1;
            const node = tXml(buffer, { pos: pos, parseNode: true });
            pos = (node as RawXmlNode).pos ?? 0;
            if (pos > buffer.length - 1 || lastPos > pos) {
                if (lastPos) {
                    buffer = buffer.slice(lastPos);
                    pos = 0;
                    lastPos = 0;
                }
                return;
            }
            stream.emit('xml', node);
            lastPos = pos;
        }
    });
    return stream;
};

export default tXml;
