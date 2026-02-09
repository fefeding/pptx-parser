/**
 * tXml - Tiny XML Parser
 * 轻量级 XML 解析器，支持简化、过滤和流式解析
 * @module tXml
 */

let order = 1;

/**
 * 解析 XML 字符串
 * @param {string} xml - XML 字符串
 * @param {Object} options - 解析选项
 * @param {number} [options.pos=0] - 起始位置
 * @param {boolean} [options.parseNode=false] - 只解析单个节点
 * @param {string} [options.attrName] - 属性名过滤
 * @param {string} [options.attrValue] - 属性值过滤
 * @param {boolean} [options.simplify=false] - 简化输出
 * @param {Function} [options.filter] - 过滤函数
 * @returns {Array|Object} 解析结果
 */
function tXml(xml, options) {
    'use strict';

    options = options || {};

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
    function parseChildren() {
        const children = [];

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
     * 解析文本节点
     * @returns {string} 文本内容
     */
    function parseText() {
        const start = pos;
        pos = xml.indexOf(CHAR_LT, pos) - 1;
        if (pos === -2) {
            pos = xml.length;
        }
        return xml.slice(start, pos + 1);
    }

    /**
     * 解析标签名
     * @returns {string} 标签名
     */
    function parseTagName() {
        const start = pos;
        while (STOP_CHARS.indexOf(xml[pos]) === -1 && xml[pos]) {
            pos++;
        }
        return xml.slice(start, pos);
    }

    /**
     * 解析属性值
     * @returns {string|null} 属性值
     */
    function parseAttributeValue() {
        const quoteChar = xml[pos];
        const start = ++pos;
        pos = xml.indexOf(quoteChar, start);
        return xml.slice(start, pos);
    }

    /**
     * 查找属性位置
     * @returns {number} 属性位置索引
     */
    function findAttributePosition() {
        const pattern = new RegExp('\\s' + options.attrName + '\\s*=[\'"]' + options.attrValue + '[\'"]');
        const match = pattern.exec(xml);
        return match ? match.index : -1;
    }

    /**
     * 解析单个 XML 节点
     * @returns {Object} 节点对象
     */
    function parseNode() {
        const node = {};
        pos++;
        node.tagName = parseTagName();

        let hasAttributes = false;

        while (xml.charCodeAt(pos) !== CODE_GT && xml[pos]) {
            const charCode = xml.charCodeAt(pos);

            // 检查是否是属性名开始
            if ((charCode > 64 && charCode < 91) || (charCode > 96 && charCode < 123)) {
                const attrName = parseTagName();
                let attrValue = null;

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
                node.attributes[attrName] = attrValue;
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

    let result;

    if (options.attrValue !== undefined) {
        options.attrName = options.attrName || 'id';
        result = [];
        let attrPos;
        while ((attrPos = findAttributePosition()) !== -1) {
            pos = xml.lastIndexOf(CHAR_LT, attrPos);
            if (pos !== -1) {
                result.push(parseNode());
            }
            xml = xml.substr(pos);
            pos = 0;
        }
    } else {
        result = options.parseNode ? parseNode() : parseChildren();
    }

    if (options.filter) {
        result = tXml.filter(result, options.filter);
    }
    if (options.simplify) {
        result = tXml.simplify(result);
    }

    result.pos = pos;
    return result;
}

/**
 * 简化解析结果
 * @param {Array} nodes - 节点数组
 * @returns {Object|string} 简化后的对象
 */
tXml.simplify = function(nodes) {
    const result = {};

    if (nodes === undefined) {
        return {};
    }

    if (nodes.length === 1 && typeof nodes[0] === 'string') {
        return nodes[0];
    }

    nodes.forEach(function(node) {
        if (typeof node !== 'object') {
            return;
        }

        if (!result[node.tagName]) {
            result[node.tagName] = [];
        }

        const simplified = tXml.simplify(node.children || []);
        result[node.tagName].push(simplified);

        // 只在对象是对象类型时设置属性
        if (typeof simplified === 'object' && simplified !== null) {
            if (node.attributes) {
                simplified.attrs = node.attributes;
            }
            if (simplified.attrs === undefined) {
                simplified.attrs = { order: order };
            } else {
                simplified.attrs.order = order;
            }
            order++;
        }
    });

    // 如果数组只有一个元素，直接返回该元素
    for (const key in result) {
        if (result[key].length === 1) {
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
tXml.filter = function(nodes, filterFn) {
    const result = [];

    nodes.forEach(function(node) {
        if (typeof node === 'object' && filterFn(node)) {
            result.push(node);
        }
        if (node.children) {
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
tXml.stringify = function(nodes) {
    let xmlString = '';

    function processNodes(nodes) {
        if (!nodes) return;
        for (let i = 0; i < nodes.length; i++) {
            if (typeof nodes[i] === 'string') {
                xmlString += nodes[i].trim();
            } else {
                processNode(nodes[i]);
            }
        }
    }

    function processNode(node) {
        xmlString += '<' + node.tagName;
        for (const attr in node.attributes) {
            const value = node.attributes[attr];
            if (value === null) {
                xmlString += ' ' + attr;
            } else if (value.indexOf('"') === -1) {
                xmlString += ' ' + attr + '="' + value.trim() + '"';
            } else {
                xmlString += ' ' + attr + "='" + value.trim() + "'";
            }
        }
        xmlString += '>';
        processNodes(node.children);
        xmlString += '</' + node.tagName + '>';
    }

    processNodes(nodes);
    return xmlString;
};

/**
 * 获取节点的文本内容
 * @param {Array|Object|string} node - 节点
 * @returns {string} 文本内容
 */
tXml.toContentString = function(node) {
    if (Array.isArray(node)) {
        let text = '';
        node.forEach(function(child) {
            text += ' ' + tXml.toContentString(child);
            text = text.trim();
        });
        return text;
    }

    if (typeof node === 'object') {
        return tXml.toContentString(node.children);
    }

    return ' ' + node;
};

/**
 * 通过 ID 获取元素
 * @param {string} xml - XML 字符串
 * @param {string} id - 元素 ID
 * @param {boolean} simplify - 是否简化结果
 * @returns {Object} 元素对象
 */
tXml.getElementById = function(xml, id, simplify) {
    const result = tXml(xml, {
        attrValue: id,
        simplify: simplify
    });
    return simplify ? result : result[0];
};

/**
 * 通过 class 名获取元素
 * @param {string} xml - XML 字符串
 * @param {string} className - 类名
 * @param {boolean} simplify - 是否简化结果
 * @returns {Array} 元素数组
 */
tXml.getElementsByClassName = function(xml, className, simplify) {
    return tXml(xml, {
        attrName: 'class',
        attrValue: '[a-zA-Z0-9-s ]*' + className + '[a-zA-Z0-9-s ]*',
        simplify: simplify
    });
};

/**
 * 流式解析 XML
 * @param {string|Stream} source - XML 源
 * @param {number|Function} chunkSize - 块大小或回调函数
 * @returns {EventEmitter} 事件发射器
 */
tXml.parseStream = function(source, chunkSize) {
    let callback;

    if (typeof chunkSize === 'function') {
        callback = chunkSize;
        chunkSize = 0;
    }

    if (typeof chunkSize === 'string') {
        chunkSize = chunkSize.length + 2;
    }

    // Node.js 流处理
    if (typeof source === 'string') {
        const fs = require('fs');
        source = fs.createReadStream(source, { start: chunkSize });
        chunkSize = 0;
    }

    let pos = chunkSize;
    let buffer = '';
    let chunkIndex = 0;

    source.on('data', function(chunk) {
        chunkIndex++;
        buffer += chunk;

        let lastPos = 0;

        while (true) {
            pos = buffer.indexOf('<', pos) + 1;
            const node = tXml(buffer, { pos: pos, parseNode: true });
            pos = node.pos;

            if (pos > buffer.length - 1 || lastPos > pos) {
                if (lastPos) {
                    buffer = buffer.slice(lastPos);
                    pos = 0;
                    lastPos = 0;
                }
                return;
            }

            source.emit('xml', node);
            lastPos = pos;
        }
    });

    source.on('end', function() {
        console.log('end');
    });

    return source;
};

export default tXml;
