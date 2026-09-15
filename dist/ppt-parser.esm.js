/**
 * @fefeding/ppt-parser v1.0.12
 * PPTX文件解析与序列化核心库，纯TS编写，支持解析PPTX为JSON结构、JSON序列化为标准PPTX文件，无框架依赖
 * MIT License
 */
import JSZip from 'jszip';
import TinyColor from 'tinycolor2';

/**
 * 常量模块
 * 提供PPTX解析过程中使用的各种常量
 */

// =============================================================================
// 尺寸转换因子
// =============================================================================

/**
 * 尺寸转换因子：将PPTX的EMU单位转换为像素
 * 914400 EMU = 1 inch = 96 pixels
 */
const SLIDE_FACTOR$1 = 96 / 914400;

/**
 * 字体大小转换因子：将PT单位转换为PX单位
 */
const FONT_SIZE_FACTOR = 4 / 3.2;

// =============================================================================
// 语言相关常量
// =============================================================================

/**
 * 右到左（RTL）语言代码数组
 * @type {string[]}
 * 
 * 支持的RTL语言列表：
 * - 阿拉伯语系：ar-* (阿拉伯联合酋长国、沙特阿拉伯、埃及、伊拉克、约旦、科威特、黎巴嫩、利比亚、摩洛哥、阿曼、巴勒斯坦、卡塔尔、苏丹、叙利亚、突尼斯、也门等)
 * - 希伯来语：he-IL (以色列)
 * - 波斯语：fa-IR (伊朗)
 * - 乌尔都语：ur-PK (巴基斯坦)
 * - 迪维希语：dv-MV (马尔代夫)
 * - 撒哈拉阿拉伯语：szq-DZ (阿尔及利亚)
 * - 普什图语：ps-AF (阿富汗)
 * - 维吾尔语：ug-CN (中国)
 * - 哈萨克语：kk-KZ (哈萨克斯坦)
 * - 吉尔吉斯语：ky-KG (吉尔吉斯斯坦)
 * - 乌兹别克语：uz-UZ (乌兹别克斯坦)
 */
const RTL_LANGS_ARRAY = [
    // 阿拉伯语变体
    "he-IL", "ar-AE", "ar-SA", "ar-EG", "ar-IQ", "ar-JO", "ar-KW", "ar-LB", "ar-LY", 
    "ar-MA", "ar-OM", "ar-PS", "ar-QA", "ar-SD", "ar-SY", "ar-TN", "ar-YE",
    // 波斯语
    "fa-IR", "fa-AF",
    // 乌尔都语
    "ur-PK", "ur-IN",
    // 迪维希语
    "dv-MV",
    // 撒哈拉阿拉伯语
    "szq-DZ",
    // 普什图语
    "ps-AF", "ps-PK",
    // 维吾尔语
    "ug-CN",
    // 中亚语言
    "kk-KZ", "ky-KG", "uz-UZ", "tg-TJ",
    // 其他RTL语言
    "yi-DE", "jpr-IL", "jrb-IL"
];

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

    if (typeof chunkSize === 'function') {
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

    source.on('data', function(chunk) {
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

    return source;
};

/**
 * XML 工具函数模块
 * 
 * 提供 XML 节点遍历和查询功能，是整个项目的核心工具模块。
 * 
 * 主要功能:
 * - getTextByPathList: 通过路径数组访问嵌套的 XML 节点
 * - getTextByPathStr: 通过路径字符串访问嵌套的 XML 节点
 * - readXmlFile: 从 ZIP 文件中读取 XML 文件
 * 
 * @module utils/xml
 */


const PPTXXmlUtils = (function() {

    /**
     * getTextByPathStr - 通过路径字符串获取XML文本
     * @param {Object} node - XML节点
     * @param {string} pathStr - 路径字符串（空格分隔）
     * @returns {*} 获取的值
     */
    function getTextByPathStr(node, pathStr) {
        return getTextByPathList(node, pathStr.trim().split(/\s+/));
    }

    /**
     * getTextByPathList - 通过路径数组获取XML文本
     * @param {Object} node - XML节点
     * @param {string[]} path - 路径数组
     * @returns {*} 获取的值
     */
    function getTextByPathList(node, path) {
        if (path.constructor !== Array) {
            throw Error("Error of path type! path is not array.");
        }

        if (node === undefined) {
            return undefined;
        }

        let l = path.length;
        for (let i = 0; i < l; i++) {
            node = node[path[i]];
            if (node === undefined) {
                return undefined;
            }
        }

        return node;
    }

    /**
     * setTextByPathList - 通过路径数组设置XML文本
     * @param {Object} node - XML节点
     * @param {string[]} path - 路径数组
     * @param {*} value - 要设置的值
     */
    function setTextByPathList(node, path, value) {
        if (path.constructor !== Array) {
            throw Error("Error of path type! path is not array.");
        }

        if (node === undefined) {
            return undefined;
        }

        let obj = node;
        const len = path.length;
        for (let i = 0; i < len; i++) {
            const p = path[i];
            if (obj[p] === undefined) {
                if (i === len - 1) {
                    obj[p] = value;
                } else {
                    obj[p] = {};
                }
            }
            obj = obj[p];
        }
        return obj;
    }

    /**
     * eachElement - 遍历节点数组或单个节点
     * @param {Object|Array} node - XML节点或节点数组
     * @param {Function} doFunction - 对每个节点执行的函数
     * @returns {string} 所有函数返回值的拼接
     */
    function eachElement(node, doFunction) {
        if (node === undefined) {
            return;
        }
        let result = "";
        if (node.constructor === Array) {
            let l = node.length;
            for (let i = 0; i < l; i++) {
                result += doFunction(node[i], i);
            }
        } else {
            result += doFunction(node, 0);
        }
        return result;
    }

    /**
     * angleToDegrees - 将角度转换为度数
     * @param {number} angle - 角度值（EMU单位）
     * @returns {number} 转换后的度数
     */
    function angleToDegrees(angle) {
        if (angle == "" || angle == null) {
            return 0;
        }
        return Math.round(angle / 60000);
    }

    /**
     * degreesToRadians - 将度数转换为弧度
     * @param {number} degrees - 度数
     * @returns {number} 弧度
     */
    function degreesToRadians(degrees) {
        if (degrees == "" || degrees == null || degrees == undefined) {
            return 0;
        }
        return degrees * (Math.PI / 180);
    }

    /**
     * escapeHtml - 转义HTML特殊字符
     * @param {string} text - 原始文本
     * @returns {string} 转义后的文本
     */
    function escapeHtml(text) {
        let map = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        };
        return text.replace(/[&<>"']/g, (m) => map[m]);
    }

    /**
     * readXmlFile - 读取XML文件并解析为对象
     * @param {Object} zip - JSZip实例
     * @param {string} filename - 文件名
     * @param {boolean} isSlideContent - 是否为幻灯片内容
     * @param {number} appVersion - 应用版本
     * @returns {Promise<Object>} 解析后的XML对象
     */
    async function readXmlFile(zip, filename, isSlideContent, appVersion) {
        try {
            const zipFile = zip.file(filename);
            if (!zipFile) return null;
            let fileContent = zipFile.async? await zipFile.async("text"): zipFile.asText();
            if (isSlideContent && appVersion <= 12) {
                //< office2007
                //remove "<!CDATA[ ... ]]>" tag
                fileContent = fileContent.replace(/<!\[CDATA\[(.*?)\]\]>/g, '$1');
            }
            let xmlData = tXml(fileContent, { simplify: 1 });
            if (xmlData["?xml"] !== undefined) {
                return xmlData["?xml"];
            } else {
                return xmlData;
            }
        } catch (e) {
            //console.log("error readXmlFile: the file '" + filename + "' not exit")
            return null;
        }
    }

    /**
     * 获取内容类型
     * @param {Object} zip - JSZip实例
     * @param {number} appVersion - Office版本
     * @returns {Promise<Object>} 包含slides和slideLayouts的对象
     */
    async function getContentTypes(zip, appVersion) {
        let ContentTypesJson = await PPTXXmlUtils.readXmlFile(zip, "[Content_Types].xml", false, appVersion);
        
        let subObj = ContentTypesJson["Types"]["Override"];
        let slidesLocArray = [];
        let slideLayoutsLocArray = [];
        for (let i = 0; i < subObj.length; i++) {
            switch (subObj[i]["attrs"]["ContentType"]) {
                case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
                    slidesLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
                    break;
                case "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml":
                    slideLayoutsLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
                    break;
            }
        }
        return {
            "slides": slidesLocArray,
            "slideLayouts": slideLayoutsLocArray
        };
    }

    /**
     * 获取幻灯片尺寸并设置默认文本样式
     * @param {Object} zip - JSZip实例
     * @param {Object} settings - 设置对象
     * @param {number} SLIDE_FACTOR - 尺寸转换因子
     * @returns {Promise<Object>} 包含width和height的对象
     */
    async function getSlideSizeAndSetDefaultTextStyle(zip, settings) {
        //get app version
        let app = await PPTXXmlUtils.readXmlFile(zip, "docProps/app.xml");
        app["Properties"]["AppVersion"];

        //get slide dimensions
        let rtenObj = {};
        let content = await PPTXXmlUtils.readXmlFile(zip, "ppt/presentation.xml");
        let sldSzAttrs = content["p:presentation"]["p:sldSz"]["attrs"];
        let sldSzWidth = parseInt(sldSzAttrs["cx"]);
        let sldSzHeight = parseInt(sldSzAttrs["cy"]);
        sldSzAttrs["type"];

        //1 inches  = 96px = 2.54cm
        // 1 EMU = 1 / 914400 inch
        // Pixel = EMUs * Resolution / 914400;  (Resolution = 96)
        //var standardHeight = 6858000;
        //console.log("SLIDE_FACTOR: ", SLIDE_FACTOR, "standardHeight:", standardHeight, (standardHeight - sldSzHeight) / standardHeight)
        
        //SLIDE_FACTOR = (96 * (1 + ((standardHeight - sldSzHeight) / standardHeight))) / 914400 ;

        //SLIDE_FACTOR = SLIDE_FACTOR + sldSzHeight*((standardHeight - sldSzHeight) / standardHeight) ;

        //var ration = sldSzWidth / sldSzHeight;
        
        //Scale
        // var viewProps = readXmlFile(zip, "ppt/viewProps.xml");
        // var scaleLoc = getTextByPathList(viewProps, ["p:viewPr", "p:slideViewPr", "p:cSldViewPr", "p:cViewPr","p:scale"]);
        // var scaleXnodes, scaleX = 1, scaleYnode, scaleY = 1;
        // if (scaleLoc !== undefined){
        //     scaleXnodes = scaleLoc["a:sx"]["attrs"];
        //     var scaleXnodesN = scaleXnodes["n"];
        //     var scaleXnodesD = scaleXnodes["d"];
        //     if (scaleXnodesN !== undefined && scaleXnodesD !== undefined && scaleXnodesN != 0){
        //         scaleX = parseInt(scaleXnodesD)/parseInt(scaleXnodesN);
        //     }
        //     scaleYnode = scaleLoc["a:sy"]["attrs"];
        //     var scaleYnodeN = scaleYnode["n"];
        //     var scaleYnodeD = scaleYnode["d"];
        //     if (scaleYnodeN !== undefined && scaleYnodeD !== undefined && scaleYnodeN != 0) {
        //         scaleY = parseInt(scaleYnodeD) / parseInt(scaleYnodeN) ;
        //     }

        // }
        //console.log("scaleX: ", scaleX, "scaleY:", scaleY)
        //SLIDE_FACTOR = SLIDE_FACTOR * scaleX;

        const defaultTextStyle = content["p:presentation"]["p:defaultTextStyle"];

        const slideWidth = sldSzWidth * SLIDE_FACTOR$1 + settings.incSlide.width|0;// * scaleX;//parseInt(sldSzAttrs["cx"]) * 96 / 914400;
        const slideHeight = sldSzHeight * SLIDE_FACTOR$1 + settings.incSlide.height|0;// * scaleY;//parseInt(sldSzAttrs["cy"]) * 96 / 914400;
        rtenObj = {
            "width": slideWidth,
            "height": slideHeight,
            defaultTextStyle
        };
        return rtenObj;
    }

    // Export to global namespace
    /**
     * 解析媒体文件路径
     * 根据PPTX标准，处理不同上下文下的媒体文件路径
     * @param {string} mediaPath - 媒体文件路径（来自resObj[target]）
     * @param {string} context - 上下文类型：'slide', 'master', 'layout'
     * @param {string} basePath - 基础路径（通常是当前XML文件所在目录）
     * @returns {string} 解析后的完整路径
     */
    function resolveMediaPath(mediaPath, context, basePath) {
        // 如果已经是绝对路径（以ppt/开头），直接返回
        if (mediaPath.startsWith('ppt/')) {
            return mediaPath;
        }
            
        // 处理相对路径
        let resolvedPath = mediaPath;
            
        // 根据上下文确定基础目录
        let baseDir = '';
        switch (context) {
            case 'slide':
                // 幻灯片中的媒体文件通常相对于ppt/slides/
                baseDir = 'ppt/slides/';
                break;
            case 'master':
                // 幻灯片母版中的媒体文件通常相对于ppt/slideMasters/
                baseDir = 'ppt/slideMasters/';
                break;
            case 'layout':
                // 版式中的媒体文件通常相对于ppt/slideLayouts/
                baseDir = 'ppt/slideLayouts/';
                break;
            default:
                // 默认情况，使用传入的基础路径
                baseDir = basePath || '';
        }
            
        // 处理路径中的../
        if (mediaPath.startsWith('../')) {
            // 移除../并构建相对于ppt/的路径
            resolvedPath = 'ppt/' + mediaPath.substring(3);
        } else if (!mediaPath.includes('/')) {
            // 如果没有路径分隔符，可能是直接在media目录下的文件
            resolvedPath = 'ppt/media/' + mediaPath;
        } else {
            // 其他相对路径，拼接基础目录
            resolvedPath = baseDir + mediaPath;
        }
            
        // 清理路径中的重复斜杠
        resolvedPath = resolvedPath.replace(/\/+/g, '/');
            
        // 移除开头的./
        if (resolvedPath.startsWith('./')) {
            resolvedPath = resolvedPath.substring(2);
        }
            
        return resolvedPath;
    }
        
    /**
     * 查找媒体文件（尝试多种可能的路径）
     * @param {Object} zip - JSZip实例
     * @param {string} originalPath - 原始路径
     * @param {string} context - 上下文类型
     * @param {string} basePath - 基础路径
     * @returns {Object|null} 找到的文件对象或null
     */
    function findMediaFile(zip, originalPath, context, basePath) {
        // 首先尝试原始路径
        let file = zip.file(originalPath);
        if (file) {
            return file;
        }
            
        // 尝试解析后的标准路径
        const resolvedPath = resolveMediaPath(originalPath, context, basePath);
        file = zip.file(resolvedPath);
        if (file) {
            return file;
        }
            
        // 尝试常见的替代路径
        const alternativePaths = [];
            
        // 如果是media目录下的文件，尝试不同的前缀
        if (originalPath.includes('media/') || !originalPath.includes('/')) {
            const fileName = originalPath.split('/').pop();
            alternativePaths.push(
                'ppt/media/' + fileName,
                'media/' + fileName,
                fileName
            );
        }
            
        // 如果包含embeddings，也尝试相关路径
        if (originalPath.includes('embeddings/')) {
            const fileName = originalPath.split('/').pop();
            alternativePaths.push(
                'ppt/embeddings/' + fileName,
                'embeddings/' + fileName
            );
        }
            
        // 尝试所有备选路径
        for (const altPath of alternativePaths) {
            file = zip.file(altPath);
            if (file) {
                return file;
            }
        }
            
        // 如果都没找到，返回null
        return null;
    }
        
    /**
     * 将ArrayBuffer转换为Base64字符串
     * @param {ArrayBuffer} arrayBuffer - 要转换的ArrayBuffer
     * @returns {string} Base64字符串
     */
    function base64ArrayBuffer(arrayBuffer) {
        // Node.js: use Buffer (fastest)
        if (typeof Buffer !== 'undefined' && Buffer.from) {
            return Buffer.from(arrayBuffer).toString('base64');
        }

        // Browser: use btoa with chunked processing for large buffers
        const bytes = new Uint8Array(arrayBuffer);
        const byteLength = bytes.byteLength;

        if (typeof btoa === 'function') {
            const CHUNK_SIZE = 0x8000;
            let binary = '';
            for (let i = 0; i < byteLength; i += CHUNK_SIZE) {
                const chunk = bytes.subarray(i, Math.min(i + CHUNK_SIZE, byteLength));
                binary += String.fromCharCode.apply(null, chunk);
            }
            return btoa(binary);
        }

        // Fallback: manual encoding
        const encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
        const byteRemainder = byteLength % 3;
        const mainLength = byteLength - byteRemainder;
        const parts = [];

        for (let i = 0; i < mainLength; i += 3) {
            const chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
            parts.push(
                encodings[(chunk & 16515072) >> 18] +
                encodings[(chunk & 258048) >> 12] +
                encodings[(chunk & 4032) >> 6] +
                encodings[chunk & 63]
            );
        }

        if (byteRemainder === 1) {
            const chunk = bytes[mainLength];
            parts.push(encodings[(chunk & 252) >> 2] + encodings[(chunk & 3) << 4] + '==');
        } else if (byteRemainder === 2) {
            const chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];
            parts.push(encodings[(chunk & 64512) >> 10] + encodings[(chunk & 1008) >> 4] + encodings[(chunk & 15) << 2] + '=');
        }

        return parts.join('');
    }

    function extractFileExtension(filename) {
            return filename.substr((~-filename.lastIndexOf(".") >>> 0) + 2);
        }
    function getMimeType(imgFileExt) {
            let mimeType = "";
            //console.log(imgFileExt)
            switch (imgFileExt.toLowerCase()) {
                case "jpg":
                case "jpeg":
                    mimeType = "image/jpeg";
                    break;
                case "png":
                    mimeType = "image/png";
                    break;
                case "gif":
                    mimeType = "image/gif";
                    break;
                case "emf": // Not native support
                    mimeType = "image/x-emf";
                    break;
                case "wmf": // Not native support
                    mimeType = "image/x-wmf";
                    break;
                case "svg":
                    mimeType = "image/svg+xml";
                    break;
                case "mp4":
                    mimeType = "video/mp4";
                    break;
                case "webm":
                    mimeType = "video/webm";
                    break;
                case "ogg":
                    mimeType = "video/ogg";
                    break;
                case "avi":
                    mimeType = "video/avi";
                    break;
                case "mpg":
                    mimeType = "video/mpg";
                    break;
                case "wmv":
                    mimeType = "video/wmv";
                    break;
                case "mp3":
                    mimeType = "audio/mpeg";
                    break;
                case "wav":
                    mimeType = "audio/wav";
                    break;
                case "bmp":
                    mimeType = "image/bmp";
                    break;
                case "webp":
                    mimeType = "image/webp";
                    break;
                case "tif":
                case "tiff":
                    mimeType = "image/tiff";
                    break;
            }
            return mimeType;
        }

    
        function getPosition(slideSpNode, pNode, slideLayoutSpNode, slideMasterSpNode, sType) {
            let off;
            let x = -1, y = -1;

            if (slideSpNode !== undefined) {
                off = slideSpNode["a:off"]["attrs"];
            }

            if (off === undefined && slideLayoutSpNode !== undefined) {
                off = slideLayoutSpNode["a:off"]["attrs"];
            } else if (off === undefined && slideMasterSpNode !== undefined) {
                off = slideMasterSpNode["a:off"]["attrs"];
            }
            let offX = 0, offY = 0;
            // 计算子元素的偏移量
            if (sType == "group" && pNode !== undefined) {
                var grpXfrmNode = PPTXXmlUtils.getTextByPathList(pNode, ["p:grpSpPr", "a:xfrm"]);
                if (grpXfrmNode !== undefined && grpXfrmNode["a:chOff"] !== undefined && grpXfrmNode["a:chOff"]["attrs"] !== undefined) {
                    offX = parseInt(grpXfrmNode["a:chOff"]["attrs"]["x"]) * SLIDE_FACTOR$1;
                    offY = parseInt(grpXfrmNode["a:chOff"]["attrs"]["y"]) * SLIDE_FACTOR$1;
                    offX = Math.round(offX * 100) / 100;
                    offY = Math.round(offY * 100) / 100;
                }
            } else if (sType == "group-abs" && pNode !== undefined) {
                // 当容器扩展时，子元素使用相对chOff的绝对定位
                // 但相对于新的容器位置(minY, minX)
                var grpXfrmNode = PPTXXmlUtils.getTextByPathList(pNode, ["p:grpSpPr", "a:xfrm"]);
                if (grpXfrmNode !== undefined && grpXfrmNode["a:chOff"] !== undefined && grpXfrmNode["a:chOff"]["attrs"] !== undefined) {
                    offX = parseInt(grpXfrmNode["a:chOff"]["attrs"]["x"]) * SLIDE_FACTOR$1;
                    offY = parseInt(grpXfrmNode["a:chOff"]["attrs"]["y"]) * SLIDE_FACTOR$1;
                    offX = Math.round(offX * 100) / 100;
                    offY = Math.round(offY * 100) / 100;
                }
            }
            if (sType == "group-rotate" && pNode["p:grpSpPr"] !== undefined) {
                var xfrmNode = pNode["p:grpSpPr"]["a:xfrm"];
                // var ox = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * SLIDE_FACTOR;
                // var oy = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * SLIDE_FACTOR;
                var chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * SLIDE_FACTOR$1;
                var chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * SLIDE_FACTOR$1;

                offX = Math.round(chx * 100) / 100;
                offY = Math.round(chy * 100) / 100;
            }
            if (off === undefined) {
                return "";
            } else {
                x = parseInt(off["x"]) * SLIDE_FACTOR$1;
                y = parseInt(off["y"]) * SLIDE_FACTOR$1;
                x = Math.round(x * 100) / 100;
                y = Math.round(y * 100) / 100;
                // 当元素在组合内时，减去chOff得到相对于组合的位置
                let finalX = Math.round((x - offX) * 100) / 100;
                let finalY = Math.round((y - offY) * 100) / 100;
                return (isNaN(finalX) || isNaN(finalY)) ? "" : "top:" + finalY + "px; left:" + finalX + "px;";
            }

        }

        function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
            let ext = undefined;
            let w = -1, h = -1;

            if (slideSpNode !== undefined) {
                ext = slideSpNode["a:ext"]["attrs"];
            } else if (slideLayoutSpNode !== undefined) {
                ext = slideLayoutSpNode["a:ext"]["attrs"];
            } else if (slideMasterSpNode !== undefined) {
                ext = slideMasterSpNode["a:ext"]["attrs"];
            }

            if (ext === undefined) {
                return "";
            } else {
                w = parseInt(ext["cx"]) * SLIDE_FACTOR$1;
                h = parseInt(ext["cy"]) * SLIDE_FACTOR$1;
                w = Math.round(w * 100) / 100;
                h = Math.round(h * 100) / 100;
                return (isNaN(w) || isNaN(h)) ? "" : "width:" + w + "px; height:" + h + "px;";
            }

        }

    function IsVideoLink(vdoFile) {
            /*
            var ext = PPTXXmlUtils.extractFileExtension(vdoFile);
            if (ext.length == 3){
                return false;
            }else{
                return true;
            }
            */
            let urlregex = /^(https?|ftp):\/\/([a-zA-Z0-9.-]+(:[a-zA-Z0-9.&%$-]+)*@)*((25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]?)(\.(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])){3}|([a-zA-Z0-9-]+\.)*[a-zA-Z0-9-]+\.(com|edu|gov|int|mil|net|org|biz|arpa|info|name|pro|aero|coop|museum|[a-zA-Z]{2}))(:[0-9]+)*(\/($|[a-zA-Z0-9.,?'\\+&%$#=~_-]+))*$/;
            return urlregex.test(vdoFile);
        }

        /**
         * 将视频URL转换为YouTube embed格式
         * 支持的URL格式：
         * - https://www.youtube.com/watch?v=VIDEO_ID
         * - https://youtube.com/watch?v=VIDEO_ID
         * - https://youtu.be/VIDEO_ID
         * - https://www.youtube.com/embed/VIDEO_ID
         * - https://www.youtube.com/v/VIDEO_ID
         * 
         * @param {string} videoUrl - 原始视频URL
         * @returns {string} 转换后的embed URL
         */
        function convertYouTubeUrl(videoUrl) {
            if (!videoUrl) return videoUrl;
            
            // YouTube视频ID正则表达式
            const youtubeIdPatterns = [
                /(?:youtube\.com\/watch\?v=|youtu\.be\/|youtube\.com\/embed\/|youtube\.com\/v\/|youtube\.com\/shorts\/)([a-zA-Z0-9_-]{11})/,
                /youtube\.com\/watch\?.*list=([a-zA-Z0-9_-]+)/,
                /youtube\.com\/playlist\?list=([a-zA-Z0-9_-]+)/
            ];
            
            // 检查是否是YouTube链接
            const isYouTubeUrl = videoUrl.includes('youtube.com') || videoUrl.includes('youtu.be');
            
            if (!isYouTubeUrl) {
                return videoUrl;
            }
            
            // 提取YouTube视频ID
            for (const pattern of youtubeIdPatterns) {
                const match = videoUrl.match(pattern);
                if (match && match[1]) {
                    const videoId = match[1];
                    // 转换为embed格式，添加必要的参数
                    const embedUrl = `https://www.youtube.com/embed/${videoId}?rel=0&modestbranding=1`;
                    return embedUrl;
                }
            }
            
            // 如果已经是embed格式，直接返回
            if (videoUrl.includes('/embed/')) {
                return videoUrl;
            }
            
            return videoUrl;
        }

        /**
         * 将视频URL转换为Vimeo embed格式
         * 支持的URL格式：
         * - https://vimeo.com/VIDEO_ID
         * - https://www.vimeo.com/VIDEO_ID
         * 
         * @param {string} videoUrl - 原始视频URL
         * @returns {string} 转换后的embed URL
         */
        function convertVimeoUrl(videoUrl) {
            if (!videoUrl) return videoUrl;
            
            // 检查是否是Vimeo链接
            const isVimeoUrl = videoUrl.includes('vimeo.com');
            
            if (!isVimeoUrl) {
                return videoUrl;
            }
            
            // 提取Vimeo视频ID
            const vimeoMatch = videoUrl.match(/vimeo\.com\/(\d+)/);
            if (vimeoMatch && vimeoMatch[1]) {
                const videoId = vimeoMatch[1];
                return `https://player.vimeo.com/video/${videoId}?title=0&byline=0&portrait=0`;
            }
            
            return videoUrl;
        }

        /**
         * 统一转换视频URL为embed格式
         * @param {string} videoUrl - 原始视频URL
         * @returns {string} 转换后的embed URL
         */
        function convertVideoToEmbed(videoUrl) {
            if (!videoUrl) return videoUrl;
            
            // 优先处理YouTube
            if (videoUrl.includes('youtube.com') || videoUrl.includes('youtu.be')) {
                return convertYouTubeUrl(videoUrl);
            }
            
            // 处理Vimeo
            if (videoUrl.includes('vimeo.com')) {
                return convertVimeoUrl(videoUrl);
            }
            
            // 其他视频链接直接返回（可能是iframe支持的直接嵌入URL）
            return videoUrl;
        }   
    return {
        getTextByPathStr: getTextByPathStr,
        getTextByPathList: getTextByPathList,
        setTextByPathList: setTextByPathList,
        eachElement: eachElement,
        angleToDegrees: angleToDegrees,
        degreesToRadians: degreesToRadians,
        escapeHtml: escapeHtml,
        readXmlFile: readXmlFile,
        getContentTypes: getContentTypes,
        getSlideSizeAndSetDefaultTextStyle: getSlideSizeAndSetDefaultTextStyle,
        resolveMediaPath: resolveMediaPath,
        findMediaFile: findMediaFile,
        base64ArrayBuffer: base64ArrayBuffer,
        extractFileExtension,
        getMimeType,
        getPosition,
        getSize,
        IsVideoLink,
        convertYouTubeUrl,
        convertVimeoUrl,
        convertVideoToEmbed,
    };
})();

/**
 * 样式处理模块
 * 
 * 处理 PPTX 文件中的各种样式属性，包括：
 * - 填充类型（纯色、渐变、图片、图案等）
 * - 边框样式
 * - 阴影效果
 * - 3D 效果
 * - 反射效果
 * 
 * @module utils/style
 */


// 创建 tinycolor 工厂函数以保持向后兼容
const tinycolor$1 = (color, opts) => new TinyColor(color, opts);



function getFillType(node) {
            //Need to test/////////////////////////////////////////////
            //SOLID_FILL
            //PIC_FILL
            //GRADIENT_FILL
            //PATTERN_FILL
            //NO_FILL
            let fillType = "";
            if (node === undefined) {
                return fillType;
            }
            if (node["a:noFill"] !== undefined) {
                fillType = "NO_FILL";
            }
            if (node["a:solidFill"] !== undefined) {
                fillType = "SOLID_FILL";
            }
            if (node["a:gradFill"] !== undefined) {
                fillType = "GRADIENT_FILL";
            }
            if (node["a:pattFill"] !== undefined) {
                fillType = "PATTERN_FILL";
            }
            if (node["a:blipFill"] !== undefined) {
                fillType = "PIC_FILL";
            }
            if (node["a:grpFill"] !== undefined) {
                fillType = "GROUP_FILL";
            }


            return fillType;
        }
    // function hexToRgbNew(hex) {
        //     let arrBuff = new ArrayBuffer(4);
        //     let vw = new DataView(arrBuff);
        //     vw.setUint32(0, parseInt(hex, 16), false);
        //     let arrByte = new Uint8Array(arrBuff);
        //     return arrByte[1] + "," + arrByte[2] + "," + arrByte[3];
        // }
        async function getShapeFill(node, pNode, isSvgMode, warpObj, source) {

            // 1. presentationML
            // p:spPr/ [a:noFill, solidFill, gradFill, blipFill, pattFill, grpFill]
            // From slide
            //Fill Type:

            let fillType = getFillType (PPTXXmlUtils.getTextByPathList(node, ["p:spPr"]));
            //let noFill = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:noFill"]);
            let fillColor;
            if (fillType === "NO_FILL") {
                return isSvgMode ? "none" : "";
            } else if (fillType === "SOLID_FILL") {
                let shpFill = node["p:spPr"]["a:solidFill"];
                fillColor = getSolidFill(shpFill, undefined, undefined, warpObj);
            } else if (fillType === "GRADIENT_FILL") {
                let shpFill = node["p:spPr"]["a:gradFill"];
                fillColor = getGradientFill(shpFill, warpObj);
            } else if (fillType === "PATTERN_FILL") {
                let shpFill = node["p:spPr"]["a:pattFill"];
                fillColor = getPatternFill(shpFill, warpObj);
            } else if (fillType === "PIC_FILL") {
                let shpFill = node["p:spPr"]["a:blipFill"];
                fillColor = await getPicFill(source, shpFill, warpObj);
            }



            // 2. drawingML namespace
            if (fillColor === undefined) {
                let clrName = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:fillRef"]);
                let idx = parseInt (PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:fillRef", "attrs", "idx"]));
                if (idx == 0 || idx == 1000) {
                    //no fill
                    return isSvgMode ? "none" : "";
                }
                fillColor = getSolidFill(clrName, undefined, undefined, warpObj);
            }
            // 3. is group fill
            if (fillColor === undefined) {
                let grpFill = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:grpFill"]);
                if (grpFill !== undefined) {
                    //fillColor = getSolidFill(clrName, undefined, undefined, undefined, warpObj);
                    //get parent fill style - TODO

                    let grpShpFill = pNode["p:grpSpPr"];
                    let spShpNode = { "p:spPr": grpShpFill };
                    return await getShapeFill(spShpNode, node, isSvgMode, warpObj, source);
                } else if (fillType === "NO_FILL") {
                    return isSvgMode ? "none" : "";
                }
            }


            if (fillColor !== undefined) {
                if (fillType === "GRADIENT_FILL") {
                    if (isSvgMode) {
    
                        return fillColor;
                    } else {
                        let colorAry = fillColor.color;
                        let rot = fillColor.rot;

                        let bgcolor = `background: linear-gradient(${rot}deg,`;
                        for (let i = 0; i < colorAry.length; i++) {
                            if (i == colorAry.length - 1) {
                                bgcolor += "#" + colorAry[i] + ");";
                            } else {
                                bgcolor += "#" + colorAry[i] + ", ";
                            }

                        }
                        return bgcolor;
                    }
                } else if (fillType === "PIC_FILL") {
                    if (isSvgMode) {
                        // 当 isSvgMode 为 true 时，返回图像 URL 而不是整个对象
                        if (typeof fillColor === 'object' && fillColor.img) {
                            return fillColor.img;
                        } else {
                            return fillColor;
                        }
                    } else {
                        if (typeof fillColor === 'object' && fillColor.img) {
                            return `background-image:url(${fillColor.img}); background-size: ${fillColor.backgroundSize}; background-position: ${fillColor.backgroundPosition}; background-repeat: ${fillColor.backgroundRepeat};`;
                        } else {
                            return `background-image:url(${fillColor});`;
                        }
                    }
                } else if (fillType === "PATTERN_FILL") {
                    /////////////////////////////////////////////////////////////Need to check -----------TODO
                    // if (isSvgMode) {
                    //     let color = tinycolor(fillColor);
                    //     fillColor = color.toRgbString();

                    //     return fillColor;
                    // } else {
                    let bgPtrn = "", bgSize = "", bgPos = "";
                    bgPtrn = fillColor[0];
                    if (fillColor[1] !== null && fillColor[1] !== undefined && fillColor[1] != "") {
                        bgSize = " background-size:" + fillColor[1] + ";";
                    }
                    if (fillColor[2] !== null && fillColor[2] !== undefined && fillColor[2] != "") {
                        bgPos = " background-position:" + fillColor[2] + ";";
                    }
                    return "background: " + bgPtrn + ";" + bgSize + bgPos;
                    //}
                } else {
                    if (isSvgMode) {
                        let color = tinycolor$1(fillColor);
                        fillColor = color.toRgbString();

                        return fillColor;
                    } else {

                        return `background-color: #${fillColor};`;
                    }
                }
            } else {
                if (isSvgMode) {
                    return "none";
                } else {
                    return "background-color: inherit;";
                }

            }

        }

        
        function getFontType(node, type, warpObj, pFontStyle) {
            let typeface = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);

            if (typeface === undefined) {
                let fontIdx = "";
                let fontGrup = "";
                if (pFontStyle !== undefined) {
                    fontIdx = PPTXXmlUtils.getTextByPathList(pFontStyle, ["attrs", "idx"]);
                }
                let fontSchemeNode = PPTXXmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:fontScheme"]);
                if (fontIdx == "") {
                    if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                        fontIdx = "major";
                    } else {
                        fontIdx = "minor";
                    }
                }
                fontGrup = "a:" + fontIdx + "Font";
                typeface = PPTXXmlUtils.getTextByPathList(fontSchemeNode, [fontGrup, "a:latin", "attrs", "typeface"]);
            }

            return (typeface === undefined) ? "inherit" : typeface;
        }

        async function getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj) {
            //text border using: text-shadow: -1px 0 black, 0 1px black, 1px 0 black, 0 -1px black;
            //{getFontColor(..) return color} -> getFontColorPr(..) return array[color,textBordr/shadow]
            //https://stackoverflow.com/questions/2570972/css-font-border
            //https://www.w3schools.com/cssref/css3_pr_text-shadow.asp
            //themeContent

            let rPrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr"]);
            let filTyp, color, textBordr, colorType = "", highlightColor = "";


            if (rPrNode !== undefined) {
                filTyp = getFillType(rPrNode);
                if (filTyp == "SOLID_FILL") {
                    let solidFillNode = rPrNode["a:solidFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                    color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
                    let highlightNode = rPrNode["a:highlight"];
                    if (highlightNode !== undefined) {
                        highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                    colorType = "solid";
                } else if (filTyp == "PATTERN_FILL") {
                    let pattFill = rPrNode["a:pattFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:pattFill"]);
                    color = getPatternFill(pattFill, warpObj);
                    colorType = "pattern";
                } else if (filTyp == "PIC_FILL") {
                    color = await getBgPicFill(rPrNode, "slideBg", warpObj, undefined);
                    //color = getPicFill("slideBg", rPrNode["a:blipFill"], warpObj);
                    colorType = "pic";
                } else if (filTyp == "GRADIENT_FILL") {
                    let shpFill = rPrNode["a:gradFill"];
                    color = getGradientFill(shpFill, warpObj);
                    colorType = "gradient";
                } 
            }
            if (color === undefined && PPTXXmlUtils.getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]) !== undefined) {
                //lstStyle
                let lstStyledefRPr = PPTXXmlUtils.getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]);
                filTyp = getFillType(lstStyledefRPr);
                if (filTyp == "SOLID_FILL") {
                    let solidFillNode = lstStyledefRPr["a:solidFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                    color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
                    let highlightNode = lstStyledefRPr["a:highlight"];
                    if (highlightNode !== undefined) {
                        highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                    colorType = "solid";
                } else if (filTyp == "PATTERN_FILL") {
                    let pattFill = lstStyledefRPr["a:pattFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:pattFill"]);
                    color = getPatternFill(pattFill, warpObj);
                    colorType = "pattern";
                } else if (filTyp == "PIC_FILL") {
                    color = await getBgPicFill(lstStyledefRPr, "slideBg", warpObj, undefined);
                    //color = getPicFill("slideBg", rPrNode["a:blipFill"], warpObj);
                    colorType = "pic";
                } else if (filTyp == "GRADIENT_FILL") {
                    let shpFill = lstStyledefRPr["a:gradFill"];
                    color = getGradientFill(shpFill, warpObj);
                    colorType = "gradient";
                }

            }
            if (color === undefined) {
                let sPstyle = PPTXXmlUtils.getTextByPathList(pNode, ["p:style", "a:fontRef"]);
                if (sPstyle !== undefined) {
                    color = getSolidFill(sPstyle, undefined, undefined, warpObj);
                    if (color !== undefined) {
                        colorType = "solid";
                    }
                    let highlightNode = sPstyle["a:highlight"]; //is "a:highlight" node in 'a:fontRef' ?
                    if (highlightNode !== undefined) {
                        highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                }
                if (color === undefined) {
                    if (pFontStyle !== undefined) {
                        color = getSolidFill(pFontStyle, undefined, undefined, warpObj);
                        if (color !== undefined) {
                            colorType = "solid";
                        }
                    }
                }
            }
            //console.log("getFontColorPr node", node, "colorType: ", colorType,"color: ",color)

            if (color === undefined) {
                let layoutMasterNode = getLayoutAndMasterNode(pNode, idx, type, warpObj);
                let pPrNodeLaout = layoutMasterNode.nodeLaout;
                let pPrNodeMaster = layoutMasterNode.nodeMaster;

                if (pPrNodeLaout !== undefined) {
                    let defRpRLaout = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:solidFill"]);
                    if (defRpRLaout !== undefined) {
                        color = getSolidFill(defRpRLaout, undefined, undefined, warpObj);
                        let highlightNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:highlight"]);
                        if (highlightNode !== undefined) {
                            highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                        }
                        colorType = "solid";
                    }
                }
                if (color === undefined) {
                    if (pPrNodeMaster !== undefined) {
                        let defRprMaster = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:solidFill"]);
                        if (defRprMaster !== undefined) {
                            color = getSolidFill(defRprMaster, undefined, undefined, warpObj);
                            let highlightNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:highlight"]);
                            if (highlightNode !== undefined) {
                                highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                            }
                            colorType = "solid";
                        }
                    }
                }
            }
            let txtEffects = [];
            let txtEffObj = {};
            //textBordr
            let txtBrdrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:ln"]);
            textBordr = "";
            if (txtBrdrNode !== undefined && txtBrdrNode["a:noFill"] === undefined) {
                let txBrd = getBorder(node, pNode, false, "text", warpObj);
                let txBrdAry = txBrd.split(" ");
                //let brdSize = (parseInt(txBrdAry[0].substring(0, txBrdAry[0].indexOf("pt")))) + "px";
                let brdSize = (parseInt(txBrdAry[0].substring(0, txBrdAry[0].indexOf("px")))) + "px";
                let brdClr = txBrdAry[2];
                //let brdTyp = txBrdAry[1]; //not in use
                //console.log("getFontColorPr txBrdAry:", txBrdAry)
                if (colorType == "solid") {
                    textBordr = `-${brdSize} 0 ${brdClr}, 0 ${brdSize} ${brdClr}, ${brdSize} 0 ${brdClr}, 0 -${brdSize} ${brdClr}`;
                    // if (oShadowStr != "") {
                    //     textBordr += "," + oShadowStr;
                    // } else {
                    //     textBordr += ";";
                    // }
                    txtEffects.push(textBordr);
                } else {
                    //textBordr = brdSize + " " + brdClr;
                    txtEffObj.border = brdSize + " " + brdClr;
                }
            }
            // else {
            //     //if no border but exist/not exist shadow
            //     if (colorType == "solid") {
            //         textBordr = oShadowStr;
            //     } else {
            //         //TODO
            //     }
            // }
            let txtGlowNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:glow"]);
            let oGlowStr = "";
            if (txtGlowNode !== undefined) {
                let glowClr = getSolidFill(txtGlowNode, undefined, undefined, warpObj);
                let rad = (txtGlowNode["attrs"]["rad"]) ? (txtGlowNode["attrs"]["rad"] * SLIDE_FACTOR$1) : 0;
                oGlowStr = "0 0 " + rad + "px #" + glowClr +
                    ", 0 0 " + rad + "px #" + glowClr +
                    ", 0 0 " + rad + "px #" + glowClr +
                    ", 0 0 " + rad + "px #" + glowClr +
                    ", 0 0 " + rad + "px #" + glowClr +
                    ", 0 0 " + rad + "px #" + glowClr +
                    ", 0 0 " + rad + "px #" + glowClr;
                if (colorType == "solid") {
                    txtEffects.push(oGlowStr);
                } else {
                    // txtEffObj.glow = {
                    //     radiuse: rad,
                    //     color: glowClr
                    // } 
                    txtEffects.push(
                        "drop-shadow(0 0 " + rad / 3 + "px #" + glowClr + ") " +
                        "drop-shadow(0 0 " + rad * 2 / 3 + "px #" + glowClr + ") " +
                        "drop-shadow(0 0 " + rad + "px #" + glowClr + ")"
                    );
                }
            }
            // Check for direct shadow effect in text run properties
            let txtShadow = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:outerShdw"]);
            let oShadowStr = "";
            
            // Check for reflection effect
            let txtReflection = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:reflection"]);
            if (txtReflection !== undefined) {
                // 解析反射效果参数
                parseInt(txtReflection.attrs?.["blurRad"] || "0") * SLIDE_FACTOR$1;
                parseInt(txtReflection.attrs?.["stA"] || "100000"); // 开始透明度 (0-100000)
                parseInt(txtReflection.attrs?.["endA"] || "0");   // 结束透明度
                parseInt(txtReflection.attrs?.["dist"] || "0") * SLIDE_FACTOR$1; // 距离
                parseInt(txtReflection.attrs?.["dir"] || "5400000"); // 方向 (0-21600000)
            }
            
            // Check for soft edge effect
            let txtSoftEdge = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:softEdge"]);
            if (txtSoftEdge !== undefined) {
                parseInt(txtSoftEdge.attrs?.["rad"] || "0") * SLIDE_FACTOR$1;
            }
            
            // If no direct shadow, check effectRef from p:style
            if (txtShadow === undefined) {
                var effectRefNode = PPTXXmlUtils.getTextByPathList(pNode, ["p:style", "a:effectRef"]);
                if (effectRefNode !== undefined) {
                    var effectIdx = PPTXXmlUtils.getTextByPathList(effectRefNode, ["attrs", "idx"]);
                    if (effectIdx !== undefined && warpObj["themeContent"] !== undefined) {
                        // Access the effect style from the theme
                        var effectStyleLst = PPTXXmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:fmtScheme", "a:effectStyleLst", "a:effectStyle"]);
                        if (effectStyleLst !== undefined) {
                            // Ensure effectStyleLst is an array
                            if (!Array.isArray(effectStyleLst)) {
                                effectStyleLst = [effectStyleLst];
                            }
                            var idx = Number(effectIdx); // idx is 0-based, not 1-based
                            if (idx >= 0 && effectStyleLst[idx] !== undefined) {
                                txtShadow = PPTXXmlUtils.getTextByPathList(effectStyleLst[idx], ["a:effectLst", "a:outerShdw"]);
                            }
                        }
                    }
                }
            }
            
            if (txtShadow !== undefined) {
                //https://developer.mozilla.org/en-US/docs/Web/CSS/filter-function/drop-shadow()
                //https://stackoverflow.com/questions/60468487/css-text-with-linear-gradient-shadow-and-text-outline
                //https://css-tricks.com/creating-playful-effects-with-css-text-shadows/
                //https://designshack.net/articles/css/12-fun-css-text-shadows-you-can-copy-and-paste/

                let shadowClr = getSolidFill(txtShadow, undefined, undefined, warpObj);
                let outerShdwAttrs = txtShadow["attrs"];
                // algn: "bl"
                // dir: "2640000"
                // dist: "38100"
                // rotWithShape: "0/1" - Specifies whether the shadow rotates with the shape if the shape is rotated.
                //blurRad (Blur Radius) - Specifies the blur radius of the shadow.
                //kx (Horizontal Skew) - Specifies the horizontal skew angle.
                //ky (Vertical Skew) - Specifies the vertical skew angle.
                //sx (Horizontal Scaling Factor) - Specifies the horizontal scaling SLIDE_FACTOR; negative scaling causes a flip.
                //sy (Vertical Scaling Factor) - Specifies the vertical scaling SLIDE_FACTOR; negative scaling causes a flip.
                outerShdwAttrs["algn"];
                let dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
                let dist = parseInt(outerShdwAttrs["dist"]) * SLIDE_FACTOR$1;//(px) //* (3 / 4); //(pt)
                outerShdwAttrs["rotWithShape"];
                let blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * SLIDE_FACTOR$1 + "px") : "";
                (outerShdwAttrs["sx"]) ? (parseInt(outerShdwAttrs["sx"]) / 100000) : 1;
                (outerShdwAttrs["sy"]) ? (parseInt(outerShdwAttrs["sy"]) / 100000) : 1;
                let vx = dist * Math.sin(dir * Math.PI / 180);
                let hx = dist * Math.cos(dir * Math.PI / 180);

                //console.log("getFontColorPr outerShdwAttrs:", outerShdwAttrs, ", shadowClr:", shadowClr, ", algn: ", algn, ",dir: ", dir, ", dist: ", dist, ",rotWithShape: ", rotWithShape, ", color: ", color)

                if (!isNaN(vx) && !isNaN(hx)) {
                    oShadowStr = hx + "px " + vx + "px " + blurRad + " #" + shadowClr;// + ";";
                    if (colorType == "solid") {
                        txtEffects.push(oShadowStr);
                    } else {

                        // txtEffObj.oShadow = {
                        //     hx: hx,
                        //     vx: vx,
                        //     radius: blurRad,
                        //     color: shadowClr
                        // }

                        //txtEffObj.oShadow = hx + "px " + vx + "px " + blurRad + " #" + shadowClr;

                        txtEffects.push("drop-shadow(" + hx + "px " + vx + "px " + blurRad + " #" + shadowClr + ")");
                    }
                }
                //console.log("getFontColorPr vx:", vx, ", hx: ", hx, ", sx: ", sx, ", sy: ", sy, ",oShadowStr: ", oShadowStr)
            }
            //console.log("getFontColorPr>>> color:", color)
            // if (color === undefined || color === "FFF") {
            //     color = "#000";
            // } else {
            //     color = "" + color;
            // }
            let text_effcts = "", txt_effects;
            if (colorType == "solid") {
                if (txtEffects.length > 0) {
                    text_effcts = txtEffects.join(",");
                }
                txt_effects = text_effcts + ";";
            } else {
                if (txtEffects.length > 0) {
                    text_effcts = txtEffects.join(" ");
                }
                txtEffObj.effcts = text_effcts;
                txt_effects = txtEffObj;
            }
            //console.log("getFontColorPr txt_effects:", txt_effects)

            //return [color, textBordr, colorType];
            return [color, txt_effects, colorType, highlightColor];
        }
        function getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj) {
            // if(type == "sldNum")
            //console.log("getFontSize node:", node, "lstStyle", lstStyle, "lvl:", lvl, 'type:', type, "warpObj:", warpObj)
            let lstStyle = (textBodyNode !== undefined)? textBodyNode["a:lstStyle"] : undefined;
            let lvlpPr = "a:lvl" + lvl + "pPr";
            let fontSize = undefined;
            let sz, kern;
            if (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"] && node["a:rPr"]["attrs"]["sz"] !== undefined) {
                fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
            }
            if (isNaN(fontSize) || fontSize === undefined && node["a:fld"] !== undefined) {
                sz = PPTXXmlUtils.getTextByPathList(node["a:fld"], ["a:rPr", "attrs", "sz"]);
                fontSize = parseInt(sz) / 100;
            }
            if ((isNaN(fontSize) || fontSize === undefined) && node["a:t"] === undefined) {
                sz = PPTXXmlUtils.getTextByPathList(node["a:endParaRPr"], [ "attrs", "sz"]);
                fontSize = parseInt(sz) / 100;
            }
            if ((isNaN(fontSize) || fontSize === undefined) && lstStyle !== undefined) {
                sz = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlpPr, "a:defRPr", "attrs", "sz"]);
                fontSize = parseInt(sz) / 100;
            }
            if (textBodyNode !== undefined){
                PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:spAutoFit"]);
            }
            if (isNaN(fontSize) || fontSize === undefined) {
                // if (type == "shape" || type == "textBox") {
                //     type = "body";
                //     lvlpPr = "a:lvl1pPr";
                // }
                        sz = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                fontSize = parseInt(sz) / 100;
                kern = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            }

            if (isNaN(fontSize) || fontSize === undefined) {
                // if (type == "shape" || type == "textBox") {
                //     type = "body";
                //     lvlpPr = "a:lvl1pPr";
                // }
                sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                if (sz === undefined) {
                    if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                        sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                        kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    } else if (type == "body" || type == "obj" || type == "dt" || type == "sldNum" || type === "textBox") {
                        sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                        kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    }
                    else if (type == "shape") {
                        //textBox and shape text does not indent
                        // 普通形状使用 otherStyle，与原始库保持一致
                        sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                        kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    }

                    if (sz === undefined) {
                        sz = PPTXXmlUtils.getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "sz"]);
                        kern = (kern === undefined)? PPTXXmlUtils.getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "kern"]) : undefined;
                    }
                    //  else if (type === undefined || type == "shape") {
                    //     sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    //     kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    // } 
                    // else if (type == "textBox") {
                    //     sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    //     kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    // }
                } 
                fontSize = parseInt(sz) / 100;
            }

            let baseline = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
            if (baseline !== undefined && !isNaN(fontSize)) {
                let baselineVl = parseInt(baseline) / 100000;
                //fontSize -= 10; 
                // fontSize = fontSize * baselineVl;
                fontSize -= baselineVl;
            }

            // 如果仍然没有找到字体大小，尝试从段落的默认样式中获取
            if (isNaN(fontSize) || fontSize === undefined) {
                let pPrNode = node.parentNode && node.parentNode["a:pPr"];
                if (pPrNode) {
                    let defRPrNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:defRPr"]);
                    if (defRPrNode && defRPrNode["attrs"] && defRPrNode["attrs"]["sz"]) {
                        fontSize = parseInt(defRPrNode["attrs"]["sz"]) / 100;
                    }
                }
            }

            // 确保字体大小有效
            if (isNaN(fontSize) || fontSize === undefined) {
                fontSize = 18; // 默认字体大小为 18pt，与第一个片段保持一致
            }

            if (!isNaN(fontSize)){
                let normAutofit = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit", "attrs", "fontScale"]);
                if (normAutofit !== undefined && normAutofit != 0){
                    //console.log("fontSize", fontSize, "normAutofit: ", normAutofit, normAutofit/100000)
                    fontSize = Math.round(fontSize * (normAutofit / 100000));
                }
            }

            return isNaN(fontSize) ? ((type == "br") ? "initial" : "inherit") : (fontSize * FONT_SIZE_FACTOR + "px");// + "pt");
        }

        function getFontBold(node, type, slideMasterTextStyles) {
            if (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"] !== undefined) {
                const boldAttr = node["a:rPr"]["attrs"]["b"];
                return (boldAttr === "1" || boldAttr === "true" || boldAttr === "on") ? "bold" : "inherit";
            }
            return "inherit";
        }

        function getFontItalic(node, type, slideMasterTextStyles) {
            return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["i"] === "1") ? "italic" : "inherit";
        }

        function getFontDecoration(node, type, slideMasterTextStyles) {
            ///////////////////////////////Amir///////////////////////////////
            if (node["a:rPr"] !== undefined) {
                let underLine = node["a:rPr"]["attrs"]["u"] !== undefined ? node["a:rPr"]["attrs"]["u"] : "none";
                let strikethrough = node["a:rPr"]["attrs"]["strike"] !== undefined ? node["a:rPr"]["attrs"]["strike"] : 'noStrike';
                //console.log("strikethrough: "+strikethrough);

                if (underLine != "none" && strikethrough == "noStrike") {
                    return "underline";
                } else if (underLine == "none" && strikethrough != "noStrike") {
                    return "line-through";
                } else if (underLine != "none" && strikethrough != "noStrike") {
                    return "underline line-through";
                } else {
                    return "inherit";
                }
            } else {
                return "inherit";
            }
            /////////////////////////////////////////////////////////////////
            //return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["u"] === "sng") ? "underline" : "inherit";
        }
        ////////////////////////////////////Amir/////////////////////////////////////
        function getTextHorizontalAlign(node, pNode, type, warpObj) {
            //console.log("getTextHorizontalAlign: type: ", type, ", node: ", node)
            let getAlgn = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
            if (getAlgn === undefined) {
                getAlgn = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "algn"]);
            }
            if (getAlgn === undefined) {
                if (type == "title" || type == "ctrTitle" || type == "subTitle") {
                    let lvlIdx = 1;
                    let lvlNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
                    if (lvlNode !== undefined) {
                        lvlIdx = parseInt(lvlNode) + 1;
                    }
                    let lvlStr = "a:lvl" + lvlIdx + "pPr";
                    getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                    if (getAlgn === undefined) {
                        getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                        if (getAlgn === undefined) {
                            getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                            if (getAlgn === undefined && type === "subTitle") {
                                getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                            }
                        }
                    }
                } else if (type == "body") {
                    getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", "a:lvl1pPr", "attrs", "algn"]);
                } else {
                    getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", "a:lvl1pPr", "attrs", "algn"]);
                }

            }

            let align = "inherit";
            if (getAlgn !== undefined) {
                switch (getAlgn) {
                    case "l":
                        align = "left";
                        break;
                    case "r":
                        align = "right";
                        break;
                    case "ctr":
                        align = "center";
                        break;
                    case "just":
                        align = "justify";
                        break;
                    case "dist":
                        align = "justify";
                        break;
                    default:
                        align = "inherit";
                }
            }
            return align;
        }
        /////////////////////////////////////////////////////////////////////
        function getTextVerticalAlign(node, type, slideMasterTextStyles) {
            let baseline = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
            return baseline === undefined ? "baseline" : (parseInt(baseline) / 1000) + "%";
        }

        function getTableBorders(node, warpObj) {
            let borderStyle = "";
            if (node["a:bottom"] !== undefined) {
                let obj = {
                    "p:spPr": {
                        "a:ln": node["a:bottom"]["a:ln"]
                    }
                };
                let borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-bottom");
            }
            if (node["a:top"] !== undefined) {
                let obj = {
                    "p:spPr": {
                        "a:ln": node["a:top"]["a:ln"]
                    }
                };
                let borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-top");
            }
            if (node["a:right"] !== undefined) {
                let obj = {
                    "p:spPr": {
                        "a:ln": node["a:right"]["a:ln"]
                    }
                };
                let borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-right");
            }
            if (node["a:left"] !== undefined) {
                let obj = {
                    "p:spPr": {
                        "a:ln": node["a:left"]["a:ln"]
                    }
                };
                let borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-left");
            }

            return borderStyle;
        }
        //////////////////////////////////////////////////////////////////
        function getBorder(node, pNode, isSvgMode, bType, warpObj) {
            //console.log("getBorder", node, pNode, isSvgMode, bType)
            let cssText, lineNode;

            if (bType == "shape") {
                cssText = "border: ";
                lineNode = node["p:spPr"]["a:ln"];
                //subNodeTxt = "p:spPr";
                //node["p:style"]["a:lnRef"] =
            } else if (bType == "text") {
                cssText = "";
                lineNode = node["a:rPr"]["a:ln"];
                //subNodeTxt = "a:rPr";
            }

            //let is_noFill = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:noFill"]);
            let is_noFill = PPTXXmlUtils.getTextByPathList(lineNode, ["a:noFill"]);
            if (is_noFill !== undefined) {
                return "hidden";
            }


            let lnRefNode;
            let phClr = undefined;
            if (lineNode == undefined) {
                lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
                if (lnRefNode !== undefined){
                    let lnIdx = PPTXXmlUtils.getTextByPathList(lnRefNode, ["attrs", "idx"]);
                    // Extract phClr from lnRef to replace placeholder colors in lnStyleLst
                    if (lnRefNode !== undefined) {
                        phClr = getSolidFill(lnRefNode, undefined, undefined, warpObj);
                    }
                    // 检查lnStyleLst的结构
                    const lnStyleLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"];
                    // 处理lnStyleLst可能是对象而不是数组的情况
                    if (Array.isArray(lnStyleLst)) {
                        lineNode = lnStyleLst[Number(lnIdx)];
                    } else {
                        // 如果是对象而不是数组，直接使用
                        lineNode = lnStyleLst;
                    }
                }
            }
            if (lineNode == undefined) {
                //is table
                cssText = "";
                lineNode = node;
            }

            let borderColor;
            let borderWidth = 0;
            let borderType = "solid";
            let strokeDasharray = "0";
            if (lineNode !== undefined) {
                // Border width: 1pt = 12700, default = 0.75pt
                let w = PPTXXmlUtils.getTextByPathList(lineNode, ["attrs", "w"]);
                borderWidth = (w !== undefined) ? parseInt(w) / 12700 : (4/3);
                if (isNaN(borderWidth) || borderWidth < 1) {
                    cssText += (4/3) + "px ";//"1pt ";
                } else {
                    cssText += borderWidth + "px ";// + "pt ";
                }
                // Border type
                borderType = PPTXXmlUtils.getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
                if (borderType === undefined) {
                    borderType = PPTXXmlUtils.getTextByPathList(lineNode, ["attrs", "cmpd"]);
                }
                strokeDasharray = "0";
                switch (borderType) {
                    case "solid":
                        cssText += "solid";
                        strokeDasharray = "0";
                        break;
                    case "dash":
                        cssText += "dashed";
                        strokeDasharray = "5";
                        break;
                    case "dashDot":
                        cssText += "dashed";
                        strokeDasharray = "5, 5, 1, 5";
                        break;
                    case "dot":
                        cssText += "dotted";
                        strokeDasharray = "1, 5";
                        break;
                    case "lgDash":
                        cssText += "dashed";
                        strokeDasharray = "10, 5";
                        break;
                    case "dbl":
                        cssText += "double";
                        strokeDasharray = "0";
                        break;
                    case "lgDashDotDot":
                        cssText += "dashed";
                        strokeDasharray = "10, 5, 1, 5, 1, 5";
                        break;
                    case "sysDash":
                        cssText += "dashed";
                        strokeDasharray = "5, 2";
                        break;
                    case "sysDashDot":
                        cssText += "dashed";
                        strokeDasharray = "5, 2, 1, 5";
                        break;
                    case "sysDashDotDot":
                        cssText += "dashed";
                        strokeDasharray = "5, 2, 1, 5, 1, 5";
                        break;
                    case "sysDot":
                        cssText += "dotted";
                        strokeDasharray = "2, 5";
                        break;
                    case undefined:
                    //console.log(borderType);
                    default:
                        cssText += "solid";
                        strokeDasharray = "0";
                }
                // Border color
                let fillTyp = getFillType(lineNode);

                if (fillTyp === "NO_FILL") {
                    borderColor = isSvgMode ? "none" : "";//"background-color: initial;";
                } else if (fillTyp === "SOLID_FILL") {
                    // 获取lnRef中的颜色作为phClr参数
                    if (!lnRefNode) {
                        lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
                    }
                    // phClr should already be extracted at line 780, but fallback here if not
                    if (phClr === undefined && lnRefNode !== undefined) {
                        phClr = getSolidFill(lnRefNode, undefined, undefined, warpObj);
                    }
                    borderColor = getSolidFill(lineNode["a:solidFill"], undefined, phClr, warpObj);
                } else if (fillTyp === "GRADIENT_FILL") {
                    borderColor = getGradientFill(lineNode["a:gradFill"], warpObj);
    
                } else if (fillTyp === "PATTERN_FILL") {
                    borderColor = getPatternFill(lineNode["a:pattFill"], warpObj);
                }

            }


            // 2. drawingML namespace
            if (borderColor === undefined) {
                //let schemeClrNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef", "a:schemeClr"]);
                // if (schemeClrNode !== undefined) {
                //     let schemeClr = "a:" + PPTXXmlUtils.getTextByPathList(schemeClrNode, ["attrs", "val"]);
                //     let borderColor = getSchemeColorFromTheme(schemeClr, undefined, undefined);
                // }
                let lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);

                if (lnRefNode !== undefined) {
                    borderColor = getSolidFill(lnRefNode, undefined, undefined, warpObj);
                }

                // if (borderColor !== undefined) {
                //     let shade = PPTXXmlUtils.getTextByPathList(schemeClrNode, ["a:shade", "attrs", "val"]);
                //     if (shade !== undefined) {
                //         shade = parseInt(shade) / 10000;
                //         let color = tinycolor("#" + borderColor);
                //         borderColor = color.darken(shade).toHex8();//.replace("#", "");
                //     }
                // }

            }


            if (borderColor === undefined) {
                if (isSvgMode) {
                    borderColor = "none";
                } else {
                    borderColor = "hidden";
                }
            } else {
                // 检查borderColor是否已经是有效的颜色值
                if (borderColor && typeof borderColor === 'string') {
                    // 如果不是以#开头的十六进制颜色，添加#前缀
                    if (!borderColor.startsWith('#') && !borderColor.startsWith('rgb') && !borderColor.startsWith('hsl') && borderColor !== 'none' && borderColor !== 'hidden') {
                        borderColor = "#" + borderColor;
                    }
                }
            }
            cssText += " " + borderColor + " ";

            if (isSvgMode) {
                let result = { "color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray };
                return result;
            } else {
                return cssText + ";";
            }
            // } else {
            //     if (isSvgMode) {
            //         return { "color": 'none', "width": '0', "type": 'none', "strokeDasharray": '0' };
            //     } else {
            //         return "hidden";
            //     }
            // }
        }
        async function getSlideBackgroundFill(warpObj, index) {
            let slideContent = warpObj["slideContent"];
            let slideLayoutContent = warpObj["slideLayoutContent"];
            let slideMasterContent = warpObj["slideMasterContent"];


            //PPTXShapeUtils.getFillType(node)
            let bgPr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgPr"]);
            let bgRef = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgRef"]);

            let bgcolor;
            if (bgPr !== undefined) {
                //bgcolor = "background-color: blue;";
                let bgFillTyp = getFillType(bgPr);

                if (bgFillTyp === "SOLID_FILL") {
                    let sldFill = bgPr["a:solidFill"];
                    let clrMapOvr;
                    let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                    if (sldClrMapOvr !== undefined) {
                        clrMapOvr = sldClrMapOvr;
                    } else {
                        let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                        if (sldClrMapOvr !== undefined) {
                            clrMapOvr = sldClrMapOvr;
                        } else {
                            clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                        }

                    }
                    let sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                    //var sldTint = getColorOpacity(sldFill);

                    //bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                    bgcolor = `background: #${sldBgClr};`;

                } else if (bgFillTyp === "GRADIENT_FILL") {
                    bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
                } else if (bgFillTyp === "PIC_FILL") {

                    bgcolor = await getBgPicFill(bgPr, "slideBg", warpObj, undefined);

                }

            } else if (bgRef !== undefined) {

                let clrMapOvr;
                let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    clrMapOvr = sldClrMapOvr;
                } else {
                    let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                    if (sldClrMapOvr !== undefined) {
                        clrMapOvr = sldClrMapOvr;
                    } else {
                        clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                    }

                }
                let phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);

                // if (bgRef["a:srgbClr"] !== undefined) {
                //     phClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:srgbClr", "attrs", "val"]); //#...
                // } else if (bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
                //     let schemeClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:schemeClr", "attrs", "val"]);
                //     phClr = getSchemeColorFromTheme("a:" + schemeClr, slideMasterContent, undefined); //#...
                // }
                let idx = Number(bgRef["attrs"]["idx"]);


                if (idx == 0 || idx == 1000) ; else if (idx > 0 && idx < 1000) ; else if (idx > 1000) {
                    //bgFillStyleLst  in themeContent
                    //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                    let trueIdx = idx - 1000;
                    // themeContent["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                    let bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                    let sortblAry = [];
                    Object.keys(bgFillLst).forEach(key => {
                        let bgFillLstTyp = bgFillLst[key];
                        if (key != "attrs") {
                            if (bgFillLstTyp.constructor === Array) {
                                for (let i = 0; i < bgFillLstTyp.length; i++) {
                                    let obj = {};
                                    obj[key] = bgFillLstTyp[i];
                                    obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                    obj["attrs"] = {
                                        "order": bgFillLstTyp[i]["attrs"]["order"]
                                    };
                                    sortblAry.push(obj);
                                }
                            } else {
                                let obj = {};
                                obj[key] = bgFillLstTyp;
                                obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                obj["attrs"] = {
                                    "order": bgFillLstTyp["attrs"]["order"]
                                };
                                sortblAry.push(obj);
                            }
                        }
                    });
                    let sortByOrder = sortblAry.slice(0);
                    sortByOrder.sort((a, b) => {
                        return a.idex - b.idex;
                    });
                    let bgFillLstIdx = sortByOrder[trueIdx - 1];
                    let bgFillTyp = getFillType(bgFillLstIdx);
                    if (bgFillTyp === "SOLID_FILL") {
                        let sldFill = bgFillLstIdx["a:solidFill"];
                        let sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                        //var sldTint = getColorOpacity(sldFill);
                        //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                        bgcolor = `background: #${sldBgClr};`;
    
                    } else if (bgFillTyp === "GRADIENT_FILL") {
                        bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
                    } else ;
                }

            }
            else {
                bgPr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr"]);
                bgRef = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgRef"]);

                let clrMapOvr;
                let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    clrMapOvr = sldClrMapOvr;
                } else {
                    clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                }
                if (bgPr !== undefined) {
                    let bgFillTyp = getFillType(bgPr);
                    if (bgFillTyp === "SOLID_FILL") {
                        let sldFill = bgPr["a:solidFill"];

                        let sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                        //var sldTint = getColorOpacity(sldFill);
                        // bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                        bgcolor = `background: #${sldBgClr};`;
                    } else if (bgFillTyp === "GRADIENT_FILL") {
                        bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
                    } else if (bgFillTyp === "PIC_FILL") {
                        bgcolor = await getBgPicFill(bgPr, "slideLayoutBg", warpObj, undefined);

                    }

                } else if (bgRef !== undefined) {
                    //bgcolor = "background: white;";
                    let phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);
                    let idx = Number(bgRef["attrs"]["idx"]);
                    //console.log("phClr=", phClr, "idx=", idx)

                    if (idx == 0 || idx == 1000) ; else if (idx > 0 && idx < 1000) ; else if (idx > 1000) {
                        //bgFillStyleLst  in themeContent
                        //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                        let trueIdx = idx - 1000;
                        let bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                        let sortblAry = [];
                        Object.keys(bgFillLst).forEach(key => {
                            //console.log("cubicBezTo[" + key + "]:");
                            let bgFillLstTyp = bgFillLst[key];
                            if (key != "attrs") {
                                if (bgFillLstTyp.constructor === Array) {
                                    for (let i = 0; i < bgFillLstTyp.length; i++) {
                                        let obj = {};
                                        obj[key] = bgFillLstTyp[i];
                                        obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                        obj["attrs"] = {
                                            "order": bgFillLstTyp[i]["attrs"]["order"]
                                        };
                                        sortblAry.push(obj);
                                    }
                                } else {
                                    let obj = {};
                                    obj[key] = bgFillLstTyp;
                                    obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                    obj["attrs"] = {
                                        "order": bgFillLstTyp["attrs"]["order"]
                                    };
                                    sortblAry.push(obj);
                                }
                            }
                        });
                        let sortByOrder = sortblAry.slice(0);
                        sortByOrder.sort((a, b) => {
                            return a.idex - b.idex;
                        });
                        let bgFillLstIdx = sortByOrder[trueIdx - 1];
                        let bgFillTyp = getFillType(bgFillLstIdx);
                        if (bgFillTyp === "SOLID_FILL") {
                            let sldFill = bgFillLstIdx["a:solidFill"];
                            //console.log("sldFill: ", sldFill)
                            //var sldTint = getColorOpacity(sldFill);
                            //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                            let sldBgClr = getSolidFill(sldFill, clrMapOvr, phClr, warpObj);
                            //console.log("bgcolor: ", bgcolor)
                            bgcolor = `background: #${sldBgClr};`;
                        } else if (bgFillTyp === "GRADIENT_FILL") {
                            //console.log("GRADIENT_FILL: ", bgFillLstIdx, phClr)
                            bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
                        } else if (bgFillTyp === "PIC_FILL") {
                            //theme rels
                            //console.log("PIC_FILL - ", bgFillTyp, bgFillLstIdx, bgFillLst, warpObj);
                            bgcolor = await getBgPicFill(bgFillLstIdx, "themeBg", warpObj, phClr);
                        } else ;
                    }
                } else {
                    bgPr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgPr"]);
                    bgRef = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgRef"]);

                    var clrMap = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);

                    if (bgPr !== undefined) {
                        let bgFillTyp = getFillType(bgPr);
                        if (bgFillTyp === "SOLID_FILL") {
                            let sldFill = bgPr["a:solidFill"];
                            let sldBgClr = getSolidFill(sldFill, clrMap, undefined, warpObj);
                            // var sldTint = getColorOpacity(sldFill);
                            // bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                            bgcolor = `background: #${sldBgClr};`;
                        } else if (bgFillTyp === "GRADIENT_FILL") {
                            bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
                        } else if (bgFillTyp === "PIC_FILL") {
                            bgcolor = await getBgPicFill(bgPr, "slideMasterBg", warpObj, undefined);
                        }
                    } else if (bgRef !== undefined) {
                        //let obj={
                        //    "a:solidFill": bgRef
                        //}
                        let phClr = getSolidFill(bgRef, clrMap, undefined, warpObj);
                        // let phClr;
                        // if (bgRef["a:srgbClr"] !== undefined) {
                        //     phClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:srgbClr", "attrs", "val"]); //#...
                        // } else if (bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
                        //     let schemeClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:schemeClr", "attrs", "val"]);

                        //     phClr = getSchemeColorFromTheme("a:" + schemeClr, slideMasterContent, undefined); //#...
                        // }
                        let idx = Number(bgRef["attrs"]["idx"]);

                        if (idx == 0 || idx == 1000) ; else if (idx > 0 && idx < 1000) ; else if (idx > 1000) {
                            //bgFillStyleLst  in themeContent
                            //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                            let trueIdx = idx - 1000;
                            let bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                            let sortblAry = [];
                            Object.keys(bgFillLst).forEach(key => {
                                //console.log("cubicBezTo[" + key + "]:");
                                let bgFillLstTyp = bgFillLst[key];
                                if (key != "attrs") {
                                    if (bgFillLstTyp.constructor === Array) {
                                        for (let i = 0; i < bgFillLstTyp.length; i++) {
                                            let obj = {};
                                            obj[key] = bgFillLstTyp[i];
                                            obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                            obj["attrs"] = {
                                                "order": bgFillLstTyp[i]["attrs"]["order"]
                                            };
                                            sortblAry.push(obj);
                                        }
                                    } else {
                                        let obj = {};
                                        obj[key] = bgFillLstTyp;
                                        obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                        obj["attrs"] = {
                                            "order": bgFillLstTyp["attrs"]["order"]
                                        };
                                        sortblAry.push(obj);
                                    }
                                }
                            });
                            let sortByOrder = sortblAry.slice(0);
                            sortByOrder.sort((a, b) => {
                                return a.idex - b.idex;
                            });
                            let bgFillLstIdx = sortByOrder[trueIdx - 1];
                            let bgFillTyp = getFillType(bgFillLstIdx);

                            if (bgFillTyp == "SOLID_FILL") {
                                let sldFill = bgFillLstIdx["a:solidFill"];
                                //var sldTint = getColorOpacity(sldFill);
                                //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                                let sldBgClr = getSolidFill(sldFill, clrMap, phClr, warpObj);
                                bgcolor = `background: #${sldBgClr}`;
                            } else if (bgFillTyp == "GRADIENT_FILL") {
                                bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
                            } else if (bgFillTyp == "PIC_FILL") {
                                bgcolor = await getBgPicFill(bgFillLstIdx, "themeBg", warpObj, phClr);
                            } else ;
                        }
                    }
                }
            }

            return bgcolor;
        }
        function getBgGradientFill(bgPr, phClr, slideMasterContent, warpObj) {
            let bgcolor = "";
            if (bgPr !== undefined) {
                let grdFill = bgPr["a:gradFill"];
                let gsLst = grdFill["a:gsLst"]["a:gs"];
                //var startColorNode, endColorNode;
                let color_ary = [];
                var pos_ary = [];
                //let tint_ary = [];
                for (let i = 0; i < gsLst.length; i++) {
                    let lo_color = getSolidFill(gsLst[i], slideMasterContent["p:sldMaster"]["p:clrMap"]["attrs"], phClr, warpObj);
                    var pos = PPTXXmlUtils.getTextByPathList(gsLst[i], ["attrs", "pos"]);
                    if (pos !== undefined) {
                        pos_ary[i] = pos / 1000 + "%";
                    } else {
                        pos_ary[i] = "";
                    }
                    color_ary[i] = "#" + lo_color;
                    //tint_ary[i] = (lo_tint !== undefined) ? parseInt(lo_tint) / 100000 : 1;
                }
                //get rot
                let lin = grdFill["a:lin"];
                let rot = 90;
                if (lin !== undefined) {
                    rot = PPTXXmlUtils.angleToDegrees(lin["attrs"]["ang"]);
                    rot = rot + 90;
                }
                bgcolor = `background: linear-gradient(${rot}deg,`;
                for (let i = 0; i < gsLst.length; i++) {
                    if (i == gsLst.length - 1) {
                        //if (phClr === undefined) {
                        //bgcolor += "rgba(" + hexToRgbNew(color_ary[i]) + "," + tint_ary[i] + ")" + ");";
                        bgcolor += color_ary[i] + " " + pos_ary[i] + ");";
                        //} else {
                        //bgcolor += "rgba(" + hexToRgbNew(phClr) + "," + tint_ary[i] + ")" + ");";
                        // bgcolor += "" + phClr + ";";;
                        //}
                    } else {
                        //if (phClr === undefined) {
                        //bgcolor += "rgba(" + hexToRgbNew(color_ary[i]) + "," + tint_ary[i] + ")" + ", ";
                        bgcolor += color_ary[i] + " " + pos_ary[i] + ", ";                        //} else {
                        //bgcolor += "rgba(" + hexToRgbNew(phClr) + "," + tint_ary[i] + ")" + ", ";
                        // bgcolor += phClr + ", ";
                        //}
                    }
                }
            } else {
                if (phClr !== undefined) {
                    //bgcolor = "rgba(" + hexToRgbNew(phClr) + ",0);";
                    //bgcolor = phClr + ");";
                    bgcolor = "background: #" + phClr + ";";
                }
            }
            return bgcolor;
        }
        async function getBgPicFill(bgPr, sorce, warpObj, phClr, index) {
            let bgcolor;
            let picFillResult = await getPicFill(sorce, bgPr["a:blipFill"], warpObj);
            let picFillBase64 = picFillResult;
            if (typeof picFillResult === 'object' && picFillResult.img) {
                picFillBase64 = picFillResult.img;
            }
            let ordr = bgPr["attrs"]["order"];
            let aBlipNode = bgPr["a:blipFill"]["a:blip"];
            //a:duotone
            let duotone = PPTXXmlUtils.getTextByPathList(aBlipNode, ["a:duotone"]);
            if (duotone !== undefined) {
                let clr_ary = [];
                // duotone.forEach(clr => {
                //     console.log("pic duotone clr: ", clr)
                // }) 
                Object.keys(duotone).forEach(clr_type => {

                    if (clr_type != "attrs") {
                        let obj = {};
                        obj[clr_type] = duotone[clr_type];
                        clr_ary.push(getSolidFill(obj, undefined, phClr, warpObj));
                    }
                    // Object.keys(duotone[clr_type]).forEach(clr => {
                    //     if (clr != "order") {
                    //         let obj = {};
                    //         obj[clr_type] = duotone[clr_type][clr];
                    //         clr_ary.push(getSolidFill(obj, undefined, phClr, warpObj));
                    //     }
                    // })
                });

                //filter: url(file.svg#filter-element-id)
                //https://codepen.io/bhenbe/pen/QEZOvd
                //https://www.w3schools.com/cssref/css3_pr_filter.asp

                // let color1 = clr_ary[0];
                // let color2 = clr_ary[1];
                // let cssName = "";

                // let styleText_before_after = "content: '';" +
                //     "display: block;" +
                //     "width: 100%;" +
                //     "height: 100%;" +
                //     // "z-index: 1;" +
                //     "position: absolute;" +
                //     "top: 0;" +
                //     "left: 0;";

                // let cssName = "slide-background-" + index + "::before," + " .slide-background-" + index + "::after";
                // styleTable[styleText_before_after] = {
                //     "name": cssName,
                //     "text": styleText_before_after
                // };


                // let styleText_after = "background-color: #" + clr_ary[1] + ";" +
                //     "mix-blend-mode: darken;";

                // cssName = "slide-background-" + index + "::after";
                // styleTable[styleText_after] = {
                //     "name": cssName,
                //     "text": styleText_after
                // };

                // let styleText_before = "background-color: #" + clr_ary[0] + ";" +
                //     "mix-blend-mode: lighten;";

                // cssName = "slide-background-" + index + "::before";
                // styleTable[styleText_before] = {
                //     "name": cssName,
                //     "text": styleText_before
                // };

            }
            //a:alphaModFix
            let aphaModFixNode = PPTXXmlUtils.getTextByPathList(aBlipNode, ["a:alphaModFix", "attrs"]);
            let imgOpacity = "";
            if (aphaModFixNode !== undefined && aphaModFixNode["amt"] !== undefined && aphaModFixNode["amt"] != "") {
                var amt = parseInt(aphaModFixNode["amt"]) / 100000;
                //let opacity = amt;
                imgOpacity = "opacity:" + amt + ";";

            }
            // 使用getPicFill函数返回的填充模式信息
            let prop_style = "";
            if (typeof picFillResult === 'object') {
                if (picFillResult.backgroundSize) {
                    prop_style += "background-size: " + picFillResult.backgroundSize + ";";
                }
                if (picFillResult.backgroundPosition) {
                    prop_style += "background-position: " + picFillResult.backgroundPosition + ";";
                }
                if (picFillResult.backgroundRepeat) {
                    prop_style += "background-repeat: " + picFillResult.backgroundRepeat + ";";
                }
            }
            bgcolor = "background: url(" + picFillBase64 + ");  z-index: " + ordr + ";" + prop_style + imgOpacity;

            return bgcolor;
        }
      
        
        function getGradientFill(node, warpObj) {
            //console.log("getGradientFill: node", node)
            let gsLst = node["a:gsLst"]["a:gs"];
            //get start color
            let color_ary = [];
            for (let i = 0; i < gsLst.length; i++) {
                let lo_color = getSolidFill(gsLst[i], undefined, undefined, warpObj);
                color_ary[i] = lo_color;
            }
            //get rot
            let lin = node["a:lin"];
            let rot = 0;
            if (lin !== undefined) {
                rot = PPTXXmlUtils.angleToDegrees(lin["attrs"]["ang"]) + 90;
            }
            return {
                "color": color_ary,
                "rot": rot
            }
        }
        async function getPicFill(type, node, warpObj) {
            //Need to test/////////////////////////////////////////////
            //rId
            let img;
            let rId = node["a:blip"]["attrs"]["r:embed"];
            let imgPath;
            if (type == "slideBg" || type == "slide") {
                imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["slideResObj", rId, "target"]);
            } else if (type == "slideLayoutBg") {
                imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["layoutResObj", rId, "target"]);
            } else if (type == "slideMasterBg") {
                imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["masterResObj", rId, "target"]);
            } else if (type == "themeBg") {
                imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["themeResObj", rId, "target"]);
            } else if (type == "diagramBg") {
                imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["diagramResObj", rId, "target"]);
            }
            if (imgPath === undefined) {
                return undefined;
            }
            img = PPTXXmlUtils.getTextByPathList(warpObj, ["loaded-images", imgPath]); //, type, rId
            if (img === undefined) {
                 // 确定上下文类型用于路径解析
                var context = 'slide';
                if (type == "slideMasterBg") {
                    context = 'master';
                } else if (type == "slideLayoutBg") {
                    context = 'layout';
                }
                imgPath = PPTXXmlUtils.resolveMediaPath(imgPath, context, '');

                let imgExt = imgPath.split(".").pop();
                if (imgExt == "xml") {
                    return undefined;
                }
                let imgFile = warpObj["zip"].file(imgPath);
                if (imgFile === null || imgFile === undefined) {
                    return undefined;
                }
                let imgArrayBuffer = await imgFile.async("arraybuffer");
                let imgMimeType = PPTXXmlUtils.getMimeType(imgExt);
                img = "data:" + imgMimeType + ";base64," + PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer);
                //warpObj["loaded-images"][imgPath] = img; //"defaultTextStyle": defaultTextStyle,
                setTextByPathList(warpObj, ["loaded-images", imgPath], img); //, type, rId
            }
            // 处理图像属性 - Tile, Stretch, or Display Portion of Image
            let tileNode = node["a:tile"];
            let stretchNode = node["a:stretch"];
            let fillMode = "stretch";
            let backgroundSize = "cover";
            let backgroundPosition = "center";
            let backgroundRepeat = "no-repeat";
            
            if (tileNode) {
                // 平铺模式
                fillMode = "tile";
                backgroundRepeat = "repeat";
                
                // 处理平铺大小
                let sx = tileNode["attrs"]["sx"];
                let sy = tileNode["attrs"]["sy"];
                if (sx && sy) {
                    let widthPercent = parseInt(sx) / 100000 * 100;
                    let heightPercent = parseInt(sy) / 100000 * 100;
                    backgroundSize = widthPercent + "% " + heightPercent + "%";
                }
                
                // 处理平铺偏移
                let tx = tileNode["attrs"]["tx"];
                let ty = tileNode["attrs"]["ty"];
                if (tx && ty) {
                    let xPercent = parseInt(tx) / 100000 * 100;
                    let yPercent = parseInt(ty) / 100000 * 100;
                    backgroundPosition = xPercent + "% " + yPercent + "%";
                }
            } else if (stretchNode) {
                // 拉伸模式
                fillMode = "stretch";
                let fillRect = stretchNode["a:fillRect"];
                if (fillRect) {
                    // 处理填充矩形
                    backgroundSize = "cover";
                }
            }
            
            // 返回包含图像和填充模式的对象
            return {
                "img": img,
                "fillMode": fillMode,
                "backgroundSize": backgroundSize,
                "backgroundPosition": backgroundPosition,
                "backgroundRepeat": backgroundRepeat
            };
        }
        function getPatternFill(node, warpObj) {
            //https://developer.mozilla.org/en-US/docs/Web/CSS/CSS_Images/Using_CSS_gradients
            //https://cssgradient.io/blog/css-gradient-text/
            //https://css-tricks.com/background-patterns-simplified-by-conic-gradients/
            //https://stackoverflow.com/questions/6705250/how-to-get-a-pattern-into-a-written-text-via-css
            //https://stackoverflow.com/questions/14072142/striped-text-in-css
            //https://css-tricks.com/stripes-css/
            //https://yuanchuan.dev/gradient-shapes/
            let fgColor = "", bgColor = "", prst = "";
            let bgClr = node["a:bgClr"];
            let fgClr = node["a:fgClr"];
            prst = node["attrs"]["prst"];
            fgColor = getSolidFill(fgClr, undefined, undefined, warpObj);
            bgColor = getSolidFill(bgClr, undefined, undefined, warpObj);
            //var angl_ary = getAnglefromParst(prst);
            //let ptrClr = "repeating-linear-gradient(" + angl + "deg,  #" + bgColor + ",#" + fgColor + " 2px);"
            //linear-gradient(0deg, black 10 %, transparent 10 %, transparent 90 %, black 90 %, black), 
            //linear-gradient(90deg, black 10 %, transparent 10 %, transparent 90 %, black 90 %, black);
            let linear_gradient = getLinerGrandient(prst, bgColor, fgColor);
            return linear_gradient;
        }

        function getLinerGrandient(prst, bgColor, fgColor) {
            // dashDnDiag (Dashed Downward Diagonal)-V
            // dashHorz (Dashed Horizontal)-V
            // dashUpDiag(Dashed Upward DIagonal)-V
            // dashVert(Dashed Vertical)-V
            // diagBrick(Diagonal Brick)-V
            // divot(Divot)-VX
            // dkDnDiag(Dark Downward Diagonal)-V
            // dkHorz(Dark Horizontal)-V
            // dkUpDiag(Dark Upward Diagonal)-V
            // dkVert(Dark Vertical)-V
            // dotDmnd(Dotted Diamond)-VX
            // dotGrid(Dotted Grid)-V
            // horzBrick(Horizontal Brick)-V
            // lgCheck(Large Checker Board)-V
            // lgConfetti(Large Confetti)-V
            // lgGrid(Large Grid)-V
            // ltDnDiag(Light Downward Diagonal)-V
            // ltHorz(Light Horizontal)-V
            // ltUpDiag(Light Upward Diagonal)-V
            // ltVert(Light Vertical)-V
            // narHorz(Narrow Horizontal)-V
            // narVert(Narrow Vertical)-V
            // openDmnd(Open Diamond)-V
            // pct10(10 %)-V
            // pct20(20 %)-V
            // pct25(25 %)-V
            // pct30(30 %)-V
            // pct40(40 %)-V
            // pct5(5 %)-V
            // pct50(50 %)-V
            // pct60(60 %)-V
            // pct70(70 %)-V
            // pct75(75 %)-V
            // pct80(80 %)-V
            // pct90(90 %)-V
            // smCheck(Small Checker Board) -V
            // smConfetti(Small Confetti)-V
            // smGrid(Small Grid) -V
            // solidDmnd(Solid Diamond)-V
            // sphere(Sphere)-V
            // trellis(Trellis)-VX
            // wave(Wave)-V
            // wdDnDiag(Wide Downward Diagonal)-V
            // wdUpDiag(Wide Upward Diagonal)-V
            // weave(Weave)-V
            // zigZag(Zig Zag)-V
            // shingle(Shingle)-V
            // plaid(Plaid)-V
            // cross (Cross)
            // diagCross(Diagonal Cross)
            // dnDiag(Downward Diagonal)
            // horz(Horizontal)
            // upDiag(Upward Diagonal)
            // vert(Vertical)
            switch (prst) {
                case "smGrid":
                    return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1px ), " +
                        "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1px)  #" + bgColor + ";", "4px 4px"];
                case "dotGrid":
                    return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1px ), " +
                        "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1px)  #" + bgColor + ";", "8px 8px"];
                case "lgGrid":
                    return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1.5px ), " +
                        "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1.5px)  #" + bgColor + ";", "8px 8px"];
                case "wdUpDiag":
                    //return ["repeating-linear-gradient(-45deg,  #" + bgColor + ", #" + bgColor + " 1px,#" + fgColor + " 5px);"];
                    return ["repeating-linear-gradient(-45deg, transparent 1px , transparent 4px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
                case "dkUpDiag":
                    return ["repeating-linear-gradient(-45deg, transparent 1px , #" + bgColor + " 5px)" + "#" + fgColor + ";"];
                case "ltUpDiag":
                    return ["repeating-linear-gradient(-45deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
                case "wdDnDiag":
                    return ["repeating-linear-gradient(45deg, transparent 1px , transparent 4px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
                case "dkDnDiag":
                    return ["repeating-linear-gradient(45deg, transparent 1px , #" + bgColor + " 5px)" + "#" + fgColor + ";"];
                case "ltDnDiag":
                    return ["repeating-linear-gradient(45deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
                case "dkHorz":
                    return ["repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #" + bgColor + " 7px)" + "#" + fgColor + ";"];
                case "ltHorz":
                    return ["repeating-linear-gradient(0deg, transparent 1px , transparent 5px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
                case "narHorz":
                    return ["repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
                case "dkVert":
                    return ["repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #" + bgColor + " 7px)" + "#" + fgColor + ";"];
                case "ltVert":
                    return ["repeating-linear-gradient(90deg, transparent 1px , transparent 5px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
                case "narVert":
                    return ["repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
                case "lgCheck":
                case "smCheck":
                    var size = "";
                    var pos = "";
                    if (prst == "lgCheck") {
                        size = "8px 8px";
                        pos = "0 0, 4px 4px, 4px 4px, 8px 8px";
                    } else {
                        size = "4px 4px";
                        pos = "0 0, 2px 2px, 2px 2px, 4px 4px";
                    }
                    return ["linear-gradient(45deg,  #" + fgColor + " 25%, transparent 0, transparent 75%,  #" + fgColor + " 0), " +
                        "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 0, transparent 75%,  #" + fgColor + " 0) " +
                        "#" + bgColor + ";", size, pos];
                // case "smCheck":
                //     return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 0.5px),  #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%), " +
                //         "linear-gradient(-45deg, transparent 0%, transparent calc(50% - 0.5px) , #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%)  " +
                //         "#" + bgColor + ";", "4px 4px"];
                //     break 

                case "dashUpDiag":
                    return ["repeating-linear-gradient(152deg, #" + fgColor + ", #" + fgColor + " 5% , transparent 0, transparent 70%)" +
                        "#" + bgColor + ";", "4px 4px"];
                case "dashDnDiag":
                    return ["repeating-linear-gradient(45deg, #" + fgColor + ", #" + fgColor + " 5% , transparent 0, transparent 70%)" +
                        "#" + bgColor + ";", "4px 4px"];
                case "diagBrick":
                    return ["linear-gradient(45deg, transparent 15%,  #" + fgColor + " 30%, transparent 30%), " +
                        "linear-gradient(-45deg, transparent 15%,  #" + fgColor + " 30%, transparent 30%), " +
                        "linear-gradient(-45deg, transparent 65%,  #" + fgColor + " 80%, transparent 0) " +
                        "#" + bgColor + ";", "4px 4px"];
                case "horzBrick":
                    return ["linear-gradient(335deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                        "linear-gradient(155deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                        "linear-gradient(335deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                        "linear-gradient(155deg, #" + bgColor + " 1.6px, transparent 1.6px) " +
                        "#" + fgColor + ";", "4px 4px", "0 0.15px, 0.3px 2.5px, 2px 2.15px, 2.35px 0.4px"];

                case "dashVert":
                    return ["linear-gradient(0deg,  #" + bgColor + " 30%, transparent 30%)," +
                        "linear-gradient(90deg,transparent, transparent 40%, #" + fgColor + " 40%, #" + fgColor + " 60% , transparent 60%)" +
                        "#" + bgColor + ";", "4px 4px"];
                case "dashHorz":
                    return ["linear-gradient(90deg,  #" + bgColor + " 30%, transparent 30%)," +
                        "linear-gradient(0deg,transparent, transparent 40%, #" + fgColor + " 40%, #" + fgColor + " 60% , transparent 60%)" +
                        "#" + bgColor + ";", "4px 4px"];
                case "solidDmnd":
                    return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%), " +
                        "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%), " +
                        "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%), " +
                        "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                        "#" + bgColor + ";", "8px 8px"];
                case "openDmnd":
                    return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 0.5px),  #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%), " +
                        "linear-gradient(-45deg, transparent 0%, transparent calc(50% - 0.5px) , #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%) " +
                        "#" + bgColor + ";", "8px 8px"];

                case "dotDmnd":
                    return ["radial-gradient(#" + fgColor + " 15%, transparent 0), " +
                        "radial-gradient(#" + fgColor + " 15%, transparent 0) " +
                        "#" + bgColor + ";", "4px 4px", "0 0, 2px 2px"];
                case "zigZag":
                case "wave":
                    var size = "";
                    if (prst == "zigZag") size = "0";
                    else size = "1px";
                    return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%) 50px " + size + ", " +
                        "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%) 50px " + size + ", " +
                        "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%), " +
                        "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                        "#" + bgColor + ";", "4px 4px"];
                case "lgConfetti":
                case "smConfetti":
                    var size = "";
                    if (prst == "lgConfetti") size = "4px 4px";
                    else size = "2px 2px";
                    return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%) 50px 1px, " +
                        "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%), " +
                        "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%) 50px 1px , " +
                        "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                        "#" + bgColor + ";", size];
                // case "weave":
                //     return ["linear-gradient(45deg,  #" + bgColor + " 5%, transparent 25%) 50px 0, " +
                //         "linear-gradient(135deg,  #" + bgColor + " 25%, transparent 25%) 50px 0, " +
                //         "linear-gradient(45deg,  #" + bgColor + " 25%, transparent 25%) " +
                //         "#" + fgColor + ";", "4px 4px"];
                //     //background: linear-gradient(45deg, #dca 12%, transparent 0, transparent 88%, #dca 0),
                //     //linear-gradient(135deg, transparent 37 %, #a85 0, #a85 63 %, transparent 0),
                //     //linear-gradient(45deg, transparent 37 %, #dca 0, #dca 63 %, transparent 0) #753;
                //     // background-size: 25px 25px;
                //     break;

                case "plaid":
                    return ["linear-gradient(0deg, transparent, transparent 25%, #" + fgColor + "33 25%, #" + fgColor + "33 50%)," +
                        "linear-gradient(90deg, transparent, transparent 25%, #" + fgColor + "66 25%, #" + fgColor + "66 50%) " +
                        "#" + bgColor + ";", "4px 4px"];
                case "sphere":
                    return ["radial-gradient(#" + fgColor + " 50%, transparent 50%)," +
                        "#" + bgColor + ";", "4px 4px"];
                case "weave":
                case "shingle":
                    return ["linear-gradient(45deg, #" + bgColor + " 1.31px , #" + fgColor + " 1.4px, #" + fgColor + " 1.5px, transparent 1.5px, transparent 4.2px, #" + fgColor + " 4.2px, #" + fgColor + " 4.3px, transparent 4.31px), " +
                        "linear-gradient(-45deg,  #" + bgColor + " 1.31px , #" + fgColor + " 1.4px, #" + fgColor + " 1.5px, transparent 1.5px, transparent 4.2px, #" + fgColor + " 4.2px, #" + fgColor + " 4.3px, transparent 4.31px) 0 4px, " +
                        "#" + bgColor + ";", "4px 8px"];
                //background:
                //linear-gradient(45deg, #708090 1.31px, #d9ecff 1.4px, #d9ecff 1.5px, transparent 1.5px, transparent 4.2px, #d9ecff 4.2px, #d9ecff 4.3px, transparent 4.31px),
                //linear-gradient(-45deg, #708090 1.31px, #d9ecff 1.4px, #d9ecff 1.5px, transparent 1.5px, transparent 4.2px, #d9ecff 4.2px, #d9ecff 4.3px, transparent 4.31px)0 4px;
                //background-color:#708090;
                //background-size: 4px 8px;
                case "pct5":
                case "pct10":
                case "pct20":
                case "pct25":
                case "pct30":
                case "pct40":
                case "pct50":
                case "pct60":
                case "pct70":
                case "pct75":
                case "pct80":
                case "pct90":
                //case "dotDmnd":
                case "trellis":
                case "divot":
                    var px_pr_ary;
                    switch (prst) {
                        case "pct5":
                            px_pr_ary = ["0.3px", "10%", "2px 2px"];
                            break
                        case "divot":
                            px_pr_ary = ["0.3px", "40%", "4px 4px"];
                            break
                        case "pct10":
                            px_pr_ary = ["0.3px", "20%", "2px 2px"];
                            break
                        case "pct20":
                            //case "dotDmnd":
                            px_pr_ary = ["0.2px", "40%", "2px 2px"];
                            break
                        case "pct25":
                            px_pr_ary = ["0.2px", "50%", "2px 2px"];
                            break
                        case "pct30":
                            px_pr_ary = ["0.5px", "50%", "2px 2px"];
                            break
                        case "pct40":
                            px_pr_ary = ["0.5px", "70%", "2px 2px"];
                            break
                        case "pct50":
                            px_pr_ary = ["0.09px", "90%", "2px 2px"];
                            break
                        case "pct60":
                            px_pr_ary = ["0.3px", "90%", "2px 2px"];
                            break
                        case "pct70":
                        case "trellis":
                            px_pr_ary = ["0.5px", "95%", "2px 2px"];
                            break
                        case "pct75":
                            px_pr_ary = ["0.65px", "100%", "2px 2px"];
                            break
                        case "pct80":
                            px_pr_ary = ["0.85px", "100%", "2px 2px"];
                            break
                        case "pct90":
                            px_pr_ary = ["1px", "100%", "2px 2px"];
                            break
                    }
                    return ["radial-gradient(#" + fgColor + " " + px_pr_ary[0] + ", transparent " + px_pr_ary[1] + ")," +
                        "#" + bgColor + ";", px_pr_ary[2]];
                default:
                    return [0, 0];
            }
        }

        function getSolidFill(node, clrMap, phClr, warpObj) {

            if (node === undefined) {
                return undefined;
            }

            let color = "";
            let clrNode;
            if (node["a:srgbClr"] !== undefined) {
                clrNode = node["a:srgbClr"];
                color = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]); //#...
            } else if (node["a:schemeClr"] !== undefined) { //a:schemeClr
                clrNode = node["a:schemeClr"];
                let schemeClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]);
                color = getSchemeColorFromTheme("a:" + schemeClr, clrMap, phClr, warpObj);
            } else if (node["a:scrgbClr"] !== undefined) {
                clrNode = node["a:scrgbClr"];
                //<a:scrgbClr r="50%" g="50%" b="50%"/>  //Need to test/////////////////////////////////////////////
                let defBultColorVals = clrNode["attrs"];
                let red = (defBultColorVals["r"].indexOf("%") != -1) ? defBultColorVals["r"].split("%").shift() : defBultColorVals["r"];
                let green = (defBultColorVals["g"].indexOf("%") != -1) ? defBultColorVals["g"].split("%").shift() : defBultColorVals["g"];
                let blue = (defBultColorVals["b"].indexOf("%") != -1) ? defBultColorVals["b"].split("%").shift() : defBultColorVals["b"];
                //let scrgbClr = red + "," + green + "," + blue;
                color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100));
                //console.log("scrgbClr: " + scrgbClr);

            } else if (node["a:prstClr"] !== undefined) {
                clrNode = node["a:prstClr"];
                //<a:prstClr val="black"/>  //Need to test/////////////////////////////////////////////
                let prstClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]); //node["a:prstClr"]["attrs"]["val"];
                color = getColorName2Hex(prstClr);
                //console.log("blip prstClr: ", prstClr, " => hexClr: ", color);
            } else if (node["a:hslClr"] !== undefined) {
                clrNode = node["a:hslClr"];
                //<a:hslClr hue="14400000" sat="100%" lum="50%"/>  //Need to test/////////////////////////////////////////////
                let defBultColorVals = clrNode["attrs"];
                let hue = Number(defBultColorVals["hue"]) / 100000;
                let sat = Number((defBultColorVals["sat"].indexOf("%") != -1) ? defBultColorVals["sat"].split("%").shift() : defBultColorVals["sat"]) / 100;
                let lum = Number((defBultColorVals["lum"].indexOf("%") != -1) ? defBultColorVals["lum"].split("%").shift() : defBultColorVals["lum"]) / 100;
                //let hslClr = defBultColorVals["hue"] + "," + defBultColorVals["sat"] + "," + defBultColorVals["lum"];
                let hsl2rgb = hslToRgb(hue, sat, lum);
                color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b);
                //defBultColor = cnvrtHslColor2Hex(hslClr); //TODO
                // console.log("hslClr: " + hslClr);
            } else if (node["a:sysClr"] !== undefined) {
                clrNode = node["a:sysClr"];
                //<a:sysClr val="windowText" lastClr="000000"/>  //Need to test/////////////////////////////////////////////
                let sysClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "lastClr"]);
                if (sysClr !== undefined) {
                    color = sysClr;
                }
            }
            //console.log("color: [%cstart]", "color: #" + color, tinycolor(color).toHslString(), color)

            //fix color -------------------------------------------------------- TODO 
            //
            //1. "alpha":
            //Specifies the opacity as expressed by a percentage value.
            // [Example: The following represents a green solid fill which is 50 % opaque
            // < a: solidFill >
            //     <a:srgbClr val="00FF00">
            //         <a:alpha val="50%" />
            //     </a:srgbClr>
            // </a: solidFill >
            let isAlpha = false;
            let alpha = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:alpha", "attrs", "val"])) / 100000;
            //console.log("alpha: ", alpha)
            if (!isNaN(alpha)) {
                // var al_color = new colz.Color(color);
                // al_color.setAlpha(alpha);
                // let ne_color = al_color.rgba.toString();
                // color = (rgba2hex(ne_color))
                let al_color = tinycolor$1(color);
                al_color.setAlpha(alpha);
                color = al_color.toHex8();
                isAlpha = true;
            }
            //2. "alphaMod":
            // Specifies the opacity as expressed by a percentage relative to the input color.
            //     [Example: The following represents a green solid fill which is 50 % opaque
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:alphaMod val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //3. "alphaOff":
            // Specifies the opacity as expressed by a percentage offset increase or decrease to the
            // input color.Increases never increase the opacity beyond 100 %, decreases never decrease
            // the opacity below 0 %.
            // [Example: The following represents a green solid fill which is 90 % opaque
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:alphaOff val="-10%" />
            //         </a:srgbClr>
            //     </a: solidFill >

            //4. "blue":
            //Specifies the value of the blue component.The assigned value is specified as a
            //percentage with 0 % indicating minimal blue and 100 % indicating maximum blue.
            //  [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //      to value RRGGBB = (00, FF, FF)
            //          <a: solidFill >
            //              <a:srgbClr val="00FF00">
            //                  <a:blue val="100%" />
            //              </a:srgbClr>
            //          </a: solidFill >
            //5. "blueMod"
            // Specifies the blue component as expressed by a percentage relative to the input color
            // component.Increases never increase the blue component beyond 100 %, decreases
            // never decrease the blue component below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
            //     to value RRGGBB = (00, 00, 80)
            //     < a: solidFill >
            //         <a:srgbClr val="0000FF">
            //             <a:blueMod val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //6. "blueOff"
            // Specifies the blue component as expressed by a percentage offset increase or decrease
            // to the input color component.Increases never increase the blue component
            // beyond 100 %, decreases never decrease the blue component below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
            // to value RRGGBB = (00, 00, CC)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:blueOff val="-20%" />
            //         </a:srgbClr>
            //     </a: solidFill >

            //7. "comp" - This element specifies that the color rendered should be the complement of its input color with the complement
            // being defined as such.Two colors are called complementary if, when mixed they produce a shade of grey.For
            // instance, the complement of red which is RGB(255, 0, 0) is cyan.(<a:comp/>)

            //8. "gamma" - This element specifies that the output color rendered by the generating application should be the sRGB gamma
            //              shift of the input color.

            //9. "gray" - This element specifies a grayscale of its input color, taking into relative intensities of the red, green, and blue
            //              primaries.

            //10. "green":
            // Specifies the value of the green component. The assigned value is specified as a
            // percentage with 0 % indicating minimal green and 100 % indicating maximum green.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
            // to value RRGGBB = (00, FF, FF)
            //     < a: solidFill >
            //         <a:srgbClr val="0000FF">
            //             <a:green val="100%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //11. "greenMod":
            // Specifies the green component as expressed by a percentage relative to the input color
            // component.Increases never increase the green component beyond 100 %, decreases
            // never decrease the green component below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            // to value RRGGBB = (00, 80, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:greenMod val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //12. "greenOff":
            // Specifies the green component as expressed by a percentage offset increase or decrease
            // to the input color component.Increases never increase the green component
            // beyond 100 %, decreases never decrease the green component below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            // to value RRGGBB = (00, CC, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:greenOff val="-20%" />
            //         </a:srgbClr>
            //     </a: solidFill >

            //13. "hue" (This element specifies a color using the HSL color model):
            // This element specifies the input color with the specified hue, but with its saturation and luminance unchanged.
            // < a: solidFill >
            //     <a:hslClr hue="14400000" sat="100%" lum="50%">
            // </a:solidFill>
            // <a:solidFill>
            //     <a:hslClr hue="0" sat="100%" lum="50%">
            //         <a:hue val="14400000"/>
            //     <a:hslClr/>
            // </a:solidFill>

            //14. "hueMod" (This element specifies a color using the HSL color model):
            // Specifies the hue as expressed by a percentage relative to the input color.
            // [Example: The following manipulates the fill color from having RGB value RRGGBB = (00, FF, 00) to value RRGGBB = (FF, FF, 00)
            //         < a: solidFill >
            //             <a:srgbClr val="00FF00">
            //                 <a:hueMod val="50%" />
            //             </a:srgbClr>
            //         </a: solidFill >

            let hueMod = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:hueMod", "attrs", "val"])) / 100000;
            if (!isNaN(hueMod)) {
                color = applyHueMod(color, hueMod, isAlpha);
            }
            //15. "hueOff"(This element specifies a color using the HSL color model):
            // Specifies the actual angular value of the shift.The result of the shift shall be between 0
            // and 360 degrees.Shifts resulting in angular values less than 0 are treated as 0. Shifts
            // resulting in angular values greater than 360 are treated as 360.
            // [Example:
            //     The following increases the hue angular value by 10 degrees.
            //     < a: solidFill >
            //         <a:hslClr hue="0" sat="100%" lum="50%"/>
            //             <a:hueOff val="600000"/>
            //     </a: solidFill >
            //let hueOff = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:hueOff", "attrs", "val"])) / 100000;
            // if (!isNaN(hueOff)) {
            //     //console.log("hueOff: ", hueOff, " (TODO)")
            //     //color = applyHueOff(color, hueOff, isAlpha);
            // }

            //16. "inv" (inverse)
            //This element specifies the inverse of its input color.
            //The inverse of red (1, 0, 0) is cyan (0, 1, 1 ).
            // The following represents cyan, the inverse of red:
            // <a:solidFill>
            //     <a:srgbClr val="FF0000">
            //         <a:inv />
            //     </a:srgbClr>
            // </a:solidFill>

            //17. "invGamma" - This element specifies that the output color rendered by the generating application should be the inverse sRGB
            //                  gamma shift of the input color.

            //18. "lum":
            // This element specifies the input color with the specified luminance, but with its hue and saturation unchanged.
            // Typically luminance values fall in the range[0 %, 100 %].
            // The following two solid fills are equivalent:
            // <a:solidFill>
            //     <a:hslClr hue="14400000" sat="100%" lum="50%">
            // </a:solidFill>
            // <a:solidFill>
            //     <a:hslClr hue="14400000" sat="100%" lum="0%">
            //         <a:lum val="50%" />
            //     <a:hslClr />
            // </a:solidFill>
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            // to value RRGGBB = (00, 66, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:lum val="20%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            // end example]
            //19. "lumMod":
            // Specifies the luminance as expressed by a percentage relative to the input color.
            // Increases never increase the luminance beyond 100 %, decreases never decrease the
            // luminance below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (00, 75, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:lumMod val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            // end example]
            let lumMod = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:lumMod", "attrs", "val"])) / 100000;
            if (!isNaN(lumMod)) {
                color = applyLumMod(color, lumMod, isAlpha);
            }
            //let lumMod_color = applyLumMod(color, 0.5);
            //console.log("lumMod_color: ", lumMod_color)
            //20. "lumOff"
            // Specifies the luminance as expressed by a percentage offset increase or decrease to the
            // input color.Increases never increase the luminance beyond 100 %, decreases never
            // decrease the luminance below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (00, 99, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:lumOff val="-20%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            let lumOff = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:lumOff", "attrs", "val"])) / 100000;
            if (!isNaN(lumOff)) {
                color = applyLumOff(color, lumOff, isAlpha);
            }


            //21. "red":
            // Specifies the value of the red component.The assigned value is specified as a percentage
            // with 0 % indicating minimal red and 100 % indicating maximum red.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (FF, FF, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:red val="100%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //22. "redMod":
            // Specifies the red component as expressed by a percentage relative to the input color
            // component.Increases never increase the red component beyond 100 %, decreases never
            // decrease the red component below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
            //     to value RRGGBB = (80, 00, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="FF0000">
            //             <a:redMod val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //23. "redOff":
            // Specifies the red component as expressed by a percentage offset increase or decrease to
            // the input color component.Increases never increase the red component beyond 100 %,
            //     decreases never decrease the red component below 0 %.
            //     [Example: The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
            //     to value RRGGBB = (CC, 00, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="FF0000">
            //             <a:redOff val="-20%" />
            //         </a:srgbClr>
            //     </a: solidFill >

            //23. "sat":
            // This element specifies the input color with the specified saturation, but with its hue and luminance unchanged.
            // Typically saturation values fall in the range[0 %, 100 %].
            // [Example:
            //     The following two solid fills are equivalent:
            //     <a:solidFill>
            //         <a:hslClr hue="14400000" sat="100%" lum="50%">
            //     </a:solidFill>
            //     <a:solidFill>
            //         <a:hslClr hue="14400000" sat="0%" lum="50%">
            //             <a:sat val="100000" />
            //         <a:hslClr />
            //     </a:solidFill>
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (40, C0, 40)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:sat val="50%" />
            //         </a:srgbClr>
            //     <a: solidFill >
            // end example]

            //24. "satMod":
            // Specifies the saturation as expressed by a percentage relative to the input color.
            // Increases never increase the saturation beyond 100 %, decreases never decrease the
            // saturation below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (66, 99, 66)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:satMod val="20%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            let satMod = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:satMod", "attrs", "val"])) / 100000;
            if (!isNaN(satMod)) {
                color = applySatMod(color, satMod, isAlpha);
            }
            //25. "satOff":
            // Specifies the saturation as expressed by a percentage offset increase or decrease to the
            // input color.Increases never increase the saturation beyond 100 %, decreases never
            // decrease the saturation below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (19, E5, 19)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:satOff val="-20%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            // let satOff = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:satOff", "attrs", "val"])) / 100000;
            // if (!isNaN(satOff)) {
            //     console.log("satOff: ", satOff, " (TODO)")
            // }

            //26. "shade":
            // This element specifies a darker version of its input color.A 10 % shade is 10 % of the input color combined with 90 % black.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (00, BC, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:shade val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            // end example]
            let shade = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:shade", "attrs", "val"])) / 100000;
            if (!isNaN(shade)) {
                color = applyShade(color, shade, isAlpha);
            }
            //27.  "tint":
            // This element specifies a lighter version of its input color.A 10 % tint is 10 % of the input color combined with
            // 90 % white.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (BC, FF, BC)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:tint val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            let tint = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:tint", "attrs", "val"])) / 100000;
            if (!isNaN(tint)) {
                color = applyTint(color, tint, isAlpha);
            }
            return color;
        }
        function toHex(n) {
            let hex = n.toString(16);
            while (hex.length < 2) { hex = "0" + hex; }
            return hex;
        }
        function hslToRgb(hue, sat, light) {
            let t1, t2, r, g, b;
            hue = hue / 60;
            if (light <= 0.5) {
                t2 = light * (sat + 1);
            } else {
                t2 = light + sat - (light * sat);
            }
            t1 = light * 2 - t2;
            r = hueToRgb(t1, t2, hue + 2) * 255;
            g = hueToRgb(t1, t2, hue) * 255;
            b = hueToRgb(t1, t2, hue - 2) * 255;
            return { r: r, g: g, b: b };
        }
        function hueToRgb(t1, t2, hue) {
            if (hue < 0) hue += 6;
            if (hue >= 6) hue -= 6;
            if (hue < 1) return (t2 - t1) * hue + t1;
            else if (hue < 3) return t2;
            else if (hue < 4) return (t2 - t1) * (4 - hue) + t1;
            else return t1;
        }
        function getColorName2Hex(name) {
            let hex;
            let colorName = ['white', 'AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk', 'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenRod', 'DarkGray', 'DarkGrey', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray', 'DarkSlateGrey', 'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DimGrey', 'DodgerBlue', 'FireBrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'GoldenRod', 'Gray', 'Grey', 'Green', 'GreenYellow', 'HoneyDew', 'HotPink', 'IndianRed', 'Indigo', 'Ivory', 'Khaki', 'Lavender', 'LavenderBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenRodYellow', 'LightGray', 'LightGrey', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray', 'LightSlateGrey', 'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquaMarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue', 'MintCream', 'MistyRose', 'Moccasin', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenRod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue', 'Purple', 'RebeccaPurple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'SlateGrey', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Turquoise', 'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen'];
            let colorHex = ['ffffff', 'f0f8ff', 'faebd7', '00ffff', '7fffd4', 'f0ffff', 'f5f5dc', 'ffe4c4', '000000', 'ffebcd', '0000ff', '8a2be2', 'a52a2a', 'deb887', '5f9ea0', '7fff00', 'd2691e', 'ff7f50', '6495ed', 'fff8dc', 'dc143c', '00ffff', '00008b', '008b8b', 'b8860b', 'a9a9a9', 'a9a9a9', '006400', 'bdb76b', '8b008b', '556b2f', 'ff8c00', '9932cc', '8b0000', 'e9967a', '8fbc8f', '483d8b', '2f4f4f', '2f4f4f', '00ced1', '9400d3', 'ff1493', '00bfff', '696969', '696969', '1e90ff', 'b22222', 'fffaf0', '228b22', 'ff00ff', 'dcdcdc', 'f8f8ff', 'ffd700', 'daa520', '808080', '808080', '008000', 'adff2f', 'f0fff0', 'ff69b4', 'cd5c5c', '4b0082', 'fffff0', 'f0e68c', 'e6e6fa', 'fff0f5', '7cfc00', 'fffacd', 'add8e6', 'f08080', 'e0ffff', 'fafad2', 'd3d3d3', 'd3d3d3', '90ee90', 'ffb6c1', 'ffa07a', '20b2aa', '87cefa', '778899', '778899', 'b0c4de', 'ffffe0', '00ff00', '32cd32', 'faf0e6', 'ff00ff', '800000', '66cdaa', '0000cd', 'ba55d3', '9370db', '3cb371', '7b68ee', '00fa9a', '48d1cc', 'c71585', '191970', 'f5fffa', 'ffe4e1', 'ffe4b5', 'ffdead', '000080', 'fdf5e6', '808000', '6b8e23', 'ffa500', 'ff4500', 'da70d6', 'eee8aa', '98fb98', 'afeeee', 'db7093', 'ffefd5', 'ffdab9', 'cd853f', 'ffc0cb', 'dda0dd', 'b0e0e6', '800080', '663399', 'ff0000', 'bc8f8f', '4169e1', '8b4513', 'fa8072', 'f4a460', '2e8b57', 'fff5ee', 'a0522d', 'c0c0c0', '87ceeb', '6a5acd', '708090', '708090', 'fffafa', '00ff7f', '4682b4', 'd2b48c', '008080', 'd8bfd8', 'ff6347', '40e0d0', 'ee82ee', 'f5deb3', 'ffffff', 'f5f5f5', 'ffff00', '9acd32'];
            let findIndx = colorName.indexOf(name);
            if (findIndx != -1) {
                hex = colorHex[findIndx];
            }
            return hex;
        }
        function getSchemeColorFromTheme(schemeClr, clrMap, phClr, warpObj) {
            //<p:clrMap ...> in slide master
            // e.g. tx2="dk2" bg2="lt2" tx1="dk1" bg1="lt1" slideLayoutClrOvride
            let color = '';
            var slideLayoutClrOvride;
            if (clrMap !== undefined) {
                slideLayoutClrOvride = clrMap;//getTextByPathList(clrMap, ["p:sldMaster", "p:clrMap", "attrs"])
            } else if (warpObj !== undefined) {
                let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(warpObj["slideContent"], ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    slideLayoutClrOvride = sldClrMapOvr;
                } else {
                    let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutContent"], ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                    if (sldClrMapOvr !== undefined) {
                        slideLayoutClrOvride = sldClrMapOvr;
                    } else {
                        slideLayoutClrOvride = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterContent"], ["p:sldMaster", "p:clrMap", "attrs"]);
                    }

                }
            }

            let schmClrName = schemeClr.substr(2);
            if (schmClrName == "phClr" && phClr !== undefined) {
                color = phClr;
            } else {
                if (slideLayoutClrOvride !== undefined) {
                    switch (schmClrName) {
                        case "tx1":
                        case "tx2":
                        case "bg1":
                        case "bg2":
                            schemeClr = "a:" + slideLayoutClrOvride[schmClrName];
                            break;
                    }
                } else {
                    switch (schmClrName) {
                        case "tx1":
                            schemeClr = "a:dk1";
                            break;
                        case "tx2":
                            schemeClr = "a:dk2";
                            break;
                        case "bg1":
                            schemeClr = "a:lt1";
                            break;
                        case "bg2":
                            schemeClr = "a:lt2";
                            break;
                    }
                }

                let refNode = PPTXXmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
                color = PPTXXmlUtils.getTextByPathList(refNode, ["a:srgbClr", "attrs", "val"]);
                if (color === undefined && refNode !== undefined) {
                    color = PPTXXmlUtils.getTextByPathList(refNode, ["a:sysClr", "attrs", "lastClr"]);
                }
            }
            return color;
        }

        function extractChartData(serNode, warpObj) {

            let dataMat = new Array();

            if (serNode === undefined) {
                return dataMat;
            }

            if (serNode["c:xVal"] !== undefined) {
                var dataRow = new Array();
                eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                    dataRow.push(parseFloat(innerNode["c:v"]));
                    return "";
                });
                dataMat.push(dataRow);
                dataRow = new Array();
                eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                    dataRow.push(parseFloat(innerNode["c:v"]));
                    return "";
                });
                dataMat.push(dataRow);
            } else {
                eachElement(serNode, function (innerNode, index) {
                    var dataRow = new Array();
                    // 提取系列名称（从c:tx中）
                    let colName;
                    const txStrRef = PPTXXmlUtils.getTextByPathList(innerNode, ["c:tx", "c:strRef"]);
                    if (txStrRef) {
                        const strCache = PPTXXmlUtils.getTextByPathList(txStrRef, ["c:strCache"]);
                        if (strCache) {
                            const pt = PPTXXmlUtils.getTextByPathList(strCache, ["c:pt"]);
                            if (pt) {
                                // pt可能是数组或单个对象
                                if (Array.isArray(pt)) {
                                    colName = pt[0]["c:v"];
                                } else {
                                    colName = pt["c:v"];
                                }
                            }
                        }
                    }
                    // 如果没有从strRef中提取到，尝试从其他方式提取
                    if (!colName) {
                        colName = PPTXXmlUtils.getTextByPathList(innerNode, ["c:tx", "c:v"]) || index;
                    }

                    // Category (string or number)
                    let rowNames = {};
                    if  (PPTXXmlUtils.getTextByPathList(innerNode, ["c:cat", "c:strRef", "c:strCache", "c:pt"]) !== undefined) {
                        eachElement(innerNode["c:cat"]["c:strRef"]["c:strCache"]["c:pt"], function (innerNode, index) {
                            rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                            return "";
                        });
                    } else if  (PPTXXmlUtils.getTextByPathList(innerNode, ["c:cat", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                        eachElement(innerNode["c:cat"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                            rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                            return "";
                        });
                    } else if (PPTXXmlUtils.getTextByPathList(innerNode, ["c:cat", "c:multiLvlStrRef", "c:multiLvlStrCache"]) !== undefined) {
                        // Handle multi-level string reference (c:multiLvlStrRef) - use first level labels
                        const multiLvlCache = PPTXXmlUtils.getTextByPathList(innerNode, ["c:cat", "c:multiLvlStrRef", "c:multiLvlStrCache"]);
                        const lvl = PPTXXmlUtils.getTextByPathList(multiLvlCache, ["c:lvl"]);
                        if (lvl) {
                            // lvl might be an array of levels; use the first level
                            const firstLvl = Array.isArray(lvl) ? lvl[0] : lvl;
                            const pts = PPTXXmlUtils.getTextByPathList(firstLvl, ["c:pt"]);
                            if (pts) {
                                eachElement(pts, function (pt, index) {
                                    rowNames[pt["attrs"]["idx"]] = pt["c:v"];
                                    return "";
                                });
                            }
                        }
                    }

                    // Value
                    if  (PPTXXmlUtils.getTextByPathList(innerNode, ["c:val", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                        eachElement(innerNode["c:val"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                            dataRow.push({ x: innerNode["attrs"]["idx"], y: parseFloat(innerNode["c:v"]) });
                            return "";
                        });
                    }

                    // Extract series style information
                    let seriesStyle = {};
                    
                    // Extract fill color if available
                    let fillType = getFillType(PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr"]));
                    if (fillType === "SOLID_FILL" && warpObj !== undefined) {
                        let fillNode = PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr", "a:solidFill"]);
                        if (fillNode !== undefined) {
                            let fillColor = getSolidFill(fillNode, undefined, undefined, warpObj);
                            if (fillColor !== undefined) {
                                if (fillColor && !fillColor.startsWith('#')) {
                                    fillColor = '#' + fillColor;
                                }
                                seriesStyle.fillColor = fillColor;
                            }
                        }
                    } else if (fillType === "GRADIENT_FILL" && warpObj !== undefined) {
                        let gradFillNode = PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr", "a:gradFill"]);
                        if (gradFillNode !== undefined) {
                            let gradientFill = getGradientFill(gradFillNode, warpObj);
                            if (gradientFill !== undefined) {
                                seriesStyle.gradientFill = gradientFill;
                            }
                        }
                    }
                    
                    // Extract line color if available
                    let lineNode = PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr", "a:ln"]);
                    if (lineNode !== undefined && warpObj !== undefined) {
                        let lineFillType = getFillType(lineNode);
                        if (lineFillType === "SOLID_FILL") {
                            let lineColor = getSolidFill(lineNode["a:solidFill"], undefined, undefined, warpObj);
                            if (lineColor !== undefined) {
                                if (lineColor && !lineColor.startsWith('#')) {
                                    lineColor = '#' + lineColor;
                                }
                                seriesStyle.lineColor = lineColor;
                            }
                        } else if (lineFillType === "GRADIENT_FILL") {
                            let lineGradFillNode = lineNode["a:gradFill"];
                            if (lineGradFillNode !== undefined) {
                                let lineGradientFill = getGradientFill(lineGradFillNode, warpObj);
                                if (lineGradientFill !== undefined) {
                                    seriesStyle.lineGradientFill = lineGradientFill;
                                }
                            }
                        }
                    }

                    dataMat.push({ key: colName, values: dataRow, xlabels: rowNames, style: seriesStyle });
                    return "";
                });
            }

            return dataMat;
        }
        /**
         * setTextByPathList
         * @param {Object} node
         * @param {string Array} path
         * @param {string} value
         */
        function setTextByPathList(node, path, value) {

            if (path.constructor !== Array) {
                throw Error("Error of path type! path is not array.");
            }

            if (node === undefined) {
                return undefined;
            }

            function setObjectPath(obj, parts, value) {
                if(!parts) return obj;
                //var parts = prop.split('.');
                let current = obj;
                let lent = parts.length;
                for (let i = 0; i < lent; i++) {
                    var p = parts[i];
                    if (current[p] === undefined) {
                        if (i == lent - 1) {
                            current[p] = value;
                        } else {
                            current[p] = {};
                        }
                    }
                    current = current[p];
                }
                return obj;
            }

            setObjectPath(node, path, value);
        }

        /**
         * eachElement
         * @param {Object} node
         * @param {function} doFunction
         */
        function eachElement(node, doFunction) {
            if (node === undefined) {
                return;
            }
            let result = "";
            if (node.constructor === Array) {
                let l = node.length;
                for (let i = 0; i < l; i++) {
                    result += doFunction(node[i], i);
                }
            } else {
                result += doFunction(node, 0);
            }
            return result;
        }

        // ===== Color functions =====
        /**
         * applyShade
         * @param {string} rgbStr
         * @param {number} shadeValue
         */
        function applyShade(rgbStr, shadeValue, isAlpha) {
            let color = tinycolor$1(rgbStr).toHsl();
            // 确保shadeValue在0-1之间
            shadeValue = Math.max(0, Math.min(1, shadeValue));
            // PPTX标准：Shade = L * shadeValue
            let cacl_l = Math.max(0, Math.min(1, color.l * shadeValue));
            if (isAlpha)
                return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
            return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
        }

        /**
         * applyTint
         * @param {string} rgbStr
         * @param {number} tintValue
         */
        function applyTint(rgbStr, tintValue, isAlpha) {
            let color = tinycolor$1(rgbStr).toHsl();
            // 确保tintValue在0-1之间
            tintValue = Math.max(0, Math.min(1, tintValue));
            // PPTX标准：Tint = L * tintValue + (1 - tintValue)
            let cacl_l = Math.max(0, Math.min(1, color.l * tintValue + (1 - tintValue)));
            if (isAlpha)
                return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
            return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
        }

        /**
         * applyLumOff
         * @param {string} rgbStr
         * @param {number} offset
         */
        function applyLumOff(rgbStr, offset, isAlpha) {
            let color = tinycolor$1(rgbStr).toHsl();
            let lum = offset + color.l;
            if (lum >= 1) {
                if (isAlpha)
                    return tinycolor$1({ h: color.h, s: color.s, l: 1, a: color.a }).toHex8();
                return tinycolor$1({ h: color.h, s: color.s, l: 1, a: color.a }).toHex();
            }
            if (isAlpha)
                return tinycolor$1({ h: color.h, s: color.s, l: lum, a: color.a }).toHex8();
            return tinycolor$1({ h: color.h, s: color.s, l: lum, a: color.a }).toHex();
        }

        /**
         * applyLumMod
         * @param {string} rgbStr
         * @param {number} multiplier
         */
        function applyLumMod(rgbStr, multiplier, isAlpha) {
            let color = tinycolor$1(rgbStr).toHsl();
            let cacl_l = color.l * multiplier;
            if (cacl_l >= 1) {
                cacl_l = 1;
            }
            if (isAlpha)
                return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
            return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
        }


        // /**
        //  * applyHueMod
        //  * @param {string} rgbStr
        //  * @param {number} multiplier
        //  */
        function applyHueMod(rgbStr, multiplier, isAlpha) {
            let color = tinycolor$1(rgbStr).toHsl();
            var cacl_h = color.h * multiplier;
            if (cacl_h >= 360) {
                cacl_h = cacl_h - 360;
            }
            if (isAlpha)
                return tinycolor$1({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex8();
            return tinycolor$1({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex();
        }


        // /**
        //  * applyHueOff
        //  * @param {string} rgbStr
        //  * @param {number} offset
        //  */
        // function applyHueOff(rgbStr, offset, isAlpha) {
        //     let color = tinycolor(rgbStr).toHsl();
        //     //console.log("applyLumMod  color.h: ", color.h, ", offset: ", offset, ", color.h * offset : ", color.h * offset)

        //     var cacl_h = color.h * offset;
        //     if (cacl_h >= 360) {
        //         cacl_h = cacl_h - 360;
        //     }
        //     if (isAlpha)
        //         return tinycolor({ h: cocacl_h, s: color.s, l: color.l, a: color.a }).toHex8();
        //     return tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex();
        // }
        // /**
        //  * applySatMod
        //  * @param {string} rgbStr
        //  * @param {number} multiplier
        //  */
        function applySatMod(rgbStr, multiplier, isAlpha) {
            let color = tinycolor$1(rgbStr).toHsl();
            let cacl_s = color.s * multiplier;
            if (cacl_s >= 1) {
                cacl_s = 1;
            }
            //return;
            // if (isAlpha)
            //     return tinycolor(rgbStr).saturate(multiplier * 100).toHex8();
            // return tinycolor(rgbStr).saturate(multiplier * 100).toHex();
            if (isAlpha)
                return tinycolor$1({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex8();
            return tinycolor$1({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex();
        }

        /**
         * rgba2hex
         * @param {string} rgbaStr
         */
        function rgba2hex(rgbaStr) {
            var a,
                rgb = rgbaStr.replace(/\s/g, '').match(/^rgba?\((\d+),(\d+),(\d+),?([^,\s)]+)?/i),
                alpha = (rgb && rgb[4] || "").trim(),
                hex = rgb ?
                    (rgb[1] | 1 << 8).toString(16).slice(1) +
                    (rgb[2] | 1 << 8).toString(16).slice(1) +
                    (rgb[3] | 1 << 8).toString(16).slice(1) : rgbaStr;

            if (alpha !== "") {
                a = alpha;
            } else {
                a = 0o1;
            }
            // multiply before convert to HEX
            a = ((a * 255) | 1 << 8).toString(16).slice(1);
            hex = hex + a;

            return hex;
        }
        // function degreesToRadians(degrees) {
        //     //Math.PI
        //     if (degrees == "" || degrees == null || degrees == undefined) {
        //         return 0;
        //     }
        //     return degrees * (Math.PI / 180);
        // }
        
        function getSvgGradient(w, h, angl, color_arry, shpId) {
            var stopsArray = getMiddleStops(color_arry - 2);

            let svgAngle = '',
                svgHeight = h,
                svgWidth = w,
                svg = '',
                xy_ary = SVGangle(angl, svgHeight, svgWidth),
                x1 = xy_ary[0],
                y1 = xy_ary[1],
                x2 = xy_ary[2],
                y2 = xy_ary[3];

            let sal = stopsArray.length,
                sr = sal < 20 ? 100 : 1000;
            svgAngle = ' gradientUnits="userSpaceOnUse" x1="' + x1 + '%" y1="' + y1 + '%" x2="' + x2 + '%" y2="' + y2 + '%"';
            svgAngle = '<linearGradient id="linGrd_' + shpId + '"' + svgAngle + '>\n';
            svg += svgAngle;

            for (let i = 0; i < sal; i++) {
                var tinClr = tinycolor$1("#" + color_arry[i]);
                let alpha = tinClr.getAlpha();
                svg += '<stop offset="' + Math.round(parseFloat(stopsArray[i]) / 100 * sr) / sr + '" style="stop-color:' + tinClr.toHexString() + '; stop-opacity:' + (alpha) + ';"';
                svg += '/>\n';
            }

            svg += '</linearGradient>\n' + '';

            return svg
        }
        function getMiddleStops(s) {
            let sArry = ['0%', '100%'];
            if (s == 0) {
                return sArry;
            } else {
                let i = s;
                while (i--) {
                    let middleStop = 100 - ((100 / (s + 1)) * (i + 1)), // AM: Ex - For 3 middle stops, progression will be 25%, 50%, and 75%, plus 0% and 100% at the ends.
                        middleStopString = middleStop + "%";
                    sArry.splice(-1, 0, middleStopString);
                } // AM: add into stopsArray before 100%
            }
            return sArry
        }
        function SVGangle(deg, svgHeight, svgWidth) {
            let w = parseFloat(svgWidth),
                h = parseFloat(svgHeight),
                ang = parseFloat(deg),
                o = 2,
                n = 2,
                wc = w / 2,
                hc = h / 2,
                tx1 = 2,
                ty1 = 2,
                tx2 = 2,
                ty2 = 2,
                k = (((ang % 360) + 360) % 360),
                j = (360 - k) * Math.PI / 180,
                i = Math.tan(j),
                l = hc - i * wc;

            if (k == 0) {
                tx1 = w,
                    ty1 = hc,
                    tx2 = 0,
                    ty2 = hc;
            } else if (k < 90) {
                n = w,
                    o = 0;
            } else if (k == 90) {
                tx1 = wc,
                    ty1 = 0,
                    tx2 = wc,
                    ty2 = h;
            } else if (k < 180) {
                n = 0,
                    o = 0;
            } else if (k == 180) {
                tx1 = 0,
                    ty1 = hc,
                    tx2 = w,
                    ty2 = hc;
            } else if (k < 270) {
                n = 0,
                    o = h;
            } else if (k == 270) {
                tx1 = wc,
                    ty1 = h,
                    tx2 = wc,
                    ty2 = 0;
            } else {
                n = w,
                    o = h;
            }
            // AM: I could not quite figure out what m, n, and o are supposed to represent from the original code on visualcsstools.com.
            let m = o + (n / i);
                tx1 = tx1 == 2 ? i * (m - l) / (Math.pow(i, 2) + 1) : tx1,
                ty1 = ty1 == 2 ? i * tx1 + l : ty1,
                tx2 = tx2 == 2 ? w - tx1 : tx2,
                ty2 = ty2 == 2 ? h - ty1 : ty2;
            let x1 = Math.round(tx2 / w * 100 * 100) / 100,
                y1 = Math.round(ty2 / h * 100 * 100) / 100,
                x2 = Math.round(tx1 / w * 100 * 100) / 100,
                y2 = Math.round(ty1 / h * 100 * 100) / 100;
            return [x1, y1, x2, y2];
        }
        function getSvgImagePattern(node, fill, shpId, warpObj) {
            // 处理 fill 参数是对象的情况
            let fillUrl = fill;
            if (typeof fill === 'object' && fill.img) {
                fillUrl = fill.img;
            }
            
            let pic_dim = getBase64ImageDimensions(fillUrl);
            let width = pic_dim[0];
            let height = pic_dim[1];
            let blipFillNode = node["p:spPr"]["a:blipFill"];
            let sx = 0, sy = 0;
            let tileNode = PPTXXmlUtils.getTextByPathList(blipFillNode, ["a:tile", "attrs"]);
            if (tileNode !== undefined && tileNode["sx"] !== undefined) {
                sx = (parseInt(tileNode["sx"]) / 100000) * width;
                sy = (parseInt(tileNode["sy"]) / 100000) * height;
            }

            let blipNode = node["p:spPr"]["a:blipFill"]["a:blip"];
            let tialphaModFixNode = PPTXXmlUtils.getTextByPathList(blipNode, ["a:alphaModFix", "attrs"]);
            let imgOpacity = "";
            if (tialphaModFixNode !== undefined && tialphaModFixNode["amt"] !== undefined && tialphaModFixNode["amt"] != "") {
                parseInt(tialphaModFixNode["amt"]) / 100000;

            }
            let ptrn = '';
            if (sx !== undefined && sx != 0) {
                ptrn = '<pattern id="imgPtrn_' + shpId + '" x="0" y="0"  width="' + sx + '" height="' + sy + '" patternUnits="userSpaceOnUse">';
            } else {
                ptrn = '<pattern id="imgPtrn_' + shpId + '"  patternContentUnits="objectBoundingBox"  width="1" height="1">';
            }
            let duotoneNode = PPTXXmlUtils.getTextByPathList(blipNode, ["a:duotone"]);
            let fillterNode = "";
            let filterUrl = "";
            if (duotoneNode !== undefined) {
                var clr_ary = [];
                Object.keys(duotoneNode).forEach(clr_type => {
                    //Object.keys(duotoneNode[clr_type]).forEach(clr => {
                    //console.log("blip pic duotone clr: ", duotoneNode[clr_type][clr], clr)
                    if (clr_type != "attrs") {
                        let obj = {};
                        obj[clr_type] = duotoneNode[clr_type];
                        let hexClr = getSolidFill(obj, undefined, undefined, warpObj);
                        //clr_ary.push();

                        let color = tinycolor$1("#" + hexClr);
                        clr_ary.push(color.toRgb()); // { r: 255, g: 0, b: 0, a: 1 }
                    }
                    // })
                });

                if (clr_ary.length == 2) {

                    fillterNode = '<filter id="svg_image_duotone"> ' +
                        '<feColorMatrix type="matrix" values=".33 .33 .33 0 0' +
                        '.33 .33 .33 0 0' +
                        '.33 .33 .33 0 0' +
                        '0 0 0 1 0">' +
                        '</feColorMatrix>' +
                        '<feComponentTransfer color-interpolation-filters="sRGB">' +
                        //clr_ary.forEach(function(clr){
                        '<feFuncR type="table" tableValues="' + clr_ary[0].r / 255 + ' ' + clr_ary[1].r / 255 + '"></feFuncR>' +
                        '<feFuncG type="table" tableValues="' + clr_ary[0].g / 255 + ' ' + clr_ary[1].g / 255 + '"></feFuncG>' +
                        '<feFuncB type="table" tableValues="' + clr_ary[0].b / 255 + ' ' + clr_ary[1].b / 255 + '"></feFuncB>' +
                        //});
                        '</feComponentTransfer>' +
                        ' </filter>';
                }

                filterUrl = 'filter="url(#svg_image_duotone)"';

                ptrn += fillterNode;
            }

            fillUrl = PPTXXmlUtils.escapeHtml(fillUrl);
            if (sx !== undefined && sx != 0) {
                ptrn += '<image  xlink:href="' + fillUrl + '" x="0" y="0" width="' + sx + '" height="' + sy + '" ' + imgOpacity + ' ' + filterUrl + '></image>';
            } else {
                ptrn += '<image  xlink:href="' + fillUrl + '" preserveAspectRatio="none" width="1" height="1" ' + imgOpacity + ' ' + filterUrl + '></image>';
            }
            ptrn += '</pattern>';

            return ptrn;
        }

        function getBase64ImageDimensions(imgSrc) {
            let image = new Image();
            image.onload = function () {
                image.width;
                image.height;
            };
            image.src = imgSrc;

            do {
                if (image.width !== undefined) {
                    return [image.width, image.height];
                }
            } while (image.width === undefined);

            //return [w, h];
        }

        function getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) {

            //X, <a:bodyPr anchor="ctr">, <a:bodyPr anchor="b">
            var anchor = PPTXXmlUtils.getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);

            if (anchor === undefined) {
                anchor = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                if (anchor === undefined) {
                    anchor = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                    if (anchor === undefined) {
                        //"If this attribute is omitted, then a value of t, or top is implied."
                        anchor = "t";//getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                    }
                }
            }

            // 对于圆形/椭圆类形状，强制使用居中对齐以确保文本在形状内正确居中显示
            let shapeType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
            const circularShapes = [
                "ellipse", "ovalCallout", "wedgeEllipseCallout",
                "pie", "pieWedge", "chord", "sector", "arc", "blockArc"
            ];
            if (circularShapes.includes(shapeType) && anchor === "t") {
                anchor = "ctr";
            }
            
            return (anchor === "ctr")?"v-mid" : ((anchor === "b") ? "v-down" : "v-up");
        }

    function getContentDir(node, type, warpObj) {
            return "content";
            //console.log("getContentDir() type:", type, "slideMasterTextStyles:", slideMasterTextStyles,"dirNode:",dirVal)
        }

        function getVerticalMargins(pNode, textBodyNode, type, idx, warpObj, totalParagraphs, paragraphIndex, anchor) {
            //margin-top ;
            //a:pPr => a:spcBef => a:spcPts (/100) | a:spcPct (/?)
            //margin-bottom
            //a:pPr => a:spcAft => a:spcPts (/100) | a:spcPct (/?)
            //+
            //a:pPr =>a:lnSpc => a:spcPts (/?) | a:spcPct (/?)
            //let lstStyle = textBodyNode["a:lstStyle"];
            let lvl = 1;
            var spcBefNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPts", "attrs", "val"]);
            var spcAftNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPts", "attrs", "val"]);
            var spcBefType = "Pts";
            var spcAftType = "Pts";
            // 标记 spcBef 是否来自段落的显式设置（而非 lstStyle 的默认值）
            var spcBefIsExplicit = (spcBefNode !== undefined);
            // 标记是否应该减少 lstStyle 的默认 spcBef
            // 当只有一个段落时，且是第一个段落，则减少 lstStyle 的默认 spcBef
            var spcBefScale = 1.0;
            if (!spcBefIsExplicit && totalParagraphs === 1 && paragraphIndex === 0) {
                spcBefScale = 0.0; // 完全忽略第一个段落的默认 spcBef
            }
            // 如果没有找到 spcPts，则查找 spcPct（百分比类型的段落间距）
            if (spcBefNode === undefined) {
                spcBefNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPct", "attrs", "val"]);
                if (spcBefNode !== undefined) {
                    spcBefType = "Pct";
                }
            }
            if (spcAftNode === undefined) {
                spcAftNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPct", "attrs", "val"]);
                if (spcAftNode !== undefined) {
                    spcAftType = "Pct";
                }
            }
            let lnSpcNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPct", "attrs", "val"]);
            let lnSpcNodeType = "Pct";
            if (lnSpcNode === undefined) {
                lnSpcNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                if (lnSpcNode !== undefined) {
                    lnSpcNodeType = "Pts";
                }
            }
            let lvlNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvl = parseInt(lvlNode) + 1;
            }
            let fontSize;
            if  (PPTXXmlUtils.getTextByPathList(pNode, ["a:r"]) !== undefined) {
                let fontSizeStr = getFontSize(pNode["a:r"], textBodyNode,undefined, lvl, type, warpObj);
                if (fontSizeStr != "inherit") {
                    // 提取数字部分（假设格式为 "35px"）
                    const fontSizeMatch = fontSizeStr.match(/(\d+(?:\.\d+)?)px/);
                    if (fontSizeMatch) {
                        fontSize = parseFloat(fontSizeMatch[1]);
                    }
                }
            }
            //var spcBef = "";
            // if(spcBefNode !== undefined){
            //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
            // }
            // else{
            //    //i did not found case with percentage 
            //     spcBefNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPct","attrs","val"]);
            //     if(spcBefNode !== undefined){
            //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
            //     }
            // }
            //var spcAft = "";
            // if(spcAftNode !== undefined){
            //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
            // }
            // else{
            //    //i did not found case with percentage 
            //     spcAftNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPct","attrs","val"]);
            //     if(spcAftNode !== undefined){
            //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
            //     }
            // }
            // if(spcAftNode !== undefined){
            //     //check in layout and then in master
            // }
            // 首先检查slide本身的lstStyle（PPTX标准：段落样式继承顺序：pPr -> lstStyle -> layout -> master）
            let lstStyle = textBodyNode["a:lstStyle"];
            if (lstStyle !== undefined && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
                let lvlKey = "a:lvl" + lvl + "pPr";
                let lstLvlNode = lstStyle[lvlKey];
                if (lstLvlNode !== undefined) {
                    if (spcBefNode === undefined) {
                        spcBefNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                        if (spcBefNode === undefined) {
                            spcBefNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:spcBef", "a:spcPct", "attrs", "val"]);
                            if (spcBefNode !== undefined) {
                                spcBefType = "Pct";
                            }
                        }
                    }
                    if (spcAftNode === undefined) {
                        spcAftNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                        if (spcAftNode === undefined) {
                            spcAftNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:spcAft", "a:spcPct", "attrs", "val"]);
                            if (spcAftNode !== undefined) {
                                spcAftType = "Pct";
                            }
                        }
                    }
                    if (lnSpcNode === undefined) {
                        lnSpcNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                        if (lnSpcNode === undefined) {
                            lnSpcNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:lnSpc", "a:spcPts", "attrs", "val"]);
                            if (lnSpcNode !== undefined) {
                                lnSpcNodeType = "Pts";
                            }
                        }
                    }
                }
            }
            let isInLayoutOrMaster = true;
            if(type == "shape" || type == "textBox"){
                isInLayoutOrMaster = false;
            }
            if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
                //check in layout
                if (idx !== undefined) {
                    let laypPrNode = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "idxTable", idx, "p:txBody", "a:p", (lvl - 1), "a:pPr"]);

                    if (spcBefNode === undefined) {
                        spcBefNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                        if (spcBefNode === undefined) {
                            spcBefNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPct", "attrs", "val"]);
                            if (spcBefNode !== undefined) {
                                spcBefType = "Pct";
                            }
                        }
                    }

                    if (spcAftNode === undefined) {
                        spcAftNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                        if (spcAftNode === undefined) {
                            spcAftNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPct", "attrs", "val"]);
                            if (spcAftNode !== undefined) {
                                spcAftType = "Pct";
                            }
                        }
                    }

                    if (lnSpcNode === undefined) {
                        lnSpcNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                        if (lnSpcNode === undefined) {
                            lnSpcNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                            if (lnSpcNode !== undefined) {
                                lnSpcNodeType = "Pts";
                            }
                        }
                    }
                }

            }
            if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
                //check in master
                //slideMasterTextStyles
                var slideMasterTextStyles = warpObj["slideMasterTextStyles"];
                let dirLoc = "";
                lvl = "a:lvl" + lvl + "pPr";
                switch (type) {
                    case "title":
                    case "ctrTitle":
                        dirLoc = "p:titleStyle";
                        break;
                    case "body":
                    case "obj":
                    case "dt":
                    case "ftr":
                    case "sldNum":
                    case "textBox":
                    // case "shape":
                        dirLoc = "p:bodyStyle";
                        break;
                    case "shape":
                    //case "textBox":
                    default:
                        dirLoc = "p:otherStyle";
                }
                // if (type == "shape" || type == "textBox") {
                //     lvl = "a:lvl1pPr";
                // }
                let inLvlNode = PPTXXmlUtils.getTextByPathList(slideMasterTextStyles, [dirLoc, lvl]);
                if (inLvlNode !== undefined) {
                    if (spcBefNode === undefined) {
                        spcBefNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                        if (spcBefNode === undefined) {
                            spcBefNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPct", "attrs", "val"]);
                            if (spcBefNode !== undefined) {
                                spcBefType = "Pct";
                            }
                        }
                    }

                    if (spcAftNode === undefined) {
                        spcAftNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                        if (spcAftNode === undefined) {
                            spcAftNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPct", "attrs", "val"]);
                            if (spcAftNode !== undefined) {
                                spcAftType = "Pct";
                            }
                        }
                    }

                    if (lnSpcNode === undefined) {
                        lnSpcNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                        if (lnSpcNode === undefined) {
                            lnSpcNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                            if (lnSpcNode !== undefined) {
                                lnSpcNodeType = "Pts";
                            }
                        }
                    }
                }
            }
            let spcBefor = 0, spcAfter = 0, spcLines = 0;
            let marginTopBottomStr = "";
            if (spcBefNode !== undefined) {
                if (spcBefType === "Pct") {
                    // 百分比类型：值以千分之一为单位，如 20000 = 20000‰ = 20倍行高
                    spcBefor = parseInt(spcBefNode) / 1000;
                } else {
                    // 点数类型：值以百分之一磅为单位，需要除以 100 转换为 pt
                    spcBefor = parseInt(spcBefNode) / 100;
                }
            }
            if (spcAftNode !== undefined) {
                if (spcAftType === "Pct") {
                    // 百分比类型
                    spcAfter = parseInt(spcAftNode) / 1000;
                } else {
                    // 点数类型
                    spcAfter = parseInt(spcAftNode) / 100;
                }
            }

            // 处理行间距（lnSpc）
            if (lnSpcNode !== undefined) {
                if (lnSpcNodeType === "Pct") {
                    // 百分比类型的行间距（如90000 = 90%）
                    spcLines = parseInt(lnSpcNode) / 100000;
                    // 将百分比转换为line-height
                    // PPTX的行间距百分比与CSS的line-height含义相同
                    let lineHeight = spcLines;
                    // 如果行间距太小（如90%），调整为更合理的值
                    // 实际测试发现，90%太挤，1.3更符合PPT的实际显示效果
                    if (lineHeight < 1.0) {
                        lineHeight = 1.3;
                    }
                    marginTopBottomStr += "line-height: " + lineHeight + ";";
                } else if (lnSpcNodeType === "Pts") {
                    // 点数类型的行间距
                    spcLines = parseInt(lnSpcNode) / 100;
                    // 将点数转换为line-height（相对于字体大小）
                    if (fontSize && fontSize > 0) {
                        let lineHeight = spcLines / fontSize;
                        // 如果行间距太小，调整为更合理的值
                        if (lineHeight < 1.0) {
                            lineHeight = 1.3;
                        }
                        marginTopBottomStr += "line-height: " + lineHeight + ";";
                    }
                }
            } else if (type === "textBox") {
                // PPTX标准：当段落没有明确设置行间距时，使用合理的默认行高
                // 实际测试发现，虽然PPTX标准说是100%，但在实际软件中默认值更接近130%
                // 使用1.3作为默认值，完全符合PPT的实际显示效果
                // bodyPr的内边距已经减少了可用空间，配合box-sizing: border-box使用更合适
                marginTopBottomStr += "line-height: 1.3;";
            }

            // 段落前间距
            // 只有当间距来自段落的显式设置（而非 lstStyle 的默认值）时才应用
            // 或者当有多个段落时，lstStyle 的默认间距也是合理的（用于段落之间）
            // 或者应用缩放比例（主要用于单个段落的情况）
            // 重要：当垂直居中时（anchor="ctr"），PPT会自动计算文本的中心位置，不应用段落间距
            if (spcBefNode !== undefined && (spcBefIsExplicit || spcBefScale > 0) && anchor !== "ctr") {
                let marginTop;
                if (spcBefType === "Pct") {
                    // 百分比类型：相对于行高
                    // PPTX中 spcPct 值以千分之一为单位
                    // 实际测试发现，某些PPT文件中的值（如20000）会设置得非常大
                    // 但在PPT中实际渲染时并不会完全按照这个值显示
                    // 可能是PPT内部有最大值限制或特殊处理
                    // 根据经验，将超过500‰（50%行高）的值限制为50%更符合实际效果
                    let lineHeightPx = fontSize || 18; // 默认 18px
                    let spcBeforLimited = Math.min(spcBefor, 0.5); // 限制最大为0.5（即50%行高）
                    // 应用缩放比例（主要用于单个段落的情况）
                    spcBeforLimited *= spcBefScale;
                    marginTop = lineHeightPx * spcBeforLimited;
                } else {
                    // 点数类型：转换为像素（假设1pt = 1.33px）
                    marginTop = spcBefor * 1.33 * spcBefScale;
                }
                // 只有当间距大于0时才应用
                if (marginTop > 0) {
                    marginTop = Math.round(marginTop * 100) / 100;
                    marginTopBottomStr += "margin-top: " + marginTop + "px;";
                }
            }

            // 段落后间距
            // 重要：当垂直居中时（anchor="ctr"），PPT会自动计算文本的中心位置，不应用段落间距
            if (spcAftNode !== undefined && anchor !== "ctr") {
                let marginBottom;
                if (spcAftType === "Pct") {
                    // 百分比类型：相对于行高
                    let lineHeightPx = fontSize || 18; // 默认 18px
                    // 同样对过大的值进行限制
                    let spcAfterLimited = Math.min(spcAfter, 0.5);
                    marginBottom = lineHeightPx * spcAfterLimited;
                } else {
                    // 点数类型：转换为像素
                    marginBottom = spcAfter * 1.33;
                }
                marginBottom = Math.round(marginBottom * 100) / 100;
                marginTopBottomStr += "margin-bottom: " + marginBottom + "px;";
            }

            //return spcAft + spcBef;
            return marginTopBottomStr;
        }
        function getHorizontalAlign(node, textBodyNode, idx, type, prg_dir, warpObj, spNode) {
            let algn = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
            if (algn === undefined) {
                let layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
                let pPrNodeLaout = layoutMasterNode.nodeLaout;
                let pPrNodeMaster = layoutMasterNode.nodeMaster;
                let lvlIdx = 1;
                let lvlNode = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "lvl"]);
                if (lvlNode !== undefined) {
                    lvlIdx = parseInt(lvlNode) + 1;
                }
                let lvlStr = "a:lvl" + lvlIdx + "pPr";

                let lstStyle = textBodyNode["a:lstStyle"];
                algn = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr, "attrs", "algn"]);

                if (algn === undefined && idx !== undefined ) {
                    //slidelayout
                    algn = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                    if (algn === undefined) {
                        algn = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
                        if (algn === undefined) {
                            algn = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvlIdx - 1), "a:pPr", "attrs", "algn"]);
                        }
                    }
                }
                if (algn === undefined) {
                    if (type !== undefined) {
                        //slidelayout
                        algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);

                        if (algn === undefined) {
                            //masterlayout
                            if (type == "title" || type == "ctrTitle") {
                                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                            } else if (type == "body" || type == "obj" || type == "subTitle") {
                                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                            } else if (type == "shape" || type == "diagram") {
                                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr, "attrs", "algn"]);
                            } else if (type == "textBox") {
                                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr, "attrs", "algn"]);
                            } else {
                                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                            }
                        }
                    } else {
                        algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                    }
                }
                // 尝试从布局和母版节点中直接获取对齐属性
                if (algn === undefined && pPrNodeLaout) {
                    algn = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
                }
                if (algn === undefined && pPrNodeMaster) {
                    algn = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
                }
            }

            if (algn === undefined) {
                // 对于特定位置的文本元素，尝试根据位置推断对齐方式
                // 例如，位于幻灯片右侧的文本元素可能是右对齐的
                if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                    return "h-mid";
                } else if (type == "sldNum") {
                    return "h-right";
                } else {
                    // 默认返回左对齐
                    return "h-left";
                }
            }
            
            // 对于圆形/椭圆类形状，强制使用居中对齐
            let shapeType = "";
            if (spNode) {
                shapeType = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
            }
            // 如果在组合形状中没有找到，尝试从数据属性中获取
            if (!shapeType && spNode && spNode["attrs"] && spNode["attrs"]["data-geom-type"]) {
                shapeType = spNode["attrs"]["data-geom-type"];
            }
            const circularShapes = [
                "ellipse", "ovalCallout", "wedgeEllipseCallout",
                "pie", "pieWedge", "chord", "sector", "arc", "blockArc"
            ];
            const isCircularShape = circularShapes.includes(shapeType);
            
            if (isCircularShape) {
                return "h-mid";
            }
            
            if (algn !== undefined) {
                switch (algn) {
                    case "l":
                        if (prg_dir == "pregraph-rtl"){
                            return "h-left-rtl";
                        }else {
                            return "h-left";
                        }
                    case "r":
                        if (prg_dir == "pregraph-rtl") {
                            return "h-right-rtl";
                        }else {
                            return "h-right";
                        }
                    case "ctr":
                        return "h-mid";
                    case "just":
                    case "dist":
                    default:
                        return "h-" + algn;
                }
            }
            //return algn === "ctr" ? "h-mid" : algn === "r" ? "h-right" : "h-left";
        }

        function getLayoutAndMasterNode(node, idx, type, warpObj) {
            let pPrNodeLaout, pPrNodeMaster;
            var pPrNode = node["a:pPr"];
            //lvl
            let lvl = 1;
            let lvlNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvl = parseInt(lvlNode) + 1;
            }
            if (idx !== undefined) {
                //slidelayout
                pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", "a:lvl" + lvl + "pPr"]);
                if (pPrNodeLaout === undefined) {
                    pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr"]);
                    if (pPrNodeLaout === undefined) {
                        pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvl - 1), "a:pPr"]);
                    }
                }
            }
            if (type !== undefined) {
                //slidelayout
                let lvlStr = "a:lvl" + lvl + "pPr";
                if (pPrNodeLaout === undefined) {
                    pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
                }
                //masterlayout
                if (type == "title" || type == "ctrTitle") {
                    pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr]);
                } else if (type == "body" || type == "obj" || type == "subTitle") {
                    pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr]);
                } else if (type == "shape" || type == "diagram") {
                    pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr]);
                } else if (type == "textBox") {
                    pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr]);
                } else {
                    pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
                }
            }
            return {
                "nodeLaout": pPrNodeLaout,
                "nodeMaster": pPrNodeMaster
            };
        }
    function getPregraphDir(node, textBodyNode, idx, type, warpObj) {
            let rtl = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "rtl"]);

            if (rtl === undefined) {
                let layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
                let pPrNodeLaout = layoutMasterNode.nodeLaout;
                let pPrNodeMaster = layoutMasterNode.nodeMaster;
                rtl = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
                if (rtl === undefined && type != "shape") {
                    rtl = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
                }
            }

            // 检查 bodyPr 的 rtlCol 属性（表格单元格和文本框的RTL设置）
            if (rtl === undefined && textBodyNode !== undefined) {
                let rtlCol = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rtlCol"]);
                if (rtlCol !== undefined) {
                    rtl = rtlCol;
                }
            }

            if (rtl == "1") {
                return "pregraph-rtl";
            } else if (rtl == "0") {
                return "pregraph-ltr";
            }
            return "pregraph-inherit";

            // var contentDir = PPTXStyleUtils.getContentDir(type, warpObj);
            // console.log("getPregraphDir node:", node["a:r"], "rtl:", rtl, "idx", idx, "type", type, "contentDir:", contentDir)

            // if (contentDir == "content"){
            //     return "pregraph-ltr";
            // } else if (contentDir == "content-rtl"){ 
            //     return "pregraph-rtl";
            // }
            // return "";
        }
    function getPregraphMargn(pNode, idx, type, isBullate, warpObj, fontSize){
            if (!isBullate){
                return ["",0];
            }
            let marLStr = "", maginVal = 0;
            let pPrNode = pNode["a:pPr"];
            let layoutMasterNode = getLayoutAndMasterNode(pNode, idx, type, warpObj);
            let pPrNodeLaout = layoutMasterNode.nodeLaout;
            let pPrNodeMaster = layoutMasterNode.nodeMaster;

            // 在 RTL 模式下，margin 和 indent 的语义保持不变
            // - marL (margin-left): 左边距（在 RTL 中是文本结束边的距离）
            // - indent: 缩进，负值表示从起始边向内缩进
            // 这些属性值直接转换为 CSS 的 padding，不需要根据 RTL 进行方向翻转
            let getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
            if (getRtlVal === undefined) {
                getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
                if (getRtlVal === undefined && type != "shape") {
                    getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
                }
            }

            //align
            let alignNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "algn"]); //"l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist
            if (alignNode === undefined) {
                alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
                if (alignNode === undefined) {
                    alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
                }
            }
            //indent?
            let indentNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "indent"]);
            if (indentNode === undefined) {
                indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
                if (indentNode === undefined) {
                    indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
                }
            }
            let indent = 0;
            if (indentNode !== undefined) {
                indent = parseInt(indentNode) * SLIDE_FACTOR$1;
            }
            //
            //marL
            let marLNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"]);
            if (marLNode === undefined) {
                marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
                if (marLNode === undefined) {
                    marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
                }
            }
            let marginLeft = 0;
            if (marLNode !== undefined) {
                marginLeft = parseInt(marLNode) * SLIDE_FACTOR$1;
            }
            if ((indentNode !== undefined || marLNode !== undefined)) {
                // let lvlIndent = defTabSz * lvl;
                // 无论 RTL 还是 LTR，marL 和 indent 都转换为 padding-left
                // 在 RTL 模式下，文本从右向左排列，但 padding-left 仍然是左边距
                // align="ctr" + marL > 0: 文本居中，但有左边距，导致整体偏右
                // align="ctr" + indent < 0: 文本居中，向左（文本起始方向）缩进，导致整体偏右
                marLStr = "padding-left: ";
                if (isBullate) {
                    maginVal = Math.abs(0 - indent);
                    // 减去项目符号数字的长度/大小，根据字体大小估算
                    let bulletSizeAdjustment = 0;
                    if (fontSize !== undefined) {
                        // 对于数字项目符号，根据字体大小估算宽度
                        bulletSizeAdjustment = fontSize * 0.8;
                    }
                    maginVal = Math.max(0, maginVal - bulletSizeAdjustment);
                    marLStr += maginVal + "px;";
                } else {
                    maginVal = Math.abs(marginLeft + indent);
                    marLStr += maginVal + "px;";
                }
            }

            //marR?
            let marRNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"]);
            if (marRNode === undefined && marLNode === undefined) {
                //need to check if this posble - TODO
                marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
                if (marRNode === undefined) {
                    marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
                }
            }


            return [marLStr, maginVal];
        }
// 提取图表标题样式
function extractChartTitleStyle(chartNode, warpObj) {
    const titleNode = PPTXXmlUtils.getTextByPathList(chartNode, ["c:title"]);
    if (!titleNode) return { text: "", style: {} };
    
    const style = {};
    let text = "";
    
    // 提取标题文本
    // 方法1: 尝试从富文本 (c:rich) 中提取
    const rich = PPTXXmlUtils.getTextByPathList(titleNode, ["c:tx", "c:rich"]);
    if (rich) {
        const p = PPTXXmlUtils.getTextByPathList(rich, ["a:p"]);
        if (p) {
            // 从段落中提取文本
            const r = PPTXXmlUtils.getTextByPathList(p, ["a:r"]);
            if (r) {
                // 可能有多个 run
                if (Array.isArray(r)) {
                    const textArray = r.map(run => PPTXXmlUtils.getTextByPathList(run, ["a:t"]));
                    text = textArray.filter(t => t).join('');
                } else {
                    text = PPTXXmlUtils.getTextByPathList(r, ["a:t"]);
                }
            }
            
            // 提取标题文本属性
            const pPr = PPTXXmlUtils.getTextByPathList(p, ["a:pPr"]);
            if (pPr) {
                const defRPr = PPTXXmlUtils.getTextByPathList(pPr, ["a:defRPr"]);
                if (defRPr) {
                    // 提取字体大小
                    if (defRPr["attrs"] && defRPr["attrs"]["sz"]) {
                        style.fontSize = parseFloat(defRPr["attrs"]["sz"]) / 100;
                    }
                    
                    // 提取字体粗细
                    if (defRPr["attrs"] && defRPr["attrs"]["b"] === "1") {
                        style.fontWeight = "bold";
                    }
                    
                    // 提取字体颜色
                    const solidFill = PPTXXmlUtils.getTextByPathList(defRPr, ["a:solidFill"]);
                    if (solidFill) {
                        let color = getColor(solidFill, undefined, undefined, warpObj);
                        if (color && !color.startsWith('#')) {
                            color = '#' + color;
                        }
                        style.color = color;
                    }
                }
            }
        }
    }
    
    // 方法2: 尝试从文本属性 (c:txPr) 中提取
    if (!text) {
        const txPr = PPTXXmlUtils.getTextByPathList(titleNode, ["c:txPr"]);
        if (txPr) {
            const p = PPTXXmlUtils.getTextByPathList(txPr, ["a:p"]);
            if (p) {
                // 从段落中提取文本
                const r = PPTXXmlUtils.getTextByPathList(p, ["a:r"]);
                if (r) {
                    // 可能有多个 run
                    if (Array.isArray(r)) {
                        const textArray = r.map(run => PPTXXmlUtils.getTextByPathList(run, ["a:t"]));
                        text = textArray.filter(t => t).join('');
                    } else {
                        text = PPTXXmlUtils.getTextByPathList(r, ["a:t"]);
                    }
                }
                
                // 提取标题文本属性
                const pPr = PPTXXmlUtils.getTextByPathList(p, ["a:pPr"]);
                if (pPr) {
                    const defRPr = PPTXXmlUtils.getTextByPathList(pPr, ["a:defRPr"]);
                    if (defRPr) {
                        // 提取字体大小
                        if (defRPr["attrs"] && defRPr["attrs"]["sz"]) {
                            style.fontSize = parseFloat(defRPr["attrs"]["sz"]) / 100;
                        }
                        
                        // 提取字体粗细
                        if (defRPr["attrs"] && defRPr["attrs"]["b"] === "1") {
                            style.fontWeight = "bold";
                        }
                        
                        // 提取字体颜色
                        const solidFill = PPTXXmlUtils.getTextByPathList(defRPr, ["a:solidFill"]);
                        if (solidFill) {
                            let color = getColor(solidFill, undefined, undefined, warpObj);
                            if (color && !color.startsWith('#')) {
                                color = '#' + color;
                            }
                            style.color = color;
                        }
                    }
                }
            }
        }
    }
    
    // 方法3: 如果没有找到文本，尝试从字符串引用中提取
    if (!text) {
        const tx = PPTXXmlUtils.getTextByPathList(titleNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]);
        if (tx) {
            text = tx;
        }
    }
    
    return { text, style };
}

// 提取图表区域样式
function extractChartAreaStyle(chartSpaceNode, warpObj) {
    const style = {};
    
    // 提取图表区域填充
    const spPr = PPTXXmlUtils.getTextByPathList(chartSpaceNode, ["c:spPr"]);
    if (spPr) {
        // 提取填充样式
        const fillType = getFillType(spPr);
        if (fillType === "SOLID_FILL") {
            const solidFill = PPTXXmlUtils.getTextByPathList(spPr, ["a:solidFill"]);
            if (solidFill) {
                let fillColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                if (fillColor && !fillColor.startsWith('#')) {
                    fillColor = '#' + fillColor;
                }
                style.fillColor = fillColor;
            }
        } else if (fillType === "GRADIENT_FILL") {
            const gradFill = PPTXXmlUtils.getTextByPathList(spPr, ["a:gradFill"]);
            if (gradFill) {
                style.gradientFill = getGradientFill(gradFill, warpObj);
            }
        }
        
        // 提取边框样式
        const ln = PPTXXmlUtils.getTextByPathList(spPr, ["a:ln"]);
        if (ln) {
            const solidFill = PPTXXmlUtils.getTextByPathList(ln, ["a:solidFill"]);
            if (solidFill) {
                let borderColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                if (borderColor && !borderColor.startsWith('#')) {
                    borderColor = '#' + borderColor;
                }
                style.borderColor = borderColor;
            }
            if (ln["attrs"] && ln["attrs"]["w"]) {
                style.borderWidth = parseFloat(ln["attrs"]["w"]) / 9525; // 转换为像素
            }
        }
    }
    
    return style;
}

// 提取图表图例样式
function extractChartLegendStyle(chartNode, warpObj) {
    const legendNode = PPTXXmlUtils.getTextByPathList(chartNode, ["c:legend"]);
    if (!legendNode) return {};
    
    const style = {};
    
    // 提取图例位置
    if (legendNode["c:legendPos"]) {
        style.position = legendNode["c:legendPos"]["attrs"]["val"];
    }
    
    // 提取图例文本属性
    const txPr = PPTXXmlUtils.getTextByPathList(legendNode, ["c:txPr"]);
    if (txPr) {
        const p = PPTXXmlUtils.getTextByPathList(txPr, ["a:p"]);
        if (p) {
            const pPr = PPTXXmlUtils.getTextByPathList(p, ["a:pPr"]);
            if (pPr) {
                const defRPr = PPTXXmlUtils.getTextByPathList(pPr, ["a:defRPr"]);
                if (defRPr) {
                    // 提取字体大小
                    if (defRPr["attrs"] && defRPr["attrs"]["sz"]) {
                        style.fontSize = parseFloat(defRPr["attrs"]["sz"]) / 100;
                    }
                    
                    // 提取字体颜色
                    const solidFill = PPTXXmlUtils.getTextByPathList(defRPr, ["a:solidFill"]);
                    if (solidFill) {
                        let color = getSolidFill(solidFill, undefined, undefined, warpObj);
                        if (color && !color.startsWith('#')) {
                            color = '#' + color;
                        }
                        style.color = color;
                    }
                }
            }
        }
    }
    
    return style;
}

// 提取图表轴样式
function extractChartAxisStyle(plotAreaNode, axisType, warpObj) {
    const axisNode = PPTXXmlUtils.getTextByPathList(plotAreaNode, [axisType]);
    if (!axisNode) return {};
    
    const style = {};
    
    // 提取轴文本属性
    const txPr = PPTXXmlUtils.getTextByPathList(axisNode, ["c:txPr"]);
    if (txPr) {
        const p = PPTXXmlUtils.getTextByPathList(txPr, ["a:p"]);
        if (p) {
            const pPr = PPTXXmlUtils.getTextByPathList(p, ["a:pPr"]);
            if (pPr) {
                const defRPr = PPTXXmlUtils.getTextByPathList(pPr, ["a:defRPr"]);
                if (defRPr) {
                    // 提取字体大小
                    if (defRPr["attrs"] && defRPr["attrs"]["sz"]) {
                        style.fontSize = parseFloat(defRPr["attrs"]["sz"]) / 100;
                    }
                    
                    // 提取字体颜色
                    const solidFill = PPTXXmlUtils.getTextByPathList(defRPr, ["a:solidFill"]);
                    if (solidFill) {
                        let color = getSolidFill(solidFill, undefined, undefined, warpObj);
                        if (color && !color.startsWith('#')) {
                            color = '#' + color;
                        }
                        style.color = color;
                    }
                }
            }
        }
    }
    
    // 提取轴线条样式
    const spPr = PPTXXmlUtils.getTextByPathList(axisNode, ["c:spPr"]);
    if (spPr) {
        const ln = PPTXXmlUtils.getTextByPathList(spPr, ["a:ln"]);
        if (ln) {
            const solidFill = PPTXXmlUtils.getTextByPathList(ln, ["a:solidFill"]);
            if (solidFill) {
                let lineColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                if (lineColor && !lineColor.startsWith('#')) {
                    lineColor = '#' + lineColor;
                }
                style.lineColor = lineColor;
            }
            if (ln["attrs"] && ln["attrs"]["w"]) {
                style.lineWidth = parseFloat(ln["attrs"]["w"]) / 9525; // 转换为像素
            }
        }
    }
    
    // 提取轴网格线样式
    if (axisType === "c:valAx") {
        const majorGridlines = PPTXXmlUtils.getTextByPathList(axisNode, ["c:majorGridlines"]);
        if (majorGridlines) {
            const spPr = PPTXXmlUtils.getTextByPathList(majorGridlines, ["c:spPr"]);
            if (spPr) {
                const ln = PPTXXmlUtils.getTextByPathList(spPr, ["a:ln"]);
                if (ln) {
                    const solidFill = PPTXXmlUtils.getTextByPathList(ln, ["a:solidFill"]);
                    if (solidFill) {
                        let gridlineColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                        if (gridlineColor && !gridlineColor.startsWith('#')) {
                            gridlineColor = '#' + gridlineColor;
                        }
                        style.gridlineColor = gridlineColor;
                    }
                    if (ln["attrs"] && ln["attrs"]["w"]) {
                        style.gridlineWidth = parseFloat(ln["attrs"]["w"]) / 9525; // 转换为像素
                    }
                }
            }
        }
    }
    
    return style;
}

// 辅助函数：获取颜色
function getColor(node, clrMap, phClr, warpObj) {
    if (node["a:solidFill"]) {
        return getSolidFill(node["a:solidFill"], clrMap, phClr, warpObj);
    }
    return "";
}

const PPTXStyleUtils = {
        getFillType,
        getShapeFill,
        getFontType,
        getFontColorPr,
        getFontSize,
        getFontBold,
        getFontItalic,
        getFontDecoration,
        getTextHorizontalAlign,
        getTextVerticalAlign,
        getTableBorders,
        getBorder,
        getSlideBackgroundFill,
        getBgGradientFill,
        getBgPicFill,
        getGradientFill,
        getPicFill,
        getPatternFill,
        getLinerGrandient,
        getSolidFill,
        toHex,
        hslToRgb,
        hueToRgb,
        getColorName2Hex,
        getSchemeColorFromTheme,
        getGradientFill,
        extractChartData,
        extractChartTitleStyle,
        extractChartAreaStyle,
        extractChartLegendStyle,
        extractChartAxisStyle,
        setTextByPathList,
        eachElement,
        applyShade,
        applyTint,
        applyLumOff,
        applyLumMod,
        applyHueMod,
        applySatMod,
        rgba2hex,
        getSvgGradient,
        getMiddleStops,
        SVGangle,
        getSvgImagePattern,
        getBase64ImageDimensions,
        getVerticalAlign,
        getContentDir,
        getHorizontalAlign,
        getVerticalMargins,
        getLayoutAndMasterNode,
        getPregraphDir,
        getPregraphMargn,
        getColor
    };

/**
 * Chart processing module
 * Handles chart generation and data processing
 */


/**
 * Generate chart HTML and data
 * @param {Object} node - Chart node
 * @param {Object} warpObj - Warp object containing context
 * @param {Object} parentNode - Parent node (for group elements coordinate calculation)
 * @returns {Promise<string>} Chart HTML
 */
async function genChart(node, warpObj, parentNode) {
    const order = node["attrs"]["order"];
    let xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);

    // 处理组合缩放 - 当chart在group-abs类型组合中时需要应用缩放
    let workingXfrmNode = xfrmNode;
    if (warpObj.currentGroupScale && xfrmNode) {
        const { scaleX, scaleY, childX, childY } = warpObj.currentGroupScale;

        // 创建缩放后的xfrmNode
        workingXfrmNode = JSON.parse(JSON.stringify(xfrmNode));

        // 缩放尺寸
        if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
            const originalCx = parseInt(xfrmNode['a:ext'].attrs.cx);
            const originalCy = parseInt(xfrmNode['a:ext'].attrs.cy);
            workingXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
            workingXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
        }

        // 调整位置(相对于childX/childY)
        if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
            const originalOffX = parseInt(xfrmNode['a:off'].attrs.x);
            const originalOffY = parseInt(xfrmNode['a:off'].attrs.y);

            // 计算相对于childOff的偏移
            const relativeX = originalOffX - (childX / SLIDE_FACTOR$1);
            const relativeY = originalOffY - (childY / SLIDE_FACTOR$1);

            // 应用缩放
            workingXfrmNode['a:off'].attrs.x = Math.round(childX / SLIDE_FACTOR$1 + relativeX * scaleX);
            workingXfrmNode['a:off'].attrs.y = Math.round(childY / SLIDE_FACTOR$1 + relativeY * scaleY);
        }
    }

    // 提取位置和尺寸信息
    let offX = 0, offY = 0, extCx = 0, extCy = 0;
    if (workingXfrmNode !== undefined) {
        if (workingXfrmNode['a:off'] && workingXfrmNode['a:off'].attrs) {
            offX = workingXfrmNode['a:off'].attrs.x || 0;
            offY = workingXfrmNode['a:off'].attrs.y || 0;
        }
        if (workingXfrmNode['a:ext'] && workingXfrmNode['a:ext'].attrs) {
            extCx = workingXfrmNode['a:ext'].attrs.cx || 0;
            extCy = workingXfrmNode['a:ext'].attrs.cy || 0;
        }
    }

    // 生成 data- 属性
    const dataAttrs = ` data-node-type="chart" data-off-x="${offX}" data-off-y="${offY}" data-ext-cx="${extCx}" data-ext-cy="${extCy}"`;

    const result = "<div id='chart" + warpObj.chartId.value + "' class='block content' style='" +
        PPTXXmlUtils.getPosition(workingXfrmNode, parentNode || node, undefined, undefined) + PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined) +
        ` z-index: ${order};'${dataAttrs}></div>`;

    const rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
    const refName = warpObj["slideResObj"][rid]["target"];
    const content = await PPTXXmlUtils.readXmlFile(warpObj["zip"], refName);
    // Guard: chart XML file may be missing or unreadable
    if (!content) {
        return result;
    }
    const chartSpace = PPTXXmlUtils.getTextByPathList(content, ["c:chartSpace"]);
    if (!chartSpace) {
        return result;
    }
    const chart = PPTXXmlUtils.getTextByPathList(chartSpace, ["c:chart"]);
    const plotArea = PPTXXmlUtils.getTextByPathList(chart, ["c:plotArea"]);

    // 提取3D视图属性
    const view3D = PPTXXmlUtils.getTextByPathList(chart, ["c:view3D"]);
    const view3DProps = {};
    if (view3D) {
        if (view3D["attrs"]?.rotX !== undefined) view3DProps.rotX = parseFloat(view3D["attrs"].rotX);
        if (view3D["attrs"]?.rotY !== undefined) view3DProps.rotY = parseFloat(view3D["attrs"].rotY);
        if (view3D["attrs"]?.depthPercent !== undefined) view3DProps.depthPercent = parseFloat(view3D["attrs"].depthPercent);
        if (view3D["attrs"]?.rAngAx !== undefined) view3DProps.rAngAx = view3D["attrs"].rAngAx === "1";
    }

    // 提取图表类型特定属性
    const chartType = Object.keys(plotArea).find(key => key.startsWith('c:') && key.endsWith('Chart'));
    const varyColors = chartType ? PPTXXmlUtils.getTextByPathList(plotArea[chartType], ["c:varyColors", "attrs", "val"]) : undefined;

    // 提取系列数据点的样式（dPt）和爆炸效果（explosion）
    let dataPointStyles = [];
    if (chartType && plotArea[chartType]["c:ser"]) {
        const serArray = Array.isArray(plotArea[chartType]["c:ser"]) 
            ? plotArea[chartType]["c:ser"] 
            : [plotArea[chartType]["c:ser"]];
        
        serArray.forEach(ser => {
            const dPtArray = ser["c:dPt"];
            if (dPtArray) {
                const dpStyles = {};
                const dpList = Array.isArray(dPtArray) ? dPtArray : [dPtArray];
                dpList.forEach(dp => {
                    const idx = dp["c:idx"]?.["attrs"]?.val;
                    const explosion = dp["c:explosion"]?.["attrs"]?.val;
                    const spPr = dp["c:spPr"];
                    
                    if (idx !== undefined) {
                        const dpStyle = {};
                        if (explosion !== undefined) {
                            dpStyle.explosion = parseFloat(explosion);
                        }
                        if (spPr) {
                            const gradFill = spPr["a:gradFill"];
                            if (gradFill) {
                                dpStyle.gradientFill = PPTXStyleUtils.getGradientFill(gradFill, warpObj);
                            }
                        }
                        dpStyles[idx] = dpStyle;
                    }
                });
                dataPointStyles.push(dpStyles);
            }
        });
    }

    // 提取图表标题
    const chartTitleObj = PPTXStyleUtils.extractChartTitleStyle(chart, warpObj);
    const chartTitle = chartTitleObj.text;

    // 提取图表样式信息
    const chartStyle = {
        chartArea: PPTXStyleUtils.extractChartAreaStyle(chartSpace, warpObj),
        legend: PPTXStyleUtils.extractChartLegendStyle(chart, warpObj),
        categoryAxis: PPTXStyleUtils.extractChartAxisStyle(plotArea, "c:catAx", warpObj),
        valueAxis: PPTXStyleUtils.extractChartAxisStyle(plotArea, "c:valAx", warpObj),
        view3D: view3DProps,
        varyColors: varyColors === "1",
        dataPointStyles: dataPointStyles,
        title: chartTitleObj.style
    };

    let chartData = null;
    for (const key in plotArea) {
        switch (key) {
            case "c:lineChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "lineChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:barChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "barChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:pieChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "pieChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:pie3DChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "pie3DChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:areaChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "areaChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:scatterChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "scatterChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
        }
    }

    return result;
}

/**
 * Process message queue for charts
 * @param {Array} queue - Message queue
 * @param {Object} result - Result object to store chart data
 */
function processMsgQueue(queue, result) {
    for (const msg of queue) {
        if (msg.type === "chart" || msg.type === "createChart") {
            const chartObj = msg.data;
            result.charts.push({
                chartId: chartObj.chartId,
                type: chartObj.chartType,
                data: chartObj.chartData,
                style: chartObj.style,
                title: chartObj.title
            });
        }
    }
}

/**
 * 文本处理模块
 * 
 * 处理 PPTX 中的文本内容，包括：
 * - 文本样式解析（字体、大小、颜色、对齐等）
 * - 段落和文本运行处理
 * - 项目符号和编号
 * - 超链接处理
 * - 文本宽度计算
 * - RTL（从右到左）语言支持
 * 
 * @module utils/text
 */


// 创建 tinycolor 工厂函数以保持向后兼容
const tinycolor = (color, opts) => new TinyColor(color, opts);
let is_first_br = false;



function getTextWidth(html) {
        let div = document.createElement('div');
        div.style.position = 'absolute';
        div.style.float = 'left';
        div.style.whiteSpace = 'nowrap';
        div.style.visibility = 'hidden';
        div.innerHTML = html;
        document.body.appendChild(div);
        let width = div.offsetWidth;
        document.body.removeChild(div);
        return width;
    }

    async function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, tbl_col_width) {
            let text = "";
            warpObj["slideMasterTextStyles"];

            if (textBodyNode === undefined) {
                return text;
            }
            //rtl : <p:txBody>
            //          <a:bodyPr wrap="square" rtlCol="1">

            // 获取anchor属性（垂直对齐方式）
            let anchor = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "anchor"]);
            if (anchor === undefined) {
                anchor = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                if (anchor === undefined) {
                    anchor = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                    if (anchor === undefined) {
                        anchor = "t";
                    }
                }
            }

            // 对于圆形/椭圆类形状，强制使用居中对齐以确保文本在形状内正确居中显示
            if (type === "shape") {
                let shapeType = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
                const circularShapes = [
                    "ellipse", "ovalCallout", "wedgeEllipseCallout",
                    "pie", "pieWedge", "chord", "sector", "arc", "blockArc"
                ];
                if (circularShapes.includes(shapeType) && anchor === "t") {
                    anchor = "ctr";
                }
            }

            // 获取bodyPr的内边距设置
            let bodyPrPadding = getBodyPrPadding(textBodyNode, type, anchor);
            text += bodyPrPadding;

            let pFontStyle = PPTXXmlUtils.getTextByPathList(spNode, ["p:style", "a:fontRef"]);
            let wrapAttr = PPTXXmlUtils.getTextByPathList(textBodyNode["a:bodyPr"], ["attrs", "wrap"]);
            let spAutoFitNode = PPTXXmlUtils.getTextByPathList(textBodyNode["a:bodyPr"], ["a:spAutoFit"]);
            PPTXXmlUtils.getTextByPathList(textBodyNode["a:bodyPr"], ["attrs", "rtlCol"]);
            let isNoWrap = (wrapAttr === "none");
            let isAutoFit = (spAutoFitNode !== undefined);

            
            let apNode = textBodyNode["a:p"];
            if (apNode.constructor !== Array) {
                apNode = [apNode];
            }

            for (let i = 0; i < apNode.length; i++) {
                let pNode = apNode[i];
                let rNode = pNode["a:r"];
                let fldNode = pNode["a:fld"];
                let brNode = pNode["a:br"];
                if (rNode !== undefined) {
                    rNode = (rNode.constructor === Array) ? rNode : [rNode];
                }
                if (rNode !== undefined && fldNode !== undefined) {
                    fldNode = (fldNode.constructor === Array) ? fldNode : [fldNode];
                    rNode = rNode.concat(fldNode);
                }
                if (rNode !== undefined && brNode !== undefined) {
                    is_first_br = true;
                    brNode = (brNode.constructor === Array) ? brNode : [brNode];
                    brNode.forEach((item, indx) => {
                        item.type = "br";
                    });
                    if (brNode.length > 1) {
                        brNode.shift();
                    }
                    rNode = rNode.concat(brNode);
                    rNode.sort((a, b) => {
                        return a.attrs.order - b.attrs.order;
                    });
                }
                //rtlStr = "";//`dir='${isRTL}'`;
                let styleText = "";
                let marginsVer = PPTXStyleUtils.getVerticalMargins(pNode, textBodyNode, type, idx, warpObj, apNode.length, i, anchor);
                if (marginsVer != "") {
                    styleText = marginsVer;
                }
                // 移除 font-size: 0px 设置，避免影响文本显示
                // if (type == "body" || type == "obj" || type == "shape") {
                //     styleText += "font-size: 0px;";
                //     //styleText += "line-height: 0;";
                //     styleText += "font-weight: 100;";
                //     styleText += "font-style: normal;";
                // }
                let cssName = "";

                if (styleText in warpObj.styleTable) {
                    cssName = warpObj.styleTable[styleText]["name"];
                } else {
                    cssName = "_css_" + (Object.keys(warpObj.styleTable).length + 1);
                    warpObj.styleTable[styleText] = {
                        "name": cssName,
                        "text": styleText
                    };
                }

                let prg_width_node = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
                // 占位符可能没有自己的xfrm（<p:spPr/>为空），尺寸继承自版式/母版中的同名占位符，需回退获取
                if (prg_width_node === undefined || prg_width_node === null) {
                    prg_width_node = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
                }
                if (prg_width_node === undefined || prg_width_node === null) {
                    prg_width_node = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
                }
                
                // 获取bodyPr的内边距属性，用于计算可用宽度
                let lIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "lIns"]);
                let rIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rIns"]);
                            
                // 计算内边距像素值
                let lInsPx, rInsPx;
                if (type === "table") {
                    // 对于表格，如果没有明确设置内边距，则使用0
                    lInsPx = lIns ? (parseInt(lIns) * SLIDE_FACTOR$1) : 0;
                    rInsPx = rIns ? (parseInt(rIns) * SLIDE_FACTOR$1) : 0;
                } else {
                    lInsPx = lIns ? (parseInt(lIns) * SLIDE_FACTOR$1) : (type === "diagram" ? 0.04 * 96 : 0.1 * 96);
                    rInsPx = rIns ? (parseInt(rIns) * SLIDE_FACTOR$1) : (type === "diagram" ? 0.04 * 96 : 0.1 * 96);
                }
                            
                // 如果明确设置了lIns="0"或rIns="0"，则不应用默认值
                if (lIns === "0") lInsPx = 0;
                if (rIns === "0") rInsPx = 0;
                            
                // 检查是否为圆形/椭圆类形状，如果是则应用额外的安全边距
                let shapeType = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
                const circularShapes = [
                    "ellipse", "ovalCallout", "wedgeEllipseCallout",
                    "pie", "pieWedge", "chord", "sector", "arc", "blockArc"
                ];
                const isCircularShape = circularShapes.includes(shapeType);
                
                // 处理组合缩放 - 如果形状在group-abs组合中,需要应用缩放到段落宽度
                let sld_prg_width_val = null;
                if (prg_width_node !== undefined && prg_width_node !== null) {
                    let parsedWidth = parseInt(prg_width_node);
                    if (!isNaN(parsedWidth) && parsedWidth > 0) {
                        sld_prg_width_val = Math.round(parsedWidth * SLIDE_FACTOR$1 * 100) / 100;
                    }
                }
                if (sld_prg_width_val !== null && warpObj.currentGroupScale) {
                    const { scaleX, scaleY } = warpObj.currentGroupScale;
                    sld_prg_width_val = Math.round(sld_prg_width_val * scaleX * 100) / 100;
                }
                
                let sld_prg_width = "";
                if (sld_prg_width_val !== null && !isNoWrap) {
                    // 减去内边距宽度，得到实际可用宽度
                    let availableWidth = sld_prg_width_val - lInsPx - rInsPx;
                                
                    // 对于圆形/椭圆类形状，应用额外的安全边距（减少5%）以确保正确换行
                    if (isCircularShape) {
                        availableWidth = availableWidth * 0.95;
                    }
                                
                    sld_prg_width = "width:" + Math.max(0, Math.round(availableWidth * 100) / 100) + "px;";
                } else if (sld_prg_width_val === null) {
                    sld_prg_width = "width:inherit;";
                }
                let sld_prg_height = ""; // 移除高度设置，避免段落叠加
                let prg_dir = PPTXStyleUtils.getPregraphDir(pNode, textBodyNode, idx, type, warpObj);
                let isRTL = (prg_dir == "pregraph-rtl");
                let directionStyle = isRTL ? "direction: rtl;" : "direction: ltr;";
                let horizontalAlign = PPTXStyleUtils.getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir, warpObj, spNode);
                // 在外层div上也设置justify-content，确保对齐正确
                // 注意：表格单元格的对齐由td元素的text-align控制，这里不设置justify-content
                let outerFlexStyle = "";
                if (type !== "table") {
                    if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                        // 如果是 RTL，则 flex-start 是右对齐；如果是 LTR，则 flex-end 是右对齐
                        outerFlexStyle = isRTL ? "justify-content: flex-start;" : "justify-content: flex-end;";
                    } else if (horizontalAlign === "h-mid") {
                        outerFlexStyle = "justify-content: center;";
                    } else if (horizontalAlign === "h-left-rtl") {
                        // RTL 模式下，左对齐使用 flex-end
                        outerFlexStyle = "justify-content: flex-end;";
                    } else {
                        outerFlexStyle = "justify-content: flex-start;";
                    }
                }
                text += "<div style='display: flex;" + sld_prg_width + sld_prg_height + outerFlexStyle + directionStyle + "' class='slide-prgrph " + horizontalAlign + ` ${prg_dir} ` + cssName + "' >";
                let buText_ary = await genBuChar(pNode, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj);
                let isBullate = (buText_ary[0] !== undefined && buText_ary[0] !== null && buText_ary[0] != "" ) ? true : false;
                let bu_width = (buText_ary[1] !== undefined && buText_ary[1] !== null && isBullate) ? (Number(buText_ary[1]) + Number(buText_ary[2])) : 0;

                // 在 RTL 模式下，项目符号在右边，所以先添加文本，再添加项目符号
                if (isRTL && isBullate) ; else {
                    text += (buText_ary[0] !== undefined) ? buText_ary[0]:"";
                }
                //get text margin 
                // 获取段落的字体大小，用于计算项目符号边距
                let fontSize = undefined;
                if (rNode !== undefined && rNode.length > 0) {
                    // 使用第一个文本运行的字体大小作为参考
                    fontSize = PPTXStyleUtils.getFontSize(rNode[0], textBodyNode, pFontStyle, 1, type, warpObj);
                    if (fontSize && fontSize.endsWith('px')) {
                        fontSize = parseFloat(fontSize);
                    }
                }
                let margin_ary = PPTXStyleUtils.getPregraphMargn(pNode, idx, type, isBullate, warpObj, fontSize);
                let margin = margin_ary[0];
                let mrgin_val = margin_ary[1];
                if (prg_width_node === undefined && tbl_col_width !== undefined && prg_width_node != 0){
                    //sorce : table text
                    prg_width_node = tbl_col_width;
                }

                let prgrph_text = "";
                //let prgr_txt_art = [];
                let total_text_len = 0;
                

                if (rNode === undefined && pNode !== undefined) {
                    // without r
                    let prgr_text = await genSpanElement(pNode, undefined, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, 1, warpObj);
                    if (isBullate) {
                        total_text_len += getTextWidth(prgr_text);
                    }
                    prgrph_text += prgr_text;
                } else if (rNode !== undefined) {
                    // with multi r
                    let previousStyle = {};
                    for (let j = 0; j < rNode.length; j++) {
                        // 如果当前元素没有sz属性，使用前面元素的样式
                        if (rNode[j]["a:rPr"] && !rNode[j]["a:rPr"]["attrs"] && previousStyle["sz"]) {
                            rNode[j]["a:rPr"]["attrs"] = { "sz": previousStyle["sz"] };
                        } else if (rNode[j]["a:rPr"] && rNode[j]["a:rPr"]["attrs"] && !rNode[j]["a:rPr"]["attrs"]["sz"] && previousStyle["sz"]) {
                            rNode[j]["a:rPr"]["attrs"]["sz"] = previousStyle["sz"];
                        }
                        
                        let prgr_text = await genSpanElement(rNode[j], j, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNode.length, warpObj);
                        if (isBullate) {
                            total_text_len += getTextWidth(prgr_text);
                        }
                        

                        prgrph_text += prgr_text;
                        
                        // 保存当前元素的样式，供后面元素继承
                        if (rNode[j]["a:rPr"] && rNode[j]["a:rPr"]["attrs"] && rNode[j]["a:rPr"]["attrs"]["sz"]) {
                            previousStyle["sz"] = rNode[j]["a:rPr"]["attrs"]["sz"];
                        }
                    }
                }

                prg_width_node = parseInt(prg_width_node) * SLIDE_FACTOR$1 - bu_width - mrgin_val;
                prg_width_node = Math.round(prg_width_node * 100) / 100;
                let textContainerWidth = ""; // 默认不设置宽度，让文本容器自适应
                // 如果是 noAutofit（文本框宽度固定），需要设置文本容器宽度以限制换行
                // 但对于表格，不设置内层容器宽度，让其自适应
                // 只有拿到了有效的文本框宽度才设置内层容器宽度，避免 width:NaNpx/0px 导致逐字换行（视觉上变竖排）
                if (!isAutoFit && !isNoWrap && sld_prg_width_val !== null && !isNaN(sld_prg_width_val) && type !== "table") {
                    // 使用与外层段落相同的可用宽度
                    let availableWidthForTextContainer = sld_prg_width_val - lInsPx - rInsPx;
                    // 对于圆形/椭圆类形状，应用额外的安全边距
                    if (isCircularShape) {
                        availableWidthForTextContainer = availableWidthForTextContainer * 0.95;
                    }
                    textContainerWidth = "width:" + Math.max(0, Math.round(availableWidthForTextContainer * 100) / 100) + "px;";
                }
                if (isRTL && isBullate) {
                    // RTL 模式下有项目符号时，文本容器不设宽度，让内容自适应
                    textContainerWidth = "";
                }
                // 对于圆形/椭圆类形状，使用normal white-space和break-word以确保正确换行
                let whiteSpaceStyle;
                if (isCircularShape) {
                    whiteSpaceStyle = isNoWrap ? "white-space: nowrap;" : "white-space: normal; overflow-wrap: break-word;";
                } else {
                    whiteSpaceStyle = isNoWrap ? "white-space: nowrap;" : "white-space: pre-wrap;";
                }


                let textAlignStyle = "";
                if (horizontalAlign === "h-mid") {
                    textAlignStyle = "text-align: center;";
                } else if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                    textAlignStyle = "text-align: right;";
                } else if (horizontalAlign === "h-left-rtl") {
                    textAlignStyle = "text-align: left;";
                } else {
                    textAlignStyle = "text-align: left;";
                }
                
                // 添加强制换行样式以确保中文文本正确换行
                textAlignStyle += " word-break: break-all;";
                // 为了确保右对齐生效，添加flex布局的justify-content属性
                // 注意：表格单元格的对齐由td元素的text-align控制，这里不设置justify-content
                let flexStyle = "";
                if (type !== "table") {
                    if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                        // 如果是 RTL，则 flex-start 是右对齐；如果是 LTR，则 flex-end 是右对齐
                        flexStyle = isRTL ? "justify-content: flex-start;" : "justify-content: flex-end;";
                    } else if (horizontalAlign === "h-mid") {
                        flexStyle = "justify-content: center;";
                    } else if (horizontalAlign === "h-left-rtl") {
                        // RTL 模式下，左对齐使用 flex-end
                        flexStyle = "justify-content: flex-end;";
                    } else {
                        flexStyle = "justify-content: flex-start;";
                    }
                }
                text += "<div style='display: flex;" + flexStyle + textContainerWidth + directionStyle + "'>";
                // 在 RTL 模式下，项目符号应该和文本在同一个容器中
                if (isRTL && isBullate && buText_ary[0] !== undefined) {
                    // 先添加项目符号，再添加文本（在 RTL 容器中，第一个子元素显示在最右边）
                    text += buText_ary[0];
                }
                text += "<div style='" + styleText + directionStyle + whiteSpaceStyle + margin + textAlignStyle + "'>";
                text += prgrph_text;
                text += "</div>";
                text += "</div>";
                text += "</div>";
            }

            // 关闭bodyPr内边距div（如果存在）
            // 与 getBodyPrPadding 中的条件保持一致：type !== "table"
            if (type !== "table") {
                text += "</div>";
            }

            return text;
    }

    /**
     * 获取bodyPr的内边距设置
     * @param {Object} textBodyNode - 文本体节点
     * @param {string} type - 形状类型
     * @param {string} anchor - 垂直对齐方式（t=顶部, ctr=居中, b=底部）
     * @returns {string} CSS padding字符串
     */
    function getBodyPrPadding(textBodyNode, type, anchor) {
        let paddingStyle = "";

        // 获取bodyPr的各个内边距属性
        let lIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "lIns"]);
        let tIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "tIns"]);
        let rIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rIns"]);
        let bIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "bIns"]);

        // 文本框、形状、diagram 或普通对象时都应用内边距
        // 注意：type 可能为 "textBox", "shape", "diagram", "obj" 或 undefined
        // 只要不是 table 类型，就应该应用内边距
        if (type !== "table") {
            // 根据PPTX规范，bodyPr的ins属性单位是EMU（English Metric Units）
            // 1 inch = 914400 EMU, 1 inch = 96px, 所以 1 EMU = 96/914400 px ≈ 0.000105 px
            // 如果没有设置内边距，使用默认值：
            // - diagram 类型使用更小的默认值，使其更接近原始 PPT 效果
            // - 其他类型：tIns 和 bIns 默认值为 0.05 inch = 45720 EMU ≈ 4.8px
            // - 其他类型：lIns 和 rIns 默认值为 0.1 inch = 91440 EMU ≈ 9.6px
            let defaultLIns, defaultTIns, defaultRIns, defaultBIns;
            if (type === "diagram") {
                // diagram 类型使用更小的默认内边距
                defaultLIns = 0.04 * 96;  // 0.04 inch ≈ 3.84px
                defaultTIns = 0.02 * 96;  // 0.02 inch ≈ 1.92px
                defaultRIns = 0.04 * 96;  // 0.04 inch ≈ 3.84px
                defaultBIns = 0.02 * 96;  // 0.02 inch ≈ 1.92px
            } else {
                defaultLIns = 0.1 * 96;   // 0.1 inch ≈ 9.6px
                defaultTIns = 0.05 * 96;  // 0.05 inch ≈ 4.8px
                defaultRIns = 0.1 * 96;   // 0.1 inch ≈ 9.6px
                defaultBIns = 0.05 * 96;  // 0.05 inch ≈ 4.8px
            }

            let lInsPx = lIns ? (parseInt(lIns) * SLIDE_FACTOR$1).toFixed(2) : defaultLIns.toFixed(2);
            let tInsPx = tIns ? (parseInt(tIns) * SLIDE_FACTOR$1).toFixed(2) : defaultTIns.toFixed(2);
            let rInsPx = rIns ? (parseInt(rIns) * SLIDE_FACTOR$1).toFixed(2) : defaultRIns.toFixed(2);
            let bInsPx = bIns ? (parseInt(bIns) * SLIDE_FACTOR$1).toFixed(2) : defaultBIns.toFixed(2);

            // 如果明确设置了lIns="0"或rIns="0"，则不应用默认值
            if (lIns === "0") lInsPx = "0";
            if (rIns === "0") rInsPx = "0";

            // 根据anchor决定padding div的高度设置
            // 如果是垂直居中（anchor="ctr"），不设置height: 100%，让内容自然撑开
            // 这样外层的v-mid类的justify-content: center才能生效
            let heightStyle = "";
            if (anchor !== "ctr") {
                heightStyle = "height: 100%;";
            }

            paddingStyle = `<div style="padding: ${tInsPx}px ${rInsPx}px ${bInsPx}px ${lInsPx}px; box-sizing: border-box; ${heightStyle}">`;
        }

        return paddingStyle;
    }
        
        async function genBuChar(node, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj) {

            ///////////////////////////////////////Amir///////////////////////////////
            warpObj["slideMasterTextStyles"];
            let lstStyle = textBodyNode["a:lstStyle"];

            let rNode = PPTXXmlUtils.getTextByPathList(node, ["a:r"]);
            if (rNode !== undefined && rNode.constructor === Array) {
                rNode = rNode[0]; //bullet only to first "a:r"
            }

            let lvl = parseInt (PPTXXmlUtils.getTextByPathList(node["a:pPr"], ["attrs", "lvl"])) + 1;
            if (isNaN(lvl)) {
                lvl = 1;
            }
            let lvlStr = `a:lvl${lvl}pPr`;
            let dfltBultColor, dfltBultSize, bultColor, bultSize, color_tye;

            if (rNode !== undefined) {
                dfltBultColor = await PPTXStyleUtils.getFontColorPr(rNode, spNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
                color_tye = dfltBultColor[2];
                dfltBultSize = PPTXStyleUtils.getFontSize(rNode, textBodyNode, pFontStyle, lvl, type, warpObj);

            } else {
                return "";
            }


            let bullet = "", marRStr = "", marLStr = "", margin_val=0, font_val=0;
            /////////////////////////////////////////////////////////////////


            let pPrNode = node["a:pPr"];
            let BullNONE = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buNone"]);
            if (BullNONE !== undefined) {
                return "";
            }

            let buType = "TYPE_NONE";

            let layoutMasterNode = PPTXStyleUtils.getLayoutAndMasterNode(node, idx, type, warpObj);
            let pPrNodeLaout = layoutMasterNode.nodeLaout;
            let pPrNodeMaster = layoutMasterNode.nodeMaster;

            let buChar = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
            let buNum = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buAutoNum", "attrs", "type"]);
            let buPic = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buBlip"]);
            if (buChar !== undefined) {
                buType = "TYPE_BULLET";
            }
            if (buNum !== undefined) {
                buType = "TYPE_NUMERIC";
            }
            if (buPic !== undefined) {
                buType = "TYPE_BULPIC";
            }

            let buFontSize = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buSzPts", "attrs", "val"]);
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buSzPct", "attrs", "val"]);
                if (buFontSize !== undefined) {
                    let prcnt = parseInt(buFontSize) / 100000;
                    //dfltBultSize = XXpt
                    //let dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                    let dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                    bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                }
            } else {
                bultSize = (parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR + "px";
            }

            //get definde bullet COLOR
            let buClrNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buClr"]);


            if (buChar === undefined && buNum === undefined && buPic === undefined) {

                if (lstStyle !== undefined) {
                    BullNONE = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buNone"]);
                    if (BullNONE !== undefined) {
                        return "";
                    }
                    buType = "TYPE_NONE";
                    buChar = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buChar", "attrs", "char"]);
                    buNum = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buAutoNum", "attrs", "type"]);
                    buPic = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buBlip"]);
                    if (buChar !== undefined) {
                        buType = "TYPE_BULLET";
                    }
                    if (buNum !== undefined) {
                        buType = "TYPE_NUMERIC";
                    }
                    if (buPic !== undefined) {
                        buType = "TYPE_BULPIC";
                    }
                    if (buChar !== undefined || buNum !== undefined || buPic !== undefined) {
                        pPrNode = lstStyle[lvlStr];
                    }
                }
            }
            if (buChar === undefined && buNum === undefined && buPic === undefined) {
                //check in slidelayout and masterlayout - TODO
                if (pPrNodeLaout !== undefined) {
                    BullNONE = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buNone"]);
                    if (BullNONE !== undefined) {
                        return "";
                    }
                    buType = "TYPE_NONE";
                    buChar = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buChar", "attrs", "char"]);
                    buNum = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buAutoNum", "attrs", "type"]);
                    buPic = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buBlip"]);
                    if (buChar !== undefined) {
                        buType = "TYPE_BULLET";
                    }
                    if (buNum !== undefined) {
                        buType = "TYPE_NUMERIC";
                    }
                    if (buPic !== undefined) {
                        buType = "TYPE_BULPIC";
                    }
                }
                if (buChar === undefined && buNum === undefined && buPic === undefined) {
                    //masterlayout

                    if (pPrNodeMaster !== undefined) {
                        BullNONE = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buNone"]);
                        if (BullNONE !== undefined) {
                            return "";
                        }
                        buType = "TYPE_NONE";
                        buChar = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buChar", "attrs", "char"]);
                        buNum = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buAutoNum", "attrs", "type"]);
                        buPic = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buBlip"]);
                        if (buChar !== undefined) {
                            buType = "TYPE_BULLET";
                        }
                        if (buNum !== undefined) {
                            buType = "TYPE_NUMERIC";
                        }
                        if (buPic !== undefined) {
                            buType = "TYPE_BULPIC";
                        }
                    }

                }

            }
            //rtl
            let getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
            if (getRtlVal === undefined) {
                getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
                if (getRtlVal === undefined && type != "shape") {
                    getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
                }
            }
            let isRTL = false;
            if (getRtlVal !== undefined && getRtlVal == "1") {
                isRTL = true;
            }
            //align
            let alignNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "algn"]); //"l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist
            if (alignNode === undefined) {
                alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
                if (alignNode === undefined) {
                    alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
                }
            }
            //indent?
            let indentNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "indent"]);
            if (indentNode === undefined) {
                indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
                if (indentNode === undefined) {
                    indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
                }
            }
            let indent = 0;
            if (indentNode !== undefined) {
                indent = parseInt(indentNode) * SLIDE_FACTOR$1;
            }
            //marL
            let marLNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"]);
            if (marLNode === undefined) {
                marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
                if (marLNode === undefined) {
                    marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
                }
            }

            if (marLNode !== undefined) {
                let marginLeft = parseInt(marLNode) * SLIDE_FACTOR$1;
                if (isRTL) {// && alignNode == "r") {
                    marLStr = "padding-right:";// "margin-right: ";
                } else {
                    marLStr = "padding-left:";//"margin-left: ";
                }
                margin_val = ((marginLeft + indent < 0) ? 0 : (marginLeft + indent));
                marLStr += margin_val + "px;";
            }
            
            //marR?
            let marRNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"]);
            if (marRNode === undefined && marLNode === undefined) {
                //need to check if this posble - TODO
                marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
                if (marRNode === undefined) {
                    marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
                }
            }
            if (marRNode !== undefined) {
                let marginRight = parseInt(marRNode) * SLIDE_FACTOR$1;
                if (isRTL) {// && alignNode == "r") {
                    marLStr = "padding-right:";// "margin-right: ";
                } else {
                    marLStr = "padding-left:";//"margin-left: ";
                }
                marRStr += ((marginRight + indent < 0) ? 0 : (marginRight + indent)) + "px;";
            }

            //get definde bullet COLOR
            if (buClrNode === undefined){
                //lstStyle
                buClrNode = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr, "a:buClr"]);
            }
            if (buClrNode === undefined) {
                buClrNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buClr"]);
                if (buClrNode === undefined) {
                    buClrNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buClr"]);
                }
            }
            let defBultColor;
            if (buClrNode !== undefined) {
                defBultColor = PPTXStyleUtils.getSolidFill(buClrNode, undefined, undefined, warpObj);
            }
            if (defBultColor === undefined || defBultColor == "NONE") {
                bultColor = dfltBultColor;
            } else {
                bultColor = [defBultColor, "", "solid"];
                color_tye = "solid";
            }


            //get definde bullet SIZE
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPts", "attrs", "val"]);
                if (buFontSize === undefined) {
                    buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPct", "attrs", "val"]);
                    if (buFontSize !== undefined) {
                        let prcnt = parseInt(buFontSize) / 100000;
                        //let dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                        let dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                        bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                    }
                }else {
                    bultSize = (parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR + "px";
                }
            }
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPts", "attrs", "val"]);
                if (buFontSize === undefined) {
                    buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPct", "attrs", "val"]);
                    if (buFontSize !== undefined) {
                        let prcnt = parseInt(buFontSize) / 100000;
                        //dfltBultSize = XXpt
                        //let dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                        let dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                        bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                    }
                } else {
                    bultSize = (parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR + "px";
                }
            }
            if (buFontSize === undefined) {
                bultSize = dfltBultSize;
            }
            font_val = parseInt(bultSize, "px");
            ////////////////////////////////////////////////////////////////////////
            if (buType == "TYPE_BULLET") {
                let typefaceNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs", "typeface"]);
                let typeface = "";
                let isWingdingsFont = false;
                if (typefaceNode !== undefined) {
                    isWingdingsFont = (typefaceNode == "Wingdings" || typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3" || typefaceNode == "Webdings");
                    typeface = "font-family: " + typefaceNode;
                }
                // let marginLeft = parseInt (PPTXXmlUtils.getTextByPathList(marLNode)) * SLIDE_FACTOR;
                // let marginRight = parseInt (PPTXXmlUtils.getTextByPathList(marRNode)) * SLIDE_FACTOR;
                // if (isNaN(marginLeft)) {
                //     marginLeft = 328600 * SLIDE_FACTOR;
                // }
                // if (isNaN(marginRight)) {
                //     marginRight = 0;
                // }

                bullet = `<div style='${typeface};` +
                    marLStr + marRStr +
                    `font-size:${bultSize};` ;
                
                //bullet += "display: table-cell;";
                //"line-height: 0px;";
                if (color_tye == "solid") {
                    if (bultColor[0] !== undefined && bultColor[0] != "") {
                        let bulletColorValue = bultColor[0];
                        if (bulletColorValue.length === 8) {
                            let colorObj = tinycolor(bulletColorValue);
                            bulletColorValue = colorObj.toRgbString();
                        } else {
                            bulletColorValue = "#" + bulletColorValue;
                        }
                        bullet += "color:" + bulletColorValue + "; ";
                    }
                    if (bultColor[1] !== undefined && bultColor[1] != "" && bultColor[1] != ";") {
                        bullet += "text-shadow:" + bultColor[1] + ";";
                    }
                    //no highlight/background-color to bullet
                    // if (bultColor[3] !== undefined && bultColor[3] != "") {
                    //     styleText += "background-color: #" + bultColor[3] + ";";
                    // }
                } else if (color_tye == "pattern" || color_tye == "pic" || color_tye == "gradient") {
                    if (color_tye == "pattern") {
                        bullet += "background:" + bultColor[0][0] + ";";
                        if (bultColor[0][1] !== null && bultColor[0][1] !== undefined && bultColor[0][1] != "") {
                            bullet += "background-size:" + bultColor[0][1] + ";";//" 2px 2px;" +
                        }
                        if (bultColor[0][2] !== null && bultColor[0][2] !== undefined && bultColor[0][2] != "") {
                            bullet += "background-position:" + bultColor[0][2] + ";";//" 2px 2px;" +
                        }
                        // bullet += "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "color: transparent;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";" +
                        //     "filter: " + bultColor[1].effcts + ";";
                    } else if (color_tye == "pic") {
                        bullet += bultColor[0] + ";";
                        // bullet += "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "color: transparent;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";";

                    } else if (color_tye == "gradient") {

                        let colorAry = bultColor[0].color;
                        let rot = bultColor[0].rot;

                        bullet += `background: linear-gradient(${rot}deg,`;
                        for (let i = 0; i < colorAry.length; i++) {
                            if (i == colorAry.length - 1) {
                                bullet += "#" + colorAry[i] + ");";
                            } else {
                                bullet += "#" + colorAry[i] + ", ";
                            }
                        }
                        // bullet += "color: transparent;" +
                        //     "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";";
                    }
                    bullet += "-webkit-background-clip: text;" +
                        "background-clip: text;" +
                        "color: transparent;";
                    if (bultColor[1].border !== undefined && bultColor[1].border !== "") {
                        bullet += "-webkit-text-stroke: " + bultColor[1].border + ";";
                    }
                    if (bultColor[1].effcts !== undefined && bultColor[1].effcts !== "") {
                        bullet += "filter: " + bultColor[1].effcts + ";";
                    }
                }

                if (isRTL) {
                    //bullet += "display: inline-block;white-space: nowrap ;direction:rtl"; // float: right;  
                    bullet += "white-space: nowrap ;direction:rtl"; // display: table-cell;;
                }
                let isIE11 = !!window.MSInputMethodContext && !!document.documentMode;
                let htmlBu = buChar;
                let useUnicodeFont = false;

                // 只有在非 Wingdings 字体时才进行 Unicode 转换
                if (!isIE11 && !isWingdingsFont) {
                    //ie11 does not support unicode ?
                    htmlBu = getHtmlBullet(typefaceNode, buChar);
                    useUnicodeFont = (htmlBu !== buChar);
                }
                
                // 如果使用了 Unicode 转换且是 Wingdings 字体，则使用标准字体
                if (useUnicodeFont && isWingdingsFont && typefaceNode !== undefined) {
                    // 使用正则表达式替换所有可能的 Wingdings 字体变体
                    bullet = bullet.replace(/font-family:\s*(Wingdings|Wingdings\s*2|Wingdings\s*3|Webdings)\s*/gi, "font-family: Arial, sans-serif");
                }
                
                bullet += "display: flex; align-items: center;'><div>" + htmlBu + "</div></div>";
                //} 
                // else {
                //     marginLeft = 328600 * SLIDE_FACTOR * lvl;

                //     bullet = `<div style='${marLStr}'>` + buChar + "</div>";
                // }
            } else if (buType == "TYPE_NUMERIC") {
                // 初始化项目符号计数器
                if (!warpObj.bulletCounter) {
                    warpObj.bulletCounter = {};
                }
                
                // 生成项目符号的唯一键
                const bulletKey = `${buNum}_${lvl}`;
                
                // 初始化或获取当前计数器
                if (!warpObj.bulletCounter[bulletKey]) {
                    warpObj.bulletCounter[bulletKey] = {
                        index: 0,
                        type: buNum,
                        level: lvl
                    };
                }
                
                // 增加计数器
                warpObj.bulletCounter[bulletKey].index++;
                
                // 生成数字编号
                const bulletIndex = warpObj.bulletCounter[bulletKey].index;
                const bulletText = getNumTypeNum(buNum, bulletIndex);

                bullet = "<div style='" + marLStr + marRStr;
                if (bultColor && bultColor[0] !== undefined && bultColor[0] != "") {
                    let bulletNumColorValue = bultColor[0];
                    if (bulletNumColorValue.length === 8) {
                        let colorObj = tinycolor(bulletNumColorValue);
                        bulletNumColorValue = colorObj.toRgbString();
                    } else {
                        bulletNumColorValue = "#" + bulletNumColorValue;
                    }
                    bullet += "color:" + bulletNumColorValue + ";";
                }
                bullet += `font-size:${bultSize};`;
                if (isRTL) {
                    bullet += "white-space: nowrap ;direction:rtl;";
                } else {
                    bullet += "white-space: nowrap ;direction:ltr;";
                }
                bullet += `display: flex; align-items: center;'><div>${bulletText}</div></div>`;

            } else if (buType == "TYPE_BULPIC") { //PIC BULLET
                // let marginLeft = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"])) * SLIDE_FACTOR;
                // let marginRight = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"])) * SLIDE_FACTOR;

                // if (isNaN(marginRight)) {
                //     marginRight = 0;
                // }
                // 
                // //buPic
                // if (isNaN(marginLeft)) {
                //     marginLeft = 328600 * SLIDE_FACTOR;
                // } else {
                //     marginLeft = 0;
                // }
                //let buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip","a:extLst","a:ext","asvg:svgBlip" , "attrs", "r:embed"]);
                let buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                let buImg;
                if (buPicId !== undefined) {
                    //svgPicPath = warpObj["slideResObj"][buPicId]["target"];
                    //buImg = warpObj["zip"].file(svgPicPath).asText();
                    //}else{
                    //buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                    let imgPath = (warpObj["slideResObj"][buPicId] !== undefined) ? warpObj["slideResObj"][buPicId]["target"] : undefined;

                    if (imgPath === undefined) {
                        buImg = "";
                    } else {
                        let imgFile = warpObj["zip"].file(imgPath);
                        if (imgFile === null) {
                            buImg = "";
                        } else {
                            let imgArrayBuffer = await imgFile.async("arraybuffer");
                            let imgExt = imgPath.split(".").pop();
                            let imgMimeType = PPTXXmlUtils.getMimeType(imgExt);
                            buImg = `<img src='data:${imgMimeType};base64,` + PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%;'/>";// height: 100%
        
                        }
                    }
                }
                if (buPicId === undefined) {
                    buImg = "&#8227;";
                }
                bullet = "<div style='" + marLStr + marRStr +
                    `width:${bultSize};display: flex; align-items: center;`;// +
                //"line-height: 0px;";
                if (isRTL) {
                    bullet += "white-space: nowrap ;direction:rtl;"; //direction:rtl; float: right;
                }
                bullet += `'>${buImg}  </div>`;
                //////////////////////////////////////////////////////////////////////////////////////
            }
            // else {
            //     bullet = "<div style='margin-left: " + 328600 * SLIDE_FACTOR * lvl + "px" +
            //         `; margin-right: ${0}px;'></div>`;
            // }

            return [bullet, margin_val, font_val];//$(bullet).outerWidth()];
        }
        function getHtmlBullet(typefaceNode, buChar) {
            //http://www.alanwood.net/demos/wingdings.html
            //not work for IE11
            //console.log("genBuChar typefaceNode:", typefaceNode, " buChar:", buChar, "charCodeAt:", buChar.charCodeAt(0))
            switch (buChar) {
                case "§":
                    return "&#9632;";//"■"; //9632 | U+25A0 | Black square
                case "q":
                    return "&#10065;";//"❑"; // 10065 | U+2751 | Lower right shadowed white square
                case "v":
                    return "&#10070;";//"❖"; //10070 | U+2756 | Black diamond minus white X
                case "Ø":
                    return "&#11162;";//"⮚"; //11162 | U+2B9A | Three-D top-lighted rightwards equilateral arrowhead
                case "ü":
                    return "&#10004;";//"✔";  //10004 | U+2714 | Heavy check mark
                case "o":
                    return "&#9679;";//"●"; //9679 | U+25CF | Black circle
                case "O":
                    return "&#9675;";//"○"; //9675 | U+25CB | White circle
                case "a":
                    return "&#9650;";//"▲"; //9650 | U+25B2 | Black up-pointing triangle
                case "A":
                    return "&#9651;";//"△"; //9651 | U+25B3 | White up-pointing triangle
                case "b":
                    return "&#9660;";//"▼"; //9660 | U+25BC | Black down-pointing triangle
                case "B":
                    return "&#9661;";//"▽"; //9661 | U+25BD | White down-pointing triangle
                case "c":
                    return "&#9654;";//"▶"; //9654 | U+25B6 | Black right-pointing triangle
                case "C":
                    return "&#9655;";//"▷"; //9655 | U+25B7 | White right-pointing triangle
                case "d":
                    return "&#9664;";//"◀"; //9664 | U+25C0 | Black left-pointing triangle
                case "D":
                    return "&#9665;";//"◁"; //9665 | U+25C1 | White left-pointing triangle
                case "e":
                    return "&#9670;";//"◆"; //9670 | U+25C6 | Black diamond
                case "E":
                    return "&#9671;";//"◇"; //9671 | U+25C7 | White diamond
                case "f":
                    return "&#10003;";//"✓"; //10003 | U+2713 | Check mark
                case "F":
                    return "&#10007;";//"✗"; //10007 | U+2717 | Ballot X
                case "g":
                    return "&#10002;";//"✔"; //10002 | U+2714 | Heavy check mark
                case "G":
                    return "&#10008;";//"✘"; //10008 | U+2718 | Heavy ballot X
                case "h":
                    return "&#9899;";//"★"; //9899 | U+2605 | Black star
                case "H":
                    return "&#9734;";//"☆"; //9734 | U+2606 | White star
                case "i":
                    return "&#10052;";//"✤"; //10052 | U+2724 | Heavy four-pointed star
                case "I":
                    return "&#10053;";//"✥"; //10053 | U+2725 | Four-pointed star
                case "j":
                    return "&#10022;";//"✶"; //10022 | U+2736 | Six-pointed star
                case "J":
                    return "&#10023;";//"✷"; //10023 | U+2737 | Eight-pointed star
                case "k":
                    return "&#10016;";//"✈"; //10016 | U+2708 | Airplane
                case "K":
                    return "&#10024;";//"✈"; //10024 | U+2708 | Airplane
                case "l":
                    return "&#10038;";//"✦"; //10038 | U+2726 | Black four-pointed star
                case "L":
                    return "&#10039;";//"✧"; //10039 | U+2727 | White four-pointed star
                case "m":
                    return "&#10017;";//"✉"; //10017 | U+2709 | Envelope
                case "M":
                    return "&#9993;";//"✉"; //9993 | U+2709 | Envelope
                case "n":
                    return "&#10084;";//"❤"; //10084 | U+2764 | Heavy black heart
                case "N":
                    return "&#9829;";//"♥"; //9829 | U+2665 | Black heart suit
                case "p":
                    return "&#9830;";//"♦"; //9830 | U+2666 | Black diamond suit
                case "P":
                    return "&#9826;";//"♢"; //9826 | U+2662 | White diamond suit
                case "r":
                    return "&#9827;";//"♣"; //9827 | U+2663 | Black club suit
                case "R":
                    return "&#9827;";//"♣"; //9827 | U+2663 | Black club suit
                case "s":
                    return "&#9824;";//"♠"; //9824 | U+2660 | Black spade suit
                case "S":
                    return "&#9824;";//"♠"; //9824 | U+2660 | Black spade suit
                case "t":
                    return "&#9828;";//"♣"; //9828 | U+2664 | White club suit
                case "T":
                    return "&#9825;";//"♥"; //9825 | U+2661 | White heart suit
                case "u":
                    return "&#9829;";//"♥"; //9829 | U+2665 | Black heart suit
                case "U":
                    return "&#9825;";//"♥"; //9825 | U+2661 | White heart suit
                case "w":
                    return "&#10071;";//"❗"; //10071 | U+2757 | Heavy exclamation mark symbol
                case "W":
                    return "&#10071;";//"❗"; //10071 | U+2757 | Heavy exclamation mark symbol
                case "x":
                    return "&#10062;";//"❞"; //10062 | U+275E | Heavy right-pointing angle quotation mark ornament
                case "X":
                    return "&#10063;";//"❟"; //10063 | U+275F | Heavy low single comma quotation mark ornament
                case "y":
                    return "&#10064;";//"❠"; //10064 | U+2760 | Heavy low double comma quotation mark ornament
                case "Y":
                    return "&#10064;";//"❠"; //10064 | U+2760 | Heavy low double comma quotation mark ornament
                case "z":
                    return "&#10061;";//"❝"; //10061 | U+275D | Heavy double turned comma quotation mark ornament
                case "Z":
                    return "&#10061;";//"❝"; //10061 | U+275D | Heavy double turned comma quotation mark ornament
                default:
                    if (typefaceNode == "Wingdings" || typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3" || typefaceNode == "Webdings"){
                        let wingCharCode =  getDingbatToUnicode(typefaceNode, buChar);
                        if (wingCharCode !== null){
                            return `&#${wingCharCode};`;
                        }
                    }
                    return "&#" + (buChar.charCodeAt(0)) + ";";
            }
        }
        function getDingbatToUnicode(typefaceNode, buChar){
            if (dingbatUnicode){
                let dingbat_code = buChar.codePointAt(0) & 0xFFF;
                let char_unicode = null;
                let len = dingbatUnicode.length;
                let i = 0;
                while (len--) {
                    // blah blah
                    let item = dingbatUnicode[i];
                    if (item.f == typefaceNode && item.code == dingbat_code) {
                        char_unicode = item.unicode;
                        break;
                    }
                    i++;
                }
                return char_unicode
        }
    }

    /**
     * alphaNumeric - 将数字转换为字母数字格式
     * @param {number} num - 数字
     * @param {string} upperLower - 大小写选项（upperCase或lowerCase）
     * @returns {string} 字母数字格式的字符串
     */
    function alphaNumeric(num, upperLower) {
        num = Number(num) - 1;
        let aNum = "";
        if (upperLower == "upperCase") {
            aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toUpperCase();
        } else if (upperLower == "lowerCase") {
            aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toLowerCase();
        }
        return aNum;
    }

    /**
     * hebrewAlphaNumeric - 将数字转换为希伯来字母编号格式
     * @param {number} num - 数字
     * @returns {string} 希伯来字母编号字符串
     */
    function hebrewAlphaNumeric(num) {
        num = Number(num) - 1;
        // 希伯来字母表（22个字母）
        const hebrewLetters = [
            'א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט',
            'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ',
            'ק', 'ר', 'ש', 'ת'
        ];
        const hebrewLength = hebrewLetters.length;

        if (num < hebrewLength) {
            // 单字母（1-22）
            return hebrewLetters[num];
        } else if (num < hebrewLength * (hebrewLength + 1)) {
            // 双字母（23-506）
            const first = Math.floor(num / hebrewLength);
            const second = num % hebrewLength;
            return hebrewLetters[first] + hebrewLetters[second];
        } else {
            // 三字母（507+）
            const third = num % hebrewLength;
            const remaining = Math.floor(num / hebrewLength);
            const second = remaining % hebrewLength;
            const first = Math.floor(remaining / hebrewLength);
            return hebrewLetters[first] + hebrewLetters[second] + hebrewLetters[third];
        }
    }

    /**
     * archaicNumbers - 处理古数字格式
     * @param {Array} arr - 数字映射数组
     * @returns {Object} 包含format方法的对象
     */
    function archaicNumbers(arr) {
        arr.slice().sort((a, b) => { return b[1].length - a[1].length });
        return {
            format: (n) => {
                let ret = '';
                for (let i = 0; i < arr.length; i++) {
                    let num = arr[i][0];
                    if (parseInt(num) > 0) {
                        for (; n >= num; n -= num) ret += arr[i][1];
                    } else {
                        ret = ret.replace(num, arr[i][1]);
                    }
                }
                return ret;
            }
        }
    }

    /**
     * romanize - 将数字转换为罗马数字
     * @param {number} num - 数字
     * @returns {string} 罗马数字字符串
     */
    function romanize(num) {
        if (!+num)
            return false;
        let digits = String(+num).split(""),
            key = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM",
                "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC",
                "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"],
            roman = "",
            i = 3;
        while (i--)
            roman = (key[+digits.pop() + (i * 10)] || "") + roman;
        return Array(+digits.join("") + 1).join("M") + roman;
    }
    archaicNumbers([
            [1000, ''],
            [400, 'ת'],
            [300, 'ש'],
            [200, 'ר'],
            [100, 'ק'],
            [90, 'צ'],
            [80, 'פ'],
            [70, 'ע'],
            [60, 'ס'],
            [50, 'נ'],
            [40, 'מ'],
            [30, 'ל'],
            [20, 'כ'],
            [10, 'י'],
            [9, 'ט'],
            [8, 'ח'],
            [7, 'ז'],
            [6, 'ו'],
            [5, 'ה'],
            [4, 'ד'],
            [3, 'ג'],
            [2, 'ב'],
            [1, 'א'],
            [/יה/, 'ט״ו'],
            [/יו/, 'ט״ז'],
            [/([א-ת])([א-ת])$/, '$1״$2'],
            [/^([א-ת])$/, "$1׳"]
        ]);
    /**
     * getNumTypeNum - 根据数字类型获取格式化的数字
     * @param {string} numTyp - 数字类型
     * @param {number} num - 数字
     * @returns {string} 格式化的数字字符串
     */
    function getNumTypeNum(numTyp, num) {
        let rtrnNum = "";
        switch (numTyp) {
            case "arabicPeriod":
                rtrnNum = num + ". ";
                break;
            case "arabicParenR":
                rtrnNum = num + ") ";
                break;
            case "alphaLcParenR":
                rtrnNum = alphaNumeric(num, "lowerCase") + ") ";
                break;
            case "alphaLcPeriod":
                rtrnNum = alphaNumeric(num, "lowerCase") + ". ";
                break;

            case "alphaUcParenR":
                rtrnNum = alphaNumeric(num, "upperCase") + ") ";
                break;
            case "alphaUcPeriod":
                rtrnNum = alphaNumeric(num, "upperCase") + ". ";
                break;

            case "romanUcPeriod":
                rtrnNum = romanize(num) + ". ";
                break;
            case "romanLcParenR":
                rtrnNum = romanize(num) + ") ";
                break;
            case "hebrew2Minus":
                // 希伯来字母编号：使用现代希伯来字母（א, ב, ג, ד, ...）类似英文字母编号
                rtrnNum = hebrewAlphaNumeric(num) + "-";
                break;
            default:
                rtrnNum = num;
        }
        return rtrnNum;
    }

    async function genSpanElement(node, rIndex, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNodeLength, warpObj, isBullate) {
            //https://codepen.io/imdunn/pen/GRgwaye ?
            let text_style = "";
            let lstStyle = textBodyNode["a:lstStyle"];
            let slideMasterTextStyles = warpObj["slideMasterTextStyles"];

            let text = node["a:t"];

            // 检查是否启用 rtlCol 模式（每个单词单独一行）
            // 只有在没有 wrap 属性时，rtlCol="1" 才表示每个单词单独一行
            let rtlColAttr = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rtlCol"]);
            let wrapAttr = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "wrap"]);
            let isRTLCol = (rtlColAttr === "1" && wrapAttr === undefined);
            //let text_count = text.length;

            let openElemnt = "<span";//"<bdi";
            let closeElemnt = "</span>";// "</bdi>";
            let styleText = "";
            if (text === undefined && node["type"] !== undefined) {
                if (is_first_br) {
                    //openElemnt = "<br";
                    //closeElemnt = "";
                    //return "<br style='font-size: initial'>"
                    is_first_br = false;
                    return "<span class='line-break-br' ></span>";
                }

                styleText += "display: block;";
                //openElemnt = "<span";
                //closeElemnt = "</span>";
            } else {

                is_first_br = true;
            }
            if (typeof text !== 'string') {
                text = PPTXXmlUtils.getTextByPathList(node, ["a:fld", "a:t"]);
                if (typeof text !== 'string') {
                    text = "&nbsp;";
                    //return "<span class='text-block '>&nbsp;</span>";
                }
                // if (text === undefined) {
                //     return "";
                // }
            }

            let pPrNode = pNode["a:pPr"];
            //lvl
            let lvl = 1;
            let lvlNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvl = parseInt(lvlNode) + 1;
            }
            //console.log("genSpanElement node: ", node, "rIndex: ", rIndex, ", pNode: ", pNode, ",pPrNode: ", pPrNode, "pFontStyle:", pFontStyle, ", idx: ", idx, "type:", type, warpObj);
            let layoutMasterNode = PPTXStyleUtils.getLayoutAndMasterNode(pNode, idx, type, warpObj);
            let pPrNodeLaout = layoutMasterNode.nodeLaout;
            let pPrNodeMaster = layoutMasterNode.nodeMaster;

            //Language
            let lang = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
            let isRtlLan = (lang !== undefined && RTL_LANGS_ARRAY.indexOf(lang) !== -1)?true:false;
            //rtl
            let getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
            if (getRtlVal === undefined) {
                getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
                if (getRtlVal === undefined && type != "shape") {
                    getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
                }
            }

            let linkID = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
            let linkTooltip = "";
            let defLinkClr;
            if (linkID !== undefined) {
                linkTooltip = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]);
                if (linkTooltip !== undefined) {
                    linkTooltip = `title='${linkTooltip}'`;
                }
                defLinkClr = PPTXStyleUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);
            } else {
                // Fallback to hover hyperlink (a:hlinkHover)
                linkID = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkHover", "attrs", "r:id"]);
                if (linkID !== undefined) {
                    linkTooltip = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkHover", "attrs", "tooltip"]);
                    if (linkTooltip !== undefined) {
                        linkTooltip = `title='${linkTooltip}'`;
                    }
                    defLinkClr = PPTXStyleUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);
                }
            }

            if (linkID !== undefined) {
                let linkClrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                PPTXStyleUtils.getSolidFill(linkClrNode, undefined, undefined, warpObj);


                //console.log("genSpanElement defLinkClr: ", defLinkClr, "rPrlinkClr:", rPrlinkClr)
                // 对于超链接，优先使用主题中的超链接颜色，而不是文本运行中的颜色定义
                // 注释掉下面的覆盖逻辑，让超链接始终使用主题颜色
                // if (rPrlinkClr !== undefined && rPrlinkClr != "") {
                //     defLinkClr = rPrlinkClr;
                // }
            }
            /////////////////////////////////////////////////////////////////////////////////////
            //getFontColor
            let fontClrPr = await PPTXStyleUtils.getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
            let fontClrType = fontClrPr[2];
            //console.log("genSpanElement fontClrPr: ", fontClrPr, "linkID", linkID);
            if (fontClrType == "solid") {
                if (linkID === undefined && fontClrPr[0] !== undefined && fontClrPr[0] != "") {
                    let colorValue = fontClrPr[0];
                    if (colorValue.length === 8) {
                        let colorObj = tinycolor(colorValue);
                        colorValue = colorObj.toRgbString();
                    } else {
                        colorValue = "#" + colorValue;
                    }
                    styleText += "color: " + colorValue + ";";
                }
                else if (linkID !== undefined && defLinkClr !== undefined) {
                    styleText += `color: #${defLinkClr};`;
                }

                if (fontClrPr[1] !== undefined && fontClrPr[1] != "" && fontClrPr[1] != ";") {
                    styleText += "text-shadow:" + fontClrPr[1] + ";";
                }
                if (fontClrPr[3] !== undefined && fontClrPr[3] != "") {
                    let highlightColorValue = fontClrPr[3];
                    if (highlightColorValue.length === 8) {
                        let colorObj = tinycolor(highlightColorValue);
                        highlightColorValue = colorObj.toRgbString();
                    } else {
                        highlightColorValue = "#" + highlightColorValue;
                    }
                    styleText += "background-color: " + highlightColorValue + ";";
                }
            } else if (fontClrType == "pattern" || fontClrType == "pic" || fontClrType == "gradient") {
                if (fontClrType == "pattern") {
                    styleText += "background:" + fontClrPr[0][0] + ";";
                    if (fontClrPr[0][1] !== null && fontClrPr[0][1] !== undefined && fontClrPr[0][1] != "") {
                        styleText += "background-size:" + fontClrPr[0][1] + ";";//" 2px 2px;" +
                    }
                    if (fontClrPr[0][2] !== null && fontClrPr[0][2] !== undefined && fontClrPr[0][2] != "") {
                        styleText += "background-position:" + fontClrPr[0][2] + ";";//" 2px 2px;" +
                    }
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";" +
                    //     "filter: " + fontClrPr[1].effcts + ";";
                } else if (fontClrType == "pic") {
                    styleText += fontClrPr[0] + ";";
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";
                } else if (fontClrType == "gradient") {

                    let colorAry = fontClrPr[0].color;
                    let rot = fontClrPr[0].rot;

                    styleText += `background: linear-gradient(${rot}deg,`;
                    for (let i = 0; i < colorAry.length; i++) {
                        if (i == colorAry.length - 1) {
                            styleText += "#" + colorAry[i] + ");";
                        } else {
                            styleText += "#" + colorAry[i] + ", ";
                        }
                    }
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";

                }
                styleText += "-webkit-background-clip: text;" +
                    "background-clip: text;" +
                    "color: transparent;";
                if (fontClrPr[1].border !== undefined && fontClrPr[1].border !== "") {
                    styleText += "-webkit-text-stroke: " + fontClrPr[1].border + ";";
                }
                if (fontClrPr[1].effcts !== undefined && fontClrPr[1].effcts !== "") {
                    styleText += "filter: " + fontClrPr[1].effcts + ";";
                }
            }
            let font_size = PPTXStyleUtils.getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj);
            //text_style += `font-size:${font_size};`
            
            text_style += `font-size:${font_size};` +
                // marLStr +
                "font-family:" + PPTXStyleUtils.getFontType(node, type, warpObj, pFontStyle) + ";" +
                "font-weight:" + PPTXStyleUtils.getFontBold(node, type, slideMasterTextStyles) + ";" +
                "font-style:" + PPTXStyleUtils.getFontItalic(node, type, slideMasterTextStyles) + ";" +
                "text-decoration:" + PPTXStyleUtils.getFontDecoration(node, type, slideMasterTextStyles) + ";" +
                "text-align:" + PPTXStyleUtils.getTextHorizontalAlign(node, pNode, type, warpObj) + ";" +
                "vertical-align:" + PPTXStyleUtils.getTextVerticalAlign(node, type, slideMasterTextStyles) + ";";
            
            // Merge styleText into text_style
            text_style += styleText;
            //rNodeLength
            //console.log("genSpanElement node:", node, "lang:", lang, "isRtlLan:", isRtlLan, "span parent dir:", dirStr)
            if (isRtlLan) { //|| rIndex === undefined
                styleText += "direction:rtl;";
            }else { //|| rIndex === undefined
                styleText += "direction:ltr;";
            }
            // } else if (dirStr == "rtl" && isRtlLan ) {
            //     styleText += "direction:rtl;";

            // } else if (dirStr == "ltr" && !isRtlLan ) {
            //     styleText += "direction:ltr;";
            // } else if (dirStr == "ltr" && isRtlLan){
            //     styleText += "direction:ltr;";
            // }else{
            //     styleText += "direction:inherit;";
            // }

            // if (dirStr == "rtl" && !isRtlLan) { //|| rIndex === undefined
            //     styleText += "direction:ltr;";
            // } else if (dirStr == "rtl" && isRtlLan) {
            //     styleText += "direction:rtl;";
            // } else if (dirStr == "ltr" && !isRtlLan) {
            //     styleText += "direction:ltr;";
            // } else if (dirStr == "ltr" && isRtlLan) {
            //     styleText += "direction:rtl;";
            // } else {
            //     styleText += "direction:inherit;";
            // }

            //     //`direction:${dirStr};`;
            //if (rNodeLength == 1 || rIndex == 0 ){
            //styleText += "display: table-cell;white-space: nowrap;";
            //}
            let highlight = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:highlight"]);
            if (highlight !== undefined) {
                let highlightColor = PPTXStyleUtils.getSolidFill(highlight, undefined, undefined, warpObj);
                if (highlightColor !== undefined && highlightColor != "") {
                    if (highlightColor.length === 8) {
                        let colorObj = tinycolor(highlightColor);
                        highlightColor = colorObj.toRgbString();
                    } else {
                        highlightColor = "#" + highlightColor;
                    }
                    styleText += "background-color:" + highlightColor + ";";
                }
                //styleText += "Opacity:" + getColorOpacity(highlight) + ";";
            }

            //letter-spacing:
            let spcNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "spc"]);
            if (spcNode === undefined) {
                spcNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "spc"]);
                if (spcNode === undefined) {
                    spcNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "spc"]);
                }
            }
            if (spcNode !== undefined) {
                let ltrSpc = parseInt(spcNode) / 100; //pt
                styleText += `letter-spacing: ${ltrSpc}px;`;// + "pt;";
            }

            //Text Cap Types
            let capNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "cap"]);
            if (capNode === undefined) {
                capNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "cap"]);
                if (capNode === undefined) {
                    capNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "cap"]);
                }
            }
            if (capNode == "small" || capNode == "all") {
                styleText += "text-transform: uppercase";
            }
            //styleText += "word-break: break-word;";
            //console.log("genSpanElement node: ", node, ", capNode: ", capNode, ",pPrNodeLaout: ", pPrNodeLaout, ", pPrNodeMaster: ", pPrNodeMaster, "warpObj:", warpObj);

            let cssName = "";

            if (styleText in warpObj.styleTable) {
                cssName = warpObj.styleTable[styleText]["name"];
            } else {
                cssName = "_css_" + (Object.keys(warpObj.styleTable).length + 1);
                warpObj.styleTable[styleText] = {
                    "name": cssName,
                    "text": styleText
                };
            }
            let linkColorSyle = "";
            if (fontClrType == "solid" && linkID !== undefined) {
                // 对于超链接，始终使用主题中的超链接颜色，而不是文本运行中的颜色
                if (defLinkClr !== undefined) {
                    linkColorSyle = `style='color: #${defLinkClr};'`;
                }
            }

            if (linkID !== undefined && linkID != "") {
                let linkURL = warpObj["slideResObj"][linkID]["target"];
                linkURL = PPTXXmlUtils.escapeHtml(linkURL);
                // 处理文本：制表符、换行符、多个连续空格
                let processedText = text
                    .replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;')  // 制表符转4个空格
                    .replace(/\n/g, "<br>")                      // 换行符转<br>
                    .replace(/  +/g, (spaces) => '&nbsp;'.repeat(spaces.length));  // 多个空格转&nbsp;

                // 在 rtlCol 模式下，每个单词单独一行
                if (isRTLCol) {
                    processedText = processedText.split(/\s+/).filter(word => word.length > 0).join("<br>");
                }

                return openElemnt + ` class='text-block ${cssName}' style='` + text_style + `'><a href='${linkURL}' ` + linkColorSyle + `  ${linkTooltip} target='_blank'>` +
                        processedText + "</a>" + closeElemnt;
            } else {
                // 处理文本：制表符、换行符、多个连续空格
                let processedText = text
                    .replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;')  // 制表符转4个空格
                    .replace(/\n/g, "<br>")                      // 换行符转<br>
                    .replace(/  +/g, (spaces) => '&nbsp;'.repeat(spaces.length));  // 多个空格转&nbsp;

                // 在 rtlCol 模式下，每个单词单独一行
                if (isRTLCol) {
                    processedText = processedText.split(/\s+/).filter(word => word.length > 0).join("<br>");
                }

                return openElemnt + ` class='text-block ${cssName}' style='` + text_style + "'>" + processedText + closeElemnt;//"</bdi>";
            }

        }


        async function genTable(node, warpObj, shapeType) {
            let order = node["attrs"]["order"];
            let tableNode = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
            let xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);

            // 处理组合缩放 - 当table在group-abs类型组合中时需要应用缩放
            let workingXfrmNode = xfrmNode;
            if (shapeType === 'group-abs' && warpObj.currentGroupScale && xfrmNode) {
                const { scaleX, scaleY, childX, childY } = warpObj.currentGroupScale;

                // 创建缩放后的xfrmNode
                workingXfrmNode = JSON.parse(JSON.stringify(xfrmNode));

                // 缩放尺寸
                if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
                    const originalCx = parseInt(xfrmNode['a:ext'].attrs.cx);
                    const originalCy = parseInt(xfrmNode['a:ext'].attrs.cy);
                    workingXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
                    workingXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
                }

                // 调整位置(相对于childX/childY)
                if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
                    const originalOffX = parseInt(xfrmNode['a:off'].attrs.x);
                    const originalOffY = parseInt(xfrmNode['a:off'].attrs.y);

                    // 计算相对于childOff的偏移
                    const relativeX = originalOffX - (childX / SLIDE_FACTOR$1);
                    const relativeY = originalOffY - (childY / SLIDE_FACTOR$1);

                    // 应用缩放
                    workingXfrmNode['a:off'].attrs.x = Math.round(childX / SLIDE_FACTOR$1 + relativeX * scaleX);
                    workingXfrmNode['a:off'].attrs.y = Math.round(childY / SLIDE_FACTOR$1 + relativeY * scaleY);
                }
            }
            /////////////////////////////////////////Amir////////////////////////////////////////////////
            let getTblPr = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
            let getColsGrid = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
            // PPTX中的rtl属性通常表示文本方向，而不是表格布局方向
            // 不在table标签上设置dir属性，避免列顺序反转
            // 单元格内部的文本会根据自身的RTL设置正确显示
            let tblDir = "";
            // if (getTblPr !== undefined) {
            //     let isRTL = getTblPr["attrs"]["rtl"];
            //     tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
            // }
            let firstRowAttr = getTblPr["attrs"]["firstRow"]; //associated element <a:firstRow> in the table styles
            let firstColAttr = getTblPr["attrs"]["firstCol"]; //associated element <a:firstCol> in the table styles
            let lastRowAttr = getTblPr["attrs"]["lastRow"]; //associated element <a:lastRow> in the table styles
            let lastColAttr = getTblPr["attrs"]["lastCol"]; //associated element <a:lastCol> in the table styles
            let bandRowAttr = getTblPr["attrs"]["bandRow"]; //associated element <a:band1H>, <a:band2H> in the table styles
            let bandColAttr = getTblPr["attrs"]["bandCol"]; //associated element <a:band1V>, <a:band2V> in the table styles
            //console.log("getTblPr: ", getTblPr);
            let tblStylAttrObj = {
                isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
                isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
                isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
                isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
                isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
                isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
            };

            let thisTblStyle;
            let tbleStyleId = getTblPr["a:tableStyleId"];
            if (tbleStyleId !== undefined) {
                let tbleStylList = warpObj.tableStyles["a:tblStyleLst"]["a:tblStyle"];
                if (tbleStylList !== undefined) {
                    if (tbleStylList.constructor === Array) {
                        for (let k = 0; k < tbleStylList.length; k++) {
                            if (tbleStylList[k]["attrs"]["styleId"] == tbleStyleId) {
                                thisTblStyle = tbleStylList[k];
                            }
                        }
                    } else {
                        if (tbleStylList["attrs"]["styleId"] == tbleStyleId) {
                            thisTblStyle = tbleStylList;
                        }
                    }
                }
            }
            if (thisTblStyle !== undefined) {
                thisTblStyle["tblStylAttrObj"] = tblStylAttrObj;
                warpObj["thisTbiStyle"] = thisTblStyle;
            }
            let tblStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
            let tblBorderStyl = PPTXXmlUtils.getTextByPathList(tblStyl, ["a:tcBdr"]);
            let tbl_borders = "";
            if (tblBorderStyl !== undefined) {
                tbl_borders = PPTXStyleUtils.getTableBorders(tblBorderStyl, warpObj);
            }
            let tbl_bgcolor = "";
            let tbl_bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
            //console.log( "thisTblStyle:", thisTblStyle, "warpObj:", warpObj)
            if (tbl_bgFillschemeClr !== undefined) {
                tbl_bgcolor = PPTXStyleUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
            }
            if (tbl_bgFillschemeClr === undefined) {
                tbl_bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                tbl_bgcolor = PPTXStyleUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
            }
            if (tbl_bgcolor !== "" && typeof tbl_bgcolor === 'string') {
                if (tbl_bgcolor.length === 8) {
                    let colorObj = tinycolor(tbl_bgcolor);
                    tbl_bgcolor = colorObj.toRgbString();
                } else {
                    tbl_bgcolor = "#" + tbl_bgcolor;
                }
                tbl_bgcolor = `background-color: ${tbl_bgcolor};`;
            }
            ////////////////////////////////////////////////////////////////////////////////////////////
            let tableHtml = `<table ${tblDir} style='border-collapse: collapse;` +
                PPTXXmlUtils.getPosition(workingXfrmNode, node, undefined, undefined, shapeType) +
                PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined) +
                ` z-index: ${order};` +
                tbl_borders + `;${tbl_bgcolor}'>`;

            let trNodes = tableNode["a:tr"];
            if (trNodes.constructor !== Array) {
                trNodes = [trNodes];
            }
                let rowSpanAry = [];
                for (let i = 0; i < trNodes.length; i++) {
                    //////////////rows Style ////////////Amir
                    let rowHeightParam = trNodes[i]["attrs"]["h"];
                    let rowHeight = 0;
                    let rowsStyl = "";
                    if (rowHeightParam !== undefined) {
                        rowHeight = parseInt(rowHeightParam) * SLIDE_FACTOR$1;
                        rowHeight = Math.round(rowHeight * 100) / 100;
                        rowsStyl += `height:${rowHeight}px;`;
                    }
                    let fillColor = "";
                    let row_borders = "";
                    let fontClrPr = "";
                    let fontWeight = "";

                    if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
                        let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                        }
                        let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            let local_fontColor = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontColor !== undefined) {
                                fontClrPr = local_fontColor;
                            }

                            let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight != "") {
                                fontWeight = local_fontWeight;
                            }
                        }
                    }

                    if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {

                        let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                        }
                        let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }
                            let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight !== "") {
                                fontWeight = local_fontWeight;
                            }
                        }

                    } else if (i > 0 && tblStylAttrObj["isBandRowAttr"] == 1 && thisTblStyle !== undefined) {
                        fillColor = "";
                        row_borders = undefined;
                        if ((i % 2) == 0 && thisTblStyle["a:band2H"] !== undefined) {
                            // console.log("i: ", i, 'thisTblStyle["a:band2H"]:', thisTblStyle["a:band2H"])
                            //check if there is a row bg
                            let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:fill", "a:solidFill"]);
                            if (bgFillschemeClr !== undefined) {
                                let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                                if (local_fillColor !== "") {
                                    fillColor = local_fillColor;
                                }
                            }


                            let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:tcBdr"]);
                            if (borderStyl !== undefined) {
                                let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                                if (local_row_borders != "") {
                                    row_borders = local_row_borders;
                                }
                            }
                            let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcTxStyle"]);
                            if (rowTxtStyl !== undefined) {
                                let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                                if (local_fontClrPr !== undefined) {
                                    fontClrPr = local_fontClrPr;
                                }
                            }

                            let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");

                            if (local_fontWeight !== "") {
                                fontWeight = local_fontWeight;
                            }
                        }
                        if ((i % 2) != 0 && thisTblStyle["a:band1H"] !== undefined) {
                            let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:fill", "a:solidFill"]);
                            if (bgFillschemeClr !== undefined) {
                                let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                                if (local_fillColor !== undefined) {
                                    fillColor = local_fillColor;
                                }
                            }
                            let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:tcBdr"]);
                            if (borderStyl !== undefined) {
                                let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                                if (local_row_borders != "") {
                                    row_borders = local_row_borders;
                                }
                            }
                            let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcTxStyle"]);
                            if (rowTxtStyl !== undefined) {
                                let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                                if (local_fontClrPr !== undefined) {
                                    fontClrPr = local_fontClrPr;
                                }
                                let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                                if (local_fontWeight != "") {
                                    fontWeight = local_fontWeight;
                                }
                            }
                        }

                    }
                    //last row
                    if (i == (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1 && thisTblStyle !== undefined) {
                        let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                            // let local_colorOpacity = getColorOpacity(bgFillschemeClr);
                            // if(local_colorOpacity !== undefined){
                            //     colorOpacity = local_colorOpacity;
                            // }
                        }
                        let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }

                            let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight !== "") {
                                fontWeight = local_fontWeight;
                            }
                        }
                    }
                    rowsStyl += ((row_borders !== undefined) ? row_borders : "");
                    if (fontClrPr !== undefined && typeof fontClrPr === 'string') {
                        let tableColorValue = fontClrPr;
                        if (tableColorValue.length === 8) {
                            let colorObj = tinycolor(tableColorValue);
                            tableColorValue = colorObj.toRgbString();
                        } else {
                            tableColorValue = "#" + tableColorValue;
                        }
                        rowsStyl += ` color: ${tableColorValue};`;
                    }
                    rowsStyl += ((fontWeight != "") ? ` font-weight:${fontWeight};` : "");
                    if (fillColor !== undefined && fillColor != "" && typeof fillColor === 'string') {
                        if (fillColor.length === 8) {
                            let colorObj = tinycolor(fillColor);
                            fillColor = colorObj.toRgbString();
                        } else {
                            fillColor = "#" + fillColor;
                        }
                        //rowsStyl += "background-color: rgba(" + hexToRgbNew(fillColor) + `,${colorOpacity});`;
                        rowsStyl += `background-color: ${fillColor};`;
                    }
                    tableHtml += `<tr style='${rowsStyl}'>`;
                    ////////////////////////////////////////////////

                    let tcNodes = trNodes[i]["a:tc"];
                    if (tcNodes !== undefined) {
                        if (tcNodes.constructor === Array) {
                            //multi columns
                            let j = 0;
                            if (rowSpanAry.length == 0) {
                                rowSpanAry = Array.apply(null, Array(tcNodes.length)).map(() => { return 0 });
                            }
                            let totalColSpan = 0;
                            while (j < tcNodes.length) {
                                if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                                    let a_sorce;
                                    //j=0 : first col
                                    if (j == 0 && tblStylAttrObj["isFrstColAttr"] == 1) {
                                        a_sorce = "a:firstCol";
                                        if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) &&
                                            PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:seCell"]) !== undefined) {
                                            a_sorce = "a:seCell";
                                        } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 &&
                                            PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:neCell"]) !== undefined) {
                                            a_sorce = "a:neCell";
                                        }
                                    } else if ((j > 0 && tblStylAttrObj["isBandColAttr"] == 1) &&
                                        !(tblStylAttrObj["isFrstColAttr"] == 1 && i == 0) &&
                                        !(tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1)) &&
                                        j != (tcNodes.length - 1)) {

                                        if ((j % 2) != 0) {

                                            let aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                                            if (aBandNode === undefined) {
                                                aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
                                                if (aBandNode !== undefined) {
                                                    a_sorce = "a:band2V";
                                                }
                                            } else {
                                                a_sorce = "a:band2V";
                                            }

                                        }
                                    }

                                    if (j == (tcNodes.length - 1) && tblStylAttrObj["isLstColAttr"] == 1) {
                                        a_sorce = "a:lastCol";
                                        if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) && PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:swCell"]) !== undefined) {
                                            a_sorce = "a:swCell";
                                        } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 && PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:nwCell"]) !== undefined) {
                                            a_sorce = "a:nwCell";
                                        }
                                    }

                                    let cellParmAry = await getTableCellParams(tcNodes[j], getColsGrid, i , j , thisTblStyle, a_sorce, warpObj);
                                    let text = cellParmAry[0];
                                    let colStyl = cellParmAry[1];
                                    let cssName = cellParmAry[2];
                                    let rowSpan = cellParmAry[3];
                                    let colSpan = cellParmAry[4];



                                    if (rowSpan !== undefined) {
                                        rowSpanAry[j] = parseInt(rowSpan) - 1;
                                        tableHtml += `<td class='${cssName}' data-row='` + i + `,${j}' rowspan ='` +
                                            parseInt(rowSpan) + `' style='${colStyl}'>` + text + "</td>";
                                    } else if (colSpan !== undefined) {
                                        tableHtml += `<td class='${cssName}' data-row='` + i + `,${j}' colspan = '` +
                                            parseInt(colSpan) + `' style='${colStyl}'>` + text + "</td>";
                                        totalColSpan = parseInt(colSpan) - 1;
                                    } else {
                                        tableHtml += `<td class='${cssName}' data-row='` + i + `,${j}' style = '` + colStyl + `'>${text}</td>`;
                                    }

                                } else {
                                    if (rowSpanAry[j] != 0) {
                                        rowSpanAry[j] -= 1;
                                    }
                                    if (totalColSpan != 0) {
                                        totalColSpan--;
                                    }
                                }
                                j++;
                            }
                        } else {
                            //single column 

                            let a_sorce;
                            if (tblStylAttrObj["isFrstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                                a_sorce = "a:firstCol";

                            } else if ((tblStylAttrObj["isBandColAttr"] == 1) && !(tblStylAttrObj["isLstRowAttr"] == 1)) {

                                let aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                                if (aBandNode === undefined) {
                                    aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
                                    if (aBandNode !== undefined) {
                                        a_sorce = "a:band2V";
                                    }
                                } else {
                                    a_sorce = "a:band2V";
                                }
                            }

                            if (tblStylAttrObj["isLstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                                a_sorce = "a:lastCol";
                            }


                            let cellParmAry = await getTableCellParams(tcNodes, getColsGrid , i , undefined , thisTblStyle, a_sorce, warpObj);
                            let text = cellParmAry[0];
                            let colStyl = cellParmAry[1];
                            let cssName = cellParmAry[2];
                            let rowSpan = cellParmAry[3];

                            if (rowSpan !== undefined) {
                                tableHtml += `<td  class='${cssName}' rowspan='` + parseInt(rowSpan) + `' style = '${colStyl}'>` + text + "</td>";
                            } else {
                                tableHtml += `<td class='${cssName}' style='` + colStyl + `'>${text}</td>`;
                            }
                        }
                    }
                    tableHtml += "</tr>";
                }
                //////////////////////////////////////////////////////////////////////////////////
            

            return tableHtml;
        }
        
        async function getTableCellParams(tcNodes, getColsGrid , row_idx , col_idx , thisTblStyle, cellSource, warpObj) {
            //thisTblStyle["a:band1V"] => thisTblStyle[cellSource]
            //text, cell-width, cell-borders, 
            //let text = PPTXTextUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj);//tableStyles
            let rowSpan = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "rowSpan"]);
            let colSpan = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "gridSpan"]);
            PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "vMerge"]);
            PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "hMerge"]);
            let colStyl = "word-wrap: break-word;";
            let colWidth;
            let celFillColor = "";
            let colFontClrPr = "";
            let colFontWeight = "";
            let lin_bottm = "",
                lin_top = "",
                lin_left = "",
                lin_right = "";
            
            let colSapnInt = parseInt(colSpan);
            let total_col_width = 0;
            if (!isNaN(colSapnInt) && colSapnInt > 1){
                for (let k = 0; k < colSapnInt ; k++) {
                    total_col_width += parseInt (PPTXXmlUtils.getTextByPathList(getColsGrid[col_idx + k], ["attrs", "w"]));
                }
            }else {
                total_col_width = PPTXXmlUtils.getTextByPathList((col_idx === undefined) ? getColsGrid : getColsGrid[col_idx], ["attrs", "w"]);
            }
            

            let text = await PPTXTextUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, "table", undefined, warpObj, total_col_width);//tableStyles

            if (total_col_width != 0 /*&& row_idx == 0*/) {
                colWidth = parseInt(total_col_width) * SLIDE_FACTOR$1;
                colWidth = Math.round(colWidth * 100) / 100;
                colStyl += `width:${colWidth}px;`;
            }

            //cell horizontal alignment
            let cellAlign = "";
            // 获取表格单元格文本的RTL状态，以确定正确的对齐方式
            let textBodyNode = tcNodes["a:txBody"];
            let prg_dir = "";
            if (textBodyNode !== undefined) {
                // 获取段落的RTL状态
                let pNodes = textBodyNode["a:p"];
                if (pNodes !== undefined) {
                    if (Array.isArray(pNodes) && pNodes.length > 0) {
                        prg_dir = PPTXStyleUtils.getPregraphDir(pNodes[0], textBodyNode, 0, "table", warpObj);
                    } else {
                        prg_dir = PPTXStyleUtils.getPregraphDir(pNodes, textBodyNode, 0, "table", warpObj);
                    }
                }
            }
            let isRTL = (prg_dir == "pregraph-rtl");

            // 获取水平对齐属性
            if (textBodyNode !== undefined) {
                let pNodes = textBodyNode["a:p"];
                if (pNodes !== undefined) {
                    let firstP = Array.isArray(pNodes) ? pNodes[0] : pNodes;
                    let horizontalAlign = PPTXStyleUtils.getHorizontalAlign(firstP, textBodyNode, 0, "table", prg_dir, warpObj);
                    if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                        cellAlign = isRTL ? "text-align: left;" : "text-align: right;";
                    } else if (horizontalAlign === "h-mid") {
                        cellAlign = "text-align: center;";
                    } else if (horizontalAlign === "h-left-rtl") {
                        cellAlign = "text-align: right;";
                    } else {
                        cellAlign = isRTL ? "text-align: right;" : "text-align: left;";
                    }
                }
            }
            if (cellAlign !== "") {
                colStyl += cellAlign;
            }

            //cell bords
            lin_bottm = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnB"]);
            if (lin_bottm === undefined && cellSource !== undefined) {
                if (cellSource !== undefined)
                    lin_bottm = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
                if (lin_bottm === undefined) {
                    lin_bottm = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
                }
            }
            lin_top = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnT"]);
            if (lin_top === undefined) {
                if (cellSource !== undefined)
                    lin_top = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
                if (lin_top === undefined) {
                    lin_top = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
                }
            }
            lin_left = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnL"]);
            if (lin_left === undefined) {
                if (cellSource !== undefined)
                    lin_left = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
                if (lin_left === undefined) {
                    lin_left = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
                }
            }
            lin_right = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnR"]);
            if (lin_right === undefined) {
                if (cellSource !== undefined)
                    lin_right = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
                if (lin_right === undefined) {
                    lin_right = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
                }
            }
            PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnBlToTr"]);
            PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:InTlToBr"]);

            if (lin_bottm !== undefined && lin_bottm != "") {
                let bottom_line_border = PPTXStyleUtils.getBorder(lin_bottm, undefined, false, "", warpObj);
                if (bottom_line_border != "") {
                    colStyl += `border-bottom:${bottom_line_border};`;
                }
            }
            if (lin_top !== undefined && lin_top != "") {
                let top_line_border = PPTXStyleUtils.getBorder(lin_top, undefined, false, "", warpObj);
                if (top_line_border != "") {
                    colStyl += `border-top: ${top_line_border};`;
                }
            }
            if (lin_left !== undefined && lin_left != "") {
                let left_line_border = PPTXStyleUtils.getBorder(lin_left, undefined, false, "", warpObj);
                if (left_line_border != "") {
                    colStyl += `border-left: ${left_line_border};`;
                }
            }
            if (lin_right !== undefined && lin_right != "") {
                let right_line_border = PPTXStyleUtils.getBorder(lin_right, undefined, false, "", warpObj);
                if (right_line_border != "") {
                    colStyl += `border-right:${right_line_border};`;
                }
            }

            //cell fill color custom
            let getCelFill = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr"]);
            if (getCelFill !== undefined && getCelFill != "") {
                let cellObj = {
                    "p:spPr": getCelFill
                };
                celFillColor = await PPTXStyleUtils.getShapeFill(cellObj, undefined, false, warpObj, "slide");
            }

            //cell fill color theme
            if (celFillColor == "" || celFillColor == "background-color: inherit;") {
                let bgFillschemeClr;
                if (cellSource !== undefined)
                    bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        celFillColor = ` background-color: #${local_fillColor};`;
                    }
                }
            }
            let cssName = "";
            if (celFillColor !== undefined && celFillColor != "") {
                if (celFillColor in warpObj.styleTable) {
                    cssName = warpObj.styleTable[celFillColor]["name"];
                } else {
                    cssName = "_tbl_cell_css_" + (Object.keys(warpObj.styleTable).length + 1);
                    warpObj.styleTable[celFillColor] = {
                        "name": cssName,
                        "text": celFillColor
                    };
                }

            }

            //border
            // let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:tcBdr"]);
            // if (borderStyl !== undefined) {
            //     let local_col_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
            //     if (local_col_borders != "") {
            //         col_borders = local_col_borders;
            //     }
            // }
            // if (col_borders != "") {
            //     colStyl += col_borders;
            // }

            //Text style
            let rowTxtStyl;
            if (cellSource !== undefined) {
                rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcTxStyle"]);
            }
            // if (rowTxtStyl === undefined) {
            //     rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
            // }
            if (rowTxtStyl !== undefined) {
                let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontClrPr !== undefined) {
                    colFontClrPr = local_fontClrPr;
                }
                let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight !== "") {
                    colFontWeight = local_fontWeight;
                }
            }
            colStyl += ((colFontClrPr !== "" && typeof colFontClrPr === 'string') ?
                ((colFontClrPr.length === 8) ?
                    (() => {
                        let colorObj = tinycolor(colFontClrPr);
                        return `color: ${colorObj.toRgbString()};`;
                    })() :
                    `color: #${colFontClrPr};`) : "");
            colStyl += ((colFontWeight != "") ? ` font-weight:${colFontWeight};` : "");

            return [text, colStyl, cssName, rowSpan, colSpan];
        }
const PPTXTextUtils = {
        genTextBody,
        genBuChar,
        getHtmlBullet,
        getDingbatToUnicode,
        genSpanElement,
        genTable,
        getTableCellParams,
        alphaNumeric,
        archaicNumbers,
        romanize,
        getNumTypeNum,
    };

/**
 * 路径生成器模块
 * 纯数学计算函数,无外部依赖,无副作用
 * 从 shape.js 提取的独立工具函数
 */

/**
 * polarToCartesian - 将极坐标转换为笛卡尔坐标
 * @param {number} cx - 圆心X坐标
 * @param {number} cy - 圆心Y坐标
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} angleInDegrees - 角度
 * @returns {Object} 笛卡尔坐标对象 {x, y}
 */
function polarToCartesian(cx, cy, w, h, angleInDegrees) {
    var angleInRadians = (angleInDegrees - 90) * Math.PI / 180.0;
    // 格式化数字为2位小数
    function fmt(num) {
        return parseFloat(num.toFixed(2));
    }
    return {
        x: fmt(cx + (w / 2) * Math.cos(angleInRadians)),
        y: fmt(cy + (h / 2) * Math.sin(angleInRadians))
    };
}

/**
 * shapeArc - 生成圆弧路径
 * @param {number} cx - 圆心X坐标
 * @param {number} cy - 圆心Y坐标
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} startAngle - 起始角度
 * @param {number} endAngle - 结束角度
 * @param {boolean} clockwise - 是否顺时针
 * @returns {string} SVG路径字符串
 */
function shapeArc(cx, cy, w, h, startAngle, endAngle, clockwise) {
    var start = polarToCartesian(cx, cy, w, h, endAngle);
    var end = polarToCartesian(cx, cy, w, h, startAngle);
    var largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";
    // 格式化数字为2位小数
    function fmt(num) {
        return parseFloat(num.toFixed(2));
    }
    var d = [
        "M", start.x, start.y,
        "A", fmt(w), fmt(h), 0, largeArcFlag, clockwise ? "0" : "1", end.x, end.y
    ].join(" ");
    return d;
}

/**
 * shapeArcAlt - 生成圆弧路径(逐点计算实现)
 * @param {number} cX - 圆心X坐标
 * @param {number} cY - 圆心Y坐标
 * @param {number} rX - X半径
 * @param {number} rY - Y半径
 * @param {number} stAng - 起始角度
 * @param {number} endAng - 结束角度
 * @param {boolean} isClose - 是否闭合
 * @returns {string} SVG路径字符串
 */
function shapeArcAlt(cX, cY, rX, rY, stAng, endAng, isClose) {
    var dData;
    var angle = stAng;
    // 辅助函数：格式化数字为2位小数
    function fmt(num) {
        return parseFloat(num.toFixed(2));
    }
    if (endAng >= stAng) {
        while (angle <= endAng) {
            var radians = angle * (Math.PI / 180);
            var x = cX + Math.cos(radians) * rX;
            var y = cY + Math.sin(radians) * rY;
            if (angle == stAng) {
                dData = " M" + fmt(x) + " " + fmt(y);
            }
            dData += " L" + fmt(x) + " " + fmt(y);
            angle++;
        }
    } else {
        while (angle > endAng) {
            var radians = angle * (Math.PI / 180);
            var x = cX + Math.cos(radians) * rX;
            var y = cY + Math.sin(radians) * rY;
            if (angle == stAng) {
                dData = " M " + fmt(x) + " " + fmt(y);
            }
            dData += " L " + fmt(x) + " " + fmt(y);
            angle--;
        }
    }
    dData += (isClose ? " z" : "");
    return dData;
}

/**
 * shapeSnipRoundRect - 生成圆角或裁剪矩形路径
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} sAdj1_val - 调整值1
 * @param {number} sAdj2_val - 调整值2
 * @param {string} shpTyp - 形状类型 ("round" 或 "snip")
 * @param {string} adjTyp - 调整类型 ("cornr1", "cornr2", "cornrAll", "diag")
 * @returns {string} SVG路径字符串
 */
function shapeSnipRoundRect(w, h, sAdj1_val, sAdj2_val, shpTyp, adjTyp) {
    var d = "";
    var sAdj1 = 0;
    var sAdj2 = 0;

    if (shpTyp == "round") {
        sAdj1 = w * sAdj1_val;
        if (adjTyp == "cornrAll") {
            d = "M0," + sAdj1 + " Q0,0 " + sAdj1 + ",0 L" + (w - sAdj1) + ",0 Q" + w + ",0 " + w + "," + sAdj1 + " L" + w + "," + (h - sAdj1) + " Q" + w + "," + h + " " + (w - sAdj1) + "," + h + " L" + sAdj1 + "," + h + " Q0," + h + " 0," + (h - sAdj1) + " z";
        } else if (adjTyp == "cornr1") {
            d = "M0,0 L" + (w - sAdj1) + ",0 Q" + w + ",0 " + w + "," + sAdj1 + " L" + w + "," + h + " L0," + h + " z";
        } else if (adjTyp == "diag") {
            sAdj2 = h * sAdj2_val;
            d = "M0,0 L" + (w - sAdj1) + ",0 Q" + w + ",0 " + w + "," + sAdj1 + " L" + w + "," + (h - sAdj2) + " Q" + w + "," + h + " " + (w - sAdj2) + "," + h + " L" + sAdj1 + "," + h + " Q0," + h + " 0," + (h - sAdj1) + " L0," + sAdj2 + " Q0,0 " + sAdj2 + ",0 z";
        } else if (adjTyp == "cornr2") {
            sAdj2 = w * sAdj2_val;
            d = "M0,0 L" + (w - sAdj1) + ",0 Q" + w + ",0 " + w + "," + sAdj1 + " L" + w + "," + (h - sAdj2) + " Q" + w + "," + h + " " + (w - sAdj2) + "," + h + " L0," + h + " z";
        }
    } else if (shpTyp == "snip") {
        sAdj1 = w * sAdj1_val;
        if (adjTyp == "cornr1") {
            d = "M" + sAdj1 + ",0 L" + w + ",0 L" + w + "," + h + " L0," + h + " L0," + sAdj1 + " z";
        } else if (adjTyp == "diag") {
            sAdj2 = h * sAdj2_val;
            d = "M" + sAdj1 + ",0 L" + w + ",0 L" + w + "," + (h - sAdj2) + " L" + sAdj2 + "," + h + " L0," + h + " L0," + sAdj1 + " z";
        } else if (adjTyp == "cornr2") {
            sAdj2 = w * sAdj2_val;
            d = "M" + sAdj1 + ",0 L" + w + ",0 L" + w + "," + (h - sAdj2) + " L" + (w - sAdj2) + "," + h + " L0," + h + " z";
        }
    }

    return d;
}

/**
 * shapeSnipRoundRectAlt - 生成圆角或裁剪矩形路径(备选实现)
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} adj1 - 调整值1
 * @param {number} adj2 - 调整值2
 * @param {string} shapeType - 形状类型 ("snip" 或 "round")
 * @param {string} adjType - 调整类型 ("cornr1", "cornr2", "cornrAll", "diag")
 * @returns {string} SVG路径字符串
 */
function shapeSnipRoundRectAlt(w, h, adj1, adj2, shapeType, adjType) {
    var adjA, adjB, adjC, adjD;
    if (adjType == "cornr1") {
        adjA = 0;
        adjB = 0;
        adjC = 0;
        adjD = adj1;
    } else if (adjType == "cornr2") {
        adjA = adj1;
        adjB = adj2;
        adjC = adj2;
        adjD = adj1;
    } else if (adjType == "cornrAll") {
        adjA = adj1;
        adjB = adj1;
        adjC = adj1;
        adjD = adj1;
    } else if (adjType == "diag") {
        adjA = adj1;
        adjB = adj2;
        adjC = adj1;
        adjD = adj2;
    }

    var d;
    if (shapeType == "round") {
        d = "M0" + "," + (h / 2 + (1 - adjB) * (h / 2)) + " Q" + 0 + "," + h + " " + adjB * (w / 2) + "," + h + " L" + (w / 2 + (1 - adjC) * (w / 2)) + "," + h +
            " Q" + w + "," + h + " " + w + "," + (h / 2 + (h / 2) * (1 - adjC)) + "L" + w + "," + (h / 2) * adjD +
            " Q" + w + "," + 0 + " " + (w / 2 + (w / 2) * (1 - adjD)) + ",0 L" + (w / 2) * adjA + ",0" +
            " Q" + 0 + "," + 0 + " 0," + (h / 2) * (adjA) + " z";
    } else if (shapeType == "snip") {
        d = "M0" + "," + adjA * (h / 2) + " L0" + "," + (h / 2 + (h / 2) * (1 - adjB)) + "L" + adjB * (w / 2) + "," + h +
            " L" + (w / 2 + (w / 2) * (1 - adjC)) + "," + h + "L" + w + "," + (h / 2 + (h / 2) * (1 - adjC)) +
            " L" + w + "," + adjD * (h / 2) + "L" + (w / 2 + (w / 2) * (1 - adjD)) + ",0 L" + ((w / 2) * adjA) + ",0 z";
    }
    return d;
}

/**
 * shapePie - 生成饼图路径
 * @param {number} H - 高度
 * @param {number} w - 宽度
 * @param {number} adj1 - 调整值1(起始角度)
 * @param {number} adj2 - 调整值2(结束角度)
 * @param {boolean} isClose - 是否闭合
 * @returns {Array} [路径字符串, 旋转字符串]
 */
function shapePie(H, w, adj1, adj2, isClose) {
    var pieVal = parseInt(adj2);
    var piAngle = parseInt(adj1);
    var size = parseInt(H),
        radius = (size / 2),
        value = pieVal - piAngle;
    if (value < 0) {
        value = 360 + value;
    }
    value = Math.min(Math.max(value, 0), 360);

    var x = Math.cos((2 * Math.PI) / (360 / value));
    var y = Math.sin((2 * Math.PI) / (360 / value));

    var longArc, d, rot;
    if (isClose) {
        longArc = (value <= 180) ? 0 : 1;
        d = "M" + radius + "," + radius + " L" + radius + "," + 0 + " A" + radius + "," + radius + " 0 " + longArc + ",1 " + (radius + y * radius) + "," + (radius - x * radius) + " z";
        rot = "rotate(" + (piAngle - 270) + ", " + radius + ", " + radius + ")";
    } else {
        longArc = (value <= 180) ? 0 : 1;
        var radius1 = radius;
        var radius2 = w / 2;
        d = "M" + radius1 + "," + 0 + " A" + radius2 + "," + radius1 + " 0 " + longArc + ",1 " + (radius2 + y * radius2) + "," + (radius1 - x * radius1);
        rot = "rotate(" + (piAngle + 90) + ", " + radius + ", " + radius + ")";
    }

    return [d, rot];
}

/**
 * shapeGear - 生成齿轮形状路径
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} points - 点数(齿轮齿数)
 * @returns {string} SVG路径字符串
 */
function shapeGear(w, h, points) {
    var innerRadius = h;
    var outerRadius = 1.5 * innerRadius;
    var cx = outerRadius;
    var cy = outerRadius;
    var notches = points;
    var radiusO = outerRadius;
    var radiusI = innerRadius;
    var taperO = 50;
    var taperI = 35;
    var pi2 = 2 * Math.PI;
    var angle = pi2 / (notches * 2);
    var taperAI = angle * taperI * 0.005;
    var taperAO = angle * taperO * 0.005;
    var a = angle;
    var toggle = false;

    var d = " M" + (cx + radiusO * Math.cos(taperAO)) + " " + (cy + radiusO * Math.sin(taperAO));

    for (; a <= pi2 + angle; a += angle) {
        if (toggle) {
            d += " L" + (cx + radiusI * Math.cos(a - taperAI)) + "," + (cy + radiusI * Math.sin(a - taperAI));
            d += " L" + (cx + radiusO * Math.cos(a + taperAO)) + "," + (cy + radiusO * Math.sin(a + taperAO));
        } else {
            d += " L" + (cx + radiusO * Math.cos(a - taperAO)) + "," + (cy + radiusO * Math.sin(a - taperAO));
            d += " L" + (cx + radiusI * Math.cos(a + taperAI)) + "," + (cy + radiusI * Math.sin(a + taperAI));
        }
        toggle = !toggle;
    }
    d += " ";
    return d;
}

/**
 * 自定义形状 (custGeom) 渲染模块
 * 处理 PowerPoint 中的自定义几何形状
 * 参考: http://officeopenxml.com/drwSp-custGeom.php
 */


/**
 * 渲染自定义形状
 * @param {Object} custShapType - 自定义形状数据
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {boolean} imgFillFlg - 是否图片填充
 * @param {boolean} grndFillFlg - 是否渐变填充
 * @param {string} fillColor - 填充颜色
 * @param {Object} border - 边框样式
 * @param {string} shpId - 形状ID
 * @param {Function} shapeArcFn - 圆弧路径生成函数
 * @returns {string} SVG路径元素
 */
function renderCustomShape(custShapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcFn) {
    var pathLstNode = PPTXXmlUtils.getTextByPathList(custShapType, ["a:pathLst"]);
    var pathNodes = PPTXXmlUtils.getTextByPathList(pathLstNode, ["a:path"]);

    // 验证 maxX 和 maxY 防止 NaN
    var maxX = 0;
    var maxY = 0;
    if (pathNodes && pathNodes["attrs"]) {
        maxX = parseInt(pathNodes["attrs"]["w"]) || 0;
        maxY = parseInt(pathNodes["attrs"]["h"]) || 0;
    }
    // 确保 maxX 和 maxY 为正数以避免除零
    if (maxX <= 0) maxX = 1;
    if (maxY <= 0) maxY = 1;
    var cX = (1 / maxX) * w;
    var cY = (1 / maxY) * h;

    var moveToNode = PPTXXmlUtils.getTextByPathList(pathNodes, ["a:moveTo"]);
    moveToNode.length;

    var lnToNodes = pathNodes["a:lnTo"];
    var cubicBezToNodes = pathNodes["a:cubicBezTo"];
    var arcToNodes = pathNodes["a:arcTo"];
    var closeNode = PPTXXmlUtils.getTextByPathList(pathNodes, ["a:close"]);

    if (!Array.isArray(moveToNode)) {
        moveToNode = [moveToNode];
    }

    var multiSapeAry = [];
    if (moveToNode.length > 0) {
        // a:moveTo
        Object.keys(moveToNode).forEach(function (key) {
            var moveToPtNode = moveToNode[key]["a:pt"];
            if (moveToPtNode !== undefined) {
                Object.keys(moveToPtNode).forEach(function (key2) {
                    var ptObj = {};
                    var moveToNoPt = moveToPtNode[key2];
                    var spX = moveToNoPt["x"];
                    var spY = moveToNoPt["y"];
                    var ptOrdr = moveToNoPt["order"];
                    ptObj.type = "movto";
                    ptObj.order = ptOrdr;
                    ptObj.x = spX;
                    ptObj.y = spY;
                    multiSapeAry.push(ptObj);
                });
            }
        });

        // a:lnTo
        if (lnToNodes !== undefined) {
            Object.keys(lnToNodes).forEach(function (key) {
                var lnToPtNode = lnToNodes[key]["a:pt"];
                if (lnToPtNode !== undefined) {
                    Object.keys(lnToPtNode).forEach(function (key2) {
                        var ptObj = {};
                        var lnToNoPt = lnToPtNode[key2];
                        var ptX = lnToNoPt["x"];
                        var ptY = lnToNoPt["y"];
                        var ptOrdr = lnToNoPt["order"];
                        ptObj.type = "lnto";
                        ptObj.order = ptOrdr;
                        ptObj.x = ptX;
                        ptObj.y = ptY;
                        multiSapeAry.push(ptObj);
                    });
                }
            });
        }

        // a:cubicBezTo
        if (cubicBezToNodes !== undefined) {
            var cubicBezToPtNodesAry = [];
            if (!Array.isArray(cubicBezToNodes)) {
                cubicBezToNodes = [cubicBezToNodes];
            }
            Object.keys(cubicBezToNodes).forEach(function (key) {
                cubicBezToPtNodesAry.push(cubicBezToNodes[key]["a:pt"]);
            });

            cubicBezToPtNodesAry.forEach(function (key2) {
                var nodeObj = {};
                nodeObj.type = "cubicBezTo";
                nodeObj.order = key2[0]["attrs"]["order"];
                var pts_ary = [];
                key2.forEach(function (pt) {
                    var pt_obj = {
                        x: pt["attrs"]["x"],
                        y: pt["attrs"]["y"]
                    };
                    pts_ary.push(pt_obj);
                });
                nodeObj.cubBzPt = pts_ary;
                multiSapeAry.push(nodeObj);
            });
        }

        // a:quadBezTo
        var quadBezToNodes = pathNodes["a:quadBezTo"];
        if (quadBezToNodes !== undefined) {
            var quadBezToPtNodesAry = [];
            if (!Array.isArray(quadBezToNodes)) {
                quadBezToNodes = [quadBezToNodes];
            }
            Object.keys(quadBezToNodes).forEach(function (key) {
                quadBezToPtNodesAry.push(quadBezToNodes[key]["a:pt"]);
            });

            quadBezToPtNodesAry.forEach(function (key2) {
                var nodeObj = {};
                nodeObj.type = "quadBezTo";
                nodeObj.order = key2[0]["attrs"]["order"];
                var pts_ary = [];
                key2.forEach(function (pt) {
                    var pt_obj = {
                        x: pt["attrs"]["x"],
                        y: pt["attrs"]["y"]
                    };
                    pts_ary.push(pt_obj);
                });
                nodeObj.quadBzPt = pts_ary;
                multiSapeAry.push(nodeObj);
            });
        }

        // a:arcTo
        if (arcToNodes !== undefined) {
            var arcToNodesAttrs = arcToNodes["attrs"];
            var arcOrder = arcToNodesAttrs["order"];
            var hR = arcToNodesAttrs["hR"];
            var wR = arcToNodesAttrs["wR"];
            var stAng = arcToNodesAttrs["stAng"];
            var swAng = arcToNodesAttrs["swAng"];
            var shftX = 0;
            var shftY = 0;
            var arcToPtNode = PPTXXmlUtils.getTextByPathList(arcToNodes, ["a:pt", "attrs"]);
            if (arcToPtNode !== undefined) {
                shftX = arcToPtNode["x"];
                shftY = arcToPtNode["y"];
            }
            var ptObj = {};
            ptObj.type = "arcTo";
            ptObj.order = arcOrder;
            ptObj.hR = hR;
            ptObj.wR = wR;
            ptObj.stAng = stAng;
            ptObj.swAng = swAng;
            ptObj.shftX = shftX;
            ptObj.shftY = shftY;
            multiSapeAry.push(ptObj);
        }

        // a:close
        if (closeNode !== undefined) {
            if (!Array.isArray(closeNode)) {
                closeNode = [closeNode];
            }
            Object.keys(closeNode).forEach(function (key) {
                var clsAttrs = closeNode[key]["attrs"];
                var clsOrder = clsAttrs["order"];
                var ptObj = {};
                ptObj.type = "close";
                ptObj.order = clsOrder;
                multiSapeAry.push(ptObj);
            });
        }

        // 按 order 排序
        multiSapeAry.sort(function (a, b) {
            return a.order - b.order;
        });

        // 生成路径字符串
        var k = 0;
        if (isNaN(cX)) cX = 0;
        if (isNaN(cY)) cY = 0;
        var d = "";
        while (k < multiSapeAry.length) {
            if (multiSapeAry[k].type == "movto") {
                var xVal = parseInt(multiSapeAry[k].x) || 0;
                var yVal = parseInt(multiSapeAry[k].y) || 0;
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                var spX = xVal * cX;
                var spY = yVal * cY;
                d += " M" + spX + "," + spY;
            } else if (multiSapeAry[k].type == "lnto") {
                var xVal = parseInt(multiSapeAry[k].x) || 0;
                var yVal = parseInt(multiSapeAry[k].y) || 0;
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                var Lx = xVal * cX;
                var Ly = yVal * cY;
                d += " L" + Lx + "," + Ly;
            } else if (multiSapeAry[k].type == "cubicBezTo") {
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                var Cx1 = (parseInt(multiSapeAry[k].cubBzPt[0].x) || 0) * cX;
                var Cy1 = (parseInt(multiSapeAry[k].cubBzPt[0].y) || 0) * cY;
                var Cx2 = (parseInt(multiSapeAry[k].cubBzPt[1].x) || 0) * cX;
                var Cy2 = (parseInt(multiSapeAry[k].cubBzPt[1].y) || 0) * cY;
                var Cx3 = (parseInt(multiSapeAry[k].cubBzPt[2].x) || 0) * cX;
                var Cy3 = (parseInt(multiSapeAry[k].cubBzPt[2].y) || 0) * cY;
                d += " C" + Cx1 + "," + Cy1 + " " + Cx2 + "," + Cy2 + " " + Cx3 + "," + Cy3;
            } else if (multiSapeAry[k].type == "arcTo") {
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                var hR = (parseInt(multiSapeAry[k].hR) || 0) * cX;
                var wR = (parseInt(multiSapeAry[k].wR) || 0) * cY;
                var stAng = (parseInt(multiSapeAry[k].stAng) || 0) / 60000;
                var swAng = (parseInt(multiSapeAry[k].swAng) || 0) / 60000;
                if (isNaN(stAng)) stAng = 0;
                if (isNaN(swAng)) swAng = 0;
                var endAng = stAng + swAng;

                if (!isNaN(hR) && !isNaN(wR) && !isNaN(stAng) && !isNaN(swAng)) {
                    d += shapeArcFn(wR, hR, wR, hR, stAng, endAng, false);
                }
            } else if (multiSapeAry[k].type == "quadBezTo") {
                // Quadratic Bezier curve: Q controlPoint endPoint
                // PPTX quadBezTo has 2 points: control point and end point
                // The start point is the previous point in the path
                var quadBzPt = multiSapeAry[k].quadBzPt;
                if (quadBzPt && quadBzPt.length >= 2) {
                    var ctrlX = quadBzPt[0].x * cX;
                    var ctrlY = quadBzPt[0].y * cY;
                    var endX = quadBzPt[1].x * cX;
                    var endY = quadBzPt[1].y * cY;
                    d += "Q" + ctrlX + "," + ctrlY + " " + endX + "," + endY;
                }
            } else if (multiSapeAry[k].type == "close") {
                d += "z";
            }
            k++;
        }

        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' />";
    }

    return "";
}

const SLIDE_FACTOR = 0.0001;

/**
 * Render star shapes (star4, star5, star6, star7, star8, star10, star12, star16, star24, star32)
 * @param {string} shapType - Shape type
 * @param {number} w - Width
 * @param {number} h - Height
 * @param {boolean} imgFillFlg - Image fill flag
 * @param {boolean} grndFillFlg - Gradient fill flag
 * @param {string} fillColor - Fill color
 * @param {object} border - Border object with color, width, strokeDasharray
 * @param {string} shpId - Shape ID
 * @param {object} shapeArcAlt - Shape arc alt
 * @param {object} node - XML node for shape adjustments
 * @returns {string} SVG string
 */
function renderStar(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt, node) {
    let result = '';
    const hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
    const fill = !imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")";

    switch (shapType) {
        case "star4": {
            const adj = getAdjValue(node, "adj", 19098);
            const cnstVal1 = 50000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, cnstVal1);
            const iwd2 = wd2 * a / cnstVal1;
            const ihd2 = hd2 * a / cnstVal1;
            const sdx = iwd2 * Math.cos(0.7853981634);
            const sdy = ihd2 * Math.sin(0.7853981634);
            const sx1 = hc - sdx;
            const sx2 = hc + sdx;
            const sy1 = vc - sdy;
            const sy2 = vc + sdy;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx2 + "," + sy1 +
                " L" + w + "," + vc +
                " L" + sx2 + "," + sy2 +
                " L" + hc + "," + h +
                " L" + sx1 + "," + sy2 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star5": {
            const adj = getAdjValue(node, "adj", 19098);
            const hf = getAdjValue(node, "hf", 105146);
            const vf = getAdjValue(node, "vf", 110557);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const shd2 = hd2 * vf / cnstVal1;
            const svc = vc * vf / cnstVal1;
            const dx1 = swd2 * Math.cos(0.31415926536);
            const dx2 = swd2 * Math.cos(5.3407075111);
            const dy1 = shd2 * Math.sin(0.31415926536);
            const dy2 = shd2 * Math.sin(5.3407075111);
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc + dx2;
            const x4 = hc + dx1;
            const y1 = svc - dy1;
            const y2 = svc - dy2;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = shd2 * a / maxAdj;
            const sdx1 = iwd2 * Math.cos(5.9690260418);
            const sdx2 = iwd2 * Math.cos(0.94247779608);
            const sdy1 = ihd2 * Math.sin(0.94247779608);
            const sdy2 = ihd2 * Math.sin(5.9690260418);
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc + sdx2;
            const sx4 = hc + sdx1;
            const sy1 = svc - sdy1;
            const sy2 = svc - sdy2;
            const sy3 = svc + ihd2;

            const d = "M" + x1 + "," + y1 +
                " L" + sx2 + "," + sy1 +
                " L" + hc + "," + 0 +
                " L" + sx3 + "," + sy1 +
                " L" + x4 + "," + y1 +
                " L" + sx4 + "," + sy2 +
                " L" + x3 + "," + y2 +
                " L" + hc + "," + sy3 +
                " L" + x2 + "," + y2 +
                " L" + sx1 + "," + sy2 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star6": {
            const adj = getAdjValue(node, "adj", 28868);
            const hf = getAdjValue(node, "hf", 115470);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const hd4 = h / 4;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const dx1 = swd2 * Math.cos(0.5235987756);
            const x1 = hc - dx1;
            const x2 = hc + dx1;
            const y2 = vc + hd4;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx2 = iwd2 / 2;
            const sx1 = hc - iwd2;
            const sx2 = hc - sdx2;
            const sx3 = hc + sdx2;
            const sx4 = hc + iwd2;
            const sdy1 = ihd2 * Math.sin(1.0471975512);
            const sy1 = vc - sdy1;
            const sy2 = vc + sdy1;

            const d = "M" + x1 + "," + hd4 +
                " L" + sx2 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx3 + "," + sy1 +
                " L" + x2 + "," + hd4 +
                " L" + sx4 + "," + vc +
                " L" + x2 + "," + y2 +
                " L" + sx3 + "," + sy2 +
                " L" + hc + "," + h +
                " L" + sx2 + "," + sy2 +
                " L" + x1 + "," + y2 +
                " L" + sx1 + "," + vc +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star7": {
            const adj = getAdjValue(node, "adj", 34601);
            const hf = getAdjValue(node, "hf", 102572);
            const vf = getAdjValue(node, "vf", 105210);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const shd2 = hd2 * vf / cnstVal1;
            const svc = vc * vf / cnstVal1;
            const dx1 = swd2 * 97493 / 100000;
            const dx2 = swd2 * 78183 / 100000;
            const dx3 = swd2 * 43388 / 100000;
            const dy1 = shd2 * 62349 / 100000;
            const dy2 = shd2 * 22252 / 100000;
            const dy3 = shd2 * 90097 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc + dx3;
            const x5 = hc + dx2;
            const x6 = hc + dx1;
            const y1 = svc - dy1;
            const y2 = svc + dy2;
            const y3 = svc + dy3;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = shd2 * a / maxAdj;
            const sdx1 = iwd2 * 97493 / 100000;
            const sdx2 = iwd2 * 78183 / 100000;
            const sdx3 = iwd2 * 43388 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc + sdx3;
            const sx5 = hc + sdx2;
            const sx6 = hc + sdx1;
            const sdy1 = ihd2 * 90097 / 100000;
            const sdy2 = ihd2 * 22252 / 100000;
            const sdy3 = ihd2 * 62349 / 100000;
            const sy1 = svc - sdy1;
            const sy2 = svc - sdy2;
            const sy3 = svc + sdy3;
            const sy4 = svc + ihd2;

            const d = "M" + x1 + "," + y2 +
                " L" + sx1 + "," + sy2 +
                " L" + x2 + "," + y1 +
                " L" + sx3 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx4 + "," + sy1 +
                " L" + x5 + "," + y1 +
                " L" + sx6 + "," + sy2 +
                " L" + x6 + "," + y2 +
                " L" + sx5 + "," + sy3 +
                " L" + x4 + "," + y3 +
                " L" + hc + "," + sy4 +
                " L" + x3 + "," + y3 +
                " L" + sx2 + "," + sy3 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star8": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * Math.cos(0.7853981634);
            const x1 = hc - dx1;
            const x2 = hc + dx1;
            const dy1 = hd2 * Math.sin(0.7853981634);
            const y1 = vc - dy1;
            const y2 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 92388 / 100000;
            const sdx2 = iwd2 * 38268 / 100000;
            const sdy1 = ihd2 * 92388 / 100000;
            const sdy2 = ihd2 * 38268 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc + sdx2;
            const sx4 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc + sdy2;
            const sy4 = vc + sdy1;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy2 +
                " L" + x1 + "," + y1 +
                " L" + sx2 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx3 + "," + sy1 +
                " L" + x2 + "," + y1 +
                " L" + sx4 + "," + sy2 +
                " L" + w + "," + vc +
                " L" + sx4 + "," + sy3 +
                " L" + x2 + "," + y2 +
                " L" + sx3 + "," + sy4 +
                " L" + hc + "," + h +
                " L" + sx2 + "," + sy4 +
                " L" + x1 + "," + y2 +
                " L" + sx1 + "," + sy3 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star10": {
            const adj = getAdjValue(node, "adj", 42533);
            const hf = getAdjValue(node, "hf", 105146);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const dx1 = swd2 * 95106 / 100000;
            const dx2 = swd2 * 58779 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc + dx2;
            const x4 = hc + dx1;
            const dy1 = hd2 * 80902 / 100000;
            const dy2 = hd2 * 30902 / 100000;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc + dy2;
            const y4 = vc + dy1;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 80902 / 100000;
            const sdx2 = iwd2 * 30902 / 100000;
            const sdy1 = ihd2 * 95106 / 100000;
            const sdy2 = ihd2 * 58779 / 100000;
            const sx1 = hc - iwd2;
            const sx2 = hc - sdx1;
            const sx3 = hc - sdx2;
            const sx4 = hc + sdx2;
            const sx5 = hc + sdx1;
            const sx6 = hc + iwd2;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc + sdy2;
            const sy4 = vc + sdy1;

            const d = "M" + x1 + "," + y2 +
                " L" + sx2 + "," + sy2 +
                " L" + x2 + "," + y1 +
                " L" + sx3 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx4 + "," + sy1 +
                " L" + x3 + "," + y1 +
                " L" + sx5 + "," + sy2 +
                " L" + x4 + "," + y2 +
                " L" + sx6 + "," + vc +
                " L" + x4 + "," + y3 +
                " L" + sx5 + "," + sy3 +
                " L" + x3 + "," + y4 +
                " L" + sx4 + "," + sy4 +
                " L" + hc + "," + h +
                " L" + sx3 + "," + sy4 +
                " L" + x2 + "," + y4 +
                " L" + sx2 + "," + sy3 +
                " L" + x1 + "," + y3 +
                " L" + sx1 + "," + vc +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star12": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const hd4 = h / 4;
            const wd4 = w / 4;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * Math.cos(0.5235987756);
            const dy1 = hd2 * Math.sin(1.0471975512);
            const x1 = hc - dx1;
            const x3 = w * 3 / 4;
            const x4 = hc + dx1;
            const y1 = vc - dy1;
            const y3 = h * 3 / 4;
            const y4 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * Math.cos(0.2617993878);
            const sdx2 = iwd2 * Math.cos(0.7853981634);
            const sdx3 = iwd2 * Math.cos(1.308996939);
            const sdy1 = ihd2 * Math.sin(1.308996939);
            const sdy2 = ihd2 * Math.sin(0.7853981634);
            const sdy3 = ihd2 * Math.sin(0.2617993878);
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc + sdx3;
            const sx5 = hc + sdx2;
            const sx6 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc + sdy3;
            const sy5 = vc + sdy2;
            const sy6 = vc + sdy1;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy3 +
                " L" + x1 + "," + hd4 +
                " L" + sx2 + "," + sy2 +
                " L" + wd4 + "," + y1 +
                " L" + sx3 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx4 + "," + sy1 +
                " L" + x3 + "," + y1 +
                " L" + sx5 + "," + sy2 +
                " L" + x4 + "," + hd4 +
                " L" + sx6 + "," + sy3 +
                " L" + w + "," + vc +
                " L" + sx6 + "," + sy4 +
                " L" + x4 + "," + y3 +
                " L" + sx5 + "," + sy5 +
                " L" + x3 + "," + y4 +
                " L" + sx4 + "," + sy6 +
                " L" + hc + "," + h +
                " L" + sx3 + "," + sy6 +
                " L" + wd4 + "," + y4 +
                " L" + sx2 + "," + sy5 +
                " L" + x1 + "," + y3 +
                " L" + sx1 + "," + sy4 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star16": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * 92388 / 100000;
            const dx2 = wd2 * 70711 / 100000;
            const dx3 = wd2 * 38268 / 100000;
            const dy1 = hd2 * 92388 / 100000;
            const dy2 = hd2 * 70711 / 100000;
            const dy3 = hd2 * 38268 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc + dx3;
            const x5 = hc + dx2;
            const x6 = hc + dx1;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc - dy3;
            const y4 = vc + dy3;
            const y5 = vc + dy2;
            const y6 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 98079 / 100000;
            const sdx2 = iwd2 * 83147 / 100000;
            const sdx3 = iwd2 * 55557 / 100000;
            const sdx4 = iwd2 * 19509 / 100000;
            const sdy1 = ihd2 * 98079 / 100000;
            const sdy2 = ihd2 * 83147 / 100000;
            const sdy3 = ihd2 * 55557 / 100000;
            const sdy4 = ihd2 * 19509 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc - sdx4;
            const sx5 = hc + sdx4;
            const sx6 = hc + sdx3;
            const sx7 = hc + sdx2;
            const sx8 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc - sdy4;
            const sy5 = vc + sdy4;
            const sy6 = vc + sdy3;
            const sy7 = vc + sdy2;
            const sy8 = vc + sdy1;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy4 +
                " L" + x1 + "," + y3 +
                " L" + sx2 + "," + sy3 +
                " L" + x2 + "," + y2 +
                " L" + sx3 + "," + sy2 +
                " L" + x3 + "," + y1 +
                " L" + sx4 + "," + sy1 +
                " L" + hc + ",0" +
                " L" + sx5 + "," + sy1 +
                " L" + x4 + "," + y1 +
                " L" + sx6 + "," + sy2 +
                " L" + x5 + "," + y2 +
                " L" + sx7 + "," + sy3 +
                " L" + x6 + "," + y3 +
                " L" + sx8 + "," + sy4 +
                " L" + w + "," + vc +
                " L" + sx8 + "," + sy5 +
                " L" + x6 + "," + y4 +
                " L" + sx7 + "," + sy6 +
                " L" + x5 + "," + y5 +
                " L" + sx6 + "," + sy7 +
                " L" + x4 + "," + y6 +
                " L" + sx5 + "," + sy8 +
                " L" + hc + "," + h +
                " L" + sx4 + "," + sy8 +
                " L" + x3 + "," + y6 +
                " L" + sx3 + "," + sy7 +
                " L" + x2 + "," + y5 +
                " L" + sx2 + "," + sy6 +
                " L" + x1 + "," + y4 +
                " L" + sx1 + "," + sy5 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star24": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const hd4 = h / 4;
            const wd4 = w / 4;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * Math.cos(0.2617993878);
            const dx2 = wd2 * Math.cos(0.5235987756);
            const dx3 = wd2 * Math.cos(0.7853981634);
            const dx4 = wd4;
            const dx5 = wd2 * Math.cos(1.308996939);
            const dy1 = hd2 * Math.sin(1.308996939);
            const dy2 = hd2 * Math.sin(1.0471975512);
            const dy3 = hd2 * Math.sin(0.7853981634);
            const dy4 = hd4;
            const dy5 = hd2 * Math.sin(0.2617993878);
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc - dx4;
            const x5 = hc - dx5;
            const x6 = hc + dx5;
            const x7 = hc + dx4;
            const x8 = hc + dx3;
            const x9 = hc + dx2;
            const x10 = hc + dx1;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc - dy3;
            const y4 = vc - dy4;
            const y5 = vc - dy5;
            const y6 = vc + dy5;
            const y7 = vc + dy4;
            const y8 = vc + dy3;
            const y9 = vc + dy2;
            const y10 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 99144 / 100000;
            const sdx2 = iwd2 * 92388 / 100000;
            const sdx3 = iwd2 * 79335 / 100000;
            const sdx4 = iwd2 * 60876 / 100000;
            const sdx5 = iwd2 * 38268 / 100000;
            const sdx6 = iwd2 * 13053 / 100000;
            const sdy1 = ihd2 * 99144 / 100000;
            const sdy2 = ihd2 * 92388 / 100000;
            const sdy3 = ihd2 * 79335 / 100000;
            const sdy4 = ihd2 * 60876 / 100000;
            const sdy5 = ihd2 * 38268 / 100000;
            const sdy6 = ihd2 * 13053 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc - sdx4;
            const sx5 = hc - sdx5;
            const sx6 = hc - sdx6;
            const sx7 = hc + sdx6;
            const sx8 = hc + sdx5;
            const sx9 = hc + sdx4;
            const sx10 = hc + sdx3;
            const sx11 = hc + sdx2;
            const sx12 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc - sdy4;
            const sy5 = vc - sdy5;
            const sy6 = vc - sdy6;
            const sy7 = vc + sdy6;
            const sy8 = vc + sdy5;
            const sy9 = vc + sdy4;
            const sy10 = vc + sdy3;
            const sy11 = vc + sdy2;
            const sy12 = vc + sdy1;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy6 +
                " L" + x1 + "," + y5 +
                " L" + sx2 + "," + sy5 +
                " L" + x2 + "," + y4 +
                " L" + sx3 + "," + sy4 +
                " L" + x3 + "," + y3 +
                " L" + sx4 + "," + sy3 +
                " L" + x4 + "," + y2 +
                " L" + sx5 + "," + sy2 +
                " L" + x5 + "," + y1 +
                " L" + sx6 + "," + sy1 +
                " L" + hc + "," + 0 +
                " L" + sx7 + "," + sy1 +
                " L" + x6 + "," + y1 +
                " L" + sx8 + "," + sy2 +
                " L" + x7 + "," + y2 +
                " L" + sx9 + "," + sy3 +
                " L" + x8 + "," + y3 +
                " L" + sx10 + "," + sy4 +
                " L" + x9 + "," + y4 +
                " L" + sx11 + "," + sy5 +
                " L" + x10 + "," + y5 +
                " L" + sx12 + "," + sy6 +
                " L" + w + "," + vc +
                " L" + sx12 + "," + sy7 +
                " L" + x10 + "," + y6 +
                " L" + sx11 + "," + sy8 +
                " L" + x9 + "," + y7 +
                " L" + sx10 + "," + sy9 +
                " L" + x8 + "," + y8 +
                " L" + sx9 + "," + sy10 +
                " L" + x7 + "," + y9 +
                " L" + sx8 + "," + sy11 +
                " L" + x6 + "," + y10 +
                " L" + sx7 + "," + sy12 +
                " L" + hc + "," + h +
                " L" + sx6 + "," + sy12 +
                " L" + x5 + "," + y10 +
                " L" + sx5 + "," + sy11 +
                " L" + x4 + "," + y9 +
                " L" + sx4 + "," + sy10 +
                " L" + x3 + "," + y8 +
                " L" + sx3 + "," + sy9 +
                " L" + x2 + "," + y7 +
                " L" + sx2 + "," + sy8 +
                " L" + x1 + "," + y6 +
                " L" + sx1 + "," + sy7 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
        case "star32": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * 98079 / 100000;
            const dx2 = wd2 * 92388 / 100000;
            const dx3 = wd2 * 83147 / 100000;
            const dx4 = wd2 * Math.cos(0.7853981634);
            const dx5 = wd2 * 55557 / 100000;
            const dx6 = wd2 * 38268 / 100000;
            const dx7 = wd2 * 19509 / 100000;
            const dy1 = hd2 * 98079 / 100000;
            const dy2 = hd2 * 92388 / 100000;
            const dy3 = hd2 * 83147 / 100000;
            const dy4 = hd2 * Math.sin(0.7853981634);
            const dy5 = hd2 * 55557 / 100000;
            const dy6 = hd2 * 38268 / 100000;
            const dy7 = hd2 * 19509 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc - dx4;
            const x5 = hc - dx5;
            const x6 = hc - dx6;
            const x7 = hc - dx7;
            const x8 = hc + dx7;
            const x9 = hc + dx6;
            const x10 = hc + dx5;
            const x11 = hc + dx4;
            const x12 = hc + dx3;
            const x13 = hc + dx2;
            const x14 = hc + dx1;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc - dy3;
            const y4 = vc - dy4;
            const y5 = vc - dy5;
            const y6 = vc - dy6;
            const y7 = vc - dy7;
            const y8 = vc + dy7;
            const y9 = vc + dy6;
            const y10 = vc + dy5;
            const y11 = vc + dy4;
            const y12 = vc + dy3;
            const y13 = vc + dy2;
            const y14 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 99518 / 100000;
            const sdx2 = iwd2 * 95694 / 100000;
            const sdx3 = iwd2 * 88192 / 100000;
            const sdx4 = iwd2 * 77301 / 100000;
            const sdx5 = iwd2 * 63439 / 100000;
            const sdx6 = iwd2 * 47140 / 100000;
            const sdx7 = iwd2 * 29028 / 100000;
            const sdx8 = iwd2 * 9802 / 100000;
            const sdy1 = ihd2 * 99518 / 100000;
            const sdy2 = ihd2 * 95694 / 100000;
            const sdy3 = ihd2 * 88192 / 100000;
            const sdy4 = ihd2 * 77301 / 100000;
            const sdy5 = ihd2 * 63439 / 100000;
            const sdy6 = ihd2 * 47140 / 100000;
            const sdy7 = ihd2 * 29028 / 100000;
            const sdy8 = ihd2 * 9802 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc - sdx4;
            const sx5 = hc - sdx5;
            const sx6 = hc - sdx6;
            const sx7 = hc - sdx7;
            const sx8 = hc - sdx8;
            const sx9 = hc + sdx8;
            const sx10 = hc + sdx7;
            const sx11 = hc + sdx6;
            const sx12 = hc + sdx5;
            const sx13 = hc + sdx4;
            const sx14 = hc + sdx3;
            const sx15 = hc + sdx2;
            const sx16 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc - sdy4;
            const sy5 = vc - sdy5;
            const sy6 = vc - sdy6;
            const sy7 = vc - sdy7;
            const sy8 = vc - sdy8;
            const sy9 = vc + sdy8;
            const sy10 = vc + sdy7;
            const sy11 = vc + sdy6;
            const sy12 = vc + sdy5;
            const sy13 = vc + sdy4;
            const sy14 = vc + sdy3;
            const sy15 = vc + sdy2;
            const sy16 = vc + sdy1;

            const d = "M0" + "," + vc +
                " L" + sx1 + "," + sy8 +
                " L" + x1 + "," + y7 +
                " L" + sx2 + "," + sy7 +
                " L" + x2 + "," + y6 +
                " L" + sx3 + "," + sy6 +
                " L" + x3 + "," + y5 +
                " L" + sx4 + "," + sy5 +
                " L" + x4 + "," + y4 +
                " L" + sx5 + "," + sy4 +
                " L" + x5 + "," + y3 +
                " L" + sx6 + "," + sy3 +
                " L" + x6 + "," + y2 +
                " L" + sx7 + "," + sy2 +
                " L" + x7 + "," + y1 +
                " L" + sx8 + "," + sy1 +
                " L" + hc + "," + 0 +
                " L" + sx9 + "," + sy1 +
                " L" + x8 + "," + y1 +
                " L" + sx10 + "," + sy2 +
                " L" + x9 + "," + y2 +
                " L" + sx11 + "," + sy3 +
                " L" + x10 + "," + y3 +
                " L" + sx12 + "," + sy4 +
                " L" + x11 + "," + y4 +
                " L" + sx13 + "," + sy5 +
                " L" + x12 + "," + y5 +
                " L" + sx14 + "," + sy6 +
                " L" + x13 + "," + y6 +
                " L" + sx15 + "," + sy7 +
                " L" + x14 + "," + y7 +
                " L" + sx16 + "," + sy8 +
                " L" + w + "," + vc +
                " L" + sx16 + "," + sy9 +
                " L" + x14 + "," + y8 +
                " L" + sx15 + "," + sy10 +
                " L" + x13 + "," + y9 +
                " L" + sx14 + "," + sy11 +
                " L" + x12 + "," + y10 +
                " L" + sx13 + "," + sy12 +
                " L" + x11 + "," + y11 +
                " L" + sx12 + "," + sy13 +
                " L" + x10 + "," + y12 +
                " L" + sx11 + "," + sy14 +
                " L" + x9 + "," + y13 +
                " L" + sx10 + "," + sy15 +
                " L" + x8 + "," + y14 +
                " L" + sx9 + "," + sy16 +
                " L" + hc + "," + h +
                " L" + sx8 + "," + sy16 +
                " L" + x7 + "," + y14 +
                " L" + sx7 + "," + sy15 +
                " L" + x6 + "," + y13 +
                " L" + sx6 + "," + sy14 +
                " L" + x5 + "," + y12 +
                " L" + sx5 + "," + sy13 +
                " L" + x4 + "," + y11 +
                " L" + sx4 + "," + sy12 +
                " L" + x3 + "," + y10 +
                " L" + sx3 + "," + sy11 +
                " L" + x2 + "," + y9 +
                " L" + sx2 + "," + sy10 +
                " L" + x1 + "," + y8 +
                " L" + sx1 + "," + sy9 +
                " z";

            result += "<path d='" + d + "' fill='" + fill + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            break;
        }
    }

    return result;
}

/**
 * Get adjustment value from node
 * @param {object} node - XML node
 * @param {string} name - Adjustment name
 * @param {number} defaultValue - Default value
 * @returns {number}
 */
function getAdjValue(node, name, defaultValue) {
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    if (shapAdjst !== undefined) {
        if (Array.isArray(shapAdjst)) {
            for (let key of Object.keys(shapAdjst)) {
                if (shapAdjst[key] && shapAdjst[key]["attrs"] && shapAdjst[key]["attrs"]["name"] === name) {
                    return parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * SLIDE_FACTOR;
                }
            }
        } else if (shapAdjst["attrs"] && shapAdjst["attrs"]["name"] === name) {
            return parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * SLIDE_FACTOR;
        }
    }
    return defaultValue * SLIDE_FACTOR;
}

/**
 * Clamp value between min and max
 * @param {number} value - Value to clamp
 * @param {number} min - Minimum value
 * @param {number} max - Maximum value
 * @returns {number}
 */
function clamp(value, min, max) {
    return value < min ? min : value > max ? max : value;
}

/**
 * 数学符号形状渲染模块
 * 提供数学符号（加减乘除等）的生成和渲染功能
 */


/**
 * 渲染数学符号形状
 */
function renderMathSymbol(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let result = "";

    // 获取形状调整参数
    var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    var sAdj1, adj1;
    var sAdj2, adj2;
    var sAdj3, adj3;
    if (shapAdjst_ary !== undefined) {
        if (shapAdjst_ary.constructor === Array) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4));
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4));
                } else if (sAdj_name == "adj3") {
                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj3 = parseInt(sAdj3.substr(4));
                }
            }
        } else {
            sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
            adj1 = parseInt(sAdj1.substr(4));
        }
    }
    var cnstVal1 = 50000 * SLIDE_FACTOR$1;
    var cnstVal2 = 100000 * SLIDE_FACTOR$1;
    var cnstVal3 = 200000 * SLIDE_FACTOR$1;
    var dVal;
    var hc = w / 2, vc = h / 2, hd2 = h / 2;

    // mathNotEqual (不等于符号)
    if (shapType == "mathNotEqual") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
            adj2 = 110 * Math.PI / 180;
            adj3 = 11760 * SLIDE_FACTOR$1;
        } else {
            adj1 = adj1 * SLIDE_FACTOR$1;
            adj2 = (adj2 / 60000) * Math.PI / 180;
            adj3 = adj3 * SLIDE_FACTOR$1;
        }
        var a1, crAng, a2a1, maxAdj3, a3, dy1, dy2, dx1, x1, x8, y2, y3, y1, y4,
            cadj2, xadj2, len, bhw, bhw2, x7, dx67, x6, dx57, x5, dx47, x4, dx37,
            x3, dx27, x2, rx7, rx6, rx5, rx4, rx3, dx7, rxt, lxt, rx, lx,
            dy3, dy4, ry, ly, dlx, drx, dly, dry;
        var angVal1 = 70 * Math.PI / 180, angVal2 = 110 * Math.PI / 180;
        var cnstVal4 = 73490 * SLIDE_FACTOR$1;
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
        crAng = (adj2 < angVal1) ? angVal1 : (adj2 > angVal2) ? angVal2 : adj2;
        a2a1 = a1 * 2;
        maxAdj3 = cnstVal2 - a2a1;
        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
        dy1 = h * a1 / cnstVal2;
        dy2 = h * a3 / cnstVal3;
        dx1 = w * cnstVal4 / cnstVal3;
        x1 = hc - dx1;
        x8 = hc + dx1;
        y2 = vc - dy2;
        y3 = vc + dy2;
        y1 = y2 - dy1;
        y4 = y3 + dy1;
        cadj2 = crAng - Math.PI / 2;
        xadj2 = hd2 * Math.tan(cadj2);
        len = Math.sqrt(xadj2 * xadj2 + hd2 * hd2);
        bhw = len * dy1 / hd2;
        bhw2 = bhw / 2;
        x7 = hc + xadj2 - bhw2;
        dx67 = xadj2 * y1 / hd2;
        x6 = x7 - dx67;
        dx57 = xadj2 * y2 / hd2;
        x5 = x7 - dx57;
        dx47 = xadj2 * y3 / hd2;
        x4 = x7 - dx47;
        dx37 = xadj2 * y4 / hd2;
        x3 = x7 - dx37;
        dx27 = xadj2 * 2;
        x2 = x7 - dx27;
        rx7 = x7 + bhw;
        rx6 = x6 + bhw;
        rx5 = x5 + bhw;
        rx4 = x4 + bhw;
        rx3 = x3 + bhw;
        dx7 = dy1 * hd2 / len;
        rxt = x7 + dx7;
        lxt = rx7 - dx7;
        rx = (cadj2 > 0) ? rxt : rx7;
        lx = (cadj2 > 0) ? x7 : lxt;
        dy3 = dy1 * xadj2 / len;
        dy4 = -dy3;
        ry = (cadj2 > 0) ? dy3 : 0;
        ly = (cadj2 > 0) ? 0 : dy4;
        dlx = w - rx;
        drx = w - lx;
        dly = h - ry;
        dry = h - ly;

        dVal = "M" + x1 + "," + y1 +
            " L" + x6 + "," + y1 +
            " L" + lx + "," + ly +
            " L" + rx + "," + ry +
            " L" + rx6 + "," + y1 +
            " L" + x8 + "," + y1 +
            " L" + x8 + "," + y2 +
            " L" + rx5 + "," + y2 +
            " L" + rx4 + "," + y3 +
            " L" + x8 + "," + y3 +
            " L" + x8 + "," + y4 +
            " L" + rx3 + "," + y4 +
            " L" + drx + "," + dry +
            " L" + dlx + "," + dly +
            " L" + x3 + "," + y4 +
            " L" + x1 + "," + y4 +
            " L" + x1 + "," + y3 +
            " L" + x4 + "," + y3 +
            " L" + x5 + "," + y2 +
            " L" + x1 + "," + y2 +
            " z";
    } 
    // mathDivide (除号)
    else if (shapType == "mathDivide") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
            adj2 = 5880 * SLIDE_FACTOR$1;
            adj3 = 11760 * SLIDE_FACTOR$1;
        } else {
            adj1 = adj1 * SLIDE_FACTOR$1;
            adj2 = adj2 * SLIDE_FACTOR$1;
            adj3 = adj3 * SLIDE_FACTOR$1;
        }
        var a1, ma1, ma3h, ma3w, maxAdj3, a3, m4a3, maxAdj2, a2, dy1, yg, rad, dx1,
            y3, y4, a, y2, y1, y5, x1, x3, x2;
        var cnstVal4 = 1000 * SLIDE_FACTOR$1;
        var cnstVal5 = 36745 * SLIDE_FACTOR$1;
        var cnstVal6 = 73490 * SLIDE_FACTOR$1;
        a1 = (adj1 < cnstVal4) ? cnstVal4 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
        ma1 = -a1;
        ma3h = (cnstVal6 + ma1) / 4;
        ma3w = cnstVal5 * w / h;
        maxAdj3 = (ma3h < ma3w) ? ma3h : ma3w;
        a3 = (adj3 < cnstVal4) ? cnstVal4 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
        m4a3 = -4 * a3;
        maxAdj2 = cnstVal6 + m4a3 - a1;
        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
        dy1 = h * a1 / cnstVal3;
        yg = h * a2 / cnstVal2;
        rad = h * a3 / cnstVal2;
        dx1 = w * cnstVal6 / cnstVal3;
        y3 = vc - dy1;
        y4 = vc + dy1;
        a = yg + rad;
        y2 = y3 - a;
        y1 = y2 - rad;
        y5 = h - y1;
        x1 = hc - dx1;
        x3 = hc + dx1;
        x2 = hc - rad;
        var cd4 = 90, c3d4 = 270;
        var cX1 = hc - Math.cos(c3d4 * Math.PI / 180) * rad;
        var cY1 = y1 - Math.sin(c3d4 * Math.PI / 180) * rad;
        var cX2 = hc - Math.cos(Math.PI / 2) * rad;
        var cY2 = y5 - Math.sin(Math.PI / 2) * rad;
            dVal = "M" + hc + "," + y1 +
                shapeArc(cX1, cY1, rad, rad, c3d4, c3d4 + 360, false).replace("M", "L") +
                " z" +
                " M" + hc + "," + y5 +
                shapeArc(cX2, cY2, rad, rad, cd4, cd4 + 360, false).replace("M", "L") +
                " z" +
            " M" + x1 + "," + y3 +
            " L" + x3 + "," + y3 +
            " L" + x3 + "," + y4 +
            " L" + x1 + "," + y4 +
            " z";
    } 
    // mathEqual (等号)
    else if (shapType == "mathEqual") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
            adj2 = 11760 * SLIDE_FACTOR$1;
        } else {
            adj1 = adj1 * SLIDE_FACTOR$1;
            adj2 = adj2 * SLIDE_FACTOR$1;
        }
        var cnstVal5 = 36745 * SLIDE_FACTOR$1;
        var cnstVal6 = 73490 * SLIDE_FACTOR$1;
        var a1, a2a1, mAdj2, a2, dy1, dy2, dx1, y2, y3, y1, y4, x1, x2;

        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
        a2a1 = a1 * 2;
        mAdj2 = cnstVal2 - a2a1;
        a2 = (adj2 < 0) ? 0 : (adj2 > mAdj2) ? mAdj2 : adj2;
        dy1 = h * a1 / cnstVal2;
        dy2 = h * a2 / cnstVal3;
        dx1 = w * cnstVal6 / cnstVal3;
        y2 = vc - dy2;
        y3 = vc + dy2;
        y1 = y2 - dy1;
        y4 = y3 + dy1;
        x1 = hc - dx1;
        x2 = hc + dx1;
        dVal = "M" + x1 + "," + y1 +
            " L" + x2 + "," + y1 +
            " L" + x2 + "," + y2 +
            " L" + x1 + "," + y2 +
            " z" +
            "M" + x1 + "," + y3 +
            " L" + x2 + "," + y3 +
            " L" + x2 + "," + y4 +
            " L" + x1 + "," + y4 +
            " z";
    } 
    // mathMinus (减号)
    else if (shapType == "mathMinus") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
        } else {
            adj1 = adj1 * SLIDE_FACTOR$1;
        }
        var cnstVal6 = 73490 * SLIDE_FACTOR$1;
        var a1, dy1, dx1, y1, y2, x1, x2;
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
        dy1 = h * a1 / cnstVal3;
        dx1 = w * cnstVal6 / cnstVal3;
        y1 = vc - dy1;
        y2 = vc + dy1;
        x1 = hc - dx1;
        x2 = hc + dx1;

        dVal = "M" + x1 + "," + y1 +
            " L" + x2 + "," + y1 +
            " L" + x2 + "," + y2 +
            " L" + x1 + "," + y2 +
            " z";
    } 
    // mathMultiply (乘号)
    else if (shapType == "mathMultiply") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
        } else {
            adj1 = adj1 * SLIDE_FACTOR$1;
        }
        var cnstVal6 = 51965 * SLIDE_FACTOR$1;
        var a1, th, a, sa, ca, ta, dl, rw, lM, xM, yM, dxAM, dyAM,
            xA, yA, xB, yB, xBC, yBC, yC, xD, xE, yFE, xFE, xF, xL, yG, yH, yI;
        var ss = Math.min(w, h);
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal6) ? cnstVal6 : adj1;
        th = ss * a1 / cnstVal2;
        a = Math.atan(h / w);
        sa = 1 * Math.sin(a);
        ca = 1 * Math.cos(a);
        ta = 1 * Math.tan(a);
        dl = Math.sqrt(w * w + h * h);
        rw = dl * cnstVal6 / cnstVal2;
        lM = dl - rw;
        xM = ca * lM / 2;
        yM = sa * lM / 2;
        dxAM = sa * th / 2;
        dyAM = ca * th / 2;
        xA = xM - dxAM;
        yA = yM + dyAM;
        xB = xM + dxAM;
        yB = yM - dyAM;
        xBC = hc - xB;
        yBC = xBC * ta;
        yC = yBC + yB;
        xD = w - xB;
        xE = w - xA;
        yFE = vc - yA;
        xFE = yFE / ta;
        xF = xE - xFE;
        xL = xA + xFE;
        yG = h - yA;
        yH = h - yB;
        yI = h - yC;

        dVal = "M" + xA + "," + yA +
            " L" + xB + "," + yB +
            " L" + hc + "," + yC +
            " L" + xD + "," + yB +
            " L" + xE + "," + yA +
            " L" + xF + "," + vc +
            " L" + xE + "," + yG +
            " L" + xD + "," + yH +
            " L" + hc + "," + yI +
            " L" + xB + "," + yH +
            " L" + xA + "," + yG +
            " L" + xL + "," + vc +
            " z";
    } 
    // mathPlus (加号)
    else if (shapType == "mathPlus") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
        } else {
            adj1 = adj1 * SLIDE_FACTOR$1;
        }
        var cnstVal6 = 73490 * SLIDE_FACTOR$1;
        var ss = Math.min(w, h);
        var a1, dx1, dy1, dx2, x1, x2, x3, x4, y1, y2, y3, y4;

        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal6) ? cnstVal6 : adj1;
        dx1 = w * cnstVal6 / cnstVal3;
        dy1 = h * cnstVal6 / cnstVal3;
        dx2 = ss * a1 / cnstVal3;
        x1 = hc - dx1;
        x2 = hc - dx2;
        x3 = hc + dx2;
        x4 = hc + dx1;
        y1 = vc - dy1;
        y2 = vc - dx2;
        y3 = vc + dx2;
        y4 = vc + dy1;

        dVal = "M" + x1 + "," + y2 +
            " L" + x2 + "," + y2 +
            " L" + x2 + "," + y1 +
            " L" + x3 + "," + y1 +
            " L" + x3 + "," + y2 +
            " L" + x4 + "," + y2 +
            " L" + x4 + "," + y3 +
            " L" + x3 + "," + y3 +
            " L" + x3 + "," + y4 +
            " L" + x2 + "," + y4 +
            " L" + x2 + "," + y3 +
            " L" + x1 + "," + y3 +
            " z";
    }

    result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

    return result;
}

/**
 * 括号形状渲染模块
 * 提供各种括号形状（大括号、方括号等）的生成和渲染功能
 */


/**
 * 渲染括号形状
 */
function renderBracket(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let result = "";
    let dVal = "";

    if (shapType === "bracePair") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var adj = 8333 * SLIDE_FACTOR$1;
        var cnstVal1 = 25000 * SLIDE_FACTOR$1;
        var cnstVal2 = 50000 * SLIDE_FACTOR$1;
        var cnstVal3 = 100000 * SLIDE_FACTOR$1;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
        }
        var vc = h / 2, cd = 360, cd2 = 180, cd4 = 90, c3d4 = 270, a, x1, x2, x3, x4, y2, y3, y4;
        if (adj < 0) a = 0;
        else if (adj > cnstVal1) a = cnstVal1;
        else a = adj;
        var minWH = Math.min(w, h);
        x1 = minWH * a / cnstVal3;
        x2 = minWH * a / cnstVal2;
        x3 = w - x2;
        x4 = w - x1;
        y2 = vc - x1;
        y3 = vc + x1;
        y4 = h - x1;
        dVal = "M" + x2 + "," + h +
            shapeArc(x2, y4, x1, x1, cd4, cd2, false).replace("M", "L") +
            " L" + x1 + "," + y3 +
            shapeArc(0, y3, x1, x1, 0, (-cd4), false).replace("M", "L") +
            shapeArc(0, y2, x1, x1, cd4, 0, false).replace("M", "L") +
            " L" + x1 + "," + x1 +
            shapeArc(x2, x1, x1, x1, cd2, c3d4, false).replace("M", "L") +
            " M" + x3 + "," + 0 +
            shapeArc(x3, x1, x1, x1, c3d4, cd, false).replace("M", "L") +
            " L" + x4 + "," + y2 +
            shapeArc(w, y2, x1, x1, cd2, cd4, false).replace("M", "L") +
            shapeArc(w, y3, x1, x1, c3d4, cd2, false).replace("M", "L") +
            " L" + x4 + "," + y4 +
            shapeArc(x3, y4, x1, x1, 0, cd4, false).replace("M", "L");
    }
    else if (shapType === "leftBrace") {
        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, adj1 = 8333 * SLIDE_FACTOR$1;
        var sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                }
            }
        }
        var vc = h / 2, cd2 = 180, cd4 = 90, c3d4 = 270, a1, a2, q1, q2, q3, y1, y2, y3, y4;
        if (adj2 < 0) a2 = 0;
        else if (adj2 > cnstVal2) a2 = cnstVal2;
        else a2 = adj2;
        var minWH = Math.min(w, h);
        q1 = cnstVal2 - a2;
        if (q1 < a2) q2 = q1;
        else q2 = a2;
        q3 = q2 / 2;
        var maxAdj1 = q3 * h / minWH;
        if (adj1 < 0) a1 = 0;
        else if (adj1 > maxAdj1) a1 = maxAdj1;
        else a1 = adj1;
        y1 = minWH * a1 / cnstVal2;
        y3 = h * a2 / cnstVal2;
        y2 = y3 - y1;
        y4 = y3 + y1;
        dVal = "M" + w + "," + h +
            shapeArc(w, h - y1, w / 2, y1, cd4, cd2, false).replace("M", "L") +
            " L" + w / 2 + "," + y4 +
            shapeArc(0, y4, w / 2, y1, 0, (-cd4), false).replace("M", "L") +
            shapeArc(0, y2, w / 2, y1, cd4, 0, false).replace("M", "L") +
            " L" + w / 2 + "," + y1 +
            shapeArc(w, y1, w / 2, y1, cd2, c3d4, false).replace("M", "L");
    }
    else if (shapType === "rightBrace") {
        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, adj1 = 8333 * SLIDE_FACTOR$1;
        var sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                }
            }
        }
        var vc = h / 2, cd = 360, cd2 = 180, cd4 = 90, c3d4 = 270, a1, a2, q1, q2, q3, y1, y2, y3, y4;
        if (adj2 < 0) a2 = 0;
        else if (adj2 > cnstVal2) a2 = cnstVal2;
        else a2 = adj2;
        var minWH = Math.min(w, h);
        q1 = cnstVal2 - a2;
        if (q1 < a2) q2 = q1;
        else q2 = a2;
        q3 = q2 / 2;
        var maxAdj1 = q3 * h / minWH;
        if (adj1 < 0) a1 = 0;
        else if (adj1 > maxAdj1) a1 = maxAdj1;
        else a1 = adj1;
        y1 = minWH * a1 / cnstVal2;
        y3 = h * a2 / cnstVal2;
        y2 = y3 - y1;
        y4 = h - y1;
        dVal = "M" + 0 + "," + 0 +
            shapeArc(0, y1, w / 2, y1, c3d4, cd, false).replace("M", "L") +
            " L" + w / 2 + "," + y2 +
            shapeArc(w, y2, w / 2, y1, cd2, cd4, false).replace("M", "L") +
            shapeArc(w, y3 + y1, w / 2, y1, c3d4, cd2, false).replace("M", "L") +
            " L" + w / 2 + "," + y4 +
            shapeArc(0, y4, w / 2, y1, 0, cd4, false).replace("M", "L");
    }
    else if (shapType === "bracketPair") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var adj = 16667 * SLIDE_FACTOR$1;
        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
        }
        var r = w, b = h, cd2 = 180, cd4 = 90, c3d4 = 270, a, x1, x2, y2;
        if (adj < 0) a = 0;
        else if (adj > cnstVal1) a = cnstVal1;
        else a = adj;
        x1 = Math.min(w, h) * a / cnstVal2;
        x2 = r - x1;
        y2 = b - x1;
        dVal = shapeArc(x1, x1, x1, x1, c3d4, cd2, false) +
            shapeArc(x1, y2, x1, x1, cd2, cd4, false).replace("M", "L") +
            shapeArc(x2, x1, x1, x1, c3d4, (c3d4 + cd4), false) +
            shapeArc(x2, y2, x1, x1, 0, cd4, false).replace("M", "L");
    }
    else if (shapType === "leftBracket") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var adj = 8333 * SLIDE_FACTOR$1;
        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
        var maxAdj = cnstVal1 * h / Math.min(w, h);
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
        }
        var r = w, b = h, cd2 = 180, cd4 = 90, c3d4 = 270, a, y1, y2;
        if (adj < 0) a = 0;
        else if (adj > maxAdj) a = maxAdj;
        else a = adj;
        y1 = Math.min(w, h) * a / cnstVal2;
        if (y1 > w) y1 = w;
        y2 = b - y1;
        dVal = "M" + r + "," + b +
            shapeArc(y1, y2, y1, y1, cd4, cd2, false).replace("M", "L") +
            " L" + 0 + "," + y1 +
            shapeArc(y1, y1, y1, y1, cd2, c3d4, false).replace("M", "L") +
            " L" + r + "," + 0;
    }
    else if (shapType === "rightBracket") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var adj = 8333 * SLIDE_FACTOR$1;
        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
        var maxAdj = cnstVal1 * h / Math.min(w, h);
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
        }
        var cd = 360, cd2 = 180, cd4 = 90, c3d4 = 270, a, y1, y2, y3;
        if (adj < 0) a = 0;
        else if (adj > maxAdj) a = maxAdj;
        else a = adj;
        y1 = Math.min(w, h) * a / cnstVal2;
        y2 = h - y1;
        y3 = w - y1;
        dVal = "M" + 0 + "," + h +
            shapeArc(y3, y2, y1, y1, cd4, 0, false).replace("M", "L") +
            " L" + w + "," + h / 2 +
            shapeArc(y3, y1, y1, y1, cd, c3d4, false).replace("M", "L") +
            " L" + 0 + "," + 0;
    }

    result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

    return result;
}

/**
 * 杂项形状渲染模块
 * 包含 smileyFace、scroll 等独立形状
 */


/**
 * 渲染杂项形状
 */
function renderMiscShape(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let result = "";
    let dVal = "";

    if (shapType === "smileyFace") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var refr = SLIDE_FACTOR$1;
        var adj = 4653 * refr;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * refr;
        }
        var cnstVal1 = 50000 * refr;
        var cnstVal2 = 100000 * refr;
        var cnstVal3 = 4653 * refr;
        var ss = Math.min(w, h);
        var a, x1, x2, x3, x4, y1, y3, dy2, y2, y4, dy3, y5, wR, hR, wd2, hd2;
        wd2 = w / 2;
        hd2 = h / 2;
        a = (adj < -cnstVal3) ? -cnstVal3 : (adj > cnstVal3) ? cnstVal3 : adj;
        x1 = w * 4969 / 21699;
        x2 = w * 6215 / 21600;
        x3 = w * 13135 / 21600;
        x4 = w * 16640 / 21600;
        y1 = h * 7570 / 21600;
        y3 = h * 16515 / 21600;
        dy2 = h * a / cnstVal2;
        y2 = y3 - dy2;
        y4 = y3 + dy2;
        dy3 = h * a / cnstVal1;
        y5 = y4 + dy3;
        wR = w * 1125 / 21600;
        hR = h * 1125 / 21600;
        var cX1 = x2 - wR * Math.cos(Math.PI);
        var cY1 = y1 - hR * Math.sin(Math.PI);
        var cX2 = x3 - wR * Math.cos(Math.PI);
        dVal = //eyes
            shapeArc(cX1, cY1, wR, hR, 180, 540, false) +
            shapeArc(cX2, cY1, wR, hR, 180, 540, false) +
            //mouth
            " M" + x1 + "," + y2 +
            " Q" + wd2 + "," + y5 + " " + x4 + "," + y2 +
            " Q" + wd2 + "," + y5 + " " + x1 + "," + y2 +
            //head
            " M" + 0 + "," + hd2 +
            shapeArc(wd2, hd2, wd2, hd2, 180, 540, false).replace("M", "L") +
            " z";
    }
    else if (shapType === "verticalScroll" || shapType === "horizontalScroll") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var refr = SLIDE_FACTOR$1;
        var adj = 12500 * refr;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * refr;
        }
        var cnstVal1 = 25000 * refr;
        var cnstVal2 = 100000 * refr;
        var ss = Math.min(w, h);
        var t = 0, l = 0, b = h, r = w;
        var a, ch, ch2, ch4;
        a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
        ch = ss * a / cnstVal2;
        ch2 = ch / 2;
        ch4 = ch / 4;
        if (shapType === "verticalScroll") {
            var x3, x4, x6, x7, x5, y3, y4;
            x3 = ch + ch2;
            x4 = ch + ch;
            x6 = r - ch;
            x7 = r - ch2;
            x5 = x6 - ch2;
            y3 = b - ch;
            y4 = b - ch2;

            dVal = "M" + ch + "," + y3 +
                " L" + ch + "," + ch2 +
                shapeArc(x3, ch2, ch2, ch2, 180, 270, false).replace("M", "L") +
                " L" + x7 + "," + t +
                shapeArc(x7, ch2, ch2, ch2, 270, 450, false).replace("M", "L") +
                " L" + x6 + "," + ch +
                " L" + x6 + "," + y4 +
                shapeArc(x5, y4, ch2, ch2, 0, 90, false).replace("M", "L") +
                " L" + ch2 + "," + b +
                shapeArc(ch2, y4, ch2, ch2, 90, 270, false).replace("M", "L") +
                " z" +
                " M" + x3 + "," + t +
                shapeArc(x3, ch2, ch2, ch2, 270, 450, false).replace("M", "L") +
                shapeArc(x3, x3 / 2, ch4, ch4, 90, 270, false).replace("M", "L") +
                " L" + x4 + "," + ch2 +
                " M" + x6 + "," + ch +
                " L" + x3 + "," + ch +
                " M" + ch + "," + y4 +
                shapeArc(ch2, y4, ch2, ch2, 0, 270, false).replace("M", "L") +
                shapeArc(ch2, (y4 + y3) / 2, ch4, ch4, 270, 450, false).replace("M", "L") +
                " z" +
                " M" + ch + "," + y4 +
                " L" + ch + "," + y3;
        } else if (shapType === "horizontalScroll") {
            var y3, y4, y6, y7, y5, x3, x4;
            y3 = ch + ch2;
            y4 = ch + ch;
            y6 = b - ch;
            y7 = b - ch2;
            y5 = y6 - ch2;
            x3 = r - ch;
            x4 = r - ch2;

            dVal = "M" + l + "," + y3 +
                shapeArc(ch2, y3, ch2, ch2, 180, 270, false).replace("M", "L") +
                " L" + x3 + "," + ch +
                " L" + x3 + "," + ch2 +
                shapeArc(x4, ch2, ch2, ch2, 180, 360, false).replace("M", "L") +
                " L" + r + "," + y5 +
                shapeArc(x4, y5, ch2, ch2, 0, 90, false).replace("M", "L") +
                " L" + ch + "," + y6 +
                " L" + ch + "," + y7 +
                shapeArc(ch2, y7, ch2, ch2, 0, 180, false).replace("M", "L") +
                " z" +
                "M" + x4 + "," + ch +
                shapeArc(x4, ch2, ch2, ch2, 90, -180, false).replace("M", "L") +
                shapeArc((x3 + x4) / 2, ch2, ch4, ch4, 180, 0, false).replace("M", "L") +
                " z" +
                " M" + x4 + "," + ch +
                " L" + x3 + "," + ch +
                " M" + ch2 + "," + y4 +
                " L" + ch2 + "," + y3 +
                shapeArc(y3 / 2, y3, ch4, ch4, 180, 360, false).replace("M", "L") +
                shapeArc(ch2, y3, ch2, ch2, 0, 180, false).replace("M", "L") +
                " M" + ch + "," + y3 +
                " L" + ch + "," + y6;
        }
    }

    result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

    return result;
}

/**
 * 饼图/弧形形状渲染模块
 * 提供饼图、弧形、扇形、弦形等形状的生成和渲染功能
 */


/**
 * 渲染饼图/弧形形状
 */
function renderPieShape(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node, oShadowSvgUrlStr) {
    let result = "";
    let dVal = "";

    if (shapType === "pie" || shapType === "pieWedge" || shapType === "arc") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var adj1, adj2, H, shapAdjst1, shapAdjst2, isClose;
        if (shapType === "pie") {
            adj1 = 0;
            adj2 = 270;
            H = h;
            isClose = true;
        } else if (shapType === "pieWedge") {
            adj1 = 180;
            adj2 = 270;
            H = 2 * h;
            isClose = true;
        } else if (shapType === "arc") {
            adj1 = 270;
            adj2 = 0;
            H = h;
            isClose = false;
        }
        if (shapAdjst !== undefined) {
            shapAdjst1 = PPTXXmlUtils.getTextByPathList(shapAdjst, ["attrs", "fmla"]);
            shapAdjst2 = shapAdjst1;
            if (shapAdjst1 === undefined) {
                shapAdjst1 = shapAdjst[0]["attrs"]["fmla"];
                shapAdjst2 = shapAdjst[1]["attrs"]["fmla"];
            }
            if (shapAdjst1 !== undefined) {
                adj1 = parseInt(shapAdjst1.substr(4)) / 60000;
            }
            if (shapAdjst2 !== undefined) {
                adj2 = parseInt(shapAdjst2.substr(4)) / 60000;
            }
        }
        var pieVals = shapePie(H, w, adj1, adj2, isClose);
        result += "<path d='" + pieVals[0] + "' transform='" + pieVals[1] + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' " + (oShadowSvgUrlStr || "") + " />";
    }
    else if (shapType === "chord") {
        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, sAdj1_val = 45;
        var sAdj2, sAdj2_val = 270;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name === "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = parseInt(sAdj1.substr(4)) / 60000;
                } else if (sAdj_name === "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj2_val = parseInt(sAdj2.substr(4)) / 60000;
                }
            }
        }
        var hR = h / 2;
        var wR = w / 2;
        dVal = shapeArc(wR, hR, wR, hR, sAdj1_val, sAdj2_val, true);
        result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' " + (oShadowSvgUrlStr || "") + " />";
    }
    else if (shapType === "blockArc") {
        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, adj1 = 180;
        var sAdj2, adj2 = 0;
        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name === "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) / 60000;
                } else if (sAdj_name === "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) / 60000;
                } else if (sAdj_name === "adj3") {
                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                }
            }
        }

        var stAng, istAng, a3, sw11, sw12, swAng, iswAng;
        var cd1 = 360;
        if (adj1 < 0) stAng = 0;
        else if (adj1 > cd1) stAng = cd1;
        else stAng = adj1;

        if (adj2 < 0) istAng = 0;
        else if (adj2 > cd1) istAng = cd1;
        else istAng = adj2;

        if (adj3 < 0) a3 = 0;
        else if (adj3 > cnstVal1) a3 = cnstVal1;
        else a3 = adj3;

        sw11 = istAng - stAng;
        sw12 = sw11 + cd1;
        swAng = (sw11 > 0) ? sw11 : sw12;
        iswAng = -swAng;

        var endAng = stAng + swAng;
        var iendAng = istAng + iswAng;

        var wt1, ht1, dx1, dy1, x1, y1, stRd, istRd, wd2, hd2, hc, vc;
        stRd = stAng * (Math.PI) / 180;
        istRd = istAng * (Math.PI) / 180;
        wd2 = w / 2;
        hd2 = h / 2;
        hc = w / 2;
        vc = h / 2;
        if (stAng > 90 && stAng < 270) {
            wt1 = wd2 * (Math.sin((Math.PI) / 2 - stRd));
            ht1 = hd2 * (Math.cos((Math.PI) / 2 - stRd));

            dx1 = wd2 * (Math.cos(Math.atan(ht1 / wt1)));
            dy1 = hd2 * (Math.sin(Math.atan(ht1 / wt1)));

            x1 = hc - dx1;
            y1 = vc - dy1;
        } else {
            wt1 = wd2 * (Math.sin(stRd));
            ht1 = hd2 * (Math.cos(stRd));

            dx1 = wd2 * (Math.cos(Math.atan(wt1 / ht1)));
            dy1 = hd2 * (Math.sin(Math.atan(wt1 / ht1)));

            x1 = hc + dx1;
            y1 = vc + dy1;
        }
        var dr, iwd2, ihd2, wt2, ht2, dx2, dy2, x2, y2;
        dr = Math.min(w, h) * a3 / cnstVal2;
        iwd2 = wd2 - dr;
        ihd2 = hd2 - dr;
        if ((endAng <= 450 && endAng > 270) || ((endAng >= 630 && endAng < 720))) {
            wt2 = iwd2 * (Math.sin(istRd));
            ht2 = ihd2 * (Math.cos(istRd));
            dx2 = iwd2 * (Math.cos(Math.atan(wt2 / ht2)));
            dy2 = ihd2 * (Math.sin(Math.atan(wt2 / ht2)));
            x2 = hc + dx2;
            y2 = vc + dy2;
        } else {
            wt2 = iwd2 * (Math.sin((Math.PI) / 2 - istRd));
            ht2 = ihd2 * (Math.cos((Math.PI) / 2 - istRd));

            dx2 = iwd2 * (Math.cos(Math.atan(ht2 / wt2)));
            dy2 = ihd2 * (Math.sin(Math.atan(ht2 / wt2)));
            x2 = hc - dx2;
            y2 = vc - dy2;
        }
        dVal = "M" + x1 + "," + y1 +
            shapeArc(wd2, hd2, wd2, hd2, stAng, endAng, false).replace("M", "L") +
            " L" + x2 + "," + y2 +
            shapeArc(wd2, hd2, iwd2, ihd2, istAng, iendAng, false).replace("M", "L") +
            " z";
        result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' " + (oShadowSvgUrlStr || "") + " />";
    }

    return result;
}

/**
 * 箭头形状渲染模块
 * 处理各种箭头形状的 SVG 生成
 * 
 * 箭头分类:
 * - 基础箭头: rightArrow, leftArrow, upArrow, downArrow
 * - 双向箭头: leftRightArrow, upDownArrow  
 * - 复杂箭头: quadArrow, bentArrow, curvedArrow, circularArrow 等
 * - 标注箭头: xxxArrowCallout
 */


/**
 * 渲染箭头形状
 * @param {string} shapType - 箭头类型
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {boolean} imgFillFlg - 是否使用图片填充
 * @param {boolean} grndFillFlg - 是否使用渐变填充
 * @param {string} fillColor - 填充颜色
 * @param {Object} border - 边框配置 {color, width, strokeDasharray}
 * @param {string} shpId - 形状ID
 * @param {Object} node - 形状节点
 * @returns {string} SVG 字符串
 */
function renderArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    // 基础箭头形状（rightArrow, leftArrow, upArrow, downArrow）
    if (["rightArrow", "leftArrow", "upArrow", "downArrow"].includes(shapType)) {
        return renderBasicArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
    }
    
    // 双向箭头
    if (["leftRightArrow", "upDownArrow"].includes(shapType)) {
        return renderDoubleArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
    }
    
    // 复杂箭头形状暂未实现，返回空字符串由 shape.js 处理
    return "";
}

// ==================== 内部函数 ====================

/**
 * 读取形状调整参数
 * @param {Object} node - 形状节点
 * @returns {Object} 包含 adj1, adj2 值的对象
 */
function readAdjustmentParams(node, w, h) {
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    let sAdj1, sAdj1_val = 0.25;
    let sAdj2, sAdj2_val = 0.5;
    
    if (shapAdjst) {
        for (let i = 0; i < shapAdjst.length; i++) {
            const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "name"]);
            if (sAdjName === "adj1") {
                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "fmla"]);
                sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
            } else if (sAdjName === "adj2") {
                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "fmla"]);
                const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                const maxConst = w / h;  // 会在调用处重新计算
                sAdj2_val = sAdj2Val2 / maxConst;
            }
        }
    }
    
    return { sAdj1_val, sAdj2_val };
}

/**
 * 渲染基础箭头形状
 * @param {string} shapType - 箭头类型
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {boolean} imgFillFlg - 是否使用图片填充
 * @param {boolean} grndFillFlg - 是否使用渐变填充
 * @param {string} fillColor - 填充颜色
 * @param {Object} border - 边框配置
 * @param {string} shpId - 形状ID
 * @param {Object} node - 形状节点
 * @returns {string} SVG 字符串
 */
function renderBasicArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let { sAdj1_val, sAdj2_val } = readAdjustmentParams(node, w, h);
    const max_sAdj2_const = w / h;
    
    // 重新读取并计算 sAdj2_val（因为需要正确的 max_sAdj2_const）
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    if (shapAdjst) {
        for (let i = 0; i < shapAdjst.length; i++) {
            const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "name"]);
            if (sAdjName === "adj2") {
                const sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "fmla"]);
                const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                sAdj2_val = sAdj2Val2 / max_sAdj2_const;
            }
        }
    }
    
    let points;
    if (shapType === "rightArrow") {
        points = `${w} ${h / 2},${sAdj2_val * w} 0,${sAdj2_val * w} ${sAdj1_val * h},0 ${sAdj1_val * h},0 ${(1 - sAdj1_val) * h},${sAdj2_val * w} ${(1 - sAdj1_val) * h}, ${sAdj2_val * w} ${h}`;
    } else if (shapType === "leftArrow") {
        points = `0 ${h / 2},${sAdj2_val * w} ${h},${sAdj2_val * w} ${(1 - sAdj1_val) * h},${w} ${(1 - sAdj1_val) * h},${w} ${sAdj1_val * h},${sAdj2_val * w} ${sAdj1_val * h}, ${sAdj2_val * w} 0`;
    } else if (shapType === "upArrow") {
        // upArrow 使用不同的宽高比计算
        const max_sAdj2_const_up = h / w;
        const { sAdj1_val: sAdj1_up} = readAdjustmentParams(node, w, h);
        const shapAdjst_up = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj2_val_up = 0.5;
        
        if (shapAdjst_up) {
            for (let i = 0; i < shapAdjst_up.length; i++) {
                const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst_up[i], ["attrs", "name"]);
                if (sAdjName === "adj2") {
                    const sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_up[i], ["attrs", "fmla"]);
                    const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val_up = sAdj2Val2 / max_sAdj2_const_up;
                }
            }
        }
        
        points = `${w / 2} 0,0 ${sAdj2_val_up * h},${(0.5 - sAdj1_up) * w} ${sAdj2_val_up * h},${(0.5 - sAdj1_up) * w} ${h},${(0.5 + sAdj1_up) * w} ${h},${(0.5 + sAdj1_up) * w} ${sAdj2_val_up * h}, ${w} ${sAdj2_val_up * h}`;
    } else { // downArrow
        const max_sAdj2_const_down = h / w;
        const { sAdj1_val: sAdj1_down} = readAdjustmentParams(node, w, h);
        const shapAdjst_down = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj2_val_down = 0.5;
        
        if (shapAdjst_down) {
            for (let i = 0; i < shapAdjst_down.length; i++) {
                const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst_down[i], ["attrs", "name"]);
                if (sAdjName === "adj2") {
                    const sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_down[i], ["attrs", "fmla"]);
                    const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val_down = sAdj2Val2 / max_sAdj2_const_down;
                }
            }
        }
        
        points = `${(0.5 - sAdj1_down) * w} 0,${(0.5 - sAdj1_down) * w} ${(1 - sAdj2_val_down) * h},0 ${(1 - sAdj2_val_down) * h},${w / 2} ${h},${w} ${(1 - sAdj2_val_down) * h},${(0.5 + sAdj1_down) * w} ${(1 - sAdj2_val_down) * h}, ${(0.5 + sAdj1_down) * w} 0`;
    }
    
    return buildPolygon(points, imgFillFlg, grndFillFlg, fillColor, border, shpId);
}

/**
 * 渲染双向箭头
 * @param {string} shapType - 箭头类型
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {boolean} imgFillFlg - 是否使用图片填充
 * @param {boolean} grndFillFlg - 是否使用渐变填充
 * @param {string} fillColor - 填充颜色
 * @param {Object} border - 边框配置
 * @param {string} shpId - 形状ID
 * @param {Object} node - 形状节点
 * @returns {string} SVG 字符串
 */
function renderDoubleArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let sAdj1_val = 0.25;
    let sAdj2_val = 0.5;
    const max_sAdj2_const = w / h;
    
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    if (shapAdjst) {
        for (let i = 0; i < shapAdjst.length; i++) {
            const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "name"]);
            if (sAdjName === "adj1") {
                const sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "fmla"]);
                sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
            } else if (sAdjName === "adj2") {
                const sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "fmla"]);
                const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                sAdj2_val = sAdj2Val2 / max_sAdj2_const;
            }
        }
    }
    
    let points;
    if (shapType === "leftRightArrow") {
        points = `0 ${h / 2},${sAdj2_val * w} 0,${sAdj2_val * w} ${h},0 ${h},${w} ${h / 2},${sAdj2_val * w} ${w},${sAdj2_val * w} ${h},${sAdj2_val * w} 0`;
    } else { // upDownArrow
        // upDownArrow 使用不同的宽高比计算
        sAdj1_val = 0.25;
        sAdj2_val = 0.5;
        const max_sAdj2_const_ud = h / w;
        
        const shapAdjst_ud = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ud) {
            for (let i = 0; i < shapAdjst_ud.length; i++) {
                const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst_ud[i], ["attrs", "name"]);
                if (sAdjName === "adj1") {
                    const sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ud[i], ["attrs", "fmla"]);
                    sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
                } else if (sAdjName === "adj2") {
                    const sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ud[i], ["attrs", "fmla"]);
                    const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = sAdj2Val2 / max_sAdj2_const_ud;
                }
            }
        }
        
        points = `${w / 2} 0,${w} ${sAdj2_val * h},${w} ${h}, ${sAdj2_val * w} ${h},${w / 2} ${h},0 ${sAdj2_val * h},0 ${sAdj2_val * h},${sAdj1_val * w} 0, ${sAdj1_val * w} 0`;
    }
    
    return buildPolygon(points, imgFillFlg, grndFillFlg, fillColor, border, shpId);
}

/**
 * 构建 polygon 元素字符串
 * @param {string} points - 点坐标字符串
 * @param {boolean} imgFillFlg - 是否使用图片填充
 * @param {boolean} grndFillFlg - 是否使用渐变填充
 * @param {string} fillColor - 填充颜色
 * @param {Object} border - 边框配置
 * @param {string} shpId - 形状ID
 * @returns {string} SVG 字符串
 */
function buildPolygon(points, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const fillUrl = !imgFillFlg 
        ? (grndFillFlg ? `url(#linGrd_${shpId})` : fillColor) 
        : `url(#imgPtrn_${shpId})`;
    
    return ` <polygon points='${points}' fill='${fillUrl}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}

/**
 * 按钮类形状渲染模块
 * 处理所有 actionButton 类型的形状渲染
 */

/**
 * 渲染 actionButtonBackPrevious 形状
 * 返回按钮（左箭头）
 */
function renderBackPrevious(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;

    const d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + g11 + "," + vc +
        " L" + g12 + "," + g9 +
        " L" + g12 + "," + g10 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonBeginning 形状
 * 开始按钮（双竖线+左箭头）
 */
function renderBeginning(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 8;
    const g15 = g13 / 4;
    const g16 = g11 + g14;
    const g17 = g11 + g15;

    const d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + g17 + "," + vc +
        " L" + g12 + "," + g9 +
        " L" + g12 + "," + g10 +
        " z" +
        "M" + g16 + "," + g9 +
        " L" + g11 + "," + g9 +
        " L" + g11 + "," + g10 +
        " L" + g16 + "," + g10 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonDocument 形状
 * 文档按钮
 */
function renderDocument(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const dx1 = ss * 9 / 32;
    const g11 = hc - dx1;
    const g12 = hc + dx1;
    const g13 = ss * 3 / 16;
    const g14 = g12 - g13;
    const g15 = g9 + g13;

    const d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + g11 + "," + g9 +
        " L" + g14 + "," + g9 +
        " L" + g12 + "," + g15 +
        " L" + g12 + "," + g10 +
        " L" + g11 + "," + g10 +
        " z" +
        "M" + g14 + "," + g9 +
        " L" + g14 + "," + g15 +
        " L" + g12 + "," + g15 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonEnd 形状
 * 结束按钮（双竖线+右箭头）
 */
function renderEnd(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 * 3 / 4;
    const g15 = g13 * 7 / 8;
    const g16 = g11 + g14;
    const g17 = g11 + g15;

    const d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        " z" +
        " M" + g17 + "," + g9 +
        " L" + g12 + "," + g9 +
        " L" + g12 + "," + g10 +
        " L" + g17 + "," + g10 +
        " z" +
        " M" + g16 + "," + vc +
        " L" + g11 + "," + g9 +
        " L" + g11 + "," + g10 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonForwardNext 形状
 * 前进按钮（右箭头）
 */
function renderForwardNext(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;

    const d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        " z" +
        " M" + g12 + "," + vc +
        " L" + g11 + "," + g9 +
        " L" + g11 + "," + g10 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonHelp 形状
 * 帮助按钮（问号）
 */
function renderHelp(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g11 = hc - dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 7;
    const g15 = g13 * 3 / 14;
    const g16 = g13 * 2 / 7;
    const g19 = g13 * 3 / 7;
    const g20 = g13 * 4 / 7;
    const g21 = g13 * 17 / 28;
    const g23 = g13 * 21 / 28;
    const g24 = g13 * 11 / 14;
    const g27 = g9 + g16;
    const g29 = g9 + g21;
    const g30 = g9 + g23;
    const g31 = g9 + g24;
    const g33 = g11 + g15;
    const g36 = g11 + g19;
    const g37 = g11 + g20;
    const g41 = g13 / 14;
    const g42 = g13 * 3 / 28;
    const cX1 = g33 + g16;
    const cX2 = g36 + g14;
    const cY3 = g31 + g42;
    const cX4 = (g37 + g36 + g16) / 2;

    const d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + g33 + "," + g27 +
        shapeArcAlt(cX1, g27, g16, g16, 180, 360, false).replace("M", "L") +
        shapeArcAlt(cX4, g27, g14, g15, 0, 90, false).replace("M", "L") +
        shapeArcAlt(cX4, g29, g41, g42, 270, 180, false).replace("M", "L") +
        " L" + g37 + "," + g30 +
        " L" + g36 + "," + g30 +
        " L" + g36 + "," + g29 +
        shapeArcAlt(cX2, g29, g14, g15, 180, 270, false).replace("M", "L") +
        shapeArcAlt(g37, g27, g41, g42, 90, 0, false).replace("M", "L") +
        shapeArcAlt(cX1, g27, g14, g14, 0, -180, false).replace("M", "L") +
        " z" +
        "M" + hc + "," + g31 +
        shapeArcAlt(hc, cY3, g42, g42, 270, 630, false).replace("M", "L") +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonHome 形状
 * 主页按钮（房子图标）
 */
function renderHome(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 16;
    const g15 = g13 / 8;
    const g16 = g13 * 3 / 16;
    const g17 = g13 * 5 / 16;
    const g18 = g13 * 7 / 16;
    const g19 = g13 * 9 / 16;
    const g20 = g13 * 11 / 16;
    const g21 = g13 * 3 / 4;
    const g22 = g13 * 13 / 16;
    const g23 = g13 * 7 / 8;
    const g24 = g9 + g14;
    const g25 = g9 + g16;
    const g26 = g9 + g17;
    const g27 = g9 + g21;
    const g28 = g11 + g15;
    const g29 = g11 + g18;
    const g30 = g11 + g19;
    const g31 = g11 + g20;
    const g32 = g11 + g22;
    const g33 = g11 + g23;

    const d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        " M" + hc + "," + g9 +
        " L" + g11 + "," + vc +
        " L" + g28 + "," + vc +
        " L" + g28 + "," + g10 +
        " L" + g33 + "," + g10 +
        " L" + g33 + "," + vc +
        " L" + g12 + "," + vc +
        " L" + g32 + "," + g26 +
        " L" + g32 + "," + g24 +
        " L" + g31 + "," + g24 +
        " L" + g31 + "," + g25 +
        " z" +
        " M" + g29 + "," + g27 +
        " L" + g30 + "," + g27 +
        " L" + g30 + "," + g10 +
        " L" + g29 + "," + g10 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonInformation 形状
 * 信息按钮（感叹号+圆点）
 */
function renderInformation(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g11 = hc - dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 32;
    const g17 = g13 * 5 / 16;
    const g18 = g13 * 3 / 8;
    const g19 = g13 * 13 / 32;
    const g20 = g13 * 19 / 32;
    const g22 = g13 * 11 / 16;
    const g23 = g13 * 13 / 16;
    const g24 = g13 * 7 / 8;
    const g25 = g9 + g14;
    const g28 = g9 + g17;
    const g29 = g9 + g18;
    const g30 = g9 + g23;
    const g31 = g9 + g24;
    const g32 = g11 + g17;
    const g34 = g11 + g19;
    const g35 = g11 + g20;
    const g37 = g11 + g22;
    const g38 = g13 * 3 / 32;
    const cY1 = g9 + dx2;
    const cY2 = g25 + g38;

    const d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + hc + "," + g9 +
        shapeArcAlt(hc, cY1, dx2, dx2, 270, 630, false).replace("M", "L") +
        " z" +
        "M" + hc + "," + g25 +
        shapeArcAlt(hc, cY2, g38, g38, 270, 630, false).replace("M", "L") +
        "M" + g32 + "," + g28 +
        " L" + g35 + "," + g28 +
        " L" + g35 + "," + g30 +
        " L" + g37 + "," + g30 +
        " L" + g37 + "," + g31 +
        " L" + g32 + "," + g31 +
        " L" + g32 + "," + g30 +
        " L" + g34 + "," + g30 +
        " L" + g34 + "," + g29 +
        " L" + g32 + "," + g29 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonMovie 形状
 * 影片按钮（胶片图标）
 */
function renderMovie(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 * 1455 / 21600;
    const g15 = g13 * 1905 / 21600;
    const g16 = g13 * 2325 / 21600;
    const g17 = g13 * 16155 / 21600;
    const g18 = g13 * 17010 / 21600;
    const g19 = g13 * 19335 / 21600;
    const g20 = g13 * 19725 / 21600;
    const g21 = g13 * 20595 / 21600;
    const g22 = g13 * 5280 / 21600;
    const g23 = g13 * 5730 / 21600;
    const g24 = g13 * 6630 / 21600;
    const g25 = g13 * 7492 / 21600;
    const g26 = g13 * 9067 / 21600;
    const g27 = g13 * 9555 / 21600;
    const g28 = g13 * 13342 / 21600;
    const g29 = g13 * 14580 / 21600;
    const g30 = g13 * 15592 / 21600;
    const g31 = g11 + g14;
    const g32 = g11 + g15;
    const g33 = g11 + g16;
    const g34 = g11 + g17;
    const g35 = g11 + g18;
    const g36 = g11 + g19;
    const g37 = g11 + g20;
    const g38 = g11 + g21;
    const g39 = g9 + g22;
    const g40 = g9 + g23;
    const g41 = g9 + g24;
    const g42 = g9 + g25;
    const g43 = g9 + g26;
    const g44 = g9 + g27;
    const g45 = g9 + g28;
    const g46 = g9 + g29;
    const g47 = g9 + g30;

    const d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        " z" +
        "M" + g11 + "," + g39 +
        " L" + g11 + "," + g44 +
        " L" + g31 + "," + g44 +
        " L" + g32 + "," + g43 +
        " L" + g33 + "," + g43 +
        " L" + g33 + "," + g47 +
        " L" + g35 + "," + g47 +
        " L" + g35 + "," + g45 +
        " L" + g36 + "," + g45 +
        " L" + g38 + "," + g46 +
        " L" + g12 + "," + g46 +
        " L" + g12 + "," + g41 +
        " L" + g38 + "," + g41 +
        " L" + g37 + "," + g42 +
        " L" + g35 + "," + g42 +
        " L" + g35 + "," + g41 +
        " L" + g34 + "," + g40 +
        " L" + g32 + "," + g40 +
        " L" + g31 + "," + g39 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonReturn 形状
 * 返回按钮（折返箭头）
 */
function renderReturn(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 * 7 / 8;
    const g15 = g13 * 3 / 4;
    const g16 = g13 * 5 / 8;
    const g17 = g13 * 3 / 8;
    const g18 = g13 / 4;
    const g19 = g9 + g15;
    const g20 = g9 + g16;
    const g21 = g9 + g18;
    const g22 = g11 + g14;
    const g23 = g11 + g15;
    const g24 = g11 + g16;
    const g25 = g11 + g17;
    const g26 = g11 + g18;
    const g27 = g13 / 8;
    const cX1 = g24 - g27;
    const cY2 = g19 - g27;
    const cX3 = g11 + g17;
    const cY4 = g10 - g17;

    const d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        " z" +
        " M" + g12 + "," + g21 +
        " L" + g23 + "," + g9 +
        " L" + hc + "," + g21 +
        " L" + g24 + "," + g21 +
        " L" + g24 + "," + g20 +
        shapeArcAlt(cX1, g20, g27, g27, 0, 90, false).replace("M", "L") +
        " L" + g25 + "," + g19 +
        shapeArcAlt(g25, cY2, g27, g27, 90, 180, false).replace("M", "L") +
        " L" + g26 + "," + g21 +
        " L" + g11 + "," + g21 +
        " L" + g11 + "," + g20 +
        shapeArcAlt(cX3, g20, g17, g17, 180, 90, false).replace("M", "L") +
        " L" + hc + "," + g10 +
        shapeArcAlt(hc, cY4, g17, g17, 90, 0, false).replace("M", "L") +
        " L" + g22 + "," + g21 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonSound 形状
 * 声音按钮（喇叭图标）
 */
function renderSound(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 8;
    const g15 = g13 * 5 / 16;
    const g16 = g13 * 5 / 8;
    const g17 = g13 * 11 / 16;
    const g18 = g13 * 3 / 4;
    const g19 = g13 * 7 / 8;
    const g20 = g9 + g14;
    const g21 = g9 + g15;
    const g22 = g9 + g17;
    const g23 = g9 + g19;
    const g24 = g11 + g15;
    const g25 = g11 + g16;
    const g26 = g11 + g18;

    const d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        " M" + g11 + "," + g21 +
        " L" + g24 + "," + g21 +
        " L" + g25 + "," + g9 +
        " L" + g25 + "," + g10 +
        " L" + g24 + "," + g22 +
        " L" + g11 + "," + g22 +
        " z" +
        " M" + g26 + "," + g21 +
        " L" + g12 + "," + g20 +
        " M" + g26 + "," + vc +
        " L" + g12 + "," + vc +
        " M" + g26 + "," + g22 +
        " L" + g12 + "," + g23;

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 按钮形状渲染器映射表
 */
const BUTTON_RENDERERS = {
    'actionButtonBackPrevious': renderBackPrevious,
    'actionButtonBeginning': renderBeginning,
    'actionButtonDocument': renderDocument,
    'actionButtonEnd': renderEnd,
    'actionButtonForwardNext': renderForwardNext,
    'actionButtonHelp': renderHelp,
    'actionButtonHome': renderHome,
    'actionButtonInformation': renderInformation,
    'actionButtonMovie': renderMovie,
    'actionButtonReturn': renderReturn,
    'actionButtonSound': renderSound
};

/**
 * 渲染按钮类形状
 * @param {string} shapeType - 按钮类型
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {boolean} imgFillFlg - 图片填充标志
 * @param {boolean} grndFillFlg - 渐变填充标志
 * @param {string} fillColor - 填充颜色
 * @param {object} border - 边框配置
 * @param {string} shpId - 形状ID
 * @param {Function} shapeArcAlt - 弧形生成函数
 * @returns {string} SVG 路径字符串
 */
function renderActionButton(shapeType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const renderer = BUTTON_RENDERERS[shapeType];
    if (!renderer) {
        return '';
    }
    return renderer(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt);
}

/**
 * 形状渲染主模块
 * 
 * 这是 PPTX 形状渲染的核心模块，负责处理所有 PowerPoint 预设形状的 SVG 生成。
 * 
 * 模块职责:
 * - 坐标变换和尺寸计算
 * - 形状类型识别和路由
 * - 基础几何形状的 SVG 生成（矩形、圆形、三角形等）
 * - 协调各子模块（箭头、星形、括号、饼图等）
 * 
 * 结构说明:
 * - 该模块是一个 IIFE，导出 PPTXShapeUtils 对象
 * - genShape() 是主入口函数，处理单个形状的完整渲染流程
 * - 使用大量的 switch-case 语句处理不同形状类型
 * - 复杂形状已拆分到独立子模块（arrow-shapes.js, star-shapes.js 等）
 * 
 * 注意事项:
 * - 代码量较大（4875行），包含约 208 个形状类型
 * - 使用 ES5 语法以保持兼容性
 * - 变量命名使用匈牙利命名法（如 shpId, imgFillFlg, grndFillFlg）
 * 
 * @module shape/shape
 */


const PPTXShapeUtils = (function() {
    /**
     * 辅助函数：生成形状的 data- 属性字符串
     * @param {Object} node - 节点
     * @param {Object} slideXfrmNode - 变换节点
     * @param {string} id - ID
     * @param {string} name - 名称
     * @param {string} idx - 索引
     * @param {string} type - 类型
     * @param {number} rotate - 旋转角度
     * @param {string} sType - 形状类型
     * @returns {string} data- 属性字符串
     */
    function genShapeDataAttributes(node, slideXfrmNode, id, name, idx, type, rotate, sType) {
        let dataAttrs = '';
        
        // 提取位置和尺寸信息
        let offX = 0, offY = 0, extCx = 0, extCy = 0, flipH = 0, flipV = 0;
        if (slideXfrmNode !== undefined) {
            if (slideXfrmNode['a:off'] && slideXfrmNode['a:off'].attrs) {
                offX = slideXfrmNode['a:off'].attrs.x || 0;
                offY = slideXfrmNode['a:off'].attrs.y || 0;
            }
            if (slideXfrmNode['a:ext'] && slideXfrmNode['a:ext'].attrs) {
                extCx = slideXfrmNode['a:ext'].attrs.cx || 0;
                extCy = slideXfrmNode['a:ext'].attrs.cy || 0;
            }
            if (slideXfrmNode['attrs']) {
                slideXfrmNode['attrs'].rot || 0;
                flipH = slideXfrmNode['attrs'].flipH || '0';
                flipV = slideXfrmNode['attrs'].flipV || '0';
            }
        }
        
        // 获取形状类型
        const shapType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
        
        // 构建 data- 属性
        dataAttrs += ` data-node-id="${id || ''}"`;
        dataAttrs += ` data-node-name="${name || ''}"`;
        dataAttrs += ` data-node-idx="${idx || ''}"`;
        dataAttrs += ` data-node-type="${type || ''}"`;
        dataAttrs += ` data-shape-type="${sType || ''}"`;
        dataAttrs += ` data-off-x="${offX}"`;
        dataAttrs += ` data-off-y="${offY}"`;
        dataAttrs += ` data-ext-cx="${extCx}"`;
        dataAttrs += ` data-ext-cy="${extCy}"`;
        dataAttrs += ` data-rotate="${rotate || 0}"`;
        dataAttrs += ` data-flip-h="${flipH}"`;
        dataAttrs += ` data-flip-v="${flipV}"`;
        if (shapType) {
            dataAttrs += ` data-geom-type="${shapType}"`;
        }
        
        return dataAttrs;
    }

    async function genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, isUserDrawnBg, sType, source, settings) {
            //var dltX = 0;
            //var dltY = 0;
            var xfrmList = ["p:spPr", "a:xfrm"];
            var slideXfrmNode = PPTXXmlUtils.getTextByPathList(node, xfrmList);
            var slideLayoutXfrmNode = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, xfrmList);
            var slideMasterXfrmNode = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, xfrmList);

            var result = "";
            var shpId = PPTXXmlUtils.getTextByPathList(node, ["attrs", "order"]);
            var shapType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);

            // 初始化3D变换样式
            let transform3dStyle = "";
            //custGeom - Amir
            var custShapType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:custGeom"]);

            var isFlipV = false;
            var isFlipH = false;
            var flip = "";
            var flipVAttr = PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipV"]);
            var flipHAttr = PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipH"]);
            if (flipVAttr === "1" || flipVAttr === "true") {
                isFlipV = true;
            }
            if (flipHAttr === "1" || flipHAttr === "true") {
                isFlipH = true;
            }
            if (isFlipH && !isFlipV) {
                flip = " scale(-1,1)";
            } else if (!isFlipH && isFlipV) {
                flip = " scale(1,-1)";
            } else if (isFlipH && isFlipV) {
                flip = " scale(-1,-1)";
            }
            /////////////////////////Amir////////////////////////
            //rotate
            var rotate = PPTXXmlUtils.angleToDegrees(PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "rot"]));
            var txtXframeNode = PPTXXmlUtils.getTextByPathList(node, ["p:txXfrm"]);
            if (txtXframeNode !== undefined) {
                PPTXXmlUtils.getTextByPathList(txtXframeNode, ["attrs", "rot"]);
            }
            //////////////////////////////////////////////////

            // 处理组合缩放 - 当形状在group-abs类型组合中时需要应用缩放
            let workingXfrmNode = slideXfrmNode;
            let drawW, drawH; // 原始未缩放的尺寸，用于SVG内部绘图

            // 先保存原始尺寸（如果存在）
            if (slideXfrmNode && slideXfrmNode['a:ext'] && slideXfrmNode['a:ext'].attrs) {
                const originalCx = parseInt(slideXfrmNode['a:ext'].attrs.cx);
                const originalCy = parseInt(slideXfrmNode['a:ext'].attrs.cy);
                drawW = (originalCx !== undefined && originalCy !== undefined)
                    ? originalCx * SLIDE_FACTOR$1
                    : undefined;
                drawH = (originalCx !== undefined && originalCy !== undefined)
                    ? originalCy * SLIDE_FACTOR$1
                    : undefined;
            }

            if (sType === 'group-abs' && warpObj.currentGroupScale && slideXfrmNode) {
                const { scaleX, scaleY, childX, childY } = warpObj.currentGroupScale;

                // 创建缩放后的xfrmNode
                workingXfrmNode = JSON.parse(JSON.stringify(slideXfrmNode));

                // 缩放尺寸
                if (slideXfrmNode['a:ext'] && slideXfrmNode['a:ext'].attrs) {
                    const originalCx = parseInt(slideXfrmNode['a:ext'].attrs.cx);
                    const originalCy = parseInt(slideXfrmNode['a:ext'].attrs.cy);
                    workingXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
                    workingXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
                }

                // 调整位置(相对于childX/childY)
                if (slideXfrmNode['a:off'] && slideXfrmNode['a:off'].attrs) {
                    const originalOffX = parseInt(slideXfrmNode['a:off'].attrs.x);
                    const originalOffY = parseInt(slideXfrmNode['a:off'].attrs.y);

                    // childX和childY已经是像素值，需要转换为EMU单位来计算
                    const childXEmu = childX / SLIDE_FACTOR$1;
                    const childYEmu = childY / SLIDE_FACTOR$1;

                    // 计算相对于childOff的偏移（EMU单位）
                    const relativeX = originalOffX - childXEmu;
                    const relativeY = originalOffY - childYEmu;

                    // 应用缩放，结果为EMU单位
                    workingXfrmNode['a:off'].attrs.x = Math.round(childXEmu + relativeX * scaleX);
                    workingXfrmNode['a:off'].attrs.y = Math.round(childYEmu + relativeY * scaleY);
                }
            }

            if (shapType !== undefined || custShapType !== undefined /*&& slideXfrmNode !== undefined*/) {
                // 使用 workingXfrmNode 而不是 slideXfrmNode，以正确处理 group-abs 情况
                var off = PPTXXmlUtils.getTextByPathList(workingXfrmNode, ["a:off", "attrs"]);
                (off !== undefined) ? parseInt(off["x"]) * SLIDE_FACTOR$1 : 0;
                (off !== undefined) ? parseInt(off["y"]) * SLIDE_FACTOR$1 : 0;

                var ext = PPTXXmlUtils.getTextByPathList(workingXfrmNode, ["a:ext", "attrs"]);

                // Fallback to slideLayoutXfrmNode if workingXfrmNode is undefined or ext is undefined
                if (ext === undefined && slideLayoutXfrmNode !== undefined) {
                    ext = PPTXXmlUtils.getTextByPathList(slideLayoutXfrmNode, ["a:ext", "attrs"]);
                }
                // Fallback to slideMasterXfrmNode if still undefined
                if (ext === undefined && slideMasterXfrmNode !== undefined) {
                    ext = PPTXXmlUtils.getTextByPathList(slideMasterXfrmNode, ["a:ext", "attrs"]);
                }

                var w = (ext !== undefined && ext["cx"] !== undefined) ? parseInt(ext["cx"]) * SLIDE_FACTOR$1 : 100;
                var h = (ext !== undefined && ext["cy"] !== undefined) ? parseInt(ext["cy"]) * SLIDE_FACTOR$1 : 100;
                w = isNaN(w) ? 100 : w;
                h = isNaN(h) ? 100 : h;

                // 如果drawW/drawH未定义（非group-abs情况），则使用w和h
                if (drawW === undefined) drawW = w;
                if (drawH === undefined) drawH = h;

                // 对于连接器类型，需要特殊处理
                var isConnector = (shapType === 'straightConnector1' || shapType === 'bentConnector2' ||
                                   shapType === 'bentConnector3' || shapType === 'bentConnector4' ||
                                   shapType === 'bentConnector5' || shapType === 'curvedConnector2' ||
                                   shapType === 'curvedConnector3' || shapType === 'curvedConnector4' ||
                                   shapType === 'curvedConnector5');





                var svgCssName = "_svg_css_" + (Object.keys(warpObj.styleTable).length + 1) + "_"  + Math.floor(Math.random() * 1001);
                var effectsClassName = svgCssName + "_effects";

                // 对于连接器，当width或height为0时，需要设置最小尺寸
                var svgSizeStyle = "";

                if (isConnector && (w === 0 || h === 0)) {
                    // 设置最小尺寸为strokeWidth的2倍（或至少4px），确保线条可见
                    var strokeWidth = 1.5; // 默认stroke-width，实际可以从border获取
                    var minSize = Math.max(strokeWidth * 2, 4);
                    // SVG容器的尺寸至少为minSize
                    var svgW = (w === 0 || w < minSize) ? minSize : w;
                    var svgH = (h === 0 || h < minSize) ? minSize : h;
                    svgSizeStyle = "width:" + svgW + "px; height:" + svgH + "px; overflow: visible;";
                    // 更新w和h为SVG容器尺寸，这样后续代码会使用正确的尺寸
                    w = svgW;
                    h = svgH;

                } else {
                    svgSizeStyle = PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined) + " overflow: visible;";
                }

                // 如果形状在组合中被缩放，SVG内容需要应用缩放
                let svgTransform = "transform: rotate(" + ((rotate !== undefined) ? rotate : 0) + "deg)" + flip + ";";
                if (sType === 'group-abs' && warpObj.currentGroupScale) {
                    // 对于自定义形状，我们已经在 renderCustomShape 中使用了缩放后的尺寸
                    // 所以不需要在这里应用 SVG transform scale()
                    // 预置形状仍然需要使用 transform scale()
                    if (custShapType === undefined) {
                        const { scaleX, scaleY } = warpObj.currentGroupScale;
                        svgTransform = `transform: rotate(${(rotate !== undefined) ? rotate : 0}deg)${flip} scale(${scaleX},${scaleY});`;
                    }
                }

                const svgTag = "<svg class='drawing " + svgCssName + "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name + "'" +
                    "' style='" +
                    PPTXXmlUtils.getPosition(workingXfrmNode, pNode, undefined, undefined, sType) +
                    svgSizeStyle +
                    " z-index: " + order + ";" +
                    svgTransform +
                    "'>";
                result += svgTag;
                result += '<defs>';
                // Fill Color
                var fillColor = await PPTXStyleUtils.getShapeFill(node, pNode, true, warpObj, source);

                var grndFillFlg = false;
                var imgFillFlg = false;
                var clrFillType = PPTXStyleUtils.getFillType (PPTXXmlUtils.getTextByPathList(node, ["p:spPr"]));
                if (clrFillType == "GROUP_FILL") {
                    clrFillType = PPTXStyleUtils.getFillType (PPTXXmlUtils.getTextByPathList(pNode, ["p:grpSpPr"]));
                }
                // if (clrFillType == "") {
                //     var clrFillType = PPTXStyleUtils.getFillType (PPTXXmlUtils.getTextByPathList(node, ["p:style","a:fillRef"]));
                // }

                /////////////////////////////////////////                    
                if (clrFillType == "GRADIENT_FILL") {
                    grndFillFlg = true;
                    var color_arry = fillColor.color;
                    var angl = fillColor.rot + 90;
                    var svgGrdnt = PPTXStyleUtils.getSvgGradient(w, h, angl, color_arry, shpId);
                    //fill="url(#linGrd)"
                    //console.log("genShape: svgGrdnt: ", svgGrdnt)
                    result += svgGrdnt;

                } else if (clrFillType == "PIC_FILL") {
                    imgFillFlg = true;
                    var svgBgImg = PPTXStyleUtils.getSvgImagePattern(node, fillColor, shpId, warpObj);
                    //fill="url(#imgPtrn)"
                    //console.log(svgBgImg)
                    result += svgBgImg;
                } else if (clrFillType == "PATTERN_FILL") {
                    var styleText = fillColor;
                    if (styleText in warpObj.styleTable) {
                        styleText += "do-nothing: " + svgCssName +";";
                    }
                    warpObj.styleTable[styleText] = {
                        "name": svgCssName,
                        "text": styleText
                    };
                    //}
                    fillColor = "none";
                } else {
                    if (clrFillType != "SOLID_FILL" && clrFillType != "PATTERN_FILL" &&
                        (shapType == "arc" ||
                            shapType == "bracketPair" ||
                            shapType == "bracePair" ||
                            shapType == "leftBracket" ||
                            shapType == "leftBrace" ||
                            shapType == "rightBrace" ||
                            shapType == "rightBracket")) { //Temp. solution  - TODO
                        fillColor = "none";
                    }
                }
                // Border Color
                var border = PPTXStyleUtils.getBorder(node, pNode, true, "shape", warpObj);

                var headEndNodeAttrs = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
                var tailEndNodeAttrs = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);
                // type: none, triangle, stealth, diamond, oval, arrow

                ////////////////////effects/////////////////////////////////////////////////////
                //p:spPr => a:effectLst =>
                //"a:blur"
                //"a:fillOverlay"
                //"a:glow"
                //"a:innerShdw"
                //"a:outerShdw"
                //"a:prstShdw"
                //"a:reflection"
                //"a:softEdge"
                //p:spPr => a:scene3d
                //"a:camera"
                //"a:lightRig"
                //"a:backdrop"
                //"a:extLst"?
                //p:spPr => a:sp3d
                //"a:bevelT"
                //"a:bevelB"
                //"a:extrusionClr"
                //"a:contourClr"
                //"a:extLst"?
                
                // 处理3D效果
                const scene3d = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:scene3d"]);
                const sp3d = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:sp3d"]);
                
                if (scene3d || sp3d) {
                    transform3dStyle = process3DEffects(scene3d, sp3d);
                }
                // Check if there's an effectRef in p:style
                var effectRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:effectRef"]);
                var effectStyleNode = undefined;
                
                if (effectRefNode !== undefined) {
                    var effectIdx = PPTXXmlUtils.getTextByPathList(effectRefNode, ["attrs", "idx"]);
                    if (effectIdx !== undefined && warpObj["themeContent"] !== undefined) {
                        // Access the effect style from the theme
                        var effectStyleLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:effectStyleLst"]["a:effectStyle"];
                        if (effectStyleLst !== undefined) {
                            // Ensure effectStyleLst is an array
                            if (!Array.isArray(effectStyleLst)) {
                                effectStyleLst = [effectStyleLst];
                            }
                            // Convert effectIdx to number and use as array index
                            // effectRef idx is 0-based (idx="0" refers to first effectStyle)
                            var idx = Number(effectIdx);
                            // Handle idx out of range
                            if (effectStyleLst.length > 0) {
                                if (idx >= 0 && idx < effectStyleLst.length) {
                                    effectStyleNode = effectStyleLst[idx];
                                } else {
                                    // When idx is out of range, try to find an effectStyle with shadow
                                    // Start from the end of the list and work backwards
                                    for (var i = effectStyleLst.length - 1; i >= 0; i--) {
                                        var testEffectStyle = effectStyleLst[i];
                                        var hasShadow = PPTXXmlUtils.getTextByPathList(testEffectStyle, ["a:effectLst", "a:outerShdw"]);
                                        if (hasShadow !== undefined) {
                                            effectStyleNode = testEffectStyle;
                                            break;
                                        }
                                    }
                                    // If no shadow found, use modulo to wrap around
                                    if (effectStyleNode === undefined) {
                                        idx = idx % effectStyleLst.length;
                                        if (idx < 0) idx += effectStyleLst.length;
                                        effectStyleNode = effectStyleLst[idx];
                                    }
                                }
                            }
                        }
                    }
                }
                
                //////////////////////////////outerShdw///////////////////////////////////////////
                //not support sizing the shadow
                var outerShdwNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:effectLst", "a:outerShdw"]);
                
                // If no direct outerShdw, check from effectStyle
                if (outerShdwNode === undefined && effectStyleNode !== undefined) {
                    outerShdwNode = PPTXXmlUtils.getTextByPathList(effectStyleNode, ["a:effectLst", "a:outerShdw"]);
                }

                var oShadowSvgUrlStr = "";
                // Check if outerShdwNode exists and has valid shadow attributes
                // A valid shadow should have at least dist defined with a non-zero value
                var hasOuterShadow = false;
                if (outerShdwNode && typeof outerShdwNode === 'object' && !Array.isArray(outerShdwNode)) {
                    // Check that outerShdwNode is not empty and is actually a valid outerShdw node
                    var nodeKeys = Object.keys(outerShdwNode);
                    if (nodeKeys.length > 0) {
                        var attrs = outerShdwNode.attrs;
                        // A valid outerShdw node should have an attrs object with shadow properties
                        if (attrs && typeof attrs === 'object') {
                            var distVal = attrs.dist;
                            var blurRadVal = attrs.blurRad;
                            // Only consider it a valid shadow if dist is defined and non-zero
                            // Also check if at least one of the required shadow attributes is present
                            var hasShadowAttrs = (distVal !== undefined || blurRadVal !== undefined ||
                                                 attrs.dir !== undefined || attrs.sx !== undefined ||
                                                 attrs.sy !== undefined || attrs.algn !== undefined);
                            hasOuterShadow = hasShadowAttrs && (distVal !== undefined && distVal !== "" && distVal !== "0" && distVal !== 0);
                        }
                    }
                }

                // Check if shape has 3D effects (sp3d or scene3d)
                // Only disable shadow if 3D effects are defined directly on the shape (p:spPr)
                // 3D effects from effectStyle (via effectRef) should not disable the shadow
                var sp3dNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:sp3d"]);
                var scene3dNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:scene3d"]);
                // Check if outerShdw is from effectStyle
                var shadowFromEffectStyle = (outerShdwNode !== undefined && effectStyleNode !== undefined);
                if ((sp3dNode !== undefined || scene3dNode !== undefined) && !shadowFromEffectStyle) {
                    // Disable shadow when 3D effects are present on the shape itself
                    hasOuterShadow = false;
                }


                if (hasOuterShadow) {
                    var chdwClrNode = PPTXStyleUtils.getSolidFill(outerShdwNode, undefined, undefined, warpObj);
                    var outerShdwAttrs = outerShdwNode["attrs"];

                    //var algn = outerShdwAttrs["algn"];
                    var dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
                    var dist = parseInt(outerShdwAttrs["dist"]) * SLIDE_FACTOR$1;//(px) //* (3 / 4); //(pt)
                    //var rotWithShape = outerShdwAttrs["rotWithShape"];
                    var blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * SLIDE_FACTOR$1) : ""; //+ "px"
                    //var sx = (outerShdwAttrs["sx"]) ? (parseInt(outerShdwAttrs["sx"]) / 100000) : 1;
                    //var sy = (outerShdwAttrs["sy"]) ? (parseInt(outerShdwAttrs["sy"]) / 100000) : 1;
                    var vx = dist * Math.sin(dir * Math.PI / 180);
                    var hx = dist * Math.cos(dir * Math.PI / 180);
                    //SVG
                    //var oShadowId = "outerhadow_" + shpId;
                    //oShadowSvgUrlStr = "filter='url(#" + oShadowId+")'";
                    //var shadowFilterStr = '<filter id="' + oShadowId + '" x="0" y="0" width="' + w * (6 / 8) + '" height="' + h + '">';
                    //1:
                    //shadowFilterStr += '<feDropShadow dx="' + vx + '" dy="' + hx + '" stdDeviation="' + blurRad * (3 / 4) + '" flood-color="#' + chdwClrNode +'" flood-opacity="1" />'
                    //2:
                    //shadowFilterStr += '<feFlood result="floodColor" flood-color="red" flood-opacity="0.5"   width="' + w * (6 / 8) + '" height="' + h + '"  />'; //#' + chdwClrNode +'
                    //shadowFilterStr += '<feOffset result="offOut" in="SourceGraph ccfsdf-+ic"  dx="' + vx + '" dy="' + hx + '"/>'; //how much to offset
                    //shadowFilterStr += '<feGaussianBlur result="blurOut" in="offOut" stdDeviation="' + blurRad*(3/4) +'"/>'; //tdDeviation is how much to blur
                    //shadowFilterStr += '<feComponentTransfer><feFuncA type="linear" slope="0.5"/></feComponentTransfer>'; //slope is the opacity of the shadow
                    //shadowFilterStr += '<feBlend in="SourceGraphic" in2="blurOut"  mode="normal" />'; //this contains the element that the filter is applied to
                    //shadowFilterStr += '</filter>'; 
                    //result += shadowFilterStr;

                    //css:
                    var svg_css_shadow = "filter:drop-shadow(" + hx + "px " + vx + "px " + blurRad + "px #" + chdwClrNode + ");";

                    if (svg_css_shadow in warpObj.styleTable) {
                        svg_css_shadow += "do-nothing: " + svgCssName + ";";
                    }

                    warpObj.styleTable[svg_css_shadow] = {
                        "name": effectsClassName,
                        "text": svg_css_shadow
                    };
                    result = result.replace("class='drawing " + svgCssName + "'", "class='drawing " + svgCssName + " " + effectsClassName + "'");
                }

                //////////////////////////////softEdge///////////////////////////////////////////
                // Soft edge effect - creates a blurred/feathered edge
                var softEdgeNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:effectLst", "a:softEdge"]);
                
                // If no direct softEdge, check from effectStyle
                if (softEdgeNode === undefined && effectStyleNode !== undefined) {
                    softEdgeNode = PPTXXmlUtils.getTextByPathList(effectStyleNode, ["a:effectLst", "a:softEdge"]);
                }
                
                var softEdgeFilterStr = "";
                if (softEdgeNode !== undefined) {
                    var softEdgeAttrs = softEdgeNode["attrs"];
                    var rad = (softEdgeAttrs["rad"]) ? (parseInt(softEdgeAttrs["rad"]) * SLIDE_FACTOR$1) : 0;
                    
                    // softEdge effect according to Office Open XML specification:
                    // Applies a Gaussian blur to the edges of the shape
                    // The radius determines how far the blur extends from the edge
                    var softEdgeId = "softedge_" + shpId;
                    var softEdgeFilter = '<filter id="' + softEdgeId + '" x="-20%" y="-20%" width="140%" height="140%">';
                    // Blur the source to create soft edge
                    softEdgeFilter += '<feGaussianBlur in="SourceGraphic" stdDeviation="' + rad + '" />';
                    softEdgeFilter += '</filter>';
                    result += softEdgeFilter;
                    softEdgeFilterStr = 'filter="url(#' + softEdgeId + ')"';
                } 
                ////////////////////////////////////////////////////////////////////////////////////////
                if ((headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) ||
                    (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow"))) {
                    // 箭头标记：refX=10 表示箭头尖端与线条端点对齐
                    var triangleMarker = "<marker id='markerTriangle_" + shpId + "' viewBox='0 0 10 10' refX='10' refY='5' markerWidth='5' markerHeight='5' stroke='" + border.color + "' fill='" + border.color +
                        "' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>";
                    result += triangleMarker;
                }
                result += '</defs>';
            }
            if (shapType !== undefined && custShapType === undefined) {
                //console.log("shapType: ", shapType)
                switch (shapType) {
                    case "rect":
                    case "flowChartProcess":
                    case "flowChartPredefinedProcess":
                    case "flowChartInternalStorage":
                    case "actionButtonBlank": {
                        result += "<rect x='0' y='0' width='" + w + "' height='" + h + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' " + oShadowSvgUrlStr + "  />";

                        if (shapType == "flowChartPredefinedProcess") {
                            result += "<rect x='" + w * (1 / 8) + "' y='0' width='" + w * (6 / 8) + "' height='" + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        } else if (shapType == "flowChartInternalStorage") {
                            result += " <polyline points='" + w * (1 / 8) + " 0," + w * (1 / 8) + " " + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                            result += " <polyline points='0 " + h * (1 / 8) + "," + w + " " + h * (1 / 8) + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        }
                        break;
                    }
                    case "flowChartCollate": {
                        var d = "M 0,0" +
                            " L" + w + "," + 0 +
                            " L" + 0 + "," + h +
                            " L" + w + "," + h +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' " + oShadowSvgUrlStr + " " + softEdgeFilterStr + " />";

                        break;
                    }
                    case "flowChartDocument": {
                        var y1, y2, y3, x1;
                        x1 = w * 10800 / 21600;
                        y1 = h * 17322 / 21600;
                        y2 = h * 20172 / 21600;
                        y3 = h * 23922 / 21600;
                        var d = "M" + 0 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + y1 +
                            " C" + x1 + "," + y1 + " " + x1 + "," + y3 + " " + 0 + "," + y2 +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "flowChartMultidocument": {
                        var y1, y2, y3, y4, y5, y6, y7, y8, y9, x1, x2, x3, x4, x5, x6, x7;
                        y1 = h * 18022 / 21600;
                        y2 = h * 3675 / 21600;
                        y3 = h * 23542 / 21600;
                        y4 = h * 1815 / 21600;
                        y5 = h * 16252 / 21600;
                        y6 = h * 16352 / 21600;
                        y7 = h * 14392 / 21600;
                        y8 = h * 20782 / 21600;
                        y9 = h * 14467 / 21600;
                        x1 = w * 1532 / 21600;
                        x2 = w * 20000 / 21600;
                        x3 = w * 9298 / 21600;
                        x4 = w * 19298 / 21600;
                        x5 = w * 18595 / 21600;
                        x6 = w * 2972 / 21600;
                        x7 = w * 20800 / 21600;
                        var d = "M" + 0 + "," + y2 +
                            " L" + x5 + "," + y2 +
                            " L" + x5 + "," + y1 +
                            " C" + x3 + "," + y1 + " " + x3 + "," + y3 + " " + 0 + "," + y8 +
                            " z" +
                            "M" + x1 + "," + y2 +
                            " L" + x1 + "," + y4 +
                            " L" + x2 + "," + y4 +
                            " L" + x2 + "," + y5 +
                            " C" + x4 + "," + y5 + " " + x5 + "," + y6 + " " + x5 + "," + y6 +
                            "M" + x6 + "," + y4 +
                            " L" + x6 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + y7 +
                            " C" + x7 + "," + y7 + " " + x2 + "," + y9 + " " + x2 + "," + y9;

                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "actionButtonBackPrevious":
                    case "actionButtonBeginning":
                    case "actionButtonDocument":
                    case "actionButtonEnd":
                    case "actionButtonForwardNext":
                    case "actionButtonHelp":
                    case "actionButtonHome":
                    case "actionButtonInformation":
                    case "actionButtonMovie":
                    case "actionButtonReturn":
                    case "actionButtonSound": {
                        result += renderActionButton(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt);
                        break;
                    }
                    case "irregularSeal1":
                    case "irregularSeal2": {
                        if (shapType == "irregularSeal1") {
                            var d = "M" + w * 10800 / 21600 + "," + h * 5800 / 21600 +
                                " L" + w * 14522 / 21600 + "," + 0 +
                                " L" + w * 14155 / 21600 + "," + h * 5325 / 21600 +
                                " L" + w * 18380 / 21600 + "," + h * 4457 / 21600 +
                                " L" + w * 16702 / 21600 + "," + h * 7315 / 21600 +
                                " L" + w * 21097 / 21600 + "," + h * 8137 / 21600 +
                                " L" + w * 17607 / 21600 + "," + h * 10475 / 21600 +
                                " L" + w + "," + h * 13290 / 21600 +
                                " L" + w * 16837 / 21600 + "," + h * 12942 / 21600 +
                                " L" + w * 18145 / 21600 + "," + h * 18095 / 21600 +
                                " L" + w * 14020 / 21600 + "," + h * 14457 / 21600 +
                                " L" + w * 13247 / 21600 + "," + h * 19737 / 21600 +
                                " L" + w * 10532 / 21600 + "," + h * 14935 / 21600 +
                                " L" + w * 8485 / 21600 + "," + h +
                                " L" + w * 7715 / 21600 + "," + h * 15627 / 21600 +
                                " L" + w * 4762 / 21600 + "," + h * 17617 / 21600 +
                                " L" + w * 5667 / 21600 + "," + h * 13937 / 21600 +
                                " L" + w * 135 / 21600 + "," + h * 14587 / 21600 +
                                " L" + w * 3722 / 21600 + "," + h * 11775 / 21600 +
                                " L" + 0 + "," + h * 8615 / 21600 +
                                " L" + w * 4627 / 21600 + "," + h * 7617 / 21600 +
                                " L" + w * 370 / 21600 + "," + h * 2295 / 21600 +
                                " L" + w * 7312 / 21600 + "," + h * 6320 / 21600 +
                                " L" + w * 8352 / 21600 + "," + h * 2295 / 21600 +
                                " z";
                        } else if (shapType == "irregularSeal2") {
                            var d = "M" + w * 11462 / 21600 + "," + h * 4342 / 21600 +
                                " L" + w * 14790 / 21600 + "," + 0 +
                                " L" + w * 14525 / 21600 + "," + h * 5777 / 21600 +
                                " L" + w * 18007 / 21600 + "," + h * 3172 / 21600 +
                                " L" + w * 16380 / 21600 + "," + h * 6532 / 21600 +
                                " L" + w + "," + h * 6645 / 21600 +
                                " L" + w * 16985 / 21600 + "," + h * 9402 / 21600 +
                                " L" + w * 18270 / 21600 + "," + h * 11290 / 21600 +
                                " L" + w * 16380 / 21600 + "," + h * 12310 / 21600 +
                                " L" + w * 18877 / 21600 + "," + h * 15632 / 21600 +
                                " L" + w * 14640 / 21600 + "," + h * 14350 / 21600 +
                                " L" + w * 14942 / 21600 + "," + h * 17370 / 21600 +
                                " L" + w * 12180 / 21600 + "," + h * 15935 / 21600 +
                                " L" + w * 11612 / 21600 + "," + h * 18842 / 21600 +
                                " L" + w * 9872 / 21600 + "," + h * 17370 / 21600 +
                                " L" + w * 8700 / 21600 + "," + h * 19712 / 21600 +
                                " L" + w * 7527 / 21600 + "," + h * 18125 / 21600 +
                                " L" + w * 4917 / 21600 + "," + h +
                                " L" + w * 4805 / 21600 + "," + h * 18240 / 21600 +
                                " L" + w * 1285 / 21600 + "," + h * 17825 / 21600 +
                                " L" + w * 3330 / 21600 + "," + h * 15370 / 21600 +
                                " L" + 0 + "," + h * 12877 / 21600 +
                                " L" + w * 3935 / 21600 + "," + h * 11592 / 21600 +
                                " L" + w * 1172 / 21600 + "," + h * 8270 / 21600 +
                                " L" + w * 5372 / 21600 + "," + h * 7817 / 21600 +
                                " L" + w * 4502 / 21600 + "," + h * 3625 / 21600 +
                                " L" + w * 8550 / 21600 + "," + h * 6382 / 21600 +
                                " L" + w * 9722 / 21600 + "," + h * 1887 / 21600 +
                                " z";
                        }
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "flowChartTerminator": {
                        var x1, x2, y1, cd2 = 180, cd4 = 90, c3d4 = 270;
                        x1 = w * 3475 / 21600;
                        x2 = w * 18125 / 21600;
                        y1 = h * 10800 / 21600;
                        //path attrs: w = 21600; h = 21600; 
                        var d = "M" + x1 + "," + 0 +
                            " L" + x2 + "," + 0 +
                            PPTXShapeUtils.shapeArcAlt(x2, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L") +
                            " L" + x1 + "," + h +
                            PPTXShapeUtils.shapeArcAlt(x1, h / 2, x1, y1, cd4, cd4 + cd2, false).replace("M", "L") +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "flowChartPunchedTape": {
                        var x1, x1, y1, y2, cd2 = 180;
                        x1 = w * 5 / 20;
                        y1 = h * 2 / 20;
                        y2 = h * 18 / 20;
                        var d = "M" + 0 + "," + y1 +
                            PPTXShapeUtils.shapeArcAlt(x1, y1, x1, y1, cd2, 0, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArcAlt(w * (3 / 4), y1, x1, y1, cd2, 360, false).replace("M", "L") +
                            " L" + w + "," + y2 +
                            PPTXShapeUtils.shapeArcAlt(w * (3 / 4), y2, x1, y1, 0, -cd2, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArcAlt(x1, y2, x1, y1, 0, cd2, false).replace("M", "L") +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "flowChartOnlineStorage": {
                        var x1, y1, c3d4 = 270, cd4 = 90;
                        x1 = w * 1 / 6;
                        y1 = h * 3 / 6;
                        var d = "M" + x1 + "," + 0 +
                            " L" + w + "," + 0 +
                            PPTXShapeUtils.shapeArcAlt(w, h / 2, x1, y1, c3d4, 90, false).replace("M", "L") +
                            " L" + x1 + "," + h +
                            PPTXShapeUtils.shapeArcAlt(x1, h / 2, x1, y1, cd4, 270, false).replace("M", "L") +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "flowChartDisplay": {
                        var x1, x2, y1, c3d4 = 270, cd2 = 180;
                        x1 = w * 1 / 6;
                        x2 = w * 5 / 6;
                        y1 = h * 3 / 6;
                        //path attrs: w = 6; h = 6; 
                        var d = "M" + 0 + "," + y1 +
                            " L" + x1 + "," + 0 +
                            " L" + x2 + "," + 0 +
                            PPTXShapeUtils.shapeArcAlt(w, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L") +
                            " L" + x1 + "," + h +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "flowChartDelay": {
                        var wd2 = w / 2, hd2 = h / 2, cd2 = 180, c3d4 = 270, cd4 = 90;
                        var d = "M" + 0 + "," + 0 +
                            " L" + wd2 + "," + 0 +
                            PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, c3d4, c3d4 + cd2, false).replace("M", "L") +
                            " L" + 0 + "," + h +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "flowChartMagneticTape": {
                        var wd2 = w / 2, hd2 = h / 2, cd2 = 180, c3d4 = 270, cd4 = 90;
                        var idy, ib, ang1;
                        idy = hd2 * Math.sin(Math.PI / 4);
                        ib = hd2 + idy;
                        ang1 = Math.atan(h / w);
                        var ang1Dg = ang1 * 180 / Math.PI;
                        var d = "M" + wd2 + "," + h +
                            PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, cd4, cd2, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, cd2, c3d4, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, c3d4, 360, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, 0, ang1Dg, false).replace("M", "L") +
                            " L" + w + "," + ib +
                            " L" + w + "," + h +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "ellipse":
                    case "flowChartConnector":
                    case "flowChartSummingJunction":
                    case "flowChartOr": {
                        result += "<ellipse cx='" + (w / 2) + "' cy='" + (h / 2) + "' rx='" + (w / 2) + "' ry='" + (h / 2) + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        if (shapType == "flowChartOr") {
                            result += " <polyline points='" + w / 2 + " " + 0 + "," + w / 2 + " " + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                            result += " <polyline points='" + 0 + " " + h / 2 + "," + w + " " + h / 2 + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        } else if (shapType == "flowChartSummingJunction") {
                            var iDx, idy, il, ir, it, ib, hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
                            var angVal = Math.PI / 4;
                            iDx = wd2 * Math.cos(angVal);
                            idy = hd2 * Math.sin(angVal);
                            il = hc - iDx;
                            ir = hc + iDx;
                            it = vc - idy;
                            ib = vc + idy;
                            result += " <polyline points='" + il + " " + it + "," + ir + " " + ib + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                            result += " <polyline points='" + ir + " " + it + "," + il + " " + ib + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        }
                        break;
                    }
                    case "roundRect":
                    case "round1Rect":
                    case "round2DiagRect":
                    case "round2SameRect":
                    case "snip1Rect":
                    case "snip2DiagRect":
                    case "snip2SameRect":
                    case "flowChartAlternateProcess":
                    case "flowChartPunchedCard": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, sAdj1_val;// = 0.33334;
                        var sAdj2, sAdj2_val;// = 0.33334;
                        var shpTyp, adjTyp;
                        if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor === Array) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) / 50000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) / 50000;
                                }
                            }
                        } else if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor !== Array) {
                            var sAdj = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                            sAdj1_val = parseInt(sAdj.substr(4)) / 50000;
                            sAdj2_val = 0;
                        }
                        //console.log("shapType: ",shapType,",node: ",node )
                        var tranglRott = "";
                        switch (shapType) {
                            case "roundRect":
                            case "flowChartAlternateProcess": {
                                shpTyp = "round";
                                adjTyp = "cornrAll";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                break;
                            }
                            case "round1Rect": {
                                shpTyp = "round";
                                adjTyp = "cornr1";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                break;
                            }
                            case "round2DiagRect": {
                                shpTyp = "round";
                                adjTyp = "diag";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                            }
                            case "round2SameRect": {
                                shpTyp = "round";
                                adjTyp = "cornr2";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                            }
                            case "snip1Rect":
                            case "flowChartPunchedCard": {
                                shpTyp = "snip";
                                adjTyp = "cornr1";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                if (shapType == "flowChartPunchedCard") {
                                    tranglRott = "transform='translate(" + w + ",0) scale(-1,1)'";
                                }
                                break;
                            }
                            case "snip2DiagRect": {
                                shpTyp = "snip";
                                adjTyp = "diag";
                                if (sAdj1_val === undefined) sAdj1_val = 0;
                                if (sAdj2_val === undefined) sAdj2_val = 0.33334;
                                break;
                            }
                            case "snip2SameRect": {
                                shpTyp = "snip";
                                adjTyp = "cornr2";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                            }
                        }
                        var d_val = PPTXShapeUtils.shapeSnipRoundRectAlt(w, h, sAdj1_val, sAdj2_val, shpTyp, adjTyp);
                        result += "<path " + tranglRott + "  d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "snipRoundRect": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, sAdj1_val = 0.33334;
                        var sAdj2, sAdj2_val = 0.33334;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) / 50000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) / 50000;
                                }
                            }
                        }
                        /**
                         * snipRoundRect: 混合形状，有两个角是圆角，有两个角是缺角
                         *
                         * 形状说明：
                         * - 左上角：凹进去的缺角（直线斜切）
                         * - 右上角：凹进去的缺角（直线斜切）
                         * - 右下角：凹进去的圆角
                         * - 左下角：凹进去的圆角
                         *
                         * 参数说明：
                         * - adj1: 控制圆角的半径（用于右下角和左下角）
                         * - adj2: 控制缺角的大小（用于左上角和右上角）
                         */
                        var radius = Math.min(w, h) * sAdj1_val;     // 圆角半径
                        var snipSize = Math.min(w, h) * sAdj2_val;   // 缺角大小

                        // 生成路径：从左下角开始，逆时针绘制
                        var d_val = "M0," + (h - radius) +           // 左下角圆弧起点
                            " Q0," + h + " " + radius + "," + h +   // 左下角圆弧（凸圆角）
                            " L" + w + "," + h +                    // 沿底边到右下角
                            " Q" + w + "," + h + " " + w + "," + (h - radius) + // 右下角圆弧（凸圆角）
                            " L" + w + "," + snipSize +             // 沿右边向下到缺角位置
                            " L" + (w - snipSize) + ",0" +          // 斜切到左上角缺角
                            " L" + snipSize + ",0" +                // 沿上边向右到右上角缺角位置
                            " L0," + (h - snipSize) +               // 斜切到左下角
                            " z";

                        result += "<path   d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "bentConnector2": {
                        var d = "";
                        // 使用drawW和drawH（原始尺寸）
                        var bendW = (drawW !== undefined) ? drawW : w;
                        var bendH = (drawH !== undefined) ? drawH : h;
                        // 路径方向（SVG容器会通过flip变换处理翻转）
                        d = "M " + bendW + " 0 L " + bendW + " " + bendH + " L 0 " + bendH;
                        result += "<path d='" + d + "' stroke='" + border.color +
                            "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' fill='none' ";
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                        }
                        result += "/>";
                        break;
                    }
                    case "rtTriangle": {
                        result += " <polygon points='0 0,0 " + h + "," + w + " " + h + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "triangle":
                    case "flowChartExtract":
                    case "flowChartMerge": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var shapAdjst_val = 0.5;
                        if (shapAdjst !== undefined) {
                            shapAdjst_val = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                            //console.log("w: "+w+"\nh: "+h+"\nshapAdjst: "+shapAdjst+"\nshapAdjst_val: "+shapAdjst_val);
                        }
                        var tranglRott = "";
                        if (shapType == "flowChartMerge") {
                            tranglRott = "transform='rotate(180 " + w / 2 + "," + h / 2 + ")'";
                        }
                        result += " <polygon " + tranglRott + " points='" + (w * shapAdjst_val) + " 0,0 " + h + "," + w + " " + h + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "diamond":
                    case "flowChartDecision":
                    case "flowChartSort": {
                        result += " <polygon points='" + (w / 2) + " 0,0 " + (h / 2) + "," + (w / 2) + " " + h + "," + w + " " + (h / 2) + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        if (shapType == "flowChartSort") {
                            result += " <polyline points='0 " + h / 2 + "," + w + " " + h / 2 + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        }
                        break;
                    }
                    case "trapezoid":
                    case "flowChartManualOperation":
                    case "flowChartManualInput": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adjst_val = 0.2;
                        var max_adj_const = 0.7407;
                        if (shapAdjst !== undefined) {
                            var adjst = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                            adjst_val = (adjst * 0.5) / max_adj_const;
                            // console.log("w: "+w+"\nh: "+h+"\nshapAdjst: "+shapAdjst+"\nadjst_val: "+adjst_val);
                        }
                        var cnstVal = 0;
                        var tranglRott = "";
                        if (shapType == "flowChartManualOperation") {
                            tranglRott = "transform='rotate(180 " + w / 2 + "," + h / 2 + ")'";
                        }
                        if (shapType == "flowChartManualInput") {
                            adjst_val = 0;
                            cnstVal = h / 5;
                        }
                        result += " <polygon " + tranglRott + " points='" + (w * adjst_val) + " " + cnstVal + ",0 " + h + "," + w + " " + h + "," + (1 - adjst_val) * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "parallelogram":
                    case "flowChartInputOutput": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adjst_val = 0.25;
                        var max_adj_const;
                        if (w > h) {
                            max_adj_const = w / h;
                        } else {
                            max_adj_const = h / w;
                        }
                        if (shapAdjst !== undefined) {
                            var adjst = parseInt(shapAdjst.substr(4)) / 100000;
                            adjst_val = adjst / max_adj_const;
                            //console.log("w: "+w+"\nh: "+h+"\nadjst: "+adjst_val+"\nmax_adj_const: "+max_adj_const);
                        }
                        result += " <polygon points='" + adjst_val * w + " 0,0 " + h + "," + (1 - adjst_val) * w + " " + h + "," + w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "pentagon": {
                        result += " <polygon points='" + (0.5 * w) + " 0,0 " + (0.375 * h) + "," + (0.15 * w) + " " + h + "," + 0.85 * w + " " + h + "," + w + " " + 0.375 * h + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "hexagon":
                    case "flowChartPreparation": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 25000 * SLIDE_FACTOR$1;
                        var vf = 115470 * SLIDE_FACTOR$1;                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var angVal1 = 60 * Math.PI / 180;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                        }
                        var maxAdj, a, shd2, x1, x2, dy1, y1, y2, vc = h / 2, hd2 = h / 2;
                        var ss = Math.min(w, h);
                        maxAdj = cnstVal1 * w / ss;
                        a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
                        shd2 = hd2 * vf / cnstVal2;
                        x1 = ss * a / cnstVal2;
                        x2 = w - x1;
                        dy1 = shd2 * Math.sin(angVal1);
                        y1 = vc - dy1;
                        y2 = vc + dy1;

                        var d = "M" + 0 + "," + vc +
                            " L" + x1 + "," + y1 +
                            " L" + x2 + "," + y1 +
                            " L" + w + "," + vc +
                            " L" + x2 + "," + y2 +
                            " L" + x1 + "," + y2 +
                            " z";

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "heptagon": {
                        result += " <polygon points='" + (0.5 * w) + " 0," + w / 8 + " " + h / 4 + ",0 " + (5 / 8) * h + "," + w / 4 + " " + h + "," + (3 / 4) * w + " " + h + "," +
                            w + " " + (5 / 8) * h + "," + (7 / 8) * w + " " + h / 4 + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "octagon": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj1 = 0.25;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) / 100000;

                        }
                        var adj2 = (1 - adj1);
                        //console.log("adj1: "+adj1+"\nadj2: "+adj2);
                        result += " <polygon points='" + adj1 * w + " 0,0 " + adj1 * h + ",0 " + adj2 * h + "," + adj1 * w + " " + h + "," + adj2 * w + " " + h + "," +
                            w + " " + adj2 * h + "," + w + " " + adj1 * h + "," + adj2 * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "decagon": {
                        result += " <polygon points='" + (3 / 8) * w + " 0," + w / 8 + " " + h / 8 + ",0 " + h / 2 + "," + w / 8 + " " + (7 / 8) * h + "," + (3 / 8) * w + " " + h + "," +
                            (5 / 8) * w + " " + h + "," + (7 / 8) * w + " " + (7 / 8) * h + "," + w + " " + h / 2 + "," + (7 / 8) * w + " " + h / 8 + "," + (5 / 8) * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "dodecagon": {
                        result += " <polygon points='" + (3 / 8) * w + " 0," + w / 8 + " " + h / 8 + ",0 " + (3 / 8) * h + ",0 " + (5 / 8) * h + "," + w / 8 + " " + (7 / 8) * h + "," + (3 / 8) * w + " " + h + "," +
                            (5 / 8) * w + " " + h + "," + (7 / 8) * w + " " + (7 / 8) * h + "," + w + " " + (5 / 8) * h + "," + w + " " + (3 / 8) * h + "," + (7 / 8) * w + " " + h / 8 + "," + (5 / 8) * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "star4":
                    case "star5":
                    case "star6":
                    case "star7":
                    case "star8":
                    case "star10":
                    case "star12":
                    case "star16":
                    case "star24":
                    case "star32": {
                        // 使用drawW和drawH（原始尺寸）进行形状计算
                        result += renderStar(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt, node);
                        break;
                    }
                    case "pie":
                    case "pieWedge":
                    case "arc":
                    case "chord": {
                        // 使用drawW和drawH（原始尺寸）进行形状计算
                        result += renderPieShape(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node, oShadowSvgUrlStr);
                        break;
                    }
                    case "frame": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj1 = 12500 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                        }
                        var a1, x1, x4, y4;
                        if (adj1 < 0) a1 = 0;
                        else if (adj1 > cnstVal1) a1 = cnstVal1;
                        else a1 = adj1;
                        x1 = Math.min(w, h) * a1 / cnstVal2;
                        x4 = w - x1;
                        y4 = h - x1;
                        var d = "M" + 0 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + h +
                            " L" + 0 + "," + h +
                            " z" +
                            "M" + x1 + "," + x1 +
                            " L" + x1 + "," + y4 +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + x1 +
                            " z";
                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "donut": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 25000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                        }
                        var a, dr, iwd2, ihd2;
                        if (adj < 0) a = 0;
                        else if (adj > cnstVal1) a = cnstVal1;
                        else a = adj;
                        dr = Math.min(w, h) * a / cnstVal2;
                        iwd2 = w / 2 - dr;
                        ihd2 = h / 2 - dr;
                        var d = "M" + 0 + "," + h / 2 +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") +
                            " z" +
                            "M" + dr + "," + h / 2 +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 180, 90, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 90, 0, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 0, -90, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 270, 180, false).replace("M", "L") +
                            " z";
                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' " + oShadowSvgUrlStr + " />";
                        break;
                    }
                    case "noSmoking": {
                        /**
                         * noSmoking: 禁止符形状
                         * 参考 pptxjs.js 实现
                         *
                         * 形状说明：
                         * - 一个完整的圆圈
                         * - 中间有一条从左上到右下的斜杠（带圆角）
                         *
                         * 参数说明：
                         * - adj: 控制斜杠的粗细 (范围: 0-50000)
                         */
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 18750 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                        }

                        // 计算调整值
                        var a, dr, iwd2, ihd2, ang, ct, st, m, n;
                        if (adj < 0) a = 0;
                        else if (adj > cnstVal1) a = cnstVal1;
                        else a = adj;

                        // 斜杠宽度
                        dr = Math.min(w, h) * a / cnstVal2;
                        iwd2 = w / 2 - dr;
                        ihd2 = h / 2 - dr;
                        ang = Math.atan(h / w);
                        ct = ihd2 * Math.cos(ang);
                        st = iwd2 * Math.sin(ang);
                        m = Math.sqrt(ct * ct + st * st);
                        n = iwd2 * ihd2 / m;
                        var drd2 = dr / 2;
                        var dang = Math.atan(drd2 / n);
                        var dang2 = dang * 2;
                        var swAng = -Math.PI + dang2;

                        // 绘制路径（参考 pptxjs.js 使用圆弧方式）
                        var stAng1 = ang - dang;
                        var stAng2 = stAng1 - Math.PI;
                        var stAng1deg = stAng1 * 180 / Math.PI;
                        var stAng2deg = stAng2 * 180 / Math.PI;
                        var swAng2deg = swAng * 180 / Math.PI;

                        var dx1 = n * Math.cos(stAng1);
                        var dy1 = n * Math.sin(stAng1);
                        var x1 = w / 2 + dx1;
                        var y1 = h / 2 + dy1;
                        var x2 = w / 2 - dx1;
                        var y2 = h / 2 - dy1;

                        var d = "M" + 0 + "," + h / 2 +
                            shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L") +
                            shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L") +
                            shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
                            shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") +
                            " z" +
                            "M" + x1 + "," + y1 +
                            shapeArcAlt(w / 2, h / 2, iwd2, ihd2, stAng1deg, (stAng1deg + swAng2deg), false).replace("M", "L") +
                            " z" +
                            "M" + x2 + "," + y2 +
                            shapeArcAlt(w / 2, h / 2, iwd2, ihd2, stAng2deg, (stAng2deg + swAng2deg), false).replace("M", "L") +
                            " z";

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "halfFrame": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, sAdj1_val = 3.5;
                        var sAdj2, sAdj2_val = 3.5;
                        var cnsVal = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var minWH = Math.min(w, h);
                        var maxAdj2 = (cnsVal * w) / minWH;
                        var a1, a2;
                        if (sAdj2_val < 0) a2 = 0;
                        else if (sAdj2_val > maxAdj2) a2 = maxAdj2;
                        else a2 = sAdj2_val;
                        var x1 = (minWH * a2) / cnsVal;
                        var g1 = h * x1 / w;
                        var g2 = h - g1;
                        var maxAdj1 = (cnsVal * g2) / minWH;
                        if (sAdj1_val < 0) a1 = 0;
                        else if (sAdj1_val > maxAdj1) a1 = maxAdj1;
                        else a1 = sAdj1_val;
                        var y1 = minWH * a1 / cnsVal;
                        var dx2 = y1 * w / h;
                        var x2 = w - dx2;
                        var dy2 = x1 * h / w;
                        var y2 = h - dy2;
                        var d = "M0,0" +
                            " L" + w + "," + 0 +
                            " L" + x2 + "," + y1 +
                            " L" + x1 + "," + y1 +
                            " L" + x1 + "," + y2 +
                            " L0," + h + " z";

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        //console.log("w: ",w,", h: ",h,", sAdj1_val: ",sAdj1_val,", sAdj2_val: ",sAdj2_val,",maxAdj1: ",maxAdj1,",maxAdj2: ",maxAdj2)
                        break;
                    }
                    case "bracePair":
                    case "bracketPair":
                    case "leftBrace":
                    case "leftBracket":
                    case "rightBrace":
                    case "rightBracket": {
                        // 使用drawW和drawH（原始尺寸）进行形状计算
                        result += renderBracket(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
                        break;
                    }
                    case "moon": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 0.5;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) / 100000;//*96/914400;;
                        }
                        var hd2, cd2, cd4;

                        hd2 = h / 2;
                        cd2 = 180;
                        cd4 = 90;

                        var adj2 = (1 - adj) * w;
                        var d = "M" + w + "," + h +
                            PPTXShapeUtils.shapeArc(w, hd2, w, hd2, cd4, (cd4 + cd2), false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w, hd2, adj2, hd2, (cd4 + cd2), cd4, false).replace("M", "L") +
                            " z";
                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "corner": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, sAdj1_val = 50000 * SLIDE_FACTOR$1;
                        var sAdj2, sAdj2_val = 50000 * SLIDE_FACTOR$1;
                        var cnsVal = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var minWH = Math.min(w, h);
                        var maxAdj1 = cnsVal * h / minWH;
                        var maxAdj2 = cnsVal * w / minWH;
                        var a1, a2, x1, dy1, y1;
                        if (sAdj1_val < 0) a1 = 0;
                        else if (sAdj1_val > maxAdj1) a1 = maxAdj1;
                        else a1 = sAdj1_val;

                        if (sAdj2_val < 0) a2 = 0;
                        else if (sAdj2_val > maxAdj2) a2 = maxAdj2;
                        else a2 = sAdj2_val;
                        x1 = minWH * a2 / cnsVal;
                        dy1 = minWH * a1 / cnsVal;
                        y1 = h - dy1;

                        var d = "M0,0" +
                            " L" + x1 + "," + 0 +
                            " L" + x1 + "," + y1 +
                            " L" + w + "," + y1 +
                            " L" + w + "," + h +
                            " L0," + h + " z";

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "diagStripe": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var sAdj1_val = 50000 * SLIDE_FACTOR$1;
                        var cnsVal = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst !== undefined) {
                            sAdj1_val = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                        }
                        var a1, x2, y2;
                        if (sAdj1_val < 0) a1 = 0;
                        else if (sAdj1_val > cnsVal) a1 = cnsVal;
                        else a1 = sAdj1_val;
                        x2 = w * a1 / cnsVal;
                        y2 = h * a1 / cnsVal;
                        var d = "M" + 0 + "," + y2 +
                            " L" + x2 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + 0 + "," + h + " z";

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "gear6":
                    case "gear9": {
                        var gearNum = shapType.substr(4), d;
                        if (gearNum == "6") {
                            d = shapeGear(w, h / 3.5, parseInt(gearNum));
                        } else { //gearNum=="9"
                            d = shapeGear(w, h / 3.5, parseInt(gearNum));
                        }
                        result += "<path   d='" + d + "' transform='rotate(20," + (3 / 7) * h + "," + (3 / 7) * h + ")' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "bentConnector3": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var shapAdjst_val = 0.5;
                        // 使用drawW和drawH（原始尺寸）
                        var connectorW = (drawW !== undefined) ? drawW : w;
                        var connectorH = (drawH !== undefined) ? drawH : h;
                        if (shapAdjst !== undefined) {
                            shapAdjst_val = parseInt(shapAdjst.substr(4)) / 100000;
                            // 路径方向（SVG容器会通过flip变换处理翻转）
                            result += " <polyline points='0 0," + (shapAdjst_val) * connectorW + " 0," + (shapAdjst_val) * connectorW + " " + connectorH + "," + connectorW + " " + connectorH + "' fill='transparent'" +
                                "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                            if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                                result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                            }
                            if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                                result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                            }
                            result += "/>";
                        }
                        break;
                    }
                    case "plus": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj1 = 0.25;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) / 100000;

                        }
                        var adj2 = (1 - adj1);
                        result += " <polygon points='" + adj1 * w + " 0," + adj1 * w + " " + adj1 * h + ",0 " + adj1 * h + ",0 " + adj2 * h + "," +
                            adj1 * w + " " + adj2 * h + "," + adj1 * w + " " + h + "," + adj2 * w + " " + h + "," + adj2 * w + " " + adj2 * h + "," + w + " " + adj2 * h + "," +
                            +w + " " + adj1 * h + "," + adj2 * w + " " + adj1 * h + "," + adj2 * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "teardrop": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj1 = 100000 * SLIDE_FACTOR$1;
                        var cnsVal1 = adj1;
                        var cnsVal2 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                        }
                        var a1, r2, tw, th, sw, sh, dx1, dy1, x1, y1, x2, y2, rd45;
                        if (adj1 < 0) a1 = 0;
                        else if (adj1 > cnsVal2) a1 = cnsVal2;
                        else a1 = adj1;
                        r2 = Math.sqrt(2);
                        tw = r2 * (w / 2);
                        th = r2 * (h / 2);
                        sw = (tw * a1) / cnsVal1;
                        sh = (th * a1) / cnsVal1;
                        rd45 = (45 * (Math.PI) / 180);
                        dx1 = sw * (Math.cos(rd45));
                        dy1 = sh * (Math.cos(rd45));
                        x1 = (w / 2) + dx1;
                        y1 = (h / 2) - dy1;
                        x2 = ((w / 2) + x1) / 2;
                        y2 = ((h / 2) + y1) / 2;

                        var d_val = PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false) +
                            "Q " + x2 + ",0 " + x1 + "," + y1 +
                            "Q " + w + "," + y2 + " " + w + "," + h / 2 +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") + " z";
                        result += "<path   d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        // console.log("shapAdjst: ",shapAdjst,", adj1: ",adj1);
                        break;
                    }
                    case "plaque": {
                        /**
                         * plaque: 凸出圆角的矩形
                         *
                         * 形状说明：
                         * - 4个角都是向外凸出的1/4圆弧
                         * - 类似凸出卡片的样式
                         * - 半圆弧的圆心在矩形的四个角上
                         *
                         * 参数说明：
                         * - adj: 控制圆角半径大小 (范围: 0-50000)
                         *
                         * 坐标系统：
                         * - (0, 0) 到 (w, h) 的矩形区域
                         * - 四个角向外延伸出1/4圆弧
                         */

                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adjVal = 25000; // 默认值
                        if (shapAdjst !== undefined) {
                            adjVal = parseInt(shapAdjst.substr(4));
                        }
                        // 限制 adj 在有效范围内 (0-50000)
                        if (adjVal < 0) adjVal = 0;
                        else if (adjVal > 50000) adjVal = 50000;

                        // 计算圆弧半径：adj/100000 * min(w, h)
                        var r = (adjVal / 100000) * Math.min(w, h);

                        /**
                         * 路径绘制顺序（逆时针从左上角圆弧开始）：
                         *
                         * 左上角：向外凸出的1/4圆弧，圆心在(0,0)
                         * - 起点: (r, 0)
                         * - 圆弧到: (0, r) - 1/4圆弧（0度到90度）
                         *
                         * 右上角：向外凸出的1/4圆弧，圆心在(w,0)
                         * - 线段到: (w - r, 0)
                         * - 圆弧到: (w, r) - 1/4圆弧（90度到180度）
                         *
                         * 右下角：向外凸出的1/4圆弧，圆心在(w,h)
                         * - 线段到: (w, h - r)
                         * - 圆弧到: (w - r, h) - 1/4圆弧（180度到270度）
                         *
                         * 左下角：向外凸出的1/4圆弧，圆心在(0,h)
                         * - 线段到: (r, h)
                         * - 圆弧到: (0, h - r) - 1/4圆弧（270度到360度）
                         * - 闭合: 回到起点
                         */

                        var d_val = "M" + r + ",0" +
                            "A" + r + " " + r + " 0 0 1 0," + r +
                            "L0," + (h - r) +
                            "A" + r + " " + r + " 0 0 1 " + r + "," + h +
                            "L" + (w - r) + "," + h +
                            "A" + r + " " + r + " 0 0 1 " + w + "," + (h - r) +
                            "L" + w + "," + r +
                            "A" + r + " " + r + " 0 0 1 " + (w - r) + ",0" +
                            " z";

                        result += "<path   d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "sun": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var refr = SLIDE_FACTOR$1;
                        var adj1 = 25000 * refr;
                        var cnstVal1 = 12500 * refr;
                        var cnstVal2 = 46875 * refr;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        var a1;
                        if (adj1 < cnstVal1) a1 = cnstVal1;
                        else if (adj1 > cnstVal2) a1 = cnstVal2;
                        else a1 = adj1;

                        var cnstVa3 = 50000 * refr;
                        var cnstVa4 = 100000 * refr;
                        var g0 = cnstVa3 - a1,
                            g1 = g0 * (30274 * refr) / (32768 * refr),
                            g2 = g0 * (12540 * refr) / (32768 * refr),
                            g3 = g1 + cnstVa3,
                            g4 = g2 + cnstVa3,
                            g5 = cnstVa3 - g1,
                            g6 = cnstVa3 - g2,
                            g7 = g0 * (23170 * refr) / (32768 * refr),
                            g8 = cnstVa3 + g7,
                            g9 = cnstVa3 - g7,
                            g10 = g5 * 3 / 4,
                            g11 = g6 * 3 / 4,
                            g12 = g10 + 3662 * refr,
                            g13 = g11 + 36620 * refr,
                            g14 = g11 + 12500 * refr,
                            g15 = cnstVa4 - g10,
                            g16 = cnstVa4 - g12,
                            g17 = cnstVa4 - g13,
                            g18 = cnstVa4 - g14,
                            ox1 = w * (18436 * refr) / (21600 * refr),
                            oy1 = h * (3163 * refr) / (21600 * refr),
                            ox2 = w * (3163 * refr) / (21600 * refr),
                            oy2 = h * (18436 * refr) / (21600 * refr),
                            x8 = w * g8 / cnstVa4,
                            x9 = w * g9 / cnstVa4,
                            x10 = w * g10 / cnstVa4,
                            x12 = w * g12 / cnstVa4,
                            x13 = w * g13 / cnstVa4,
                            x14 = w * g14 / cnstVa4,
                            x15 = w * g15 / cnstVa4,
                            x16 = w * g16 / cnstVa4,
                            x17 = w * g17 / cnstVa4,
                            x18 = w * g18 / cnstVa4,
                            x19 = w * a1 / cnstVa4,
                            wR = w * g0 / cnstVa4,
                            hR = h * g0 / cnstVa4,
                            y8 = h * g8 / cnstVa4,
                            y9 = h * g9 / cnstVa4,
                            y10 = h * g10 / cnstVa4,
                            y12 = h * g12 / cnstVa4,
                            y13 = h * g13 / cnstVa4,
                            y14 = h * g14 / cnstVa4,
                            y15 = h * g15 / cnstVa4,
                            y16 = h * g16 / cnstVa4,
                            y17 = h * g17 / cnstVa4,
                            y18 = h * g18 / cnstVa4;

                        var d_val = "M" + w + "," + h / 2 +
                            " L" + x15 + "," + y18 +
                            " L" + x15 + "," + y14 +
                            "z" +
                            " M" + ox1 + "," + oy1 +
                            " L" + x16 + "," + y17 +
                            " L" + x13 + "," + y12 +
                            "z" +
                            " M" + w / 2 + "," + 0 +
                            " L" + x18 + "," + y10 +
                            " L" + x14 + "," + y10 +
                            "z" +
                            " M" + ox2 + "," + oy1 +
                            " L" + x17 + "," + y12 +
                            " L" + x12 + "," + y17 +
                            "z" +
                            " M" + 0 + "," + h / 2 +
                            " L" + x10 + "," + y14 +
                            " L" + x10 + "," + y18 +
                            "z" +
                            " M" + ox2 + "," + oy2 +
                            " L" + x12 + "," + y13 +
                            " L" + x17 + "," + y16 +
                            "z" +
                            " M" + w / 2 + "," + h +
                            " L" + x14 + "," + y15 +
                            " L" + x18 + "," + y15 +
                            "z" +
                            " M" + ox1 + "," + oy2 +
                            " L" + x13 + "," + y16 +
                            " L" + x16 + "," + y13 +
                            " z" +
                            " M" + x19 + "," + h / 2 +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, wR, hR, 180, 540, false).replace("M", "L") +
                            " z";
                        //console.log("adj1: ",adj1,d_val);
                        result += "<path   d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";


                        break;
                    }
                    case "heart": {
                        var dx1, dx2, x1, x2, x3, x4, y1;
                        dx1 = w * 49 / 48;
                        dx2 = w * 10 / 48;
                        x1 = w / 2 - dx1;
                        x2 = w / 2 - dx2;
                        x3 = w / 2 + dx2;
                        x4 = w / 2 + dx1;
                        y1 = -h / 3;
                        var d_val = "M" + w / 2 + "," + h / 4 +
                            "C" + x3 + "," + y1 + " " + x4 + "," + h / 4 + " " + w / 2 + "," + h +
                            "C" + x1 + "," + h / 4 + " " + x2 + "," + y1 + " " + w / 2 + "," + h / 4 + " z";

                        result += "<path   d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "lightningBolt": {
                        var x1 = w * 5022 / 21600,
                            x2 = w * 11050 / 21600,
                            x3 = w * 8472 / 21600,
                            x4 = w * 8757 / 21600,
                            x5 = w * 10012 / 21600,
                            x6 = w * 14767 / 21600,
                            x7 = w * 12222 / 21600,
                            x8 = w * 12860 / 21600,
                            x9 = w * 13917 / 21600,
                            x10 = w * 7602 / 21600,
                            x11 = w * 16577 / 21600,
                            y1 = h * 3890 / 21600,
                            y2 = h * 6080 / 21600,
                            y3 = h * 6797 / 21600,
                            y4 = h * 7437 / 21600,
                            y5 = h * 12877 / 21600,
                            y6 = h * 9705 / 21600,
                            y7 = h * 12007 / 21600,
                            y8 = h * 13987 / 21600,
                            y9 = h * 8382 / 21600,
                            y10 = h * 14277 / 21600,
                            y11 = h * 14915 / 21600;

                        var d_val = "M" + x3 + "," + 0 +
                            " L" + x8 + "," + y2 +
                            " L" + x2 + "," + y3 +
                            " L" + x11 + "," + y7 +
                            " L" + x6 + "," + y5 +
                            " L" + w + "," + h +
                            " L" + x5 + "," + y11 +
                            " L" + x7 + "," + y8 +
                            " L" + x1 + "," + y6 +
                            " L" + x10 + "," + y9 +
                            " L" + 0 + "," + y1 + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "cube": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var refr = SLIDE_FACTOR$1;
                        var adj = 25000 * refr;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        var d_val;
                        var cnstVal2 = 100000 * refr;
                        var ss = Math.min(w, h);
                        var a, y1, y4, x4;
                        a = (adj < 0) ? 0 : (adj > cnstVal2) ? cnstVal2 : adj;
                        y1 = ss * a / cnstVal2;
                        y4 = h - y1;
                        x4 = w - y1;
                        d_val = "M" + 0 + "," + y1 +
                            " L" + y1 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + y4 +
                            " L" + x4 + "," + h +
                            " L" + 0 + "," + h +
                            " z" +
                            "M" + 0 + "," + y1 +
                            " L" + x4 + "," + y1 +
                            " M" + x4 + "," + y1 +
                            " L" + w + "," + 0 +
                            "M" + x4 + "," + y1 +
                            " L" + x4 + "," + h;

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "bevel": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var refr = SLIDE_FACTOR$1;
                        var adj = 12500 * refr;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        var d_val;
                        var cnstVal1 = 50000 * refr;
                        var cnstVal2 = 100000 * refr;
                        var ss = Math.min(w, h);
                        var a, x1, x2, y2;
                        a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                        x1 = ss * a / cnstVal2;
                        x2 = w - x1;
                        y2 = h - x1;
                        d_val = "M" + 0 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + h +
                            " L" + 0 + "," + h +
                            " z" +
                            " M" + x1 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + x2 + "," + y2 +
                            " L" + x1 + "," + y2 +
                            " z" +
                            " M" + 0 + "," + 0 +
                            " L" + x1 + "," + x1 +
                            " M" + 0 + "," + h +
                            " L" + x1 + "," + y2 +
                            " M" + w + "," + 0 +
                            " L" + x2 + "," + x1 +
                            " M" + w + "," + h +
                            " L" + x2 + "," + y2;

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "foldedCorner": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var refr = SLIDE_FACTOR$1;
                        var adj = 16667 * refr;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        var d_val;
                        var cnstVal1 = 50000 * refr;
                        var cnstVal2 = 100000 * refr;
                        var ss = Math.min(w, h);
                        var a, dy2, dy1, x1, x2, y2, y1;
                        a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                        dy2 = ss * a / cnstVal2;
                        dy1 = dy2 / 5;
                        x1 = w - dy2;
                        x2 = x1 + dy1;
                        y2 = h - dy2;
                        y1 = y2 + dy1;
                        d_val = "M" + x1 + "," + h +
                            " L" + x2 + "," + y1 +
                            " L" + w + "," + y2 +
                            " L" + x1 + "," + h +
                            " L" + 0 + "," + h +
                            " L" + 0 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + y2;

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "cloud":
                    case "cloudCallout": {
                        // 云形的原始设计是基于 43200x43200 的正方形
                        // 根据 Office Open XML 规范，X坐标使用w缩放，Y坐标使用h缩放

                        // 辅助函数：格式化数字为2位小数
                        function fmt(num) {
                            return parseFloat(num.toFixed(2));
                        }

                        // 生成椭圆弧路径的辅助函数（使用SVG A命令）
                        // 参数：中心点(cx,cy)，半径(rx,ry)，起始角度startAngle，扫描角度sweepAngle
                        function ellipseArc(cx, cy, rx, ry, startAngle, sweepAngle) {
                            var endAngle = startAngle + sweepAngle;
                            // 计算起点和终点
                            var startX = cx + rx * Math.cos(startAngle * Math.PI / 180);
                            var startY = cy + ry * Math.sin(startAngle * Math.PI / 180);
                            var endX = cx + rx * Math.cos(endAngle * Math.PI / 180);
                            var endY = cy + ry * Math.sin(endAngle * Math.PI / 180);
                            
                            // 确定large-arc-flag和sweep-flag
                            var largeArc = Math.abs(sweepAngle) > 180 ? 1 : 0;
                            var sweep = sweepAngle > 0 ? 1 : 0;
                            
                            return {
                                start: { x: fmt(startX), y: fmt(startY) },
                                end: { x: fmt(endX), y: fmt(endY) },
                                path: "A " + fmt(rx) + " " + fmt(ry) + " 0 " + largeArc + " " + sweep + " " + fmt(endX) + " " + fmt(endY)
                            };
                        }

                        // X坐标使用 w 缩放，Y坐标使用 h 缩放
                        var x0 = fmt(w * 3900 / 43200);
                        var y0 = fmt(h * 14370 / 43200);
                        
                        // 半径：RX使用 w 缩放，RY使用 h 缩放
                        var rX1 = fmt(w * 6753 / 43200), rY1 = fmt(h * 9190 / 43200);
                        var rX2 = fmt(w * 5333 / 43200), rY2 = fmt(h * 7267 / 43200);
                        var rX3 = fmt(w * 4365 / 43200), rY3 = fmt(h * 5945 / 43200);
                        var rX4 = fmt(w * 4857 / 43200), rY4 = fmt(h * 6595 / 43200);
                        var rY5 = fmt(h * 7273 / 43200);
                        var rX6 = fmt(w * 6775 / 43200), rY6 = fmt(h * 9220 / 43200);
                        var rX7 = fmt(w * 5785 / 43200), rY7 = fmt(h * 7867 / 43200);
                        var rX8 = fmt(w * 6752 / 43200), rY8 = fmt(h * 9215 / 43200);
                        var rX9 = fmt(w * 7720 / 43200), rY9 = fmt(h * 10543 / 43200);
                        var rX10 = fmt(w * 4360 / 43200), rY10 = fmt(h * 5918 / 43200);
                        var rX11 = fmt(w * 4345 / 43200);

                        // 角度（以度为单位）
                        var sA1 = -11429249 / 60000, wA1 = 7426832 / 60000;
                        var sA2 = -8646143 / 60000, wA2 = 5396714 / 60000;
                        var sA3 = -8748475 / 60000, wA3 = 5983381 / 60000;
                        var sA4 = -7859164 / 60000, wA4 = 7034504 / 60000;
                        var sA5 = -4722533 / 60000, wA5 = 6541615 / 60000;
                        var sA6 = -2776035 / 60000, wA6 = 7816140 / 60000;
                        var sA7 = 37501 / 60000, wA7 = 6842000 / 60000;
                        var sA8 = 1347096 / 60000, wA8 = 6910353 / 60000;
                        var sA9 = 3974558 / 60000, wA9 = 4542661 / 60000;
                        var sA10 = -16496525 / 60000, wA10 = 8804134 / 60000;
                        var sA11 = -14809710 / 60000, wA11 = 9151131 / 60000;

                        // 计算各弧线的中心点
                        // 弧线中心点 = 起点 - 半径 * cos/sin(起始角度)
                        var cX0 = fmt(x0 - rX1 * Math.cos(sA1 * Math.PI / 180));
                        var cY0 = fmt(y0 - rY1 * Math.sin(sA1 * Math.PI / 180));

                        // 生成弧线1
                        var arc1 = ellipseArc(cX0, cY0, rX1, rY1, sA1, wA1);
                        
                        // 计算弧线2的中心点（基于弧线1的终点）
                        var cX1 = fmt(arc1.end.x - rX2 * Math.cos(sA2 * Math.PI / 180));
                        var cY1 = fmt(arc1.end.y - rY2 * Math.sin(sA2 * Math.PI / 180));
                        var arc2 = ellipseArc(cX1, cY1, rX2, rY2, sA2, wA2);
                        
                        // 弧线3
                        var cX2 = fmt(arc2.end.x - rX3 * Math.cos(sA3 * Math.PI / 180));
                        var cY2 = fmt(arc2.end.y - rY3 * Math.sin(sA3 * Math.PI / 180));
                        var arc3 = ellipseArc(cX2, cY2, rX3, rY3, sA3, wA3);
                        
                        // 弧线4
                        var cX3 = fmt(arc3.end.x - rX4 * Math.cos(sA4 * Math.PI / 180));
                        var cY3 = fmt(arc3.end.y - rY4 * Math.sin(sA4 * Math.PI / 180));
                        var arc4 = ellipseArc(cX3, cY3, rX4, rY4, sA4, wA4);
                        
                        // 弧线5
                        var cX4 = fmt(arc4.end.x - rX2 * Math.cos(sA5 * Math.PI / 180));
                        var cY4 = fmt(arc4.end.y - rY5 * Math.sin(sA5 * Math.PI / 180));
                        var arc5 = ellipseArc(cX4, cY4, rX2, rY5, sA5, wA5);
                        
                        // 弧线6
                        var cX5 = fmt(arc5.end.x - rX6 * Math.cos(sA6 * Math.PI / 180));
                        var cY5 = fmt(arc5.end.y - rY6 * Math.sin(sA6 * Math.PI / 180));
                        var arc6 = ellipseArc(cX5, cY5, rX6, rY6, sA6, wA6);
                        
                        // 弧线7
                        var cX6 = fmt(arc6.end.x - rX7 * Math.cos(sA7 * Math.PI / 180));
                        var cY6 = fmt(arc6.end.y - rY7 * Math.sin(sA7 * Math.PI / 180));
                        var arc7 = ellipseArc(cX6, cY6, rX7, rY7, sA7, wA7);
                        
                        // 弧线8
                        var cX7 = fmt(arc7.end.x - rX8 * Math.cos(sA8 * Math.PI / 180));
                        var cY7 = fmt(arc7.end.y - rY8 * Math.sin(sA8 * Math.PI / 180));
                        var arc8 = ellipseArc(cX7, cY7, rX8, rY8, sA8, wA8);
                        
                        // 弧线9
                        var cX8 = fmt(arc8.end.x - rX9 * Math.cos(sA9 * Math.PI / 180));
                        var cY8 = fmt(arc8.end.y - rY9 * Math.sin(sA9 * Math.PI / 180));
                        var arc9 = ellipseArc(cX8, cY8, rX9, rY9, sA9, wA9);
                        
                        // 弧线10
                        var cX9 = fmt(arc9.end.x - rX10 * Math.cos(sA10 * Math.PI / 180));
                        var cY9 = fmt(arc9.end.y - rY10 * Math.sin(sA10 * Math.PI / 180));
                        var arc10 = ellipseArc(cX9, cY9, rX10, rY10, sA10, wA10);
                        
                        // 弧线11
                        var cX10 = fmt(arc10.end.x - rX11 * Math.cos(sA11 * Math.PI / 180));
                        var cY10 = fmt(arc10.end.y - rY3 * Math.sin(sA11 * Math.PI / 180));
                        var arc11 = ellipseArc(cX10, cY10, rX11, rY3, sA11, wA11);

                        // 构建完整路径
                        var d1 = "M" + x0 + "," + y0 + " " +
                            arc1.path + " " +
                            arc2.path + " " +
                            arc3.path + " " +
                            arc4.path + " " +
                            arc5.path + " " +
                            arc6.path + " " +
                            arc7.path + " " +
                            arc8.path + " " +
                            arc9.path + " " +
                            arc10.path + " " +
                            arc11.path + " z";

                        if (shapType == "cloudCallout") {
                            var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                            var refr = SLIDE_FACTOR$1;
                            var sAdj1, adj1 = -20833 * refr;
                            var sAdj2, adj2 = 62500 * refr;
                            if (shapAdjst_ary !== undefined) {
                                for (var i = 0; i < shapAdjst_ary.length; i++) {
                                    var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                    if (sAdj_name == "adj1") {
                                        sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        adj1 = parseInt(sAdj1.substr(4)) * refr;
                                    } else if (sAdj_name == "adj2") {
                                        sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        adj2 = parseInt(sAdj2.substr(4)) * refr;
                                    }
                                }
                            }
                            var d_val;
                            var cnstVal2 = 100000 * refr;
                            var ss = Math.min(w, h);
                            var wd2 = w / 2, hd2 = h / 2;

                            var dxPos, dyPos, xPos, yPos, ht, wt, g2, g3, g4, g5, g6, g7, g8, g9, g10, g11, g12, g13, g14, g15, g16,
                                g17, g18, g19, g20, g21, g22, g23, g24, g25, g26, x23, x24, x25;

                            dxPos = w * adj1 / cnstVal2;
                            dyPos = h * adj2 / cnstVal2;
                            xPos = wd2 + dxPos;
                            yPos = hd2 + dyPos;
                            ht = hd2 * Math.cos(Math.atan(dyPos / dxPos));
                            wt = wd2 * Math.sin(Math.atan(dyPos / dxPos));
                            g2 = wd2 * Math.cos(Math.atan(wt / ht));
                            g3 = hd2 * Math.sin(Math.atan(wt / ht));
                            //console.log("adj1: ",adj1,"adj2: ",adj2)
                            if (adj1 >= 0) {
                                g4 = wd2 + g2;
                                g5 = hd2 + g3;
                            } else {
                                g4 = wd2 - g2;
                                g5 = hd2 - g3;
                            }
                            g6 = g4 - xPos;
                            g7 = g5 - yPos;
                            g8 = Math.sqrt(g6 * g6 + g7 * g7);
                            g9 = ss * 6600 / 21600;
                            g10 = g8 - g9;
                            g11 = g10 / 3;
                            g12 = ss * 1800 / 21600;
                            g13 = g11 + g12;
                            g14 = g13 * g6 / g8;
                            g15 = g13 * g7 / g8;
                            g16 = g14 + xPos;
                            g17 = g15 + yPos;
                            g18 = ss * 4800 / 21600;
                            g19 = g11 * 2;
                            g20 = g18 + g19;
                            g21 = g20 * g6 / g8;
                            g22 = g20 * g7 / g8;
                            g23 = g21 + xPos;
                            g24 = g22 + yPos;
                            g25 = ss * 1200 / 21600;
                            g26 = ss * 600 / 21600;
                            x23 = xPos + g26;
                            x24 = g16 + g25;
                            x25 = g23 + g12;

                            d_val = //" M" + x23 + "," + yPos + 
                                PPTXShapeUtils.shapeArc(x23 - g26, yPos, g26, g26, 0, 360, false) + //.replace("M","L") +
                                " z" +
                                " M" + x24 + "," + g17 +
                                PPTXShapeUtils.shapeArc(x24 - g25, g17, g25, g25, 0, 360, false).replace("M", "L") +
                                " z" +
                                " M" + x25 + "," + g24 +
                                PPTXShapeUtils.shapeArc(x25 - g12, g24, g12, g12, 0, 360, false).replace("M", "L") +
                                " z";
                            d1 += d_val;
                        }
                        result += "<path d='" + d1 + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "smileyFace":
                    case "verticalScroll":
                    case "horizontalScroll": {
                        // 使用drawW和drawH（原始尺寸）进行形状计算
                        result += renderMiscShape(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
                        break;
                    }
                    case "wedgeEllipseCallout": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var refr = SLIDE_FACTOR$1;
                        var sAdj1, adj1 = -20833 * refr;
                        var sAdj2, adj2 = 62500 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal1 = 100000 * SLIDE_FACTOR$1;
                        var angVal1 = 11 * Math.PI / 180;
                        var ss = Math.min(w, h);
                        var dxPos, dyPos, xPos, yPos, sdx, sdy, pang, stAng, enAng, dx1, dy1, x1, y1, dx2, dy2,
                            x2, y2, stAng1, swAng2, swAng,
                            vc = h / 2, hc = w / 2;
                        dxPos = w * adj1 / cnstVal1;
                        dyPos = h * adj2 / cnstVal1;
                        xPos = hc + dxPos;
                        yPos = vc + dyPos;
                        sdx = dxPos * h;
                        sdy = dyPos * w;
                        pang = Math.atan(sdy / sdx);
                        stAng = pang + angVal1;
                        enAng = pang - angVal1;
                        dx1 = hc * Math.cos(stAng);
                        dy1 = vc * Math.sin(stAng);
                        dx2 = hc * Math.cos(enAng);
                        dy2 = vc * Math.sin(enAng);
                        if (dxPos >= 0) {
                            x1 = hc + dx1;
                            y1 = vc + dy1;
                            x2 = hc + dx2;
                            y2 = vc + dy2;
                        } else {
                            x1 = hc - dx1;
                            y1 = vc - dy1;
                            x2 = hc - dx2;
                            y2 = vc - dy2;
                        }
                        /*
                        //stAng = pang+angVal1;
                        //enAng = pang-angVal1;
                        //dx1 = hc*Math.cos(stAng);
                        //dy1 = vc*Math.sin(stAng);
                        x1 = hc+dx1;
                        y1 = vc+dy1;
                        dx2 = hc*Math.cos(enAng);
                        dy2 = vc*Math.sin(enAng);
                        x2 = hc+dx2;
                        y2 = vc+dy2;
                        stAng1 = Math.atan(dy1/dx1);
                        enAng1 = Math.atan(dy2/dx2);
                        swAng1 = enAng1-stAng1;
                        swAng2 = swAng1+2*Math.PI;
                        swAng = (swAng1 > 0)?swAng1:swAng2;
                        var stAng1Dg = stAng1*180/Math.PI;
                        var swAngDg = swAng*180/Math.PI;
                        var endAng = stAng1Dg + swAngDg;
                        */
                        d_val = "M" + x1 + "," + y1 +
                            " L" + xPos + "," + yPos +
                            " L" + x2 + "," + y2 +
                            //" z" +
                            PPTXShapeUtils.shapeArcAlt(hc, vc, hc, vc, 0, 360, true);// +
                        //PPTXShapeUtils.shapeArc(hc,vc,hc,vc,stAng1Dg,stAng1Dg+swAngDg,false).replace("M","L") +
                        //" z";
                        result += "<path d='" + d_val + "'" + cloudTransformAttr + " fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "wedgeRectCallout": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var refr = SLIDE_FACTOR$1;
                        var sAdj1, adj1 = -20833 * refr;
                        var sAdj2, adj2 = 62500 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal1 = 100000 * SLIDE_FACTOR$1;
                        var dxPos, dyPos, xPos, yPos, dx, dy, dq, ady, adq, dz, xg1, xg2, x1, x2,
                            yg1, yg2, y1, y2, t1, xl, t2, xt, t3, xr, t4, xb, t5, yl, t6, yt, t7, yr, t8, yb,
                            vc = h / 2, hc = w / 2;
                        dxPos = w * adj1 / cnstVal1;
                        dyPos = h * adj2 / cnstVal1;
                        xPos = hc + dxPos;
                        yPos = vc + dyPos;
                        dx = xPos - hc;
                        dy = yPos - vc;
                        dq = dxPos * h / w;
                        ady = Math.abs(dyPos);
                        adq = Math.abs(dq);
                        dz = ady - adq;
                        xg1 = (dxPos > 0) ? 7 : 2;
                        xg2 = (dxPos > 0) ? 10 : 5;
                        x1 = w * xg1 / 12;
                        x2 = w * xg2 / 12;
                        yg1 = (dyPos > 0) ? 7 : 2;
                        yg2 = (dyPos > 0) ? 10 : 5;
                        y1 = h * yg1 / 12;
                        y2 = h * yg2 / 12;
                        t1 = (dxPos > 0) ? 0 : xPos;
                        xl = (dz > 0) ? 0 : t1;
                        t2 = (dyPos > 0) ? x1 : xPos;
                        xt = (dz > 0) ? t2 : x1;
                        t3 = (dxPos > 0) ? xPos : w;
                        xr = (dz > 0) ? w : t3;
                        t4 = (dyPos > 0) ? xPos : x1;
                        xb = (dz > 0) ? t4 : x1;
                        t5 = (dxPos > 0) ? y1 : yPos;
                        yl = (dz > 0) ? y1 : t5;
                        t6 = (dyPos > 0) ? 0 : yPos;
                        yt = (dz > 0) ? t6 : 0;
                        t7 = (dxPos > 0) ? yPos : y1;
                        yr = (dz > 0) ? y1 : t7;
                        t8 = (dyPos > 0) ? yPos : h;
                        yb = (dz > 0) ? t8 : h;

                        d_val = "M" + 0 + "," + 0 +
                            " L" + x1 + "," + 0 +
                            " L" + xt + "," + yt +
                            " L" + x2 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + y1 +
                            " L" + xr + "," + yr +
                            " L" + w + "," + y2 +
                            " L" + w + "," + h +
                            " L" + x2 + "," + h +
                            " L" + xb + "," + yb +
                            " L" + x1 + "," + h +
                            " L" + 0 + "," + h +
                            " L" + 0 + "," + y2 +
                            " L" + xl + "," + yl +
                            " L" + 0 + "," + y1 +
                            " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "wedgeRoundRectCallout": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var refr = SLIDE_FACTOR$1;
                        var sAdj1, adj1 = -20833 * refr;
                        var sAdj2, adj2 = 62500 * refr;
                        var sAdj3, adj3 = 16667 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * refr;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal1 = 100000 * SLIDE_FACTOR$1;
                        var ss = Math.min(w, h);
                        var dxPos, dyPos, xPos, yPos, dq, ady, adq, dz, xg1, xg2, x1, x2, yg1, yg2, y1, y2,
                            t1, xl, t2, xt, t3, xr, t4, xb, t5, yl, t6, yt, t7, yr, t8, yb, u1, u2, v2,
                            vc = h / 2, hc = w / 2;
                        dxPos = w * adj1 / cnstVal1;
                        dyPos = h * adj2 / cnstVal1;
                        xPos = hc + dxPos;
                        yPos = vc + dyPos;
                        dq = dxPos * h / w;
                        ady = Math.abs(dyPos);
                        adq = Math.abs(dq);
                        dz = ady - adq;
                        xg1 = (dxPos > 0) ? 7 : 2;
                        xg2 = (dxPos > 0) ? 10 : 5;
                        x1 = w * xg1 / 12;
                        x2 = w * xg2 / 12;
                        yg1 = (dyPos > 0) ? 7 : 2;
                        yg2 = (dyPos > 0) ? 10 : 5;
                        y1 = h * yg1 / 12;
                        y2 = h * yg2 / 12;
                        t1 = (dxPos > 0) ? 0 : xPos;
                        xl = (dz > 0) ? 0 : t1;
                        t2 = (dyPos > 0) ? x1 : xPos;
                        xt = (dz > 0) ? t2 : x1;
                        t3 = (dxPos > 0) ? xPos : w;
                        xr = (dz > 0) ? w : t3;
                        t4 = (dyPos > 0) ? xPos : x1;
                        xb = (dz > 0) ? t4 : x1;
                        t5 = (dxPos > 0) ? y1 : yPos;
                        yl = (dz > 0) ? y1 : t5;
                        t6 = (dyPos > 0) ? 0 : yPos;
                        yt = (dz > 0) ? t6 : 0;
                        t7 = (dxPos > 0) ? yPos : y1;
                        yr = (dz > 0) ? y1 : t7;
                        t8 = (dyPos > 0) ? yPos : h;
                        yb = (dz > 0) ? t8 : h;
                        u1 = ss * adj3 / cnstVal1;
                        u2 = w - u1;
                        v2 = h - u1;
                        d_val = "M" + 0 + "," + u1 +
                            PPTXShapeUtils.shapeArc(u1, u1, u1, u1, 180, 270, false).replace("M", "L") +
                            " L" + x1 + "," + 0 +
                            " L" + xt + "," + yt +
                            " L" + x2 + "," + 0 +
                            " L" + u2 + "," + 0 +
                            PPTXShapeUtils.shapeArc(u2, u1, u1, u1, 270, 360, false).replace("M", "L") +
                            " L" + w + "," + y1 +
                            " L" + xr + "," + yr +
                            " L" + w + "," + y2 +
                            " L" + w + "," + v2 +
                            PPTXShapeUtils.shapeArc(u2, v2, u1, u1, 0, 90, false).replace("M", "L") +
                            " L" + x2 + "," + h +
                            " L" + xb + "," + yb +
                            " L" + x1 + "," + h +
                            " L" + u1 + "," + h +
                            PPTXShapeUtils.shapeArc(u1, v2, u1, u1, 90, 180, false).replace("M", "L") +
                            " L" + 0 + "," + y2 +
                            " L" + xl + "," + yl +
                            " L" + 0 + "," + y1 +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "accentBorderCallout1":
                    case "accentBorderCallout2":
                    case "accentBorderCallout3":
                    case "borderCallout1":
                    case "borderCallout2":
                    case "borderCallout3":
                    case "accentCallout1":
                    case "accentCallout2":
                    case "accentCallout3":
                    case "callout1":
                    case "callout2":
                    case "callout3": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var refr = SLIDE_FACTOR$1;
                        var sAdj1, adj1 = 18750 * refr;
                        var sAdj2, adj2 = -8333 * refr;
                        var sAdj3, adj3 = 18750 * refr;
                        var sAdj4, adj4 = -16667 * refr;
                        var sAdj5, adj5 = 100000 * refr;
                        var sAdj6, adj6 = -16667 * refr;
                        var sAdj7, adj7 = 112963 * refr;
                        var sAdj8, adj8 = -8333 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * refr;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * refr;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * refr;
                                } else if (sAdj_name == "adj6") {
                                    sAdj6 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj6 = parseInt(sAdj6.substr(4)) * refr;
                                } else if (sAdj_name == "adj7") {
                                    sAdj7 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj7 = parseInt(sAdj7.substr(4)) * refr;
                                } else if (sAdj_name == "adj8") {
                                    sAdj8 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj8 = parseInt(sAdj8.substr(4)) * refr;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal1 = 100000 * refr;
                        switch (shapType) {
                            case "borderCallout1":
                            case "callout1":
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 112500 * refr;
                                    adj4 = -38333 * refr;
                                }
                                var y1, x1, y2, x2;
                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +
                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2;
                                break;
                            case "borderCallout2":
                            case "callout2":
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;

                                    adj5 = 112500 * refr;
                                    adj6 = -46667 * refr;
                                }
                                var y1, x1, y2, x2, y3, x3;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;

                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +

                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +

                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2;

                                break;
                            case "borderCallout3":
                            case "callout3":
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;

                                    adj5 = 100000 * refr;
                                    adj6 = -16667 * refr;

                                    adj7 = 112963 * refr;
                                    adj8 = -8333 * refr;
                                }
                                var y1, x1, y2, x2, y3, x3, y4, x4;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;

                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;

                                y4 = h * adj7 / cnstVal1;
                                x4 = w * adj8 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +

                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +

                                    " L" + x3 + "," + y3 +

                                    " L" + x4 + "," + y4 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2;
                                break;
                            case "accentBorderCallout1":
                            case "accentCallout1":

                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 112500 * refr;
                                    adj4 = -38333 * refr;
                                }
                                var y1, x1, y2, x2;
                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +

                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +

                                    " M" + x1 + "," + 0 +
                                    " L" + x1 + "," + h;
                                break;
                            case "accentBorderCallout2":
                            case "accentCallout2":
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;
                                    adj5 = 112500 * refr;
                                    adj6 = -46667 * refr;
                                }
                                var y1, x1, y2, x2, y3, x3;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +

                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2 +

                                    " M" + x1 + "," + 0 +
                                    " L" + x1 + "," + h;

                                break;
                            case "accentBorderCallout3":
                            case "accentCallout3":
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;
                                    adj5 = 100000 * refr;
                                    adj6 = -16667 * refr;
                                    adj7 = 112963 * refr;
                                    adj8 = -8333 * refr;
                                }
                                var y1, x1, y2, x2, y3, x3, y4, x4;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                y4 = h * adj7 / cnstVal1;
                                x4 = w * adj8 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +

                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x4 + "," + y4 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2 +

                                    " M" + x1 + "," + 0 +
                                    " L" + x1 + "," + h;
                                break;
                        }

                        //console.log("shapType: ", shapType, ",isBorder:", isBorder)
                        //if(isBorder){
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        //}else{
                        //    result += "<path d='"+d_val+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                        //        "' stroke='none' />";

                        //}
                        break;
                    }
                    case "leftRightRibbon": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var refr = SLIDE_FACTOR$1;
                        var sAdj1, adj1 = 50000 * refr;
                        var sAdj2, adj2 = 50000 * refr;
                        var sAdj3, adj3 = 16667 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * refr;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal1 = 33333 * refr;
                        var cnstVal2 = 100000 * refr;
                        var cnstVal3 = 200000 * refr;
                        var cnstVal4 = 400000 * refr;
                        var ss = Math.min(w, h);
                        var a3, maxAdj1, a1, w1, maxAdj2, a2, x1, x4, dy1, dy2, ly1, ry4, ly2, ry3, ly4, ry1,
                            ly3, ry2, hR, x2, x3, y1, y2, wd32 = w / 32, vc = h / 2, hc = w / 2;

                        a3 = (adj3 < 0) ? 0 : (adj3 > cnstVal1) ? cnstVal1 : adj3;
                        maxAdj1 = cnstVal2 - a3;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        w1 = hc - wd32;
                        maxAdj2 = cnstVal2 * w1 / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        x1 = ss * a2 / cnstVal2;
                        x4 = w - x1;
                        dy1 = h * a1 / cnstVal3;
                        dy2 = h * a3 / -cnstVal3;
                        ly1 = vc + dy2 - dy1;
                        ry4 = vc + dy1 - dy2;
                        ly2 = ly1 + dy1;
                        ry3 = h - ly2;
                        ly4 = ly2 * 2;
                        ry1 = h - ly4;
                        ly3 = ly4 - ly1;
                        ry2 = h - ly3;
                        hR = a3 * ss / cnstVal4;
                        x2 = hc - wd32;
                        x3 = hc + wd32;
                        y1 = ly1 + hR;
                        y2 = ry2 - hR;

                        d_val = "M" + 0 + "," + ly2 +
                            "L" + x1 + "," + 0 +
                            "L" + x1 + "," + ly1 +
                            "L" + hc + "," + ly1 +
                            PPTXShapeUtils.shapeArcAlt(hc, y1, wd32, hR, 270, 450, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArcAlt(hc, y2, wd32, hR, 270, 90, false).replace("M", "L") +
                            "L" + x4 + "," + ry2 +
                            "L" + x4 + "," + ry1 +
                            "L" + w + "," + ry3 +
                            "L" + x4 + "," + h +
                            "L" + x4 + "," + ry4 +
                            "L" + hc + "," + ry4 +
                            PPTXShapeUtils.shapeArc(hc, ry4 - hR, wd32, hR, 90, 180, false).replace("M", "L") +
                            "L" + x2 + "," + ly3 +
                            "L" + x1 + "," + ly3 +
                            "L" + x1 + "," + ly4 +
                            " z" +
                            "M" + x3 + "," + y1 +
                            "L" + x3 + "," + ry2 +
                            "M" + x2 + "," + y2 +
                            "L" + x2 + "," + ly3;

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "ribbon":
                    case "ribbon2": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 16667 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal1 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 33333 * SLIDE_FACTOR$1;
                        var cnstVal3 = 75000 * SLIDE_FACTOR$1;
                        var cnstVal4 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal5 = 200000 * SLIDE_FACTOR$1;
                        var cnstVal6 = 400000 * SLIDE_FACTOR$1;
                        var hc = w / 2, t = 0, l = 0, b = h, r = w, wd8 = w / 8, wd32 = w / 32;
                        var a1, a2, x10, dx2, x2, x9, x3, x8, x5, x6, x4, x7, y1, y2, y4, y3, hR, y6;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        a2 = (adj2 < cnstVal1) ? cnstVal1 : (adj2 > cnstVal3) ? cnstVal3 : adj2;
                        x10 = r - wd8;
                        dx2 = w * a2 / cnstVal5;
                        x2 = hc - dx2;
                        x9 = hc + dx2;
                        x3 = x2 + wd32;
                        x8 = x9 - wd32;
                        x5 = x2 + wd8;
                        x6 = x9 - wd8;
                        x4 = x5 - wd32;
                        x7 = x6 + wd32;
                        hR = h * a1 / cnstVal6;
                        if (shapType == "ribbon2") {
                            var dy1, dy2, y7;
                            dy1 = h * a1 / cnstVal5;
                            y1 = b - dy1;
                            dy2 = h * a1 / cnstVal4;
                            y2 = b - dy2;
                            y4 = t + dy2;
                            y3 = (y4 + b) / 2;
                            y6 = b - hR;///////////////////
                            y7 = y1 - hR;

                            d_val = "M" + l + "," + b +
                                " L" + wd8 + "," + y3 +
                                " L" + l + "," + y4 +
                                " L" + x2 + "," + y4 +
                                " L" + x2 + "," + hR +
                                PPTXShapeUtils.shapeArcAlt(x3, hR, wd32, hR, 180, 270, false).replace("M", "L") +
                            " L" + x8 + "," + t +
                            PPTXShapeUtils.shapeArcAlt(x8, hR, wd32, hR, 270, 360, false).replace("M", "L") +
                                " L" + x9 + "," + y4 +
                                " L" + x9 + "," + y4 +
                                " L" + r + "," + y4 +
                                " L" + x10 + "," + y3 +
                                " L" + r + "," + b +
                                " L" + x7 + "," + b +
                                PPTXShapeUtils.shapeArc(x7, y6, wd32, hR, 90, 270, false).replace("M", "L") +
                                " L" + x8 + "," + y1 +
                                PPTXShapeUtils.shapeArc(x8, y7, wd32, hR, 90, -90, false).replace("M", "L") +
                                " L" + x3 + "," + y2 +
                                PPTXShapeUtils.shapeArc(x3, y7, wd32, hR, 270, 90, false).replace("M", "L") +
                                " L" + x4 + "," + y1 +
                                PPTXShapeUtils.shapeArc(x4, y6, wd32, hR, 270, 450, false).replace("M", "L") +
                                " z" +
                                " M" + x5 + "," + y2 +
                                " L" + x5 + "," + y6 +
                                "M" + x6 + "," + y6 +
                                " L" + x6 + "," + y2 +
                                "M" + x2 + "," + y7 +
                                " L" + x2 + "," + y4 +
                                "M" + x9 + "," + y4 +
                                " L" + x9 + "," + y7;
                        } else if (shapType == "ribbon") {
                            var y5;
                            y1 = h * a1 / cnstVal5;
                            y2 = h * a1 / cnstVal4;
                            y4 = b - y2;
                            y3 = y4 / 2;
                            y5 = b - hR; ///////////////////////
                            y6 = y2 - hR;
                            d_val = "M" + l + "," + t +
                                " L" + x4 + "," + t +
                                PPTXShapeUtils.shapeArcAlt(x4, hR, wd32, hR, 270, 450, false).replace("M", "L") +
                                " L" + x3 + "," + y1 +
                                PPTXShapeUtils.shapeArcAlt(x3, y6, wd32, hR, 270, 90, false).replace("M", "L") +
                                " L" + x8 + "," + y2 +
                                PPTXShapeUtils.shapeArcAlt(x8, y6, wd32, hR, 90, -90, false).replace("M", "L") +
                                " L" + x7 + "," + y1 +
                                PPTXShapeUtils.shapeArcAlt(x7, hR, wd32, hR, 90, 270, false).replace("M", "L") +
                                " L" + r + "," + t +
                                " L" + x10 + "," + y3 +
                                " L" + r + "," + y4 +
                                " L" + x9 + "," + y4 +
                                " L" + x9 + "," + y5 +
                                PPTXShapeUtils.shapeArc(x8, y5, wd32, hR, 0, 90, false).replace("M", "L") +
                                " L" + x3 + "," + b +
                                PPTXShapeUtils.shapeArc(x3, y5, wd32, hR, 90, 180, false).replace("M", "L") +
                                " L" + x2 + "," + y4 +
                                " L" + l + "," + y4 +
                                " L" + wd8 + "," + y3 +
                                " z" +
                                " M" + x5 + "," + hR +
                                " L" + x5 + "," + y2 +
                                "M" + x6 + "," + y2 +
                                " L" + x6 + "," + hR +
                                "M" + x2 + "," + y4 +
                                " L" + x2 + "," + y6 +
                                "M" + x9 + "," + y6 +
                                " L" + x9 + "," + y4;
                        }
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "doubleWave":
                    case "wave": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = (shapType == "doubleWave") ? 6250 * SLIDE_FACTOR$1 : 12500 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 0;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal2 = -1e4 * SLIDE_FACTOR$1;
                        var cnstVal3 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal4 = 100000 * SLIDE_FACTOR$1;
                        var hc = w / 2, t = 0, l = 0, b = h, r = w, wd8 = w / 8, wd32 = w / 32;
                        if (shapType == "doubleWave") {
                            var cnstVal1 = 12500 * SLIDE_FACTOR$1;
                            var a1, a2, y1, dy2, y2, y3, y4, y5, y6, of2, dx2, x2, dx8, x8, dx3, x3, dx4, x4, x5, x6, x7, x9, x15, x10, x11, x12, x13, x14;
                            a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
                            a2 = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
                            y1 = h * a1 / cnstVal4;
                            dy2 = y1 * 10 / 3;
                            y2 = y1 - dy2;
                            y3 = y1 + dy2;
                            y4 = b - y1;
                            y5 = y4 - dy2;
                            y6 = y4 + dy2;
                            of2 = w * a2 / cnstVal3;
                            dx2 = (of2 > 0) ? 0 : of2;
                            x2 = l - dx2;
                            dx8 = (of2 > 0) ? of2 : 0;
                            x8 = r - dx8;
                            dx3 = (dx2 + x8) / 6;
                            x3 = x2 + dx3;
                            dx4 = (dx2 + x8) / 3;
                            x4 = x2 + dx4;
                            x5 = (x2 + x8) / 2;
                            x6 = x5 + dx3;
                            x7 = (x6 + x8) / 2;
                            x9 = l + dx8;
                            x15 = r + dx2;
                            x10 = x9 + dx3;
                            x11 = x9 + dx4;
                            x12 = (x9 + x15) / 2;
                            x13 = x12 + dx3;
                            x14 = (x13 + x15) / 2;

                            d_val = "M" + x2 + "," + y1 +
                                " C" + x3 + "," + y2 + " " + x4 + "," + y3 + " " + x5 + "," + y1 +
                                " C" + x6 + "," + y2 + " " + x7 + "," + y3 + " " + x8 + "," + y1 +
                                " L" + x15 + "," + y4 +
                                " C" + x14 + "," + y6 + " " + x13 + "," + y5 + " " + x12 + "," + y4 +
                                " C" + x11 + "," + y6 + " " + x10 + "," + y5 + " " + x9 + "," + y4 +
                                " z";
                        } else if (shapType == "wave") {
                            var cnstVal5 = 20000 * SLIDE_FACTOR$1;
                            var a1, a2, y1, dy2, y2, y3, y4, y5, y6, of2, dx2, x2, dx5, x5, dx3, x3, x4, x6, x10, x7, x8;
                            a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
                            a2 = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
                            y1 = h * a1 / cnstVal4;
                            dy2 = y1 * 10 / 3;
                            y2 = y1 - dy2;
                            y3 = y1 + dy2;
                            y4 = b - y1;
                            y5 = y4 - dy2;
                            y6 = y4 + dy2;
                            of2 = w * a2 / cnstVal3;
                            dx2 = (of2 > 0) ? 0 : of2;
                            x2 = l - dx2;
                            dx5 = (of2 > 0) ? of2 : 0;
                            x5 = r - dx5;
                            dx3 = (dx2 + x5) / 3;
                            x3 = x2 + dx3;
                            x4 = (x3 + x5) / 2;
                            x6 = l + dx5;
                            x10 = r + dx2;
                            x7 = x6 + dx3;
                            x8 = (x7 + x10) / 2;

                            d_val = "M" + x2 + "," + y1 +
                                " C" + x3 + "," + y2 + " " + x4 + "," + y3 + " " + x5 + "," + y1 +
                                " L" + x10 + "," + y4 +
                                " C" + x8 + "," + y6 + " " + x7 + "," + y5 + " " + x6 + "," + y4 +
                                " z";
                        }
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "ellipseRibbon":
                    case "ellipseRibbon2": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 12500 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal1 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 75000 * SLIDE_FACTOR$1;
                        var cnstVal4 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal5 = 200000 * SLIDE_FACTOR$1;
                        var hc = w / 2, t = 0, l = 0, b = h, r = w, wd8 = w / 8;
                        var a1, a2, q10, q11, q12, minAdj3, a3, dx2, x2, x3, x4, x5, x6, dy1, f1, q1, q2,
                            cx1, cx2, q1, dy3, q3, q4, q5, rh, q8, cx4, q9, cx5;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal4) ? cnstVal4 : adj1;
                        a2 = (adj2 < cnstVal1) ? cnstVal1 : (adj2 > cnstVal3) ? cnstVal3 : adj2;
                        q10 = cnstVal4 - a1;
                        q11 = q10 / 2;
                        q12 = a1 - q11;
                        minAdj3 = (0 > q12) ? 0 : q12;
                        a3 = (adj3 < minAdj3) ? minAdj3 : (adj3 > a1) ? a1 : adj3;
                        dx2 = w * a2 / cnstVal5;
                        x2 = hc - dx2;
                        x3 = x2 + wd8;
                        x4 = r - x3;
                        x5 = r - x2;
                        x6 = r - wd8;
                        dy1 = h * a3 / cnstVal4;
                        f1 = 4 * dy1 / w;
                        q1 = x3 * x3 / w;
                        q2 = x3 - q1;
                        cx1 = x3 / 2;
                        cx2 = r - cx1;
                        q1 = h * a1 / cnstVal4;
                        dy3 = q1 - dy1;
                        q3 = x2 * x2 / w;
                        q4 = x2 - q3;
                        q5 = f1 * q4;
                        rh = b - q1;
                        q8 = dy1 * 14 / 16;
                        cx4 = x2 / 2;
                        q9 = f1 * cx4;
                        cx5 = r - cx4;
                        if (shapType == "ellipseRibbon") {
                            var y1, cy1, y3, q6, q7, cy3, y2, y5, y6,
                                cy4, cy6, y7, y8;
                            y1 = f1 * q2;
                            cy1 = f1 * cx1;
                            y3 = q5 + dy3;
                            q6 = dy1 + dy3 - y3;
                            q7 = q6 + dy1;
                            cy3 = q7 + dy3;
                            y2 = (q8 + rh) / 2;
                            y5 = q5 + rh;
                            y6 = y3 + rh;
                            cy4 = q9 + rh;
                            cy6 = cy3 + rh;
                            y7 = y1 + dy3;
                            y8 = b - dy1;
                            //
                            d_val = "M" + l + "," + t +
                                " Q" + cx1 + "," + cy1 + " " + x3 + "," + y1 +
                                " L" + x2 + "," + y3 +
                                " Q" + hc + "," + cy3 + " " + x5 + "," + y3 +
                                " L" + x4 + "," + y1 +
                                " Q" + cx2 + "," + cy1 + " " + r + "," + t +
                                " L" + x6 + "," + y2 +
                                " L" + r + "," + rh +
                                " Q" + cx5 + "," + cy4 + " " + x5 + "," + y5 +
                                " L" + x5 + "," + y6 +
                                " Q" + hc + "," + cy6 + " " + x2 + "," + y6 +
                                " L" + x2 + "," + y5 +
                                " Q" + cx4 + "," + cy4 + " " + l + "," + rh +
                                " L" + wd8 + "," + y2 +
                                " z" +
                                "M" + x2 + "," + y5 +
                                " L" + x2 + "," + y3 +
                                "M" + x5 + "," + y3 +
                                " L" + x5 + "," + y5 +
                                "M" + x3 + "," + y1 +
                                " L" + x3 + "," + y7 +
                                "M" + x4 + "," + y7 +
                                " L" + x4 + "," + y1;
                        } else if (shapType == "ellipseRibbon2") {
                            var u1, y1, cu1, cy1, q3, q5, u3, y3, q6, q7, cu3, cy3, rh, q8, u2, y2,
                                u5, y5, u6, y6, cu4, cy4, cu6, cy6, u7, y7;
                            u1 = f1 * q2;
                            y1 = b - u1;
                            cu1 = f1 * cx1;
                            cy1 = b - cu1;
                            u3 = q5 + dy3;
                            y3 = b - u3;
                            q6 = dy1 + dy3 - u3;
                            q7 = q6 + dy1;
                            cu3 = q7 + dy3;
                            cy3 = b - cu3;
                            u2 = (q8 + rh) / 2;
                            y2 = b - u2;
                            u5 = q5 + rh;
                            y5 = b - u5;
                            u6 = u3 + rh;
                            y6 = b - u6;
                            cu4 = q9 + rh;
                            cy4 = b - cu4;
                            cu6 = cu3 + rh;
                            cy6 = b - cu6;
                            u7 = u1 + dy3;
                            y7 = b - u7;
                            //
                            d_val = "M" + l + "," + b +
                                " L" + wd8 + "," + y2 +
                                " L" + l + "," + q1 +
                                " Q" + cx4 + "," + cy4 + " " + x2 + "," + y5 +
                                " L" + x2 + "," + y6 +
                                " Q" + hc + "," + cy6 + " " + x5 + "," + y6 +
                                " L" + x5 + "," + y5 +
                                " Q" + cx5 + "," + cy4 + " " + r + "," + q1 +
                                " L" + x6 + "," + y2 +
                                " L" + r + "," + b +
                                " Q" + cx2 + "," + cy1 + " " + x4 + "," + y1 +
                                " L" + x5 + "," + y3 +
                                " Q" + hc + "," + cy3 + " " + x2 + "," + y3 +
                                " L" + x3 + "," + y1 +
                                " Q" + cx1 + "," + cy1 + " " + l + "," + b +
                                " z" +
                                "M" + x2 + "," + y3 +
                                " L" + x2 + "," + y5 +
                                "M" + x5 + "," + y5 +
                                " L" + x5 + "," + y3 +
                                "M" + x3 + "," + y7 +
                                " L" + x3 + "," + y1 +
                                "M" + x4 + "," + y1 +
                                " L" + x4 + "," + y7;
                        }
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "line":
                    case "straightConnector1":
                    case "bentConnector4":
                    case "bentConnector5": {
                        // 使用drawW和drawH（原始尺寸）而不是w和h（可能被调整的SVG容器尺寸）
                        var lineW = drawW;
                        var lineH = drawH;
                        // 如果drawW或drawH未定义（非连接器情况），回退到w和h
                        if (lineW === undefined) lineW = w;
                        if (lineH === undefined) lineH = h;
                        
                        // 根据flipH和flipV确定线条的起点和终点
                        var x1 = 0, y1 = 0, x2 = lineW, y2 = lineH;
                        
                        result += "<line x1='" + x1 + "' y1='" + y1 + "' x2='" + x2 + "' y2='" + y2 + "' stroke='" + border.color +
                            "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                        }
                        result += "/>";
                        break;
                    }
                    case "curvedConnector2":
                    case "curvedConnector3":
                    case "curvedConnector4":
                    case "curvedConnector5": {
                        // 获取调整值
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var adj1 = 50000; // 默认值
                        if (shapAdjst_ary !== undefined) {
                            if (Array.isArray(shapAdjst_ary)) {
                                for (var i = 0; i < shapAdjst_ary.length; i++) {
                                    var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                    if (sAdj_name == "adj1") {
                                        var sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        adj1 = parseInt(sAdj1.substr(4));
                                        break;
                                    }
                                }
                            } else {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    var sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4));
                                }
                            }
                        }

                        // 使用drawW和drawH（原始尺寸）
                        var curveW = (drawW !== undefined) ? drawW : w;
                        var curveH = (drawH !== undefined) ? drawH : h;

                        // 计算曲线控制点
                        var cx1, cy1, cx2, cy2;
                        var pathD;
                        
                        // 路径方向（SVG容器会通过flip变换处理翻转）
                        if (shapType === "curvedConnector2" || shapType === "curvedConnector3") {
                            // 对于 curvedConnector2 和 curvedConnector3，使用简单的二次贝塞尔曲线
                            var controlPointRatio = adj1 / 100000;
                            cx1 = curveW * controlPointRatio;
                            cy1 = 0;
                            cx2 = curveW * (1 - controlPointRatio);
                            cy2 = curveH;
                        } else {
                            // 对于其他弯曲连接器，使用默认控制点
                            cx1 = curveW / 4;
                            cy1 = 0;
                            cx2 = curveW * 3 / 4;
                            cy2 = curveH;
                        }
                        // 正常路径
                        pathD = "M 0,0 Q " + cx1 + "," + cy1 + " " + curveW/2 + "," + curveH/2 + " Q " + cx2 + "," + cy2 + " " + curveW + "," + curveH;

                        // 使用 SVG 路径元素创建曲线
                        result += "<path d='" + pathD + "' stroke='" + border.color +
                            "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' fill='none' ";
                        
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                        }
                        result += "/>";
                        break;
                    }
                    case "rightArrow":
                    case "leftArrow":
                    case "downArrow":
                    case "upArrow":
                    case "leftRightArrow":
                    case "upDownArrow": {
                        // 使用drawW和drawH（原始尺寸）而不是w和h（缩放后尺寸）
                        // SVG会通过transform scale()进行缩放
                        result += renderArrow(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
                        break;
                    }
                    case "quadArrow": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 22500 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 22500 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 22500 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, q1, x1, x2, dx2, x3, dx3, x4, x5, x6, y2, y3, y4, y5, y6, maxAdj1, maxAdj3;
                        var minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0;
                        else if (adj2 > cnstVal1) a2 = cnstVal1;
                        else a2 = adj2;
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0;
                        else if (adj1 > maxAdj1) a1 = maxAdj1;
                        else a1 = adj1;
                        q1 = cnstVal2 - maxAdj1;
                        maxAdj3 = q1 / 2;
                        if (adj3 < 0) a3 = 0;
                        else if (adj3 > maxAdj3) a3 = maxAdj3;
                        else a3 = adj3;
                        x1 = minWH * a3 / cnstVal2;
                        dx2 = minWH * a2 / cnstVal2;
                        x2 = hc - dx2;
                        x5 = hc + dx2;
                        dx3 = minWH * a1 / cnstVal3;
                        x3 = hc - dx3;
                        x4 = hc + dx3;
                        x6 = w - x1;
                        y2 = vc - dx2;
                        y5 = vc + dx2;
                        y3 = vc - dx3;
                        y4 = vc + dx3;
                        y6 = h - x1;
                        var d_val = "M" + 0 + "," + vc +
                            " L" + x1 + "," + y2 +
                            " L" + x1 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + hc + "," + 0 +
                            " L" + x5 + "," + x1 +
                            " L" + x4 + "," + x1 +
                            " L" + x4 + "," + y3 +
                            " L" + x6 + "," + y3 +
                            " L" + x6 + "," + y2 +
                            " L" + w + "," + vc +
                            " L" + x6 + "," + y5 +
                            " L" + x6 + "," + y4 +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + y6 +
                            " L" + x5 + "," + y6 +
                            " L" + hc + "," + h +
                            " L" + x2 + "," + y6 +
                            " L" + x3 + "," + y6 +
                            " L" + x3 + "," + y4 +
                            " L" + x1 + "," + y4 +
                            " L" + x1 + "," + y5 + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "leftRightUpArrow": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, q1, x1, x2, dx2, x3, dx3, x4, x5, x6, y2, dy2, y3, y4, y5, maxAdj1, maxAdj3;
                        var minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0;
                        else if (adj2 > cnstVal1) a2 = cnstVal1;
                        else a2 = adj2;
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0;
                        else if (adj1 > maxAdj1) a1 = maxAdj1;
                        else a1 = adj1;
                        q1 = cnstVal2 - maxAdj1;
                        maxAdj3 = q1 / 2;
                        if (adj3 < 0) a3 = 0;
                        else if (adj3 > maxAdj3) a3 = maxAdj3;
                        else a3 = adj3;
                        x1 = minWH * a3 / cnstVal2;
                        dx2 = minWH * a2 / cnstVal2;
                        x2 = hc - dx2;
                        x5 = hc + dx2;
                        dx3 = minWH * a1 / cnstVal3;
                        x3 = hc - dx3;
                        x4 = hc + dx3;
                        x6 = w - x1;
                        dy2 = minWH * a2 / cnstVal1;
                        y2 = h - dy2;
                        y4 = h - dx2;
                        y3 = y4 - dx3;
                        y5 = y4 + dx3;
                        var d_val = "M" + 0 + "," + y4 +
                            " L" + x1 + "," + y2 +
                            " L" + x1 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + hc + "," + 0 +
                            " L" + x5 + "," + x1 +
                            " L" + x4 + "," + x1 +
                            " L" + x4 + "," + y3 +
                            " L" + x6 + "," + y3 +
                            " L" + x6 + "," + y2 +
                            " L" + w + "," + y4 +
                            " L" + x6 + "," + h +
                            " L" + x6 + "," + y5 +
                            " L" + x1 + "," + y5 +
                            " L" + x1 + "," + h + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "leftUpArrow": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, x1, x2, dx4, dx3, x3, x4, x5, y2, y3, y4, y5, maxAdj1, maxAdj3;
                        var minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0;
                        else if (adj2 > cnstVal1) a2 = cnstVal1;
                        else a2 = adj2;
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0;
                        else if (adj1 > maxAdj1) a1 = maxAdj1;
                        else a1 = adj1;
                        maxAdj3 = cnstVal2 - maxAdj1;
                        if (adj3 < 0) a3 = 0;
                        else if (adj3 > maxAdj3) a3 = maxAdj3;
                        else a3 = adj3;
                        x1 = minWH * a3 / cnstVal2;
                        dx2 = minWH * a2 / cnstVal1;
                        x2 = w - dx2;
                        y2 = h - dx2;
                        dx4 = minWH * a2 / cnstVal2;
                        x4 = w - dx4;
                        y4 = h - dx4;
                        dx3 = minWH * a1 / cnstVal3;
                        x3 = x4 - dx3;
                        x5 = x4 + dx3;
                        y3 = y4 - dx3;
                        y5 = y4 + dx3;
                        var d_val = "M" + 0 + "," + y4 +
                            " L" + x1 + "," + y2 +
                            " L" + x1 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + x4 + "," + 0 +
                            " L" + w + "," + x1 +
                            " L" + x5 + "," + x1 +
                            " L" + x5 + "," + y5 +
                            " L" + x1 + "," + y5 +
                            " L" + x1 + "," + h + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "bentUpArrow": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, dx1, x1, dx2, x2, dx3, x3, x4, y1, y2, dy2;
                        var minWH = Math.min(w, h);
                        if (adj1 < 0) a1 = 0;
                        else if (adj1 > cnstVal1) a1 = cnstVal1;
                        else a1 = adj1;
                        if (adj2 < 0) a2 = 0;
                        else if (adj2 > cnstVal1) a2 = cnstVal1;
                        else a2 = adj2;
                        if (adj3 < 0) a3 = 0;
                        else if (adj3 > maxAdj3) a3 = maxAdj3;
                        else a3 = adj3;
                        y1 = minWH * a3 / cnstVal2;
                        dx1 = minWH * a2 / cnstVal1;
                        x1 = w - dx1;
                        dx3 = minWH * a2 / cnstVal2;
                        x3 = w - dx3;
                        dx2 = minWH * a1 / cnstVal3;
                        x2 = x3 - dx2;
                        x4 = x3 + dx2;
                        dy2 = minWH * a1 / cnstVal2;
                        y2 = h - dy2;
                        var d_val = "M" + 0 + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + y1 +
                            " L" + x1 + "," + y1 +
                            " L" + x3 + "," + 0 +
                            " L" + w + "," + y1 +
                            " L" + x4 + "," + y1 +
                            " L" + x4 + "," + h +
                            " L" + 0 + "," + h + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "bentArrow": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var sAdj4, adj4 = 43750 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var a1, a2, a3, a4, x3, x4, y3, y4, y5, y6, maxAdj1, maxAdj4;
                        var minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0;
                        else if (adj2 > cnstVal1) a2 = cnstVal1;
                        else a2 = adj2;
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0;
                        else if (adj1 > maxAdj1) a1 = maxAdj1;
                        else a1 = adj1;
                        if (adj3 < 0) a3 = 0;
                        else if (adj3 > cnstVal1) a3 = cnstVal1;
                        else a3 = adj3;
                        var th, aw2, th2, dh2, ah, bw, bh, bs, bd, bd3, bd2,
                            th = minWH * a1 / cnstVal2;
                        aw2 = minWH * a2 / cnstVal2;
                        th2 = th / 2;
                        dh2 = aw2 - th2;
                        ah = minWH * a3 / cnstVal2;
                        bw = w - ah;
                        bh = h - dh2;
                        bs = (bw < bh) ? bw : bh;
                        maxAdj4 = cnstVal2 * bs / minWH;
                        if (adj4 < 0) a4 = 0;
                        else if (adj4 > maxAdj4) a4 = maxAdj4;
                        else a4 = adj4;
                        bd = minWH * a4 / cnstVal2;
                        bd3 = bd - th;
                        bd2 = (bd3 > 0) ? bd3 : 0;
                        x3 = th + bd2;
                        x4 = w - ah;
                        y3 = dh2 + th;
                        y4 = y3 + dh2;
                        y5 = dh2 + bd;
                        y6 = y3 + bd2;

                        var d_val = "M" + 0 + "," + h +
                            " L" + 0 + "," + y5 +
                            PPTXShapeUtils.shapeArc(bd, y5, bd, bd, 180, 270, false).replace("M", "L") +
                            " L" + x4 + "," + dh2 +
                            " L" + x4 + "," + 0 +
                            " L" + w + "," + aw2 +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            PPTXShapeUtils.shapeArc(x3, y6, bd2, bd2, 270, 180, false).replace("M", "L") +
                            " L" + th + "," + h + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "uturnArrow": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var sAdj4, adj4 = 43750 * SLIDE_FACTOR$1;
                        var sAdj5, adj5 = 75000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var a1, a2, a3, a4, a5, q1, q2, q3, x3, x4, x5, x6, x7, x8, x9, y4, y5, minAdj5, maxAdj1, maxAdj3, maxAdj4;
                        var minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0;
                        else if (adj2 > cnstVal1) a2 = cnstVal1;
                        else a2 = adj2;
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0;
                        else if (adj1 > maxAdj1) a1 = maxAdj1;
                        else a1 = adj1;
                        q2 = a1 * minWH / h;
                        q3 = cnstVal2 - q2;
                        maxAdj3 = q3 * h / minWH;
                        if (adj3 < 0) a3 = 0;
                        else if (adj3 > maxAdj3) a3 = maxAdj3;
                        else a3 = adj3;
                        q1 = a3 + a1;
                        minAdj5 = q1 * minWH / h;
                        if (adj5 < minAdj5) a5 = minAdj5;
                        else if (adj5 > cnstVal2) a5 = cnstVal2;
                        else a5 = adj5;

                        var th, aw2, th2, dh2, ah, bw, bs, bd, bd3, bd2,
                            th = minWH * a1 / cnstVal2;
                        aw2 = minWH * a2 / cnstVal2;
                        th2 = th / 2;
                        dh2 = aw2 - th2;
                        y5 = h * a5 / cnstVal2;
                        ah = minWH * a3 / cnstVal2;
                        y4 = y5 - ah;
                        x9 = w - dh2;
                        bw = x9 / 2;
                        bs = (bw < y4) ? bw : y4;
                        maxAdj4 = cnstVal2 * bs / minWH;
                        if (adj4 < 0) a4 = 0;
                        else if (adj4 > maxAdj4) a4 = maxAdj4;
                        else a4 = adj4;
                        bd = minWH * a4 / cnstVal2;
                        bd3 = bd - th;
                        bd2 = (bd3 > 0) ? bd3 : 0;
                        x3 = th + bd2;
                        x8 = w - aw2;
                        x6 = x8 - aw2;
                        x7 = x6 + dh2;
                        x4 = x9 - bd;
                        x5 = x7 - bd2;
                        var d_val = "M" + 0 + "," + h +
                            " L" + 0 + "," + bd +
                            shapeArcAlt(bd, bd, bd, bd, 180, 270, false).replace("M", "L") +
                            " L" + x4 + "," + 0 +
                            shapeArcAlt(x4, bd, bd, bd, 270, 360, false).replace("M", "L") +
                            " L" + x9 + "," + y4 +
                            " L" + w + "," + y4 +
                            " L" + x8 + "," + y5 +
                            " L" + x6 + "," + y4 +
                            " L" + x7 + "," + y4 +
                            " L" + x7 + "," + x3 +
                            shapeArcAlt(x5, x3, bd2, bd2, 0, -90, false).replace("M", "L") +
                            " L" + x3 + "," + th +
                            shapeArcAlt(x3, x3, bd2, bd2, 270, 180, false).replace("M", "L") +
                            " L" + th + "," + h + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "stripedRightArrow": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 50000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 200000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 84375 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var a1, a2, x4, x5, dx5, x6, y1, dy1, y2, maxAdj2, vc = h / 2;
                        var minWH = Math.min(w, h);
                        maxAdj2 = cnstVal3 * w / minWH;
                        if (adj1 < 0) a1 = 0;
                        else if (adj1 > cnstVal1) a1 = cnstVal1;
                        else a1 = adj1;
                        if (adj2 < 0) a2 = 0;
                        else if (adj2 > maxAdj2) a2 = maxAdj2;
                        else a2 = adj2;
                        x4 = minWH * 5 / 32;
                        dx5 = minWH * a2 / cnstVal1;
                        x5 = w - dx5;
                        dy1 = h * a1 / cnstVal2;
                        y1 = vc - dy1;
                        y2 = vc + dy1;
                        //dx6 = dy1*dx5/hd2;
                        //x6 = w-dx6;
                        var ssd8 = minWH / 8,
                            ssd16 = minWH / 16,
                            ssd32 = minWH / 32;
                        var d_val = "M" + 0 + "," + y1 +
                            " L" + ssd32 + "," + y1 +
                            " L" + ssd32 + "," + y2 +
                            " L" + 0 + "," + y2 + " z" +
                            " M" + ssd16 + "," + y1 +
                            " L" + ssd8 + "," + y1 +
                            " L" + ssd8 + "," + y2 +
                            " L" + ssd16 + "," + y2 + " z" +
                            " M" + x4 + "," + y1 +
                            " L" + x5 + "," + y1 +
                            " L" + x5 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x5 + "," + h +
                            " L" + x5 + "," + y2 +
                            " L" + x4 + "," + y2 + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "notchedRightArrow": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 50000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var a1, a2, x1, x2, dx2, y1, dy1, y2, maxAdj2, vc = h / 2, hd2 = vc;
                        var minWH = Math.min(w, h);
                        maxAdj2 = cnstVal1 * w / minWH;
                        if (adj1 < 0) a1 = 0;
                        else if (adj1 > cnstVal1) a1 = cnstVal1;
                        else a1 = adj1;
                        if (adj2 < 0) a2 = 0;
                        else if (adj2 > maxAdj2) a2 = maxAdj2;
                        else a2 = adj2;
                        dx2 = minWH * a2 / cnstVal1;
                        x2 = w - dx2;
                        dy1 = h * a1 / cnstVal2;
                        y1 = vc - dy1;
                        y2 = vc + dy1;
                        x1 = dy1 * dx2 / hd2;
                        var d_val = "M" + 0 + "," + y1 +
                            " L" + x2 + "," + y1 +
                            " L" + x2 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x2 + "," + h +
                            " L" + x2 + "," + y2 +
                            " L" + 0 + "," + y2 +
                            " L" + x1 + "," + vc + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "homePlate": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 50000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                        }
                        var a, x1, dx1, maxAdj, vc = h / 2;
                        var minWH = Math.min(w, h);
                        maxAdj = cnstVal1 * w / minWH;
                        if (adj < 0) a = 0;
                        else if (adj > maxAdj) a = maxAdj;
                        else a = adj;
                        dx1 = minWH * a / cnstVal1;
                        x1 = w - dx1;
                        var d_val = "M" + 0 + "," + 0 +
                            " L" + x1 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x1 + "," + h +
                            " L" + 0 + "," + h + " z";

                        result += "<path  d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "chevron": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 50000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                        }
                        var a, x1, dx1, x2, maxAdj, vc = h / 2;
                        var minWH = Math.min(w, h);
                        maxAdj = cnstVal1 * w / minWH;
                        if (adj < 0) a = 0;
                        else if (adj > maxAdj) a = maxAdj;
                        else a = adj;
                        x1 = minWH * a / cnstVal1;
                        x2 = w - x1;
                        var d_val = "M" + 0 + "," + 0 +
                            " L" + x2 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x2 + "," + h +
                            " L" + 0 + "," + h +
                            " L" + x1 + "," + vc + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";


                        break;
                    }
                    case "rightArrowCallout": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var sAdj4, adj4 = 64977 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, dx3, x3, x2, x1;
                        var vc = h / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);
                        maxAdj2 = cnstVal1 * h / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * w / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / w;
                        maxAdj4 = cnstVal - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        dy2 = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        dx3 = ss * a3 / cnstVal2;
                        x3 = r - dx3;
                        x2 = w * a4 / cnstVal2;
                        x1 = x2 / 2;
                        var d_val = "M" + l + "," + t +
                            " L" + x2 + "," + t +
                            " L" + x2 + "," + y2 +
                            " L" + x3 + "," + y2 +
                            " L" + x3 + "," + y1 +
                            " L" + r + "," + vc +
                            " L" + x3 + "," + y4 +
                            " L" + x3 + "," + y3 +
                            " L" + x2 + "," + y3 +
                            " L" + x2 + "," + b +
                            " L" + l + "," + b +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "downArrowCallout": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var sAdj4, adj4 = 64977 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx1, dx2, x1, x2, x3, x4, dy3, y3, y2, y1;
                        var hc = w / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);

                        maxAdj2 = cnstVal1 * w / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * h / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / h;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dx1 = ss * a2 / cnstVal2;
                        dx2 = ss * a1 / cnstVal3;
                        x1 = hc - dx1;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        x4 = hc + dx1;
                        dy3 = ss * a3 / cnstVal2;
                        y3 = b - dy3;
                        y2 = h * a4 / cnstVal2;
                        y1 = y2 / 2;
                        var d_val = "M" + l + "," + t +
                            " L" + r + "," + t +
                            " L" + r + "," + y2 +
                            " L" + x3 + "," + y2 +
                            " L" + x3 + "," + y3 +
                            " L" + x4 + "," + y3 +
                            " L" + hc + "," + b +
                            " L" + x1 + "," + y3 +
                            " L" + x2 + "," + y3 +
                            " L" + x2 + "," + y2 +
                            " L" + l + "," + y2 +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "leftArrowCallout": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var sAdj4, adj4 = 64977 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, x1, dx2, x2, x3;
                        var vc = h / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);

                        maxAdj2 = cnstVal1 * h / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * w / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / w;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        dy2 = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        x1 = ss * a3 / cnstVal2;
                        dx2 = w * a4 / cnstVal2;
                        x2 = r - dx2;
                        x3 = (x2 + r) / 2;
                        var d_val = "M" + l + "," + vc +
                            " L" + x1 + "," + y1 +
                            " L" + x1 + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + t +
                            " L" + r + "," + t +
                            " L" + r + "," + b +
                            " L" + x2 + "," + b +
                            " L" + x2 + "," + y3 +
                            " L" + x1 + "," + y3 +
                            " L" + x1 + "," + y4 +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "upArrowCallout": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var sAdj4, adj4 = 64977 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx1, dx2, x1, x2, x3, x4, y1, dy2, y2, y3;
                        var hc = w / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);
                        maxAdj2 = cnstVal1 * w / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * h / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / h;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dx1 = ss * a2 / cnstVal2;
                        dx2 = ss * a1 / cnstVal3;
                        x1 = hc - dx1;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        x4 = hc + dx1;
                        y1 = ss * a3 / cnstVal2;
                        dy2 = h * a4 / cnstVal2;
                        y2 = b - dy2;
                        y3 = (y2 + b) / 2;

                        var d_val = "M" + l + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + y1 +
                            " L" + x1 + "," + y1 +
                            " L" + hc + "," + t +
                            " L" + x4 + "," + y1 +
                            " L" + x3 + "," + y1 +
                            " L" + x3 + "," + y2 +
                            " L" + r + "," + y2 +
                            " L" + r + "," + b +
                            " L" + l + "," + b +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "leftRightArrowCallout": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var sAdj4, adj4 = 48123 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, x1, x4, dx2, x2, x3;
                        var vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);
                        maxAdj2 = cnstVal1 * h / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal1 * w / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / wd2;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        dy2 = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        x1 = ss * a3 / cnstVal2;
                        x4 = r - x1;
                        dx2 = w * a4 / cnstVal3;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        var d_val = "M" + l + "," + vc +
                            " L" + x1 + "," + y1 +
                            " L" + x1 + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + t +
                            " L" + x3 + "," + t +
                            " L" + x3 + "," + y2 +
                            " L" + x4 + "," + y2 +
                            " L" + x4 + "," + y1 +
                            " L" + r + "," + vc +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + b +
                            " L" + x2 + "," + b +
                            " L" + x2 + "," + y3 +
                            " L" + x1 + "," + y3 +
                            " L" + x1 + "," + y4 +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "quadArrowCallout": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 18515 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 18515 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 18515 * SLIDE_FACTOR$1;
                        var sAdj4, adj4 = 48123 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var cnstVal3 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);
                        var a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx2, dx3, ah, dx1, dy1, x8, x2, x7, x3, x6, x4, x5, y8, y2, y7, y3, y6, y4, y5;
                        a2 = (adj2 < 0) ? 0 : (adj2 > cnstVal1) ? cnstVal1 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal1 - a2;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * 2;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < a1) ? a1 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dx2 = ss * a2 / cnstVal2;
                        dx3 = ss * a1 / cnstVal3;
                        ah = ss * a3 / cnstVal2;
                        dx1 = w * a4 / cnstVal3;
                        dy1 = h * a4 / cnstVal3;
                        x8 = r - ah;
                        x2 = hc - dx1;
                        x7 = hc + dx1;
                        x3 = hc - dx2;
                        x6 = hc + dx2;
                        x4 = hc - dx3;
                        x5 = hc + dx3;
                        y8 = b - ah;
                        y2 = vc - dy1;
                        y7 = vc + dy1;
                        y3 = vc - dx2;
                        y6 = vc + dx2;
                        y4 = vc - dx3;
                        y5 = vc + dx3;
                        var d_val = "M" + l + "," + vc +
                            " L" + ah + "," + y3 +
                            " L" + ah + "," + y4 +
                            " L" + x2 + "," + y4 +
                            " L" + x2 + "," + y2 +
                            " L" + x4 + "," + y2 +
                            " L" + x4 + "," + ah +
                            " L" + x3 + "," + ah +
                            " L" + hc + "," + t +
                            " L" + x6 + "," + ah +
                            " L" + x5 + "," + ah +
                            " L" + x5 + "," + y2 +
                            " L" + x7 + "," + y2 +
                            " L" + x7 + "," + y4 +
                            " L" + x8 + "," + y4 +
                            " L" + x8 + "," + y3 +
                            " L" + r + "," + vc +
                            " L" + x8 + "," + y6 +
                            " L" + x8 + "," + y5 +
                            " L" + x7 + "," + y5 +
                            " L" + x7 + "," + y7 +
                            " L" + x5 + "," + y7 +
                            " L" + x5 + "," + y8 +
                            " L" + x6 + "," + y8 +
                            " L" + hc + "," + b +
                            " L" + x3 + "," + y8 +
                            " L" + x4 + "," + y8 +
                            " L" + x4 + "," + y7 +
                            " L" + x2 + "," + y7 +
                            " L" + x2 + "," + y5 +
                            " L" + ah + "," + y5 +
                            " L" + ah + "," + y6 +
                            " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "curvedDownArrow": {
                        // 下弧形箭头使用drawW和drawH（原始尺寸）进行形状计算
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        // 使用drawW和drawH进行形状计算
                        var cw = (drawW !== undefined) ? drawW : w;
                        var ch = (drawH !== undefined) ? drawH : h;
                        var vc = ch / 2, hc = cw / 2, wd2 = cw / 2, r = cw, b = ch, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        var ss = Math.min(cw, ch);
                        var maxAdj2, a2, a1, th, aw, q1, wR, q7, q8, q9, q10, q11, idy, maxAdj3, a3, ah, x3, q2, q3, q4, q5, dx, x5, x7, q6, dh, x4, x8, aw2, x6, y1, swAng, mswAng, q12, dang2, stAng, stAng2, swAng2, swAng3;

                        // 辅助函数：格式化数字为2位小数
                        function fmt(num) {
                            return parseFloat(num.toFixed(2));
                        }

                        maxAdj2 = cnstVal1 * cw / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        wR = wd2 - q1;
                        q7 = wR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        idy = q11 * ch / q7;
                        maxAdj3 = cnstVal2 * idy / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * adj3 / cnstVal2;
                        x3 = wR + th;
                        q2 = ch * ch;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dx = q5 * wR / ch;
                        x5 = wR + dx;
                        x7 = x3 + dx;
                        q6 = aw - th;
                        dh = q6 / 2;
                        x4 = x5 - dh;
                        x8 = x7 + dh;
                        aw2 = aw / 2;
                        x6 = r - aw2;
                        y1 = b - ah;
                        swAng = Math.atan(dx / ah);
                        var swAngDeg = swAng * 180 / Math.PI;
                        mswAng = -swAngDeg;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / idy);
                        var dang2Deg = dang2 * 180 / Math.PI;
                        stAng = c3d4 + swAngDeg;
                        stAng2 = c3d4 - dang2Deg;
                        swAng2 = dang2Deg - cd4;
                        swAng3 = cd4 + dang2Deg;
                        //var cX = x5 - Math.cos(stAng*Math.PI/180) * wR;
                        //var cY = y1 - Math.sin(stAng*Math.PI/180) * h;

                        // 格式化所有坐标值
                        x6 = fmt(x6);
                        b = fmt(b);
                        x4 = fmt(x4);
                        y1 = fmt(y1);
                        x5 = fmt(x5);
                        x3 = fmt(x3);
                        t = fmt(t);
                        th = fmt(th);
                        x8 = fmt(x8);
                        wR = fmt(wR);
                        ch = fmt(ch);

                        var d_val = "M" + x6 + "," + b +
                            " L" + x4 + "," + y1 +
                            " L" + x5 + "," + y1 +
                            PPTXShapeUtils.shapeArc(wR, ch, wR, ch, stAng, (stAng + mswAng), false).replace("M", "L") +
                            " L" + x3 + "," + t +
                            PPTXShapeUtils.shapeArc(x3, ch, wR, ch, c3d4, (c3d4 + swAngDeg), false).replace("M", "L") +
                            " L" + fmt(x5 + th) + "," + y1 +
                            " L" + x8 + "," + y1 +
                            " z" +
                            "M" + x3 + "," + t +
                            PPTXShapeUtils.shapeArc(x3, ch, wR, ch, stAng2, (stAng2 + swAng2), false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(wR, ch, wR, ch, cd2, (cd2 + swAng3), false).replace("M", "L");

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "curvedLeftArrow": {
                        // 左弧形箭头使用drawW和drawH（原始尺寸）进行形状计算
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        // 使用drawW和drawH进行形状计算
                        var cw = (drawW !== undefined) ? drawW : w;
                        var ch = (drawH !== undefined) ? drawH : h;
                        var vc = ch / 2, hc = cw / 2, hd2 = ch / 2, r = cw, b = ch, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        var ss = Math.min(cw, ch);
                        var maxAdj2, a2, a1, th, aw, q1, hR, q7, q8, q9, q10, q11, iDx, maxAdj3, a3, ah, y3, q2, q3, q4, q5, dy, y5, y7, q6, dh, y4, y8, aw2, y6, x1, swAng, mswAng, q12, dang2, swAng2, swAng3, stAng3;

                        // 辅助函数：格式化数字为2位小数
                        function fmt(num) {
                            return parseFloat(num.toFixed(2));
                        }

                        maxAdj2 = cnstVal1 * ch / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > a2) ? a2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        hR = hd2 - q1;
                        q7 = hR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        iDx = q11 * cw / q7;
                        maxAdj3 = cnstVal2 * iDx / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * a3 / cnstVal2;
                        y3 = hR + th;
                        q2 = cw * cw;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dy = q5 * hR / cw;
                        y5 = hR + dy;
                        y7 = y3 + dy;
                        q6 = aw - th;
                        dh = q6 / 2;
                        y4 = y5 - dh;
                        y8 = y7 + dh;
                        aw2 = aw / 2;
                        y6 = b - aw2;
                        x1 = l + ah;
                        swAng = Math.atan(dy / ah);
                        mswAng = -swAng;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / iDx);
                        swAng2 = dang2 - swAng;
                        swAng3 = swAng + dang2;
                        stAng3 = -dang2;
                        var swAngDg, swAng2Dg, stAng3dg;
                        swAngDg = swAng * 180 / Math.PI;
                        swAng2Dg = swAng2 * 180 / Math.PI;
                        stAng3dg = stAng3 * 180 / Math.PI;

                        // 格式化所有坐标值
                        r = fmt(r);
                        y3 = fmt(y3);
                        l = fmt(l);
                        hR = fmt(hR);
                        cw = fmt(cw);
                        t = fmt(t);
                        x1 = fmt(x1);
                        y7 = fmt(y7);
                        y8 = fmt(y8);
                        y6 = fmt(y6);
                        y4 = fmt(y4);
                        y5 = fmt(y5);

                        var d_val = "M" + r + "," + y3 +
                            PPTXShapeUtils.shapeArc(l, hR, cw, hR, 0, -cd4, false).replace("M", "L") +
                            " L" + l + "," + t +
                            PPTXShapeUtils.shapeArc(l, y3, cw, hR, c3d4, (c3d4 + cd4), false).replace("M", "L") +
                            " L" + r + "," + y3 +
                            PPTXShapeUtils.shapeArc(l, y3, cw, hR, 0, swAngDg, false).replace("M", "L") +
                            " L" + x1 + "," + y7 +
                            " L" + x1 + "," + y8 +
                            " L" + l + "," + y6 +
                            " L" + x1 + "," + y4 +
                            " L" + x1 + "," + y5 +
                            PPTXShapeUtils.shapeArc(l, hR, cw, hR, swAngDg, (swAngDg + swAng2Dg), false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(l, hR, cw, hR, 0, -cd4, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(l, y3, cw, hR, c3d4, (c3d4 + cd4), false).replace("M", "L");

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "curvedRightArrow": {
                        /**
                         * curvedRightArrow: 手杖形箭头（弯曲向右的箭头）
                         *
                         * 形状说明：
                         * - 从左侧开始，向上弯曲，最后指向右侧
                         * - 有箭头头部
                         *
                         * 参数说明：
                         * - adj1: 控制箭头尖端的高度
                         * - adj2: 控制箭头宽度
                         * - adj3: 控制弯曲程度（箭头宽度）
                         */
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        // 使用drawW和drawH进行形状计算
                        var cw = (drawW !== undefined) ? drawW : w;
                        var ch = (drawH !== undefined) ? drawH : h;
                        var vc = ch / 2, hc = cw / 2, hd2 = ch / 2, r = cw, b = ch, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        var ss = Math.min(cw, ch);
                        var maxAdj2, a2, a1, th, aw, q1, hR, q7, q8, q9, q10, q11, iDx, maxAdj3, a3, ah, y3, q2, q3, q4, q5, dy,
                            y5, y7, q6, dh, y4, y8, aw2, y6, x1, swAng, stAng, mswAng, q12, dang2, swAng2, swAng3, stAng3;

                        maxAdj2 = cnstVal1 * ch / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > a2) ? a2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        hR = hd2 - q1;
                        q7 = hR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        iDx = q11 * cw / q7;
                        maxAdj3 = cnstVal2 * iDx / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * a3 / cnstVal2;
                        y3 = hR + th;
                        q2 = cw * cw;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dy = q5 * hR / cw;
                        y5 = hR + dy;
                        y7 = y3 + dy;
                        q6 = aw - th;
                        dh = q6 / 2;
                        y4 = y5 - dh;
                        y8 = y7 + dh;
                        aw2 = aw / 2;
                        y6 = b - aw2;
                        x1 = r - ah;
                        swAng = Math.atan(dy / ah);
                        stAng = Math.PI + 0 - swAng;
                        mswAng = -swAng;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / iDx);
                        swAng2 = dang2 - Math.PI / 2;
                        swAng3 = Math.PI / 2 + dang2;
                        stAng3 = Math.PI - dang2;

                        var stAngDg, mswAngDg, swAngDg, swAng2dg;
                        stAngDg = stAng * 180 / Math.PI;
                        mswAngDg = mswAng * 180 / Math.PI;
                        swAngDg = swAng * 180 / Math.PI;
                        swAng2dg = swAng2 * 180 / Math.PI;

                        /**
                         * 路径绘制顺序（参考 pptxjs.js）：
                         * 1. 从左侧 (l, hR) 开始，画第一个大圆弧
                         * 2. 画箭头下翼（多条线段）
                         * 3. 连接到箭头上翼
                         * 4. 画第二个大圆弧（沿上边缘）
                         * 5. 画第三个圆弧（箭头部分）
                         * 6. 闭合
                         */
                        var d_val = "M" + l + "," + hR +
                            shapeArcAlt(cw, hR, cw, hR, cd2, cd2 + mswAngDg, false).replace("M", "L") +
                            " L" + x1 + "," + y5 +
                            " L" + x1 + "," + y4 +
                            " L" + r + "," + y6 +
                            " L" + x1 + "," + y8 +
                            " L" + x1 + "," + y7 +
                            shapeArcAlt(cw, y3, cw, hR, stAngDg, stAngDg + swAngDg, false).replace("M", "L") +
                            " L" + l + "," + hR +
                            shapeArcAlt(cw, hR, cw, hR, cd2, cd2 + cd4, false).replace("M", "L") +
                            " L" + r + "," + th +
                            shapeArcAlt(cw, y3, cw, hR, c3d4, c3d4 + swAng2dg, false).replace("M", "L") +
                            " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "curvedUpArrow": {
                        // 上弧形箭头使用drawW和drawH（原始尺寸）进行形状计算
                        // 这样在group-abs类型组合中不会被缩放影响
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                        var sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        // 使用drawW和drawH进行形状计算
                        var cw = (drawW !== undefined) ? drawW : w;
                        var ch = (drawH !== undefined) ? drawH : h;
                        var vc = ch / 2, hc = cw / 2, wd2 = cw / 2, r = cw, b = ch, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        var ss = Math.min(cw, ch);
                        var maxAdj2, a2, a1, th, aw, q1, wR, q7, q8, q9, q10, q11, idy, maxAdj3, a3, ah, x3, q2, q3, q4, q5, dx, x5, x7, q6, dh, x4, x8, aw2, x6, y1, swAng, mswAng, q12, dang2, swAng2, stAng3, swAng3, stAng2;

                        // 辅助函数：格式化数字为2位小数
                        function fmt(num) {
                            return parseFloat(num.toFixed(2));
                        }

                        // 辅助函数：格式化弧线路径中的所有坐标
                        function fmtArc(arcStr) {
                            return arcStr.replace(/[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?/g, function(match) {
                                return fmt(parseFloat(match)).toString();
                            });
                        }

                        maxAdj2 = cnstVal1 * cw / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        wR = wd2 - q1;
                        q7 = wR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        idy = q11 * ch / q7;
                        maxAdj3 = cnstVal2 * idy / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * adj3 / cnstVal2;
                        x3 = wR + th;
                        q2 = ch * ch;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dx = q5 * wR / ch;
                        x5 = wR + dx;
                        x7 = x3 + dx;
                        q6 = aw - th;
                        dh = q6 / 2;
                        x4 = x5 - dh;
                        x8 = x7 + dh;
                        aw2 = aw / 2;
                        x6 = r - aw2;
                        y1 = t + ah;
                        swAng = Math.atan(dx / ah);
                        mswAng = -swAng;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / idy);
                        swAng2 = dang2 - swAng;
                        stAng3 = Math.PI / 2 - swAng;
                        swAng3 = swAng + dang2;
                        stAng2 = Math.PI / 2 - dang2;

                        var stAng2dg, swAng2dg, swAngDg, swAng2dg;
                        stAng2dg = stAng2 * 180 / Math.PI;
                        swAng2dg = swAng2 * 180 / Math.PI;
                        stAng3dg = stAng3 * 180 / Math.PI;
                        swAngDg = swAng * 180 / Math.PI;

                        // 格式化所有坐标值
                        wR = fmt(wR);
                        ch = fmt(ch);
                        cw = fmt(cw);
                        x3 = fmt(x3);
                        x5 = fmt(x5);
                        x7 = fmt(x7);
                        x4 = fmt(x4);
                        x8 = fmt(x8);
                        x6 = fmt(x6);
                        y1 = fmt(y1);
                        b = fmt(b);
                        th = fmt(th);
                        t = fmt(t);

                        var d_val = //"M" + ix + "," +iy +
                            fmtArc(PPTXShapeUtils.shapeArc(wR, 0, wR, ch, stAng2dg, stAng2dg + swAng2dg, false)) + //.replace("M","L") +
                            " L" + x5 + "," + y1 +
                            " L" + x4 + "," + y1 +
                            " L" + x6 + "," + t +
                            " L" + x8 + "," + y1 +
                            " L" + x7 + "," + y1 +
                            fmtArc(PPTXShapeUtils.shapeArc(x3, 0, wR, ch, stAng3dg, stAng3dg + swAngDg, false)).replace("M", "L") +
                            " L" + wR + "," + b +
                            fmtArc(PPTXShapeUtils.shapeArc(wR, 0, wR, ch, cd4, cd2, false)).replace("M", "L") +
                            " L" + th + "," + t +
                            fmtArc(PPTXShapeUtils.shapeArc(x3, 0, wR, ch, cd2, cd4, false)).replace("M", "L") +
                            "";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "mathDivide":
                    case "mathEqual":
                    case "mathMinus":
                    case "mathMultiply":
                    case "mathNotEqual":
                    case "mathPlus": {
                        // 使用drawW和drawH（原始尺寸）进行形状计算
                        result += renderMathSymbol(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
                        break;
                    }
                    case "cylinder":
                    case "can":
                    case "flowChartMagneticDisk":
                    case "flowChartMagneticDrum": {
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 25000 * SLIDE_FACTOR$1;
                        var cnstVal1 = 50000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 200000 * SLIDE_FACTOR$1;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                        }
                        var ss = Math.min(w, h);
                        var maxAdj, a, y1, y2, y3, dVal;
                        if (shapType == "flowChartMagneticDisk" || shapType == "flowChartMagneticDrum") {
                            adj = 50000 * SLIDE_FACTOR$1;
                        }
                        maxAdj = cnstVal1 * h / ss;
                        a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
                        y1 = ss * a / cnstVal2;
                        y2 = y1 + y1;
                        y3 = h - y1;
                        var cd2 = 180, wd2 = w / 2;

                        var tranglRott = "";
                        if (shapType == "flowChartMagneticDrum") {
                            tranglRott = "transform='rotate(90 " + w / 2 + "," + h / 2 + ")'";
                        }

                        // 使用 shapeArcAlt，参数是半径而非直径（参考 pptxjs.js）
                        dVal = shapeArcAlt(wd2, y1, wd2, y1, 0, cd2, false) +
                            shapeArcAlt(wd2, y1, wd2, y1, cd2, cd2 + cd2, false).replace("M", "L") +
                            " L" + w + "," + y3 +
                            shapeArcAlt(wd2, y3, wd2, y1, 0, cd2, false).replace("M", "L") +
                            " L" + 0 + "," + y1;

                        result += "<path " + tranglRott + " d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "swooshArrow": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var refr = SLIDE_FACTOR$1;
                        var sAdj1, adj1 = 25000 * refr;
                        var sAdj2, adj2 = 16667 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                }
                            }
                        }
                        var cnstVal1 = 1 * refr;
                        var cnstVal2 = 70000 * refr;
                        var cnstVal3 = 75000 * refr;
                        var cnstVal4 = 100000 * refr;
                        var ss = Math.min(w, h);
                        var ssd8 = ss / 8;
                        var hd6 = h / 6;

                        var a1, maxAdj2, a2, ad1, ad2, xB, yB, alfa, dx0, xC, dx1, yF, xF, xE, yE, dy2, dy22, dy3, yD, dy4, yP1, xP1, dy5, yP2, xP2;

                        a1 = (adj1 < cnstVal1) ? cnstVal1 : (adj1 > cnstVal3) ? cnstVal3 : adj1;
                        maxAdj2 = cnstVal2 * w / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        ad1 = h * a1 / cnstVal4;
                        ad2 = ss * a2 / cnstVal4;
                        xB = w - ad2;
                        yB = ssd8;
                        alfa = (Math.PI / 2) / 14;
                        dx0 = ssd8 * Math.tan(alfa);
                        xC = xB - dx0;
                        dx1 = ad1 * Math.tan(alfa);
                        yF = yB + ad1;
                        xF = xB + dx1;
                        xE = xF + dx0;
                        yE = yF + ssd8;
                        dy2 = yE - 0;
                        dy22 = dy2 / 2;
                        dy3 = h / 20;
                        yD = dy22 - dy3;
                        dy4 = hd6;
                        yP1 = hd6 + dy4;
                        xP1 = w / 6;
                        dy5 = hd6 / 2;
                        yP2 = yF + dy5;
                        xP2 = w / 4;

                        var dVal = "M" + 0 + "," + h +
                            " Q" + xP1 + "," + yP1 + " " + xB + "," + yB +
                            " L" + xC + "," + 0 +
                            " L" + w + "," + yD +
                            " L" + xE + "," + yE +
                            " L" + xF + "," + yF +
                            " Q" + xP2 + "," + yP2 + " " + 0 + "," + h +
                            " z";

                        result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "circularArrow": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 12500 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = (1142319 / 60000) * Math.PI / 180;
                        var sAdj3, adj3 = (20457681 / 60000) * Math.PI / 180;
                        var sAdj4, adj4 = (10800000 / 60000) * Math.PI / 180;
                        var sAdj5, adj5 = 12500 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = (parseInt(sAdj3.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = (parseInt(sAdj4.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0, wd2 = w / 2, hd2 = h / 2;
                        var ss = Math.min(w, h);
                        var a5, maxAdj1, a1, enAng, stAng, th, thh, th2, rw1, rh1, rw2, rh2, rw3, rh3, wtH, htH, dxH,
                            dyH, xH, yH, rI, u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, u11, u12, u13, u14, u15, u16, u17,
                            u18, u19, u20, u21, maxAng, aAng, ptAng, wtA, htA, dxA, dyA, xA, yA, wtE, htE, dxE, dyE, xE, yE,
                            dxG, dyG, xG, yG, dxB, dyB, xB, yB, sx1, sy1, sx2, sy2, rO, x1O, y1O, x2O, y2O, dxO, dyO, dO,
                            q1, q2, DO, q3, q4, q5, q6, q7, q8, sdelO, ndyO, sdyO, q9, q10, q11, dxF1, q12, dxF2, adyO,
                            q13, q14, dyF1, q15, dyF2, q16, q17, q18, q19, q20, q21, q22, dxF, dyF, sdxF, sdyF, xF, yF,
                            x1I, y1I, x2I, y2I, dxI, dyI, dI, v1, v2, DI, v3, v4, v5, v6, v7, v8, sdelI, v9, v10, v11,
                            dxC1, v12, dxC2, adyI, v13, v14, dyC1, v15, dyC2, v16, v17, v18, v19, v20, v21, v22, dxC, dyC,
                            sdxC, sdyC, xC, yC, ist0, ist1, istAng, isw1, isw2, iswAng, p1, p2, p3, p4, p5, xGp, yGp,
                            xBp, yBp, en0, en1, en2, sw0, sw1, swAng;
                        var cnstVal1 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var rdAngVal1 = (1 / 60000) * Math.PI / 180;
                        var rdAngVal2 = (21599999 / 60000) * Math.PI / 180;
                        var rdAngVal3 = 2 * Math.PI;

                        a5 = (adj5 < 0) ? 0 : (adj5 > cnstVal1) ? cnstVal1 : adj5;
                        maxAdj1 = a5 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        enAng = (adj3 < rdAngVal1) ? rdAngVal1 : (adj3 > rdAngVal2) ? rdAngVal2 : adj3;
                        stAng = (adj4 < 0) ? 0 : (adj4 > rdAngVal2) ? rdAngVal2 : adj4; //////////////////////////////////////////
                        th = ss * a1 / cnstVal2;
                        thh = ss * a5 / cnstVal2;
                        th2 = th / 2;
                        rw1 = wd2 + th2 - thh;
                        rh1 = hd2 + th2 - thh;
                        rw2 = rw1 - th;
                        rh2 = rh1 - th;
                        rw3 = rw2 + th2;
                        rh3 = rh2 + th2;
                        wtH = rw3 * Math.sin(enAng);
                        htH = rh3 * Math.cos(enAng);

                        //dxH = rw3*Math.cos(Math.atan(wtH/htH));
                        //dyH = rh3*Math.sin(Math.atan(wtH/htH));
                        dxH = rw3 * Math.cos(Math.atan2(wtH, htH));
                        dyH = rh3 * Math.sin(Math.atan2(wtH, htH));

                        xH = hc + dxH;
                        yH = vc + dyH;
                        rI = (rw2 < rh2) ? rw2 : rh2;
                        u1 = dxH * dxH;
                        u2 = dyH * dyH;
                        u3 = rI * rI;
                        u4 = u1 - u3;
                        u5 = u2 - u3;
                        u6 = u4 * u5 / u1;
                        u7 = u6 / u2;
                        u8 = 1 - u7;
                        u9 = Math.sqrt(u8);
                        u10 = u4 / dxH;
                        u11 = u10 / dyH;
                        u12 = (1 + u9) / u11;

                        //u13 = Math.atan(u12/1);
                        u13 = Math.atan2(u12, 1);

                        u14 = u13 + rdAngVal3;
                        u15 = (u13 > 0) ? u13 : u14;
                        u16 = u15 - enAng;
                        u17 = u16 + rdAngVal3;
                        u18 = (u16 > 0) ? u16 : u17;
                        u19 = u18 - cd2;
                        u20 = u18 - rdAngVal3;
                        u21 = (u19 > 0) ? u20 : u18;
                        maxAng = Math.abs(u21);
                        aAng = (adj2 < 0) ? 0 : (adj2 > maxAng) ? maxAng : adj2;
                        ptAng = enAng + aAng;
                        wtA = rw3 * Math.sin(ptAng);
                        htA = rh3 * Math.cos(ptAng);
                        //dxA = rw3*Math.cos(Math.atan(wtA/htA));
                        //dyA = rh3*Math.sin(Math.atan(wtA/htA));
                        dxA = rw3 * Math.cos(Math.atan2(wtA, htA));
                        dyA = rh3 * Math.sin(Math.atan2(wtA, htA));

                        xA = hc + dxA;
                        yA = vc + dyA;
                        wtE = rw1 * Math.sin(stAng);
                        htE = rh1 * Math.cos(stAng);

                        //dxE = rw1*Math.cos(Math.atan(wtE/htE));
                        //dyE = rh1*Math.sin(Math.atan(wtE/htE));
                        dxE = rw1 * Math.cos(Math.atan2(wtE, htE));
                        dyE = rh1 * Math.sin(Math.atan2(wtE, htE));

                        xE = hc + dxE;
                        yE = vc + dyE;
                        dxG = thh * Math.cos(ptAng);
                        dyG = thh * Math.sin(ptAng);
                        xG = xH + dxG;
                        yG = yH + dyG;
                        dxB = thh * Math.cos(ptAng);
                        dyB = thh * Math.sin(ptAng);
                        xB = xH - dxB;
                        yB = yH - dyB;
                        sx1 = xB - hc;
                        sy1 = yB - vc;
                        sx2 = xG - hc;
                        sy2 = yG - vc;
                        rO = (rw1 < rh1) ? rw1 : rh1;
                        x1O = sx1 * rO / rw1;
                        y1O = sy1 * rO / rh1;
                        x2O = sx2 * rO / rw1;
                        y2O = sy2 * rO / rh1;
                        dxO = x2O - x1O;
                        dyO = y2O - y1O;
                        dO = Math.sqrt(dxO * dxO + dyO * dyO);
                        q1 = x1O * y2O;
                        q2 = x2O * y1O;
                        DO = q1 - q2;
                        q3 = rO * rO;
                        q4 = dO * dO;
                        q5 = q3 * q4;
                        q6 = DO * DO;
                        q7 = q5 - q6;
                        q8 = (q7 > 0) ? q7 : 0;
                        sdelO = Math.sqrt(q8);
                        ndyO = dyO * -1;
                        sdyO = (ndyO > 0) ? -1 : 1;
                        q9 = sdyO * dxO;
                        q10 = q9 * sdelO;
                        q11 = DO * dyO;
                        dxF1 = (q11 + q10) / q4;
                        q12 = q11 - q10;
                        dxF2 = q12 / q4;
                        adyO = Math.abs(dyO);
                        q13 = adyO * sdelO;
                        q14 = DO * dxO / -1;
                        dyF1 = (q14 + q13) / q4;
                        q15 = q14 - q13;
                        dyF2 = q15 / q4;
                        q16 = x2O - dxF1;
                        q17 = x2O - dxF2;
                        q18 = y2O - dyF1;
                        q19 = y2O - dyF2;
                        q20 = Math.sqrt(q16 * q16 + q18 * q18);
                        q21 = Math.sqrt(q17 * q17 + q19 * q19);
                        q22 = q21 - q20;
                        dxF = (q22 > 0) ? dxF1 : dxF2;
                        dyF = (q22 > 0) ? dyF1 : dyF2;
                        sdxF = dxF * rw1 / rO;
                        sdyF = dyF * rh1 / rO;
                        xF = hc + sdxF;
                        yF = vc + sdyF;
                        x1I = sx1 * rI / rw2;
                        y1I = sy1 * rI / rh2;
                        x2I = sx2 * rI / rw2;
                        y2I = sy2 * rI / rh2;
                        dxI = x2I - x1I;
                        dyI = y2I - y1I;
                        dI = Math.sqrt(dxI * dxI + dyI * dyI);
                        v1 = x1I * y2I;
                        v2 = x2I * y1I;
                        DI = v1 - v2;
                        v3 = rI * rI;
                        v4 = dI * dI;
                        v5 = v3 * v4;
                        v6 = DI * DI;
                        v7 = v5 - v6;
                        v8 = (v7 > 0) ? v7 : 0;
                        sdelI = Math.sqrt(v8);
                        v9 = sdyO * dxI;
                        v10 = v9 * sdelI;
                        v11 = DI * dyI;
                        dxC1 = (v11 + v10) / v4;
                        v12 = v11 - v10;
                        dxC2 = v12 / v4;
                        adyI = Math.abs(dyI);
                        v13 = adyI * sdelI;
                        v14 = DI * dxI / -1;
                        dyC1 = (v14 + v13) / v4;
                        v15 = v14 - v13;
                        dyC2 = v15 / v4;
                        v16 = x1I - dxC1;
                        v17 = x1I - dxC2;
                        v18 = y1I - dyC1;
                        v19 = y1I - dyC2;
                        v20 = Math.sqrt(v16 * v16 + v18 * v18);
                        v21 = Math.sqrt(v17 * v17 + v19 * v19);
                        v22 = v21 - v20;
                        dxC = (v22 > 0) ? dxC1 : dxC2;
                        dyC = (v22 > 0) ? dyC1 : dyC2;
                        sdxC = dxC * rw2 / rI;
                        sdyC = dyC * rh2 / rI;
                        xC = hc + sdxC;
                        yC = vc + sdyC;

                        //ist0 = Math.atan(sdyC/sdxC);
                        ist0 = Math.atan2(sdyC, sdxC);

                        ist1 = ist0 + rdAngVal3;
                        istAng = (ist0 > 0) ? ist0 : ist1;
                        isw1 = stAng - istAng;
                        isw2 = isw1 - rdAngVal3;
                        iswAng = (isw1 > 0) ? isw2 : isw1;
                        p1 = xF - xC;
                        p2 = yF - yC;
                        p3 = Math.sqrt(p1 * p1 + p2 * p2);
                        p4 = p3 / 2;
                        p5 = p4 - thh;
                        xGp = (p5 > 0) ? xF : xG;
                        yGp = (p5 > 0) ? yF : yG;
                        xBp = (p5 > 0) ? xC : xB;
                        yBp = (p5 > 0) ? yC : yB;

                        //en0 = Math.atan(sdyF/sdxF);
                        en0 = Math.atan2(sdyF, sdxF);

                        en1 = en0 + rdAngVal3;
                        en2 = (en0 > 0) ? en0 : en1;
                        sw0 = en2 - stAng;
                        sw1 = sw0 + rdAngVal3;
                        swAng = (sw0 > 0) ? sw0 : sw1;

                        var strtAng = stAng * 180 / Math.PI;
                        var endAng = strtAng + (swAng * 180 / Math.PI);
                        var stiAng = istAng * 180 / Math.PI;
                        var swiAng = iswAng * 180 / Math.PI;
                        var ediAng = stiAng + swiAng;

                        var d_val = PPTXShapeUtils.shapeArc(w / 2, h / 2, rw1, rh1, strtAng, endAng, false) +
                            " L" + xGp + "," + yGp +
                            " L" + xA + "," + yA +
                            " L" + xBp + "," + yBp +
                            " L" + xC + "," + yC +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, rw2, rh2, stiAng, ediAng, false).replace("M", "L") +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "leftCircularArrow": {
                        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 12500 * SLIDE_FACTOR$1;
                        var sAdj2, adj2 = (-1142319 / 60000) * Math.PI / 180;
                        var sAdj3, adj3 = (1142319 / 60000) * Math.PI / 180;
                        var sAdj4, adj4 = (10800000 / 60000) * Math.PI / 180;
                        var sAdj5, adj5 = 12500 * SLIDE_FACTOR$1;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = (parseInt(sAdj3.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = (parseInt(sAdj4.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * SLIDE_FACTOR$1;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0, wd2 = w / 2, hd2 = h / 2;
                        var ss = Math.min(w, h);
                        var cnstVal1 = 25000 * SLIDE_FACTOR$1;
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var rdAngVal1 = (1 / 60000) * Math.PI / 180;
                        var rdAngVal2 = (21599999 / 60000) * Math.PI / 180;
                        var rdAngVal3 = 2 * Math.PI;
                        var a5, maxAdj1, a1, enAng, stAng, th, thh, th2, rw1, rh1, rw2, rh2, rw3, rh3, wtH, htH, dxH, dyH, xH, yH, rI,
                            u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, u11, u12, u13, u14, u15, u16, u17, u18, u19, u20, u21, u22,
                            minAng, u23, a2, aAng, ptAng, wtA, htA, dxA, dyA, xA, yA, wtE, htE, dxE, dyE, xE, yE, wtD, htD, dxD, dyD,
                            xD, yD, dxG, dyG, xG, yG, dxB, dyB, xB, yB, sx1, sy1, sx2, sy2, rO, x1O, y1O, x2O, y2O, dxO, dyO, dO,
                            q1, q2, DO, q3, q4, q5, q6, q7, q8, sdelO, ndyO, sdyO, q9, q10, q11, dxF1, q12, dxF2, adyO, q13, q14, dyF1,
                            q15, dyF2, q16, q17, q18, q19, q20, q21, q22, dxF, dyF, sdxF, sdyF, xF, yF, x1I, y1I, x2I, y2I, dxI, dyI, dI,
                            v1, v2, DI, v3, v4, v5, v6, v7, v8, sdelI, v9, v10, v11, dxC1, v12, dxC2, adyI, v13, v14, dyC1, v15, dyC2, v16,
                            v17, v18, v19, v20, v21, v22, dxC, dyC, sdxC, sdyC, xC, yC, ist0, ist1, istAng0, isw1, isw2, iswAng0, istAng,
                            iswAng, p1, p2, p3, p4, p5, xGp, yGp, xBp, yBp, en0, en1, en2, sw0, sw1, swAng, stAng0;

                        a5 = (adj5 < 0) ? 0 : (adj5 > cnstVal1) ? cnstVal1 : adj5;
                        maxAdj1 = a5 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        enAng = (adj3 < rdAngVal1) ? rdAngVal1 : (adj3 > rdAngVal2) ? rdAngVal2 : adj3;
                        stAng = (adj4 < 0) ? 0 : (adj4 > rdAngVal2) ? rdAngVal2 : adj4;
                        th = ss * a1 / cnstVal2;
                        thh = ss * a5 / cnstVal2;
                        th2 = th / 2;
                        rw1 = wd2 + th2 - thh;
                        rh1 = hd2 + th2 - thh;
                        rw2 = rw1 - th;
                        rh2 = rh1 - th;
                        rw3 = rw2 + th2;
                        rh3 = rh2 + th2;
                        wtH = rw3 * Math.sin(enAng);
                        htH = rh3 * Math.cos(enAng);
                        dxH = rw3 * Math.cos(Math.atan2(wtH, htH));
                        dyH = rh3 * Math.sin(Math.atan2(wtH, htH));
                        xH = hc + dxH;
                        yH = vc + dyH;
                        rI = (rw2 < rh2) ? rw2 : rh2;
                        u1 = dxH * dxH;
                        u2 = dyH * dyH;
                        u3 = rI * rI;
                        u4 = u1 - u3;
                        u5 = u2 - u3;
                        u6 = u4 * u5 / u1;
                        u7 = u6 / u2;
                        u8 = 1 - u7;
                        u9 = Math.sqrt(u8);
                        u10 = u4 / dxH;
                        u11 = u10 / dyH;
                        u12 = (1 + u9) / u11;
                        u13 = Math.atan2(u12, 1);
                        u14 = u13 + rdAngVal3;
                        u15 = (u13 > 0) ? u13 : u14;
                        u16 = u15 - enAng;
                        u17 = u16 + rdAngVal3;
                        u18 = (u16 > 0) ? u16 : u17;
                        u19 = u18 - cd2;
                        u20 = u18 - rdAngVal3;
                        u21 = (u19 > 0) ? u20 : u18;
                        u22 = Math.abs(u21);
                        minAng = u22 * -1;
                        u23 = Math.abs(adj2);
                        a2 = u23 * -1;
                        aAng = (a2 < minAng) ? minAng : (a2 > 0) ? 0 : a2;
                        ptAng = enAng + aAng;
                        wtA = rw3 * Math.sin(ptAng);
                        htA = rh3 * Math.cos(ptAng);
                        dxA = rw3 * Math.cos(Math.atan2(wtA, htA));
                        dyA = rh3 * Math.sin(Math.atan2(wtA, htA));
                        xA = hc + dxA;
                        yA = vc + dyA;
                        wtE = rw1 * Math.sin(stAng);
                        htE = rh1 * Math.cos(stAng);
                        dxE = rw1 * Math.cos(Math.atan2(wtE, htE));
                        dyE = rh1 * Math.sin(Math.atan2(wtE, htE));
                        xE = hc + dxE;
                        yE = vc + dyE;
                        wtD = rw2 * Math.sin(stAng);
                        htD = rh2 * Math.cos(stAng);
                        dxD = rw2 * Math.cos(Math.atan2(wtD, htD));
                        dyD = rh2 * Math.sin(Math.atan2(wtD, htD));
                        xD = hc + dxD;
                        yD = vc + dyD;
                        dxG = thh * Math.cos(ptAng);
                        dyG = thh * Math.sin(ptAng);
                        xG = xH + dxG;
                        yG = yH + dyG;
                        dxB = thh * Math.cos(ptAng);
                        dyB = thh * Math.sin(ptAng);
                        xB = xH - dxB;
                        yB = yH - dyB;
                        sx1 = xB - hc;
                        sy1 = yB - vc;
                        sx2 = xG - hc;
                        sy2 = yG - vc;
                        rO = (rw1 < rh1) ? rw1 : rh1;
                        x1O = sx1 * rO / rw1;
                        y1O = sy1 * rO / rh1;
                        x2O = sx2 * rO / rw1;
                        y2O = sy2 * rO / rh1;
                        dxO = x2O - x1O;
                        dyO = y2O - y1O;
                        dO = Math.sqrt(dxO * dxO + dyO * dyO);
                        q1 = x1O * y2O;
                        q2 = x2O * y1O;
                        DO = q1 - q2;
                        q3 = rO * rO;
                        q4 = dO * dO;
                        q5 = q3 * q4;
                        q6 = DO * DO;
                        q7 = q5 - q6;
                        q8 = (q7 > 0) ? q7 : 0;
                        sdelO = Math.sqrt(q8);
                        ndyO = dyO * -1;
                        sdyO = (ndyO > 0) ? -1 : 1;
                        q9 = sdyO * dxO;
                        q10 = q9 * sdelO;
                        q11 = DO * dyO;
                        dxF1 = (q11 + q10) / q4;
                        q12 = q11 - q10;
                        dxF2 = q12 / q4;
                        adyO = Math.abs(dyO);
                        q13 = adyO * sdelO;
                        q14 = DO * dxO / -1;
                        dyF1 = (q14 + q13) / q4;
                        q15 = q14 - q13;
                        dyF2 = q15 / q4;
                        q16 = x2O - dxF1;
                        q17 = x2O - dxF2;
                        q18 = y2O - dyF1;
                        q19 = y2O - dyF2;
                        q20 = Math.sqrt(q16 * q16 + q18 * q18);
                        q21 = Math.sqrt(q17 * q17 + q19 * q19);
                        q22 = q21 - q20;
                        dxF = (q22 > 0) ? dxF1 : dxF2;
                        dyF = (q22 > 0) ? dyF1 : dyF2;
                        sdxF = dxF * rw1 / rO;
                        sdyF = dyF * rh1 / rO;
                        xF = hc + sdxF;
                        yF = vc + sdyF;
                        x1I = sx1 * rI / rw2;
                        y1I = sy1 * rI / rh2;
                        x2I = sx2 * rI / rw2;
                        y2I = sy2 * rI / rh2;
                        dxI = x2I - x1I;
                        dyI = y2I - y1I;
                        dI = Math.sqrt(dxI * dxI + dyI * dyI);
                        v1 = x1I * y2I;
                        v2 = x2I * y1I;
                        DI = v1 - v2;
                        v3 = rI * rI;
                        v4 = dI * dI;
                        v5 = v3 * v4;
                        v6 = DI * DI;
                        v7 = v5 - v6;
                        v8 = (v7 > 0) ? v7 : 0;
                        sdelI = Math.sqrt(v8);
                        v9 = sdyO * dxI;
                        v10 = v9 * sdelI;
                        v11 = DI * dyI;
                        dxC1 = (v11 + v10) / v4;
                        v12 = v11 - v10;
                        dxC2 = v12 / v4;
                        adyI = Math.abs(dyI);
                        v13 = adyI * sdelI;
                        v14 = DI * dxI / -1;
                        dyC1 = (v14 + v13) / v4;
                        v15 = v14 - v13;
                        dyC2 = v15 / v4;
                        v16 = x1I - dxC1;
                        v17 = x1I - dxC2;
                        v18 = y1I - dyC1;
                        v19 = y1I - dyC2;
                        v20 = Math.sqrt(v16 * v16 + v18 * v18);
                        v21 = Math.sqrt(v17 * v17 + v19 * v19);
                        v22 = v21 - v20;
                        dxC = (v22 > 0) ? dxC1 : dxC2;
                        dyC = (v22 > 0) ? dyC1 : dyC2;
                        sdxC = dxC * rw2 / rI;
                        sdyC = dyC * rh2 / rI;
                        xC = hc + sdxC;
                        yC = vc + sdyC;
                        ist0 = Math.atan2(sdyC, sdxC);
                        ist1 = ist0 + rdAngVal3;
                        istAng0 = (ist0 > 0) ? ist0 : ist1;
                        isw1 = stAng - istAng0;
                        isw2 = isw1 + rdAngVal3;
                        iswAng0 = (isw1 > 0) ? isw1 : isw2;
                        istAng = istAng0 + iswAng0;
                        iswAng = -iswAng0;
                        p1 = xF - xC;
                        p2 = yF - yC;
                        p3 = Math.sqrt(p1 * p1 + p2 * p2);
                        p4 = p3 / 2;
                        p5 = p4 - thh;
                        xGp = (p5 > 0) ? xF : xG;
                        yGp = (p5 > 0) ? yF : yG;
                        xBp = (p5 > 0) ? xC : xB;
                        yBp = (p5 > 0) ? yC : yB;
                        en0 = Math.atan2(sdyF, sdxF);
                        en1 = en0 + rdAngVal3;
                        en2 = (en0 > 0) ? en0 : en1;
                        sw0 = en2 - stAng;
                        sw1 = sw0 - rdAngVal3;
                        swAng = (sw0 > 0) ? sw1 : sw0;
                        stAng0 = stAng + swAng;

                        var strtAng = stAng0 * 180 / Math.PI;
                        var endAng = stAng * 180 / Math.PI;
                        var stiAng = istAng * 180 / Math.PI;
                        var swiAng = iswAng * 180 / Math.PI;
                        var ediAng = stiAng + swiAng;

                        var d_val = "M" + xE + "," + yE +
                            " L" + xD + "," + yD +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, rw2, rh2, stiAng, ediAng, false).replace("M", "L") +
                            " L" + xBp + "," + yBp +
                            " L" + xA + "," + yA +
                            " L" + xGp + "," + yGp +
                            " L" + xF + "," + yF +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, rw1, rh1, strtAng, endAng, false).replace("M", "L") +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    }
                    case "funnel": {
                        /**
                         * funnel: 漏斗形
                         * 
                         * 形状说明：
                         * - 上宽下窄的漏斗形状
                         * - 常用于数据分析和流程图
                         */
                        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 40000 * SLIDE_FACTOR$1;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                        }
                        var cnstVal2 = 100000 * SLIDE_FACTOR$1;
                        var a = (adj < 0) ? 0 : (adj > cnstVal2) ? cnstVal2 : adj;
                        
                        // 漏斗底部宽度
                        var bottomW = w * a / cnstVal2;
                        
                        var d = "M0,0" + // 左上角
                            " L" + w + ",0" + // 右上角
                            " L" + ((w + bottomW) / 2) + "," + h + // 右下角
                            " L" + ((w - bottomW) / 2) + "," + h + // 左下角
                            " z";
                        
                        result += "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "leftRightCircularArrow": {
                        /**
                         * leftRightCircularArrow: 双向圆形箭头
                         * 
                         * 形状说明：
                         * - 圆形路径，两端有向左和向右的箭头
                         */
                        var wd2 = w / 2;
                        var hd2 = h / 2;
                        var r = Math.min(wd2, hd2);
                        
                        var d = "M" + (wd2 - r) + "," + hd2 +
                            PPTXShapeUtils.shapeArc(wd2, hd2, r, r, 180, 360, false).replace("M", "L") +
                            // 左箭头
                            " M" + (wd2 - r - r * 0.3) + "," + (hd2 - r * 0.2) +
                            " L" + (wd2 - r) + "," + hd2 +
                            " L" + (wd2 - r - r * 0.3) + "," + (hd2 + r * 0.2) +
                            // 右箭头
                            " M" + (wd2 + r + r * 0.3) + "," + (hd2 - r * 0.2) +
                            " L" + (wd2 + r) + "," + hd2 +
                            " L" + (wd2 + r + r * 0.3) + "," + (hd2 + r * 0.2);
                        
                        result += "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "flowChartOfflineStorage": {
                        /**
                         * flowChartOfflineStorage: 流程图 - 离线存储
                         * 
                         * 形状说明：
                         * - 底部有三个向下的尖角（代表存储）
                         */
                        var d = "M0,0" +
                            " L" + w + ",0" +
                            " L" + w + "," + (h * 0.7) +
                            " L" + (w * 0.66) + "," + h +
                            " L" + (w * 0.34) + "," + h +
                            " L0," + (h * 0.7) +
                            " z";
                        
                        result += "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                }

                result += "</svg>";

                // 生成 data- 属性
                const dataAttrs1 = genShapeDataAttributes(node, workingXfrmNode, id, name, idx, type, rotate, sType);
                
                // 检测动画信息并添加到data属性
                const animationData = extractAnimationData(node, warpObj);
                let animationAttrs = "";
                if (animationData) {
                    animationAttrs = ` data-animation='${JSON.stringify(animationData)}'`;
                }

                result += "<div class='block " + PPTXStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) + //block content
                    " " + PPTXStyleUtils.getContentDir(node, type, warpObj) +
                    "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                    "' style='" +
                    PPTXXmlUtils.getPosition(workingXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                    PPTXXmlUtils.getSize(workingXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
                    transform3dStyle +
                    " z-index: " + order + ";" +
                    "'" + dataAttrs1 + animationAttrs + ">";

                // TextBody
                if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    if (type != "diagram" && type != "textBox") {
                        type = "shape";
                    }
                    result += await PPTXTextUtils.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj); //type='shape'
                }
                result += "</div>";
            } else if (custShapType !== undefined) {
                // 对于 group-abs 情况，使用缩放后的尺寸进行形状计算
                // 否则使用原始尺寸
                const renderW = (sType === 'group-abs') ? w : drawW;
                const renderH = (sType === 'group-abs') ? h : drawH;
                result += renderCustomShape(custShapType, renderW, renderH, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArc);
                //console.log(result);

                result += "</svg>";

                // 生成 data- 属性
                const dataAttrs2 = genShapeDataAttributes(node, workingXfrmNode, id, name, idx, type, rotate, sType);
                
                // 检测动画信息并添加到data属性
                const animationData2 = extractAnimationData(node, warpObj);
                let animationAttrs2 = "";
                if (animationData2) {
                    animationAttrs2 = ` data-animation='${JSON.stringify(animationData2)}'`;
                }

                result += "<div class='block " + PPTXStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) + //block content
                    " " + PPTXStyleUtils.getContentDir(node, type, warpObj) +
                    "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                    "' style='" +
                    PPTXXmlUtils.getPosition(workingXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                    PPTXXmlUtils.getSize(workingXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
                    " z-index: " + order + ";" +
                    "'" + dataAttrs2 + animationAttrs2 + ">";

                // TextBody
                if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    if (type != "diagram" && type != "textBox") {
                        type = "shape";
                    }
                    // 对于 group-abs 情况，使用 workingXfrmNode 替换原始 xfrmNode
                    // 这样 genTextBody 就能获取到正确的缩放后尺寸
                    let textNode = node;
                    if (sType === 'group-abs' && workingXfrmNode !== slideXfrmNode) {
                        textNode = JSON.parse(JSON.stringify(node));
                        if (textNode["p:spPr"] && textNode["p:spPr"]["a:xfrm"]) {
                            textNode["p:spPr"]["a:xfrm"] = workingXfrmNode;
                        }
                    }
                    result += await PPTXTextUtils.genTextBody(textNode["p:txBody"], textNode, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj); //type=shape
                }
                result += "</div>";

                // result = "";
            } else {

                // 生成 data- 属性
                const dataAttrs3 = genShapeDataAttributes(node, slideXfrmNode, id, name, idx, type, rotate, sType);
                
                // 检测动画信息并添加到data属性
                const animationData3 = extractAnimationData(node, warpObj);
                if (animationData3) {
                    ` data-animation='${JSON.stringify(animationData3)}'`;
                }

                result += "<div class='block " + PPTXStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +//block content 
                    " " + PPTXStyleUtils.getContentDir(node, type, warpObj) +
                    "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                    "' style='" +
                    PPTXXmlUtils.getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                    PPTXXmlUtils.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
                    PPTXStyleUtils.getBorder(node, pNode, false, "shape", warpObj) +
                    await PPTXStyleUtils.getShapeFill(node, pNode, false, warpObj, source) +
                    " z-index: " + order + ";" +
                    "'" + dataAttrs3 + ">";

                // TextBody
                if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    result += await PPTXTextUtils.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj);
                }
                result += "</div>";

            }
            return result;
        }

    return {
        // 重新导出路径生成函数
        shapeArc: shapeArc,
        shapeArcAlt: shapeArcAlt,
        shapePie: shapePie,
        shapeGear: shapeGear,
        shapeSnipRoundRect: shapeSnipRoundRect,
        shapeSnipRoundRectAlt: shapeSnipRoundRectAlt,
        polarToCartesian: polarToCartesian,
        // 核心形状生成函数
        genShape,
    };

/**
 * 提取形状的动画数据
 * @param {Object} node - 形状节点
 * @param {Object} warpObj - 包装对象
 * @returns {Object|null} 动画数据或null
 */
function extractAnimationData(node, warpObj) {
    // 检查形状是否有动画引用
    const nvSpPr = PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr"]);
    if (!nvSpPr) return null;
    
    const nvPr = PPTXXmlUtils.getTextByPathList(nvSpPr, ["p:nvPr"]);
    if (!nvPr) return null;
    
    // 检查是否有动画效果
    const animLst = PPTXXmlUtils.getTextByPathList(nvPr, ["p:animLst"]);
    if (animLst) {
        // 解析动画列表
        return parseAnimationList(animLst);
    }
    
    // 检查是否有动画引用
    const animRef = PPTXXmlUtils.getTextByPathList(nvPr, ["p:animRef"]);
    if (animRef && animRef.attrs) {
        const rId = animRef.attrs["r:embed"];
        if (rId && warpObj.slideAnims && warpObj.slideAnims[rId]) {
            return warpObj.slideAnims[rId];
        }
    }
    
    return null;
}

/**
 * 解析动画列表
 * @param {Object} animLst - 动画列表
 * @returns {Object} 解析后的动画数据
 */
function parseAnimationList(animLst) {
    // 处理单个动画或动画数组
    const animArray = Array.isArray(animLst["p:par"]) ? animLst["p:par"] : 
                     (animLst["p:par"] ? [animLst["p:par"]] : []);
    
    if (animArray.length === 0) return null;
    
    const par = animArray[0];
    if (par["p:cTn"]) {
        const cTn = par["p:cTn"];
        const animType = getAnimationType(cTn);
        const duration = cTn.attrs["dur"] || "1000";
        const delay = cTn.attrs["st"] || "0";
        
        return {
            type: animType,
            duration: parseInt(duration),
            delay: parseInt(delay)
        };
    }
    
    return null;
}

/**
 * 获取动画类型
 * @param {Object} cTn - 动画时间节点
 * @returns {string} 动画类型
 */
function getAnimationType(cTn) {
    // 检查子动画类型
    if (cTn["p:childTnLst"]) {
        const childTnLst = cTn["p:childTnLst"];
        
        // 检查是否有set动画（属性设置）
        if (childTnLst["p:set"]) {
            const set = childTnLst["p:set"];
            if (set["p:to"]) {
                const to = set["p:to"];
                if (to["p:strVal"] && to["p:strVal"].attrs["val"] === "visible") {
                    return "fade-in";
                }
            }
        }
        
        // 检查是否有motion动画（移动）
        if (childTnLst["p:cmd"]) {
            return "custom";
        }
    }
    
    // 默认动画类型
    return "appear";
}

/**
 * 处理3D效果
 * @param {Object} scene3d - 场景3D数据
 * @param {Object} sp3d - 形状3D数据
 * @returns {string} CSS 3D变换样式
 */
function process3DEffects(scene3d, sp3d) {
    let transform = "";
    
    // 处理相机设置
    if (scene3d && scene3d["a:camera"]) {
        const camera = scene3d["a:camera"];
        const prst = camera.attrs?.["prst"];
        
        // 根据预设相机类型应用不同的视角
        switch (prst) {
            case "orthographicFront":
                // 正交前视图 - 无3D效果
                break;
            case "orthographicTop":
                transform += " rotateX(-90deg)";
                break;
            case "orthographicBottom":
                transform += " rotateX(90deg)";
                break;
            case "orthographicLeft":
                transform += " rotateY(90deg)";
                break;
            case "orthographicRight":
                transform += " rotateY(-90deg)";
                break;
            case "perspectiveFront":
                // 透视前视图 - 轻微的3D效果
                transform += " perspective(1000px)";
                break;
            case "perspectiveTop":
                transform += " perspective(1000px) rotateX(-60deg)";
                break;
            case "perspectiveBottom":
                transform += " perspective(1000px) rotateX(60deg)";
                break;
            case "perspectiveLeft":
                transform += " perspective(1000px) rotateY(60deg)";
                break;
            case "perspectiveRight":
                transform += " perspective(1000px) rotateY(-60deg)";
                break;
            default:
                // 默认轻微透视
                transform += " perspective(800px)";
        }
    }
    
    // 处理形状3D效果（挤出、斜角等）
    if (sp3d) {
        // 挤出深度
        if (sp3d["a:extrusionH"]) {
            const extrusionH = parseInt(sp3d["a:extrusionH"].attrs?.["val"] || "0");
            if (extrusionH > 0) {
                // 转换为像素值
                const depth = Math.round(extrusionH * SLIDE_FACTOR$1);
                if (depth > 0) {
                    transform += ` translateZ(${depth}px)`;
                }
            }
        }
        
        // 顶部斜角
        if (sp3d["a:bevelT"]) {
            const bevelT = sp3d["a:bevelT"];
            const w = parseInt(bevelT.attrs?.["w"] || "0");
            const h = parseInt(bevelT.attrs?.["h"] || "0");
            if (w > 0 || h > 0) {
                // 应用轻微的3D旋转来模拟斜角效果
                transform += " rotateX(5deg) rotateY(5deg)";
            }
        }
        
        // 底部斜角
        if (sp3d["a:bevelB"]) {
            const bevelB = sp3d["a:bevelB"];
            const w = parseInt(bevelB.attrs?.["w"] || "0");
            const h = parseInt(bevelB.attrs?.["h"] || "0");
            if (w > 0 || h > 0) {
                // 应用轻微的3D旋转来模拟底部斜角
                transform += " rotateX(-3deg) rotateY(-3deg)";
            }
        }
    }
    
    if (transform) {
        return ` transform:${transform}; transform-style: preserve-3d;`;
    }
    
    return "";
}
})();

/**
 * 节点工具函数模块
 * 
 * 处理 PPTX 节点的各种操作，包括：
 * - 幻灯片节点处理
 * - 图表生成
 * - SmartArt 图表处理
 * - 节点索引和查询
 * 
 * @module utils/node
 */


/**
 * 生成 Diagram HTML
 * @param {Object} node - 节点
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源类型
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @param {Object} parentNode - 父节点（用于组合元素的坐标计算）
 * @returns {Promise<string>} 生成的HTML
 */
async function genDiagram(node, wrapObj, source, shapeType, settings, parentNode) {
    node.attrs.order;
    const zip = wrapObj.zip;
    let xfrmNode = PPTXXmlUtils.getTextByPathList(node, ['p:xfrm']);
    const dgmRelIds = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'dgm:relIds', 'attrs']);
    const dgmClrFileId = dgmRelIds['r:cs'];
    const dgmDataFileId = dgmRelIds['r:dm'];
    const dgmLayoutFileId = dgmRelIds['r:lo'];
    const dgmQuickStyleFileId = dgmRelIds['r:qs'];

    const dgmClrFileName = wrapObj.slideResObj[dgmClrFileId].target;
    const dgmDataFileName = wrapObj.slideResObj[dgmDataFileId].target;
    const dgmLayoutFileName = wrapObj.slideResObj[dgmLayoutFileId].target;
    const dgmQuickStyleFileName = wrapObj.slideResObj[dgmQuickStyleFileId].target;

    await PPTXXmlUtils.readXmlFile(zip, dgmClrFileName);
    await PPTXXmlUtils.readXmlFile(zip, dgmDataFileName);
    await PPTXXmlUtils.readXmlFile(zip, dgmLayoutFileName);
    await PPTXXmlUtils.readXmlFile(zip, dgmQuickStyleFileName);

    const dgmDrwSpArray = PPTXXmlUtils.getTextByPathList(wrapObj.diagramContent, ['p:drawing', 'p:spTree', 'p:sp']);
    let result = '';

    if (dgmDrwSpArray !== undefined) {
        const results = [];
        for (const dspSp of dgmDrwSpArray) {
            PPTXXmlUtils.getTextByPathList(dspSp, ['p:txBody', 'a:p', 'a:r', 'a:t']);
            results.push(processSpNode(dspSp, node, wrapObj, 'diagramBg', shapeType));
        }
        const resolvedResults = await Promise.all(results);
        result = resolvedResults.join('');
    }

    // 处理组合缩放 - 当diagram在group-abs类型组合中时需要应用缩放
    let workingXfrmNode = xfrmNode;
    if (shapeType === 'group-abs' && wrapObj.currentGroupScale && xfrmNode) {
        const { scaleX, scaleY, childX, childY } = wrapObj.currentGroupScale;

        // 创建缩放后的xfrmNode
        workingXfrmNode = JSON.parse(JSON.stringify(xfrmNode));

        // 缩放尺寸
        if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
            const originalCx = parseInt(xfrmNode['a:ext'].attrs.cx);
            const originalCy = parseInt(xfrmNode['a:ext'].attrs.cy);
            workingXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
            workingXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
        }

        // 调整位置(相对于childX/childY)
        if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
            const originalOffX = parseInt(xfrmNode['a:off'].attrs.x);
            const originalOffY = parseInt(xfrmNode['a:off'].attrs.y);

            // 计算相对于childOff的偏移
            const relativeX = originalOffX - (childX / SLIDE_FACTOR$1);
            const relativeY = originalOffY - (childY / SLIDE_FACTOR$1);

            // 应用缩放
            workingXfrmNode['a:off'].attrs.x = Math.round(childX / SLIDE_FACTOR$1 + relativeX * scaleX);
            workingXfrmNode['a:off'].attrs.y = Math.round(childY / SLIDE_FACTOR$1 + relativeY * scaleY);
        }
    }

    const position = PPTXXmlUtils.getPosition(workingXfrmNode, parentNode, undefined, undefined, shapeType);
    const size = PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined);

    // 提取位置和尺寸信息
    let offX = 0, offY = 0, extCx = 0, extCy = 0;
    if (workingXfrmNode !== undefined) {
        if (workingXfrmNode['a:off'] && workingXfrmNode['a:off'].attrs) {
            offX = workingXfrmNode['a:off'].attrs.x || 0;
            offY = workingXfrmNode['a:off'].attrs.y || 0;
        }
        if (workingXfrmNode['a:ext'] && workingXfrmNode['a:ext'].attrs) {
            extCx = workingXfrmNode['a:ext'].attrs.cx || 0;
            extCy = workingXfrmNode['a:ext'].attrs.cy || 0;
        }
    }

    // 生成 data- 属性
    const dataAttrs = ` data-node-type="diagram" data-off-x="${offX}" data-off-y="${offY}" data-ext-cx="${extCx}" data-ext-cy="${extCy}"`;

    return `<div class='block diagram-content' style='${position}${size}'${dataAttrs}>${result}</div>`;
}

/**
 * 索引幻灯片节点
 * @param {Object} content - 幻灯片内容
 * @returns {Object} 包含idTable、idxTable和typeTable的对象
 */
function indexNodes(content) {
    const keys = Object.keys(content);
    const spTreeNode = content[keys[0]]['p:cSld']['p:spTree'];

    const idTable = {};
    const idxTable = {};
    const typeTable = {};

    for (const key in spTreeNode) {
        if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') {
            continue;
        }

        const targetNode = spTreeNode[key];

        if (Array.isArray(targetNode)) {
            for (const node of targetNode) {
                const nvSpPrNode = node['p:nvSpPr'];
                const id = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id']);
                const idx = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx']);
                const type = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type']);

                if (id !== undefined) idTable[id] = node;
                if (idx !== undefined) idxTable[idx] = node;
                if (type !== undefined) typeTable[type] = node;
            }
        } else {
            const nvSpPrNode = targetNode['p:nvSpPr'];
            const id = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id']);
            const idx = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx']);
            const type = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type']);

            if (id !== undefined) idTable[id] = targetNode;
            if (idx !== undefined) idxTable[idx] = targetNode;
            if (type !== undefined) typeTable[type] = targetNode;
        }
    }

    return { idTable, idxTable, typeTable };
}

/**
 * 辅助函数：将对象属性转换为 data- 属性字符串
 * @param {Object} obj - 要转换的对象
 * @param {string} prefix - data- 属性前缀（可选）
 * @returns {string} data- 属性字符串
 */
function objectToDataAttributes(obj, prefix = '') {
    if (!obj || typeof obj !== 'object') {
        return '';
    }
    
    let result = '';
    for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
            const value = obj[key];
            const dataKey = prefix ? `${prefix}-${key}` : key;
            
            if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
                // 递归处理嵌套对象
                result += objectToDataAttributes(value, dataKey);
            } else if (typeof value === 'string' || typeof value === 'number') {
                // 对数值类型进行四舍五入保留2位小数
                let attrValue;
                if (typeof value === 'number') {
                    attrValue = Math.round(value * 100) / 100;
                } else {
                    attrValue = value;
                }
                // 转换为 data- 属性
                const escapedValue = String(attrValue).replace(/'/g, '&#39;').replace(/"/g, '&quot;');
                result += ` data-${dataKey}="${escapedValue}"`;
            }
            // 忽略 undefined, null, Array 等
        }
    }
    
    return result;
}

/**
 * 处理组形状节点
 * @param {Object} node - 组形状节点
 * @param {Object} parentNode - 父节点
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function processGroupSpNode(node, parentNode, wrapObj, source, settings) {
    const xfrmNode = PPTXXmlUtils.getTextByPathList(node, ['p:grpSpPr', 'a:xfrm']);

    let groupStyle = '';
    let shapeType = 'group';
    let top, left, width, height;
    let rotate = 0;

    // 初始化所有变量，避免未定义错误
    let x = 0, y = 0, cx = 0, cy = 0;
    let childX = 0, childY = 0, childCx = 0, childCy = 0;

    if (xfrmNode !== undefined) {
        x = Math.round(parseInt(xfrmNode['a:off'].attrs.x) * SLIDE_FACTOR$1 * 100) / 100;
        y = Math.round(parseInt(xfrmNode['a:off'].attrs.y) * SLIDE_FACTOR$1 * 100) / 100;

        // 计算相对位置（对于嵌套组合）
        let parentChOffX = 0, parentChOffY = 0;
        if (parentNode !== undefined) {
            const parentGrpXfrmNode = PPTXXmlUtils.getTextByPathList(parentNode, ['p:grpSpPr', 'a:xfrm']);
            if (parentGrpXfrmNode !== undefined && parentGrpXfrmNode['a:chOff'] !== undefined && parentGrpXfrmNode['a:chOff'].attrs !== undefined) {
                parentChOffX = Math.round(parseInt(parentGrpXfrmNode['a:chOff'].attrs.x) * SLIDE_FACTOR$1 * 100) / 100;
                parentChOffY = Math.round(parseInt(parentGrpXfrmNode['a:chOff'].attrs.y) * SLIDE_FACTOR$1 * 100) / 100;
            }
        }

        // 根据ECMA-376标准，a:chOff和a:chExt是可选元素
        // 当不存在时，应该使用父元素的对应值作为默认值

        if (xfrmNode['a:chOff'] !== undefined && xfrmNode['a:chOff'].attrs !== undefined) {
            childX = Math.round(parseInt(xfrmNode['a:chOff'].attrs.x) * SLIDE_FACTOR$1 * 100) / 100;
            childY = Math.round(parseInt(xfrmNode['a:chOff'].attrs.y) * SLIDE_FACTOR$1 * 100) / 100;
        } else {
            // 当a:chOff不存在时，使用a:off的值作为默认值
            childX = x;
            childY = y;
        }

        // 对于嵌套组合，计算相对位置
        if (parentChOffX > 0 || parentChOffY > 0) {
            // 调整 childX/childY 为相对于父组的坐标
            childX = childX - parentChOffX;
            childY = childY - parentChOffY;
            // 调整 off 为相对于父组的坐标（用于设置 top/left）
            x = x - parentChOffX;
            y = y - parentChOffY;
        }

        cx = Math.round(parseInt(xfrmNode['a:ext'].attrs.cx) * SLIDE_FACTOR$1 * 100) / 100;
        cy = Math.round(parseInt(xfrmNode['a:ext'].attrs.cy) * SLIDE_FACTOR$1 * 100) / 100;

        if (xfrmNode['a:chExt'] !== undefined && xfrmNode['a:chExt'].attrs !== undefined) {
            childCx = Math.round(parseInt(xfrmNode['a:chExt'].attrs.cx) * SLIDE_FACTOR$1 * 100) / 100;
            childCy = Math.round(parseInt(xfrmNode['a:chExt'].attrs.cy) * SLIDE_FACTOR$1 * 100) / 100;
        } else {
            // 当a:chExt不存在时，使用a:ext的值作为默认值
            childCx = cx;
            childCy = cy;
        }

        rotate = parseInt(xfrmNode.attrs.rot) || 0;
        let rotationStyle = '';

        // 组合容器的位置和尺寸计算
        // 根据PPTX规范：
        // - off/ext: 组合在幻灯片上的位置和裁剪区域
        // - chOff/chExt: 子元素的坐标系原点和范围
        //
        // 策略：
        // - 当子元素不超出ext时，使用off/ext作为容器
        // - 当子元素超出ext时，子元素需要按比例缩放以适应容器
        if (childCx > cx || childCy > cy) {
            // 子元素超出ext边界，PPT会缩放子元素以适应容器
            // 计算缩放比例
            const scaleX = childCx > 0 ? cx / childCx : 1;
            const scaleY = childCy > 0 ? cy / childCy : 1;

            // 存储缩放比例供子元素使用
            wrapObj.currentGroupScale = { scaleX, scaleY, childX, childY };

            // 使用off/ext作为容器尺寸
            top = y;
            left = x;
            width = cx;
            height = cy;

            // 标记子元素需要使用绝对定位和缩放
            shapeType = 'group-abs';
        } else {
            // 子元素在ext边界内，使用off/ext
            wrapObj.currentGroupScale = null;
            top = y;
            left = x;
            width = cx;
            height = cy;
        }

        if (!isNaN(rotate)) {
            const degrees = PPTXXmlUtils.angleToDegrees(rotate);
            rotationStyle = `transform: rotate(${degrees}deg); transform-origin: center;`;
            if (degrees !== 0) {
                shapeType = 'group-rotate';
            }
        }

        if (rotationStyle) groupStyle += rotationStyle;
    }

    if (top !== undefined) groupStyle += `top: ${top}px;`;
    if (left !== undefined) groupStyle += `left: ${left}px;`;
    if (width !== undefined) groupStyle += `width: ${width}px;`;
    if (height !== undefined) groupStyle += `height: ${height}px;`;

    const order = node.attrs.order;
    // 生成 data- 属性
    const dataAttrs = objectToDataAttributes({
        'node-id': PPTXXmlUtils.getTextByPathList(node, ['p:nvGrpSpPr', 'p:cNvPr', 'attrs', 'id']),
        'node-name': PPTXXmlUtils.getTextByPathList(node, ['p:nvGrpSpPr', 'p:cNvPr', 'attrs', 'name']),
        'off-x': x,
        'off-y': y,
        'ext-cx': cx,
        'ext-cy': cy,
        'ch-off-x': childX,
        'ch-off-y': childY,
        'ch-ext-cx': childCx,
        'ch-ext-cy': childCy,
        'shape-type': shapeType,
        'rotate': rotate
    });
    
    let result = `<div class='block group' style='z-index: ${order};${groupStyle}'${dataAttrs}>`;

    // 保存之前的缩放信息(处理嵌套组合)
    const previousGroupScale = wrapObj.currentGroupScale;

    // Process all child nodes
    for (const nodeKey in node) {
        if (Array.isArray(node[nodeKey])) {
            for (const childNode of node[nodeKey]) {
                result += await processNodesInSlide(nodeKey, childNode, node, wrapObj, source, shapeType, settings, node);
            }
        } else {
            result += await processNodesInSlide(nodeKey, node[nodeKey], node, wrapObj, source, shapeType, settings, node);
        }
    }

    // 清除当前组合的缩放信息,恢复之前的缩放信息
    wrapObj.currentGroupScale = previousGroupScale;

    result += '</div>';
    return result;
}

/**
 * 处理幻灯片中的节点
 * @param {string} nodeKey - 节点键
 * @param {Object} nodeValue - 节点值
 * @param {Object} nodes - 节点集合
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @param {Object} parentNode - 父节点
 * @returns {Promise<string>} 生成的HTML
 */
async function processNodesInSlide(nodeKey, nodeValue, nodes, wrapObj, source, shapeType, settings, parentNode) {
    switch (nodeKey) {
        case 'p:sp':    // Shape, Text
            return await processSpNode(nodeValue, parentNode, wrapObj, source, shapeType, settings);
        case 'p:cxnSp':    // Shape, Text (with connection)
            return await processCxnSpNode(nodeValue, parentNode, wrapObj, source, shapeType, settings);
        case 'p:pic':    // Picture
            return await processPicNode(nodeValue, parentNode, wrapObj, source, shapeType, settings);
        case 'p:graphicFrame':    // Chart, Diagram, Table
            return await processGraphicFrameNode(nodeValue, parentNode, wrapObj, source, shapeType, settings);
        case 'p:grpSp':
            return await processGroupSpNode(nodeValue, parentNode, wrapObj, source, settings);
        case 'mc:AlternateContent': // Equations and formulas as Image
            const mcFallbackNode = PPTXXmlUtils.getTextByPathList(nodeValue, ['mc:Fallback']);
            return await processGroupSpNode(mcFallbackNode, parentNode, wrapObj, source, settings);
        default:
            return '';
    }
}

/**
 * 处理形状节点
 * @param {Object} node - 形状节点
 * @param {Object} parentNode - 父节点
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function processSpNode(node, parentNode, wrapObj, source, shapeType, settings) {
    const id = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'id']);
    const name = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'name']);
    let idx = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx']);
    let type = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
    const order = PPTXXmlUtils.getTextByPathList(node, ['attrs', 'order']);

    let isUserDrawnBg;
    if (source === 'slideLayoutBg' || source === 'slideMasterBg') {
        const userDrawn = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'attrs', 'userDrawn']);
        isUserDrawnBg = userDrawn === '1';
    }

    let slideLayoutSpNode;
    let slideMasterSpNode;

    if (idx !== undefined) {
        slideLayoutSpNode = wrapObj.slideLayoutTables.idxTable[idx];
        if (type !== undefined) {
            slideMasterSpNode = wrapObj.slideMasterTables.typeTable[type];
        } else {
            slideMasterSpNode = wrapObj.slideMasterTables.idxTable[idx];
        }
    } else if (type !== undefined) {
        slideLayoutSpNode = wrapObj.slideLayoutTables.typeTable[type];
        slideMasterSpNode = wrapObj.slideMasterTables.typeTable[type];
    }

    if (type === undefined) {
        const txBoxVal = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:cNvSpPr', 'attrs', 'txBox']);
        if (txBoxVal === '1') {
            type = 'textBox';
        }
    }

    if (type === undefined) {
        type = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
        if (type === undefined) {
            type = source === 'diagramBg' ? 'diagram' : 'obj';
        }
    }

    const result = await PPTXShapeUtils.genShape(node, parentNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, wrapObj, isUserDrawnBg, shapeType, source, settings);
    return result;
}

/**
 * 处理连接形状节点
 * @param {Object} node - 连接形状节点
 * @param {Object} parentNode - 父节点
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function processCxnSpNode(node, parentNode, wrapObj, source, shapeType, settings) {
    const id = node['p:nvCxnSpPr']['p:cNvPr'].attrs.id;
    const name = node['p:nvCxnSpPr']['p:cNvPr'].attrs.name;
    const idx = node['p:nvCxnSpPr']['p:nvPr']['p:ph'] === undefined 
        ? undefined 
        : node['p:nvCxnSpPr']['p:nvPr']['p:ph'].attrs.idx;
    const type = node['p:nvCxnSpPr']['p:nvPr']['p:ph'] === undefined 
        ? undefined 
        : node['p:nvCxnSpPr']['p:nvPr']['p:ph'].attrs.type;
    const order = node.attrs.order;

    return await PPTXShapeUtils.genShape(node, parentNode, undefined, undefined, id, name, idx, type, order, wrapObj, undefined, shapeType, source, settings);
}

/**
 * 处理图片节点
 * @param {Object} node - 图片节点
 * @param {Object} parentNode - 父节点（用于组合元素的坐标计算）
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function processPicNode(node, parentNode, wrapObj, source, shapeType, settings) {
    const order = node.attrs.order;
    const rid = node['p:blipFill']['a:blip'].attrs['r:embed'];
    
    let resObj;
    if (source === 'slideMasterBg') {
        resObj = wrapObj.masterResObj;
    } else if (source === 'slideLayoutBg') {
        resObj = wrapObj.layoutResObj;
    } else {
        resObj = wrapObj.slideResObj;
    }
    
    // 如果 resObj 不存在，尝试使用 slideResObj 作为备用
    if (resObj === undefined) {
        resObj = wrapObj.slideResObj;
    }
    
    // 如果仍然为 undefined，返回空字符串
    if (resObj === undefined) {
        return '';
    }
    
    const imgName = resObj[rid]?.target;

    if (imgName === undefined) {
        return '';
    }

    const imgFileExt = PPTXXmlUtils.extractFileExtension(imgName).toLowerCase();
    const zip = wrapObj.zip;
    
    // 确定上下文类型用于路径解析
    let context = 'slide';
    if (source === 'slideMasterBg') {
        context = 'master';
    } else if (source === 'slideLayoutBg') {
        context = 'layout';
    }
    
    // 使用改进的媒体文件查找方法
    const imgFile = PPTXXmlUtils.findMediaFile(zip, imgName, context, '');
    if (imgFile === null) {
        return '';
    }
    
    const imgArrayBuffer = await imgFile.async("arraybuffer");
    let xfrmNode = node['p:spPr']?.['a:xfrm'];
    
    if (xfrmNode === undefined) {
        const idx = PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph', 'attrs', 'idx']);
        if (idx !== undefined) {
            xfrmNode = PPTXXmlUtils.getTextByPathList(wrapObj.slideLayoutTables, ['idxTable', idx, 'p:spPr', 'a:xfrm']);
        }
    }

    // 计算旋转角度
    let rotate = 0;
    const rotateNode = PPTXXmlUtils.getTextByPathList(node, ['p:spPr', 'a:xfrm', 'attrs', 'rot']);
    if (rotateNode !== undefined) {
        rotate = PPTXXmlUtils.angleToDegrees(rotateNode);
    }

    // 处理视频
    let mediaSupportFlag = false;
    let mediaPicFlag = false;
    let isVideoLink = false;
    let videoBlob, videoFile;
    
    const vdoNode = PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:videoFile']);
    const mediaProcess = settings.mediaProcess;
    
    if (vdoNode !== undefined && mediaProcess) {
        const vdoRid = vdoNode.attrs['r:link'];
        videoFile = resObj[vdoRid].target;
        const checkIfLink = PPTXXmlUtils.IsVideoLink(videoFile);
        
        if (checkIfLink) {
            // 使用新的URL转换函数将视频链接转换为embed格式
            videoFile = PPTXXmlUtils.convertVideoToEmbed(videoFile);
            videoFile = PPTXXmlUtils.escapeHtml(videoFile);
            isVideoLink = true;
            mediaSupportFlag = true;
            mediaPicFlag = true;
        } else {
            const vdoFileExt = PPTXXmlUtils.extractFileExtension(videoFile).toLowerCase();
            if (['mp4', 'webm', 'ogg'].includes(vdoFileExt)) {
                const vdoFileObj = PPTXXmlUtils.findMediaFile(zip, videoFile, context, '');
                if (vdoFileObj !== null) {
                    const uInt8Array = await vdoFileObj.async("arraybuffer");
                    const vdoMimeType = PPTXXmlUtils.getMimeType(vdoFileExt);
                    const blob = new Blob([uInt8Array], { type: vdoMimeType });
                    videoBlob = URL.createObjectURL(blob);
                    mediaSupportFlag = true;
                    mediaPicFlag = true;
                }
            }
        }
    }

    // 处理音频
    let audioPlayerFlag = false;
    let audioBlob;
    let audioObj;
    
    const audioNode = PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:audioFile']);
    
    if (audioNode !== undefined && mediaProcess) {
        const audioRid = audioNode.attrs['r:link'];
        const audioFile = resObj[audioRid].target;
        const audioFileExt = PPTXXmlUtils.extractFileExtension(audioFile).toLowerCase();
        
        if (['mp3', 'wav', 'ogg'].includes(audioFileExt)) {
            const audioFileObj = PPTXXmlUtils.findMediaFile(zip, audioFile, context, '');
            if (audioFileObj !== null) {
                const uInt8ArrayAudio = await audioFileObj.async("arraybuffer");
                const blobAudio = new Blob([uInt8ArrayAudio]);
                audioBlob = URL.createObjectURL(blobAudio);

                const cx = parseInt(xfrmNode['a:ext'].attrs.cx) * 20;
                const cy = parseInt(xfrmNode['a:ext'].attrs.cy);
                const x = parseInt(xfrmNode['a:off'].attrs.x) / 2.5;
                const y = parseInt(xfrmNode['a:off'].attrs.y);

                audioObj = {
                    'a:ext': { attrs: { cx, cy } },
                    'a:off': { attrs: { x, y } }
                };

                audioPlayerFlag = true;
                mediaSupportFlag = true;
                mediaPicFlag = true;
            }
        }
    }

    const mimeType = PPTXXmlUtils.getMimeType(imgFileExt);

    // 检查是否需要应用组合缩放
    let scaledXfrmNode = null;
    if (shapeType === 'group-abs' && wrapObj.currentGroupScale) {
        const { scaleX, scaleY, childX, childY } = wrapObj.currentGroupScale;

        // 创建缩放后的xfrmNode
        if (xfrmNode !== undefined) {
            scaledXfrmNode = JSON.parse(JSON.stringify(xfrmNode)); // 深拷贝

            // 缩放尺寸
            if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
                const originalCx = parseInt(xfrmNode['a:ext'].attrs.cx);
                const originalCy = parseInt(xfrmNode['a:ext'].attrs.cy);
                scaledXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
                scaledXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
            }

            // 调整位置(相对于childX/childY)
            if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
                const originalOffX = parseInt(xfrmNode['a:off'].attrs.x);
                const originalOffY = parseInt(xfrmNode['a:off'].attrs.y);

                // 计算相对于childOff的偏移
                const relativeX = originalOffX - (childX / SLIDE_FACTOR$1);
                const relativeY = originalOffY - (childY / SLIDE_FACTOR$1);

                // 应用缩放
                scaledXfrmNode['a:off'].attrs.x = Math.round(childX / SLIDE_FACTOR$1 + relativeX * scaleX);
                scaledXfrmNode['a:off'].attrs.y = Math.round(childY / SLIDE_FACTOR$1 + relativeY * scaleY);
            }
        }
    }

    const position = mediaProcess && audioPlayerFlag
        ? PPTXXmlUtils.getPosition(audioObj, parentNode, undefined, undefined, shapeType)
        : PPTXXmlUtils.getPosition(scaledXfrmNode || xfrmNode, parentNode, undefined, undefined, shapeType);
    const size = mediaProcess && audioPlayerFlag
        ? PPTXXmlUtils.getSize(audioObj, undefined, undefined)
        : PPTXXmlUtils.getSize(scaledXfrmNode || xfrmNode, undefined, undefined);

    // 提取图片位置信息
    let imgOffX = 0, imgOffY = 0, imgExtCx = 0, imgExtCy = 0;
    if (xfrmNode !== undefined) {
        if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
            imgOffX = parseInt(xfrmNode['a:off'].attrs.x) * SLIDE_FACTOR$1;
            imgOffY = parseInt(xfrmNode['a:off'].attrs.y) * SLIDE_FACTOR$1;
        }
        if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
            imgExtCx = parseInt(xfrmNode['a:ext'].attrs.cx) * SLIDE_FACTOR$1;
            imgExtCy = parseInt(xfrmNode['a:ext'].attrs.cy) * SLIDE_FACTOR$1;
        }
    }

    // 生成 data- 属性
    const dataAttrs = objectToDataAttributes({
        'node-id': PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:cNvPr', 'attrs', 'id']),
        'node-name': PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:cNvPr', 'attrs', 'name']),
        'node-descr': PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:cNvPr', 'attrs', 'descr']),
        'off-x': imgOffX,
        'off-y': imgOffY,
        'ext-cx': imgExtCx,
        'ext-cy': imgExtCy,
        'shape-type': shapeType,
        'rotate': rotate,
        'is-video': (vdoNode !== undefined) ? 'true' : 'false',
        'is-audio': (audioNode !== undefined) ? 'true' : 'false',
        'is-gif': (mimeType === 'image/gif') ? 'true' : 'false'
    });

    let result = `<div class='block content' style='${position}${size} z-index: ${order};transform: rotate(${rotate}deg);'${dataAttrs}>`;
    
    if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
        const base64Data = PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer);
        // 检测GIF格式，添加autoplay支持（GIF自动播放是浏览器默认行为）
        const gifAttrs = (mimeType === 'image/gif') ? 'autoplay loop muted playsinline' : '';
        result += `<img src='data:${mimeType};base64,${base64Data}' style='width: 100%; height: 100%' ${gifAttrs}/>`;
    } else if ((vdoNode !== undefined || audioNode !== undefined) && mediaProcess && mediaSupportFlag) {
        if (vdoNode !== undefined && !isVideoLink) {
            result += `<video src='${videoBlob}' autoplay loop muted controls style='width: 100%; height: 100%'>Your browser does not support the video tag.</video>`;
        } else if (vdoNode !== undefined && isVideoLink) {
            // 使用iframe嵌入视频，支持YouTube/Vimeo等
            // 添加allowfullscreen支持，并设置合适的sandbox权限
            const iframeAttrs = 'allowfullscreen allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" loading="lazy"';
            result += `<iframe src='${videoFile}' ${iframeAttrs} style='width: 100%; height: 100%; border: none;'></iframe>`;
        }
        if (audioNode !== undefined) {
            result += `<audio id="audio_player" controls><source src="${audioBlob}"></audio>`;
        }
    }
    
    if (!mediaSupportFlag && mediaPicFlag) {
        result += `<span style='color:red;font-size:40px;position: absolute;'>This media file Not supported by HTML5</span>`;
    }
    
    result += '</div>';
    return result;
}

/**
 * 处理图形框架节点
 * @param {Object} node - 图形框架节点
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function processGraphicFrameNode(node, parentNode, wrapObj, source, shapeType, settings) {
    const graphicTypeUri = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'attrs', 'uri']);

    switch (graphicTypeUri) {
        case 'http://schemas.openxmlformats.org/drawingml/2006/table':
            return await PPTXTextUtils.genTable(node, wrapObj, shapeType);
        case 'http://schemas.openxmlformats.org/drawingml/2006/chart':
            return await genChart(node, wrapObj, parentNode);
        case 'http://schemas.openxmlformats.org/drawingml/2006/diagram':
            return await genDiagram(node, wrapObj, source, shapeType, settings, parentNode);
        case 'http://schemas.openxmlformats.org/presentationml/2006/ole':
            let oleObjNode = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'mc:AlternateContent', 'mc:Fallback', 'p:oleObj']);
            if (oleObjNode === undefined) {
                oleObjNode = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'p:oleObj']);
            }
            if (oleObjNode !== undefined) {
                return await processGroupSpNode(oleObjNode, undefined, wrapObj, source, settings);
            }
            return '';
        default:
            return '';
    }
}

/**
 * 处理形状属性节点
 * @param {Object} node - 形状属性节点
 * @param {Object} wrapObj - 包装对象
 */
function processSpPrNode(node, wrapObj) {
    // TODO: Implement shape properties processing
}

/**
 * 获取幻灯片背景
 * @param {Object} wrapObj - 包装对象
 * @param {Object} slideSize - 幻灯片尺寸
 * @param {number} index - 幻灯片索引
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 背景HTML
 */
async function getBackground(wrapObj, slideSize, index, settings) {
    wrapObj.slideContent;
    const slideLayoutContent = wrapObj.slideLayoutContent;
    const slideMasterContent = wrapObj.slideMasterContent;

    const nodesSldLayout = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:spTree']);
    const nodesSldMaster = PPTXXmlUtils.getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:spTree']);
    const showMasterSp = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ['p:sldLayout', 'attrs', 'showMasterSp']);
    
    const bgColor = await PPTXStyleUtils.getSlideBackgroundFill(wrapObj, index);
    let result = `<div class='slide-background-${index}' style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;

    if (nodesSldLayout !== undefined) {
        for (const nodeKey in nodesSldLayout) {
            if (Array.isArray(nodesSldLayout[nodeKey])) {
                for (const node of nodesSldLayout[nodeKey]) {
                    const phType = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
                    if (phType !== 'pic') {
                        result += await processNodesInSlide(nodeKey, node, nodesSldLayout, wrapObj, 'slideLayoutBg', 'group', settings, undefined);
                    }
                }
            } else {
                const phType = PPTXXmlUtils.getTextByPathList(nodesSldLayout[nodeKey], ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
                if (phType !== 'pic') {
                    result += await processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], nodesSldLayout, wrapObj, 'slideLayoutBg', 'group', settings, undefined);
                }
            }
        }
    }
    
    if (nodesSldMaster !== undefined && (showMasterSp === '1' || showMasterSp === undefined)) {
        for (const nodeKey in nodesSldMaster) {
            if (Array.isArray(nodesSldMaster[nodeKey])) {
                for (const node of nodesSldMaster[nodeKey]) {
                    PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
                    result += await processNodesInSlide(nodeKey, node, nodesSldMaster, wrapObj, 'slideMasterBg', 'group', settings, undefined);
                }
            } else {
                result += await processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], nodesSldMaster, wrapObj, 'slideMasterBg', 'group', settings, undefined);
            }
        }
    }
    
    return result;
}

// 创建兼容的 PPTXNodeUtils 对象
const PPTXNodeUtils = {
    indexNodes,
    processGroupSpNode,
    processNodesInSlide,
    processSpNode,
    processCxnSpNode,
    processPicNode,
    processGraphicFrameNode,
    processSpPrNode,
    getBackground,
    genDiagram
};

/**
 * Parse PPTX file to structured JSON data (internal function)
 * @param {ArrayBuffer} file - The PPTX file data
 * @param {Object} settings - Conversion settings
 * @param {Object} callbacks - Callback functions
 * @param {Object} chartId - Chart ID tracker
 * @param {Object} styleTable - Style table
 * @param {*} defaultTextStyle - Default text style
 * @returns {Promise<Object>} Parsed result with structured data
 */
async function processToJson(file, settings, callbacks, chartId, styleTable, defaultTextStyle) {
    if (file.byteLength < 10) {
        if (callbacks.onError) {
            callbacks.onError({ type: "file_error", message: "Invalid file: file too small" });
        }
        throw new Error("Invalid file: file too small");
    }

    const msgQueue = [];
    const zip = JSZip.loadAsync ? await JSZip.loadAsync(file) : new JSZip().load(file);

    // Parse PPTX to structured data
    const parsedData = await parsePPTXInternal(zip, msgQueue, settings, chartId, styleTable, defaultTextStyle);

    // Return structured data (styleTable is populated during parsing)
    return {
        parsedData,
        msgQueue,
        zip,
        slideSize: parsedData.slideSize,
        thumbnail: parsedData.thumbnail,
        metadata: parsedData.metadata,
        executionTime: parsedData.executionTime
    };
}

/**
 * Parse PPTX to structured data (internal function)
 * @param {JSZip} zip - The JSZip instance
 * @param {Array} msgQueue - Message queue for charts
 * @param {Object} settings - Conversion settings
 * @param {Object} chartId - Chart ID tracker
 * @param {Object} styleTable - Style table
 * @param {*} defaultTextStyle - Default text style
 * @returns {Promise<Object>} Structured PPTX data
 */
async function parsePPTXInternal(zip, msgQueue, settings, chartId, styleTable, defaultTextStyle) {
    const dateBefore = new Date();

    // Extract thumbnail if exists
    const thumbFile = zip.file("docProps/thumbnail.jpeg");
    let thumbnail = null;
    if (thumbFile !== null) {
        const pptxThumbImg = PPTXXmlUtils.base64ArrayBuffer(await thumbFile.async("arraybuffer"));
        thumbnail = pptxThumbImg;
    }

    // Extract metadata from core.xml
    let metadata = {};
    try {
        const coreFile = zip.file("docProps/core.xml");
        if (coreFile !== null) {
            const coreContent = await PPTXXmlUtils.readXmlFile(zip, "docProps/core.xml");
            if (coreContent !== null) {
                const coreProperties = coreContent["cp:coreProperties"];
                if (coreProperties) {
                    // Extract common metadata fields
                    metadata = {
                        title: coreProperties["dc:title"] || undefined,
                        subject: coreProperties["dc:subject"] || undefined,
                        author: coreProperties["dc:creator"] || undefined,
                        keywords: coreProperties["cp:keywords"] || undefined,
                        description: coreProperties["dc:description"] || undefined,
                        lastModifiedBy: coreProperties["cp:lastModifiedBy"] || undefined,
                        created: coreProperties["dcterms:created"] || undefined,
                        modified: coreProperties["dcterms:modified"] || undefined,
                        category: coreProperties["cp:category"] || undefined,
                        status: coreProperties["cp:contentStatus"] || undefined,
                        contentType: coreProperties["dc:type"] || undefined,
                        language: coreProperties["dc:language"] || undefined
                    };
                }
            }
        }
    } catch (error) {
        // If error, return empty metadata object
        metadata = {};
    }

    const filesInfo = await PPTXXmlUtils.getContentTypes(zip);
    const slideSize = await PPTXXmlUtils.getSlideSizeAndSetDefaultTextStyle(zip, settings);

    const slides = [];
    const numOfSlides = filesInfo.slides.length;
    
    for (let i = 0; i < numOfSlides; i++) {
        const filename = filesInfo.slides[i];
        let fileNameNoPath = "";

        if (filename.includes("/")) {
            const pathParts = filename.split("/");
            fileNameNoPath = pathParts.pop();
        } else {
            fileNameNoPath = filename;
        }

        let fileNameNoExt = "";
        if (fileNameNoPath.includes(".")) {
            const nameParts = fileNameNoPath.split(".");
            nameParts.pop();
            fileNameNoExt = nameParts.join(".");
        }

        let slideNumber = 1;
        if (fileNameNoExt !== "" && fileNameNoPath.includes("slide")) {
            slideNumber = Number(fileNameNoExt.substring(5));
        }

        // Process slide and get structured data
        const slideData = await processSingleSlideStructured(zip, filename, i, slideSize, msgQueue, settings, chartId, styleTable, defaultTextStyle);
        
        slides.push({
            slideNum: slideNumber,
            fileName: fileNameNoExt,
            data: slideData
        });
    }

    // Sort slides by slideNum to ensure correct order
    slides.sort((a, b) => a.slideNum - b.slideNum);

    const dateAfter = new Date();

    return {
        slides,
        slideSize,
        thumbnail,
        metadata,
        executionTime: dateAfter - dateBefore
    };
}

/**
 * Process a single slide and extract structured data
 * @param {JSZip} zip - The JSZip instance
 * @param {string} slideFileName - Slide file name
 * @param {number} index - Slide index
 * @param {Object} slideSize - Slide size info
 * @param {Array} msgQueue - Message queue
 * @param {Object} settings - Conversion settings
 * @param {Object} chartId - Chart ID tracker
 * @param {Object} styleTable - Style table
 * @param {*} defaultTextStyle - Default text style
 * @returns {Promise<Object>} Structured slide data
 */
async function processSingleSlideStructured(zip, slideFileName, index, slideSize, msgQueue, settings, chartId, styleTable, defaultTextStyle) {
    // Read relationship file of the slide
    const resName = slideFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
    const resContent = await PPTXXmlUtils.readXmlFile(zip, resName);
    const relationshipArray = resContent.Relationships.Relationship;

    let layoutFilename = "";
    let diagramFilename = "";
    let notesFilename = ""; // 添加备注文件名
    const slideResObj = {};

    if (Array.isArray(relationshipArray)) {
        for (const rel of relationshipArray) {
            const relType = rel.attrs.Type;
            const target = rel.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");

            switch (relType) {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                    layoutFilename = target;
                    break;
                case "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing":
                    diagramFilename = target;
                    slideResObj[rel.attrs.Id] = {
                        type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target
                    };
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
                    // 处理备注关系
                    notesFilename = target;
                    break;
                default:
                    slideResObj[rel.attrs.Id] = {
                        type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target
                    };
            }
        }
    } else {
        const relType = relationshipArray.attrs.Type;
        const target = relationshipArray.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");
        
        if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout") {
            layoutFilename = target;
        } else if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide") {
            notesFilename = target;
        } else if (relType === "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing") {
            diagramFilename = target;
            slideResObj[relationshipArray.attrs.Id] = {
                type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                target: target
            };
        } else {
            // 默认处理为slideLayout
            layoutFilename = target;
        }
    }

    // Open slide layout
    const slideLayoutContent = await PPTXXmlUtils.readXmlFile(zip, layoutFilename);
    const slideLayoutTables = PPTXNodeUtils.indexNodes(slideLayoutContent);
    const layoutColorOverride = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);
    if (layoutColorOverride !== undefined) {
        layoutColorOverride.attrs;
    }

    // Read slide master
    const slideLayoutResFilename = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
    const slideLayoutResContent = await PPTXXmlUtils.readXmlFile(zip, slideLayoutResFilename);
    const layoutRelArray = slideLayoutResContent.Relationships.Relationship;

    let masterFilename = "";
    const layoutResObj = {};

    if (Array.isArray(layoutRelArray)) {
        for (const rel of layoutRelArray) {
            const relType = rel.attrs.Type;
            const target = rel.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");

            if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster") {
                masterFilename = target;
            } else {
                layoutResObj[rel.attrs.Id] = {
                    type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                    target
                };
            }
        }
    } else {
        masterFilename = layoutRelArray.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");
    }

    // Open slide master
    const slideMasterContent = await PPTXXmlUtils.readXmlFile(zip, masterFilename);
    const slideMasterTextStyles = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
    const slideMasterTables = PPTXNodeUtils.indexNodes(slideMasterContent);

    // Read slide master relationships
    const slideMasterResFilename = masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster") + ".rels";
    const slideMasterResContent = await PPTXXmlUtils.readXmlFile(zip, slideMasterResFilename);
    const masterRelArray = slideMasterResContent.Relationships.Relationship;

    let themeFilename = "";
    const masterResObj = {};

    if (Array.isArray(masterRelArray)) {
        for (const rel of masterRelArray) {
            const relType = rel.attrs.Type;
            const target = rel.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");

            if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
                themeFilename = target;
            } else {
                masterResObj[rel.attrs.Id] = {
                    type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                    target
                };
            }
        }
    } else {
        themeFilename = masterRelArray.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");
    }

    // Load theme file
    let themeContent;
    const themeResObj = {};

    if (themeFilename !== undefined) {
        const themeName = themeFilename.split("/").pop();
        const themeResFileName = themeFilename.replace(themeName, `_rels/${themeName}`) + ".rels";

        themeContent = await PPTXXmlUtils.readXmlFile(zip, themeFilename);
        const themeResContent = await PPTXXmlUtils.readXmlFile(zip, themeResFileName);

        if (themeResContent !== null) {
            const themeRelArray = themeResContent.Relationships.Relationship;
            if (themeRelArray !== undefined) {
                if (Array.isArray(themeRelArray)) {
                    for (const rel of themeRelArray) {
                        themeResObj[rel.attrs.Id] = {
                            type: rel.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target: rel.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "")
                        };
                    }
                } else {
                    themeResObj[themeRelArray.attrs.Id] = {
                        type: themeRelArray.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target: themeRelArray.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "")
                    };
                }
            }
        }
    }

    // Load diagram file
    let diagramContent = {};
    const diagramResObj = {};

    if (diagramFilename !== undefined) {
        const diagramName = diagramFilename.split("/").pop();
        const diagramResFileName = diagramFilename.replace(diagramName, `_rels/${diagramName}`) + ".rels";

        diagramContent = await PPTXXmlUtils.readXmlFile(zip, diagramFilename);
        if (diagramContent !== null && diagramContent !== undefined && diagramContent !== "") {
            const diagramJson = JSON.stringify(diagramContent);
            const cleanedJson = diagramJson.replace(/dsp:/g, "p:");
            diagramContent = JSON.parse(cleanedJson);
        }

        const diagramResContent = await PPTXXmlUtils.readXmlFile(zip, diagramResFileName);
        if (diagramResContent !== null) {
            const diagramRelArray = diagramResContent.Relationships.Relationship;
            if (Array.isArray(diagramRelArray)) {
                for (const rel of diagramRelArray) {
                    diagramResObj[rel.attrs.Id] = {
                        type: rel.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target: rel.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "")
                    };
                }
            } else {
                diagramResObj[diagramRelArray.attrs.Id] = {
                    type: diagramRelArray.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                    target: diagramRelArray.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "")
                };
            }
        }
    }

    // Load table styles
    const tableStyles = await PPTXXmlUtils.readXmlFile(zip, "ppt/tableStyles.xml");

    // Read slide content
    const slideContent = await PPTXXmlUtils.readXmlFile(zip, slideFileName, true);
    
    // Read notes content if available
    let notesContent = null;
    if (notesFilename) {
        notesContent = await PPTXXmlUtils.readXmlFile(zip, notesFilename);
    }

    slideContent["p:sld"]["p:cSld"]["p:spTree"];

    settings.themeProcess;

    // Return structured slide data
    return {
        slideLayoutContent,
        slideLayoutTables,
        slideMasterContent,
        slideMasterTables,
        slideContent,
        slideResObj,
        slideMasterTextStyles,
        layoutResObj,
        masterResObj,
        themeContent,
        themeResObj,
        diagramContent,
        diagramResObj,
        notesContent, // 添加备注内容
        defaultTextStyle: slideSize.defaultTextStyle || defaultTextStyle,
        tableStyles,
        styleTable,
        chartId,
        msgQueue,
        bulletCounter: {},
        slideSize,
        index
    };
}

/**
 * Convert structured slide data to HTML
 * @param {Object} slideData - Structured slide data
 * @param {Object} slideSize - Slide size info
 * @param {Object} settings - Conversion settings
 * @param {JSZip} zip - The JSZip instance
 * @returns {Promise<string>} Slide HTML
 */
async function convertSlideDataToHtml(slideData, slideSize, settings, zip) {
    const warpObj = {
        slideLayoutContent: slideData.slideLayoutContent,
        slideLayoutTables: slideData.slideLayoutTables,
        slideMasterContent: slideData.slideMasterContent,
        slideMasterTables: slideData.slideMasterTables,
        slideContent: slideData.slideContent,
        slideResObj: slideData.slideResObj,
        slideMasterTextStyles: slideData.slideMasterTextStyles,
        layoutResObj: slideData.layoutResObj,
        masterResObj: slideData.masterResObj,
        themeContent: slideData.themeContent,
        themeResObj: slideData.themeResObj,
        diagramContent: slideData.diagramContent,
        diagramResObj: slideData.diagramResObj,
        defaultTextStyle: slideData.defaultTextStyle,
        tableStyles: slideData.tableStyles,
        styleTable: slideData.styleTable,
        chartId: slideData.chartId,
        msgQueue: slideData.msgQueue,
        bulletCounter: slideData.bulletCounter,
        zip: zip
    };

    const processFullTheme = settings.themeProcess;
    let bgResult = "";
    if (processFullTheme === true) {
        bgResult = await PPTXNodeUtils.getBackground(warpObj, slideSize, slideData.index, settings, PPTXStyleUtils);
    }

    let bgColor = "";
    if (processFullTheme === "colorsAndImageOnly") {
        bgColor = await PPTXStyleUtils.getSlideBackgroundFill(warpObj, slideData.index);
    }
    
    // 检测幻灯片过渡效果
    let transitionClass = "";
    const transitionData = extractSlideTransition(slideData.slideContent);
    if (transitionData) {
        transitionClass = ` data-transition='${JSON.stringify(transitionData)}'`;
    }

    let result = `<section class='slide'${transitionClass} style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;
    result += bgResult;

    const nodes = slideData.slideContent["p:sld"]["p:cSld"]["p:spTree"];
    for (const nodeKey in nodes) {
        if (Array.isArray(nodes[nodeKey])) {
            for (const node of nodes[nodeKey]) {
                result += await PPTXNodeUtils.processNodesInSlide(nodeKey, node, nodes, warpObj, "slide", "group", settings);
            }
        } else {
            result += await PPTXNodeUtils.processNodesInSlide(nodeKey, nodes[nodeKey], nodes, warpObj, "slide", "group", settings);
        }
    }

    return `${result}</div></section>`;
}

/**
 * Generate global CSS
 * @param {Object} styleTable - Style table
 * @returns {string} CSS text
 */
function genGlobalCSS(styleTable) {
    let cssText = "";
    for (const key in styleTable) {
        const suffix = styleTable[key].suffix || "";
        cssText += ` .${styleTable[key].name}${suffix}{${styleTable[key].text}}\n`;
    }
    return cssText;
}

/**
 * PPTX to HTML converter
 * @param {ArrayBuffer} fileData - The PPTX file data
 * @param {Object} options - Conversion options
 * @returns {Promise<Object>} Parsed result
 */
async function pptxToHtml(fileData, options) {
    // Merge default settings with user options
    const settings = {
        mediaProcess: true,
        themeProcess: true,
        incSlide: {
            width: 0,
            height: 0
        },
        styleTable: {},
        ...options
    };

    // Callback functions
    const callbacks = settings.callbacks || {};

    // State variables
    let defaultTextStyle = null;
    const chartId = { value: 0 };
    const styleTable = settings.styleTable;

    // Trigger file start callback
    if (callbacks.onFileStart) {
        callbacks.onFileStart();
    }

    /**
     * Convert PPTX file to HTML
     * @param {ArrayBuffer} file - The PPTX file data
     * @returns {Promise<Object>} Parsed result
     */
    async function convertToHtml(file) {
        // Step 1: Parse PPTX to structured JSON data
        const { parsedData, msgQueue, zip, slideSize, thumbnail, metadata, executionTime } = 
            await processToJson(file, settings, callbacks, chartId, styleTable, defaultTextStyle);

        // Step 2: Convert structured data to HTML result
        const result = {
            slides: [],
            slideSize,
            thumbnail,
            styles: {
                global: ""
            },
            metadata,
            charts: []
        };

        // Step 3: Process slides and convert to HTML
        for (const slideData of parsedData.slides) {
            const slideHtml = await convertSlideDataToHtml(slideData.data, slideSize, settings, zip);
            result.slides.push({
                html: slideHtml,
                data: slideData.data,  // Keep structured data for potential reuse
                slideNum: slideData.slideNum,
                fileName: slideData.fileName
            });

            if (callbacks.onSlide) {
                callbacks.onSlide(slideHtml, {
                    slideNum: slideData.slideNum,
                    fileName: slideData.fileName
                });
            }
        }

        // Step 4: Generate global CSS after all slides are processed
        result.styles.global = genGlobalCSS(styleTable);

        // Step 5: Trigger other callbacks
        if (thumbnail && callbacks.onThumbnail) {
            callbacks.onThumbnail(thumbnail);
        }

        if (slideSize && callbacks.onSlideSize) {
            callbacks.onSlideSize(slideSize);
        }

        if (callbacks.onGlobalCSS) {
            callbacks.onGlobalCSS(result.styles.global);
        }

        // Step 6: Process message queue for charts
        processMsgQueue(msgQueue, result);

        if (callbacks.onComplete) {
            callbacks.onComplete({
                executionTime,
                slideWidth: slideSize?.width || 0,
                slideHeight: slideSize?.height || 0,
                styleTable,
                settings
            });
        }

        return result;
    }

    // Process the file data
    if (fileData) {
        return convertToHtml(fileData);
    }
    return null;
}

/**
 * PPTX to JSON converter
 * @param {ArrayBuffer} fileData - The PPTX file data
 * @param {Object} options - Conversion options
 * @returns {Promise<Object>} Parsed result
 */
async function pptxToJson(fileData, options) {
    // Merge default settings with user options
    const settings = {
        mediaProcess: true,
        themeProcess: true,
        incSlide: {
            width: 0,
            height: 0
        },
        styleTable: {},
        ...options
    };

    // Callback functions
    const callbacks = settings.callbacks || {};

    // State variables
    let defaultTextStyle = null;
    const chartId = { value: 0 };
    const styleTable = settings.styleTable;

    // Trigger file start callback
    if (callbacks.onFileStart) {
        callbacks.onFileStart();
    }

    /**
     * Convert PPTX file to JSON
     * @param {ArrayBuffer} file - The PPTX file data
     * @returns {Promise<Object>} Parsed result
     */
    async function convertToJson(file) {
        // Step 1: Parse PPTX to structured JSON data
        const { parsedData, msgQueue, slideSize, thumbnail, metadata, executionTime } = 
            await processToJson(file, settings, callbacks, chartId, styleTable, defaultTextStyle);

        // Step 2: Convert structured data to JSON result
        const result = {
            slides: [],
            slideSize,
            thumbnail,
            styles: {
                global: genGlobalCSS(styleTable)
            },
            metadata,
            charts: []
        };

        // Step 3: Process slides and keep as structured data
        for (const slideData of parsedData.slides) {
            result.slides.push({
                data: slideData.data,
                slideNum: slideData.slideNum,
                fileName: slideData.fileName
            });

            if (callbacks.onSlide) {
                callbacks.onSlide(slideData.data, {
                    slideNum: slideData.slideNum,
                    fileName: slideData.fileName
                });
            }
        }

        // Step 4: Trigger other callbacks
        if (thumbnail && callbacks.onThumbnail) {
            callbacks.onThumbnail(thumbnail);
        }

        if (slideSize && callbacks.onSlideSize) {
            callbacks.onSlideSize(slideSize);
        }

        if (callbacks.onGlobalCSS) {
            callbacks.onGlobalCSS(result.styles.global);
        }

        // Step 5: Process message queue for charts
        processMsgQueue(msgQueue, result);

        if (callbacks.onComplete) {
            callbacks.onComplete({
                executionTime,
                slideWidth: slideSize?.width || 0,
                slideHeight: slideSize?.height || 0,
                styleTable,
                settings
            });
        }

        return result;
    }

    // Process the file data
    if (fileData) {
        return convertToJson(fileData);
    }
    return null;
}

/**
 * PPTX to File Index and Content converter
 * @param {ArrayBuffer} fileData - The PPTX file data
 * @returns {Promise<Object>} File index and content result
 */
async function pptxToFiles(fileData) {
    if (fileData.byteLength < 10) {
        throw new Error("Invalid file: file too small");
    }

    const zip = JSZip.loadAsync ? await JSZip.loadAsync(fileData) : new JSZip().load(fileData);

    const result = {
        files: [],
        content: {}
    };

    // Iterate through all files in the zip
    const promises = [];
    zip.forEach((relativePath, zipEntry) => {
        result.files.push({
            name: relativePath,
            dir: zipEntry.dir,
            size: zipEntry._data.uncompressedSize
        });

        // Read file content based on type
        const promise = (async () => {
            try {
                if (zipEntry.dir) {
                    return;
                }

                const ext = relativePath.split('.').pop().toLowerCase();

                // For XML files, read as text
                if (ext === 'xml' || ext === 'rels') {
                    const content = await zipEntry.async('text');
                    result.content[relativePath] = {
                        type: 'text',
                        content: content
                    };
                }
                // For image files, read as base64
                else if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg'].includes(ext)) {
                    const base64 = await zipEntry.async('base64');
                    result.content[relativePath] = {
                        type: 'image',
                        format: ext,
                        base64: base64,
                        dataUrl: `data:image/${ext === 'jpg' ? 'jpeg' : ext};base64,${base64}`
                    };
                }
                // For other binary files, read as base64
                else {
                    const base64 = await zipEntry.async('base64');
                    result.content[relativePath] = {
                        type: 'binary',
                        base64: base64
                    };
                }
            } catch (error) {
                result.content[relativePath] = {
                    type: 'error',
                    error: error.message
                };
            }
        })();

        promises.push(promise);
    });

    await Promise.all(promises);

    return result;
}

// Export functions
/**
 * 提取幻灯片过渡效果数据
 * @param {Object} slideContent - 幻灯片内容
 * @returns {Object|null} 过渡效果数据或null
 */
function extractSlideTransition(slideContent) {
    // 检查slide中是否有transition元素
    const sld = slideContent["p:sld"];
    if (!sld) return null;
    
    const transition = PPTXXmlUtils.getTextByPathList(sld, ["p:transition"]);
    if (!transition) return null;
    
    // 解析过渡效果类型
    let transitionType = "fade"; // 默认淡入淡出
    let duration = 1000; // 默认1秒
    
    // 检查具体的过渡类型
    const transitionTypes = [
        "p:blinds", "p:checker", "p:circle", "p:comb", 
        "p:cover", "p:dissolve", "p:fade", "p:push",
        "p:random", "p:split", "p:strips", "p:wipe"
    ];
    
    for (const type of transitionTypes) {
        if (transition[type]) {
            transitionType = type.replace("p:", "");
            break;
        }
    }
    
    // 获取持续时间
    if (transition.attrs && transition.attrs["spd"]) {
        // spd值: slow=3, med=2, fast=1
        const speedMap = { "1": 500, "2": 1000, "3": 2000 };
        duration = speedMap[transition.attrs["spd"]] || 1000;
    }
    
    return {
        type: transitionType,
        duration: duration
    };
}

export { pptxToHtml as default, pptxToFiles, pptxToHtml, pptxToJson };
//# sourceMappingURL=ppt-parser.esm.js.map
