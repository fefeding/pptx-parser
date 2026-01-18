/**
 * PPTX 核心工具函数
 * 提供坐标转换、文件处理、路径解析等基础功能
 */

// ============================================================================
// 核心工具函数 (Core Utils)
// ============================================================================

// PPTX 坐标转换因子 (96 pixels per inch / 914400 EMUs per inch)
const slideFactor = 96 / 914400;

/**
 * 角度转度数（EMU 单位 1/60000 度 -> 标准度）
 * @param {string|number} angle - EMU角度值
 * @returns {number} 标准角度值
 */
function angleToDegrees(angle) {
    if (!angle || angle === "") return 0;
    return Math.round(angle / 60000);
}

/**
 * 度转弧度
 * @param {number} degrees - 角度值
 * @returns {number} 弧度值
 */
function degreesToRadians(degrees) {
    return degrees * (Math.PI / 180);
}

/**
 * ArrayBuffer 转 Base64（用于缩略图）
 * @param {ArrayBuffer} arrayBuffer - 二进制数据
 * @returns {string} Base64字符串
 */
function base64ArrayBuffer(arrayBuffer) {
    const encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
    const bytes = new Uint8Array(arrayBuffer);
    const byteLength = bytes.byteLength;
    const byteRemainder = byteLength % 3;
    const mainLength = byteLength - byteRemainder;
    
    let base64 = '';
    
    // 处理3字节一组的数据
    for (let i = 0; i < mainLength; i += 3) {
        const chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
        const a = (chunk & 16515072) >> 18;
        const b = (chunk & 258048) >> 12;
        const c = (chunk & 4032) >> 6;
        const d = chunk & 63;
        base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d];
    }
    
    // 处理剩余字节
    if (byteRemainder === 1) {
        let chunk = bytes[mainLength];
        const a = (chunk & 252) >> 2;
        const b = (chunk & 3) << 4;
        base64 += encodings[a] + encodings[b] + '==';
    } else if (byteRemainder === 2) {
        let chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];
        const a = (chunk & 64512) >> 10;
        const b = (chunk & 1008) >> 4;
        const c = (chunk & 15) << 2;
        base64 += encodings[a] + encodings[b] + encodings[c] + '=';
    }
    
    return base64;
}

/**
 * Base64 或 ArrayBuffer 转 Blob URL
 * @param {ArrayBuffer} arrayBuffer - 二进制数据
 * @param {string} mimeType - MIME类型
 * @returns {string} Blob URL
 */
function arrayBufferToBlobUrl(arrayBuffer, mimeType) {
    const blob = new Blob([arrayBuffer], { type: mimeType });
    return URL.createObjectURL(blob);
}

/**
 * Base64 字符串转 Blob URL
 * @param {string} base64Data - Base64数据
 * @param {string} mimeType - MIME类型
 * @returns {string} Blob URL
 */
function base64ToBlobUrl(base64Data, mimeType) {
    const binaryString = window.atob(base64Data);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: mimeType });
    return URL.createObjectURL(blob);
}

/**
 * 根据文件扩展名获取 MIME 类型
 * @param {string} imgFileExt - 文件扩展名
 * @returns {string} MIME类型
 */
function getMimeType(imgFileExt) {
    const mimeMap = {
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'png': 'image/png',
        'gif': 'image/gif',
        'emf': 'image/x-emf',
        'wmf': 'image/x-wmf',
        'svg': 'image/svg+xml',
        'mp4': 'video/mp4',
        'webm': 'video/webm',
        'ogg': 'video/ogg',
        'avi': 'video/avi',
        'mpg': 'video/mpg',
        'wmv': 'video/wmv',
        'mp3': 'audio/mpeg',
        'wav': 'audio/wav',
        'tif': 'image/tiff',
        'tiff': 'image/tiff'
    };
    return mimeMap[imgFileExt.toLowerCase()] || '';
}

/**
 * 判断是否为视频链接
 * @param {string} vdoFile - 视频文件路径或URL
 * @returns {boolean} 是否为视频链接
 */
function IsVideoLink(vdoFile) {
    const urlRegex = /^(https?|ftp):\/\/([a-zA-Z0-9.-]+(:[a-zA-Z0-9.&%$-]+)*@)*((25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]?)(\.(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])){3}|([a-zA-Z0-9-]+\.)*[a-zA-Z0-9-]+\.(com|edu|gov|int|mil|net|org|biz|arpa|info|name|pro|aero|coop|museum|[a-zA-Z]{2}))(:[0-9]+)*(\/($|[a-zA-Z0-9.,?'\\+&%$#=~_-]+))*$/;
    return urlRegex.test(vdoFile);
}

/**
 * 解析相对路径
 * @param {string} basePath - 基础路径
 * @param {string} relativePath - 相对路径
 * @returns {string} 解析后的路径
 */
function resolvePath(basePath, relativePath) {
    if (relativePath.startsWith("ppt/") || relativePath.startsWith("[Content_Types].xml") || relativePath.startsWith("docProps/")) {
        return relativePath;
    }
    
    const baseDir = basePath.substring(0, basePath.lastIndexOf("/") + 1);
    const parts = relativePath.split("/");
    const resultParts = baseDir.split("/").filter(part => part !== "");
    
    for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        if (part === "..") {
            if (resultParts.length > 0) {
                resultParts.pop();
            }
        } else if (part === "." || part === "") {
            continue;
        } else {
            resultParts.push(part);
        }
    }
    
    return resultParts.join("/");
}

/**
 * 解析关系文件目标路径
 * @param {string} relFilePath - 关系文件路径
 * @param {string} target - 目标路径
 * @returns {string} 解析后的目标路径
 */
function resolveRelationshipTarget(relFilePath, target) {
    const basePath = relFilePath;
    if (basePath.indexOf("/_rels/") !== -1) {
        return resolvePath(basePath.substring(0, basePath.indexOf("/_rels/")) + "/", target);
    }
    return resolvePath(basePath, target);
}

/**
 * 提取文件扩展名
 * @param {string} filename - 文件名
 * @returns {string} 扩展名
 */
function extractFileExtension(filename) {
    return filename.substr((~-filename.lastIndexOf(".") >>> 0) + 2);
}

/**
 * 转义 HTML 特殊字符
 * @param {string} text - 文本内容
 * @returns {string} 转义后的文本
 */
function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}

/**
 * 通过路径列表获取节点值
 * @param {Object} node - 节点对象
 * @param {Array<string>} path - 路径数组
 * @returns {any} 节点值
 */
function getTextByPathList(node, path) {
    if (!Array.isArray(path)) {
        throw new Error("Error of path type! path is not array.");
    }

    if (node === undefined || node === null) {
        return undefined;
    }

    let currentNode = node;
    for (let i = 0; i < path.length; i++) {
        currentNode = currentNode[path[i]];
        if (currentNode === undefined || currentNode === null) {
            return undefined;
        }
    }

    return currentNode;
}

/**
 * 通过路径列表设置节点值
 * @param {Object} node - 节点对象
 * @param {Array<string>} path - 路径数组
 * @param {any} value - 要设置的值
 * @returns {any} 设置后的节点
 */
function setTextByPathList(node, path, value) {
    if (!Array.isArray(path)) {
        throw new Error("Error of path type! path is not array.");
    }

    if (node === undefined) {
        return undefined;
    }

    // 为对象添加set方法
    Object.prototype.set = function(parts, value) {
        let obj = this;
        const lent = parts.length;
        for (let i = 0; i < lent; i++) {
            const p = parts[i];
            if (obj[p] === undefined) {
                if (i === lent - 1) {
                    obj[p] = value;
                } else {
                    obj[p] = {};
                }
            }
            obj = obj[p];
        }
        return obj;
    };

    return node.set(path, value);
}

/**
 * 遍历数组或对象
 * @param {Array|Object} node - 节点
 * @param {Function} doFunction - 处理函数
 * @returns {string} 处理结果
 */
function eachElement(node, doFunction) {
    if (node === undefined) {
        return '';
    }

    if (Array.isArray(node)) {
        return node.map((item, index) => doFunction(item, index)).join('');
    }
    
    return doFunction(node, 0);
}

/**
 * 计算元素位置 CSS（top, left）- 支持从多个来源获取位置
 * @param {Object} xfrmNode - 变换节点
 * @param {Object} pNode - 父节点
 * @param {Object} parentOff - 父偏移
 * @param {Object} parentExt - 父扩展
 * @param {string} sType - 类型
 * @returns {string} CSS位置字符串
 */
function getPosition(xfrmNode, pNode, parentOff, parentExt, sType) {
    // 简单版本：只处理 xfrmNode
    if (xfrmNode && arguments.length <= 5) {
        if (!xfrmNode) return '';
        
        const offNode = xfrmNode['a:off'];
        if (!offNode || !offNode['attrs']) return '';
        
        const offAttrs = offNode['attrs'];
        const x = parseInt(offAttrs['x']) * slideFactor;
        const y = parseInt(offAttrs['y']) * slideFactor;
        if (isNaN(x) || isNaN(y)) return '';

        let css = '';
        if (sType === 'group-rotate') {
            css += `top: ${y}px; left: ${x}px;`;
        } else {
            const chOff = xfrmNode['a:chOff'];
            const chExt = xfrmNode['a:chExt'];
            if (chOff && chExt && chOff['attrs'] && chExt['attrs']) {
                const chx = parseInt(chOff['attrs']['x']) * slideFactor;
                const chy = parseInt(chOff['attrs']['y']) * slideFactor;
                if (!isNaN(chx) && !isNaN(chy)) {
                    css += `top: ${y - chy}px; left: ${x - chx}px;`;
                } else {
                    css += `top: ${y}px; left: ${x}px;`;
                }
            } else {
                css += `top: ${y}px; left: ${x}px;`;
            }
        }
        return css;
    }

    // 复杂版本：支持从多个来源获取位置
    let off;
    let x = -1, y = -1;

    if (xfrmNode?.['a:off']?.['attrs']) {
        off = xfrmNode['a:off']['attrs'];
    } else if (parentOff?.['a:off']?.['attrs']) {
        off = parentOff['a:off']['attrs'];
    } else if (parentExt?.['a:off']?.['attrs']) {
        off = parentExt['a:off']['attrs'];
    }

    let offX = 0, offY = 0;
    let grpX = 0, grpY = 0;
    let offAttrs, chx, chy;
    
    if (sType === 'group') {
        const grpXfrmNode = getTextByPathList(pNode, ['p:grpSpPr', 'a:xfrm']);
        if (grpXfrmNode?.['a:off']?.['attrs']) {
offAttrs = grpXfrmNode['a:off']['attrs'];
            const tmpX = parseInt(offAttrs['x']) * slideFactor;
            const tmpY = parseInt(offAttrs['y']) * slideFactor;
            if (!isNaN(tmpX) && !isNaN(tmpY)) {
                grpX = tmpX;
                grpY = tmpY;
            }
        }
    }
    
    if (sType === 'group-rotate' && pNode?.['p:grpSpPr'] !== undefined) {
        const grpXfrmNode2 = pNode['p:grpSpPr']['a:xfrm'];
        if (grpXfrmNode2?.['a:chOff']?.['attrs']) {
            const chAttrs = grpXfrmNode2['a:chOff']['attrs'];
chx = parseInt(chAttrs['x']) * slideFactor;
chy = parseInt(chAttrs['y']) * slideFactor;
            if (!isNaN(chx) && !isNaN(chy)) {
                offX = chx;
                offY = chy;
            }
        }
    }
    
    if (off === undefined) {
        return '';
    }
    
    x = parseInt(off['x']) * slideFactor;
    y = parseInt(off['y']) * slideFactor;
    return (isNaN(x) || isNaN(y)) ? '' : `top:${y - offY + grpY}px; left:${x - offX + grpX}px;`;
}

/**
 * 计算元素尺寸 CSS（width, height）- 支持从多个来源获取尺寸
 * @param {Object} xfrmNode - 变换节点
 * @param {Object} parentExt - 父扩展
 * @param {string} sType - 类型
 * @returns {string} CSS尺寸字符串
 */
function getSize(xfrmNode, parentExt, sType) {
    // 简单版本：只处理 xfrmNode
    if (xfrmNode && arguments.length <= 3) {
        if (!xfrmNode) return '';
        
        const ext = xfrmNode['a:ext'];
        if (!ext || !ext['attrs']) return '';

        const attrs = ext['attrs'];
        const w = parseInt(attrs['cx']) * slideFactor;
        const h = parseInt(attrs['cy']) * slideFactor;
        if (isNaN(w) || isNaN(h)) return '';

        return `width: ${w}px; height: ${h}px;`;
    }

    // 复杂版本：支持从多个来源获取尺寸
    let ext;
    let w = -1, h = -1;

    if (xfrmNode?.['a:ext']?.['attrs']) {
        ext = xfrmNode['a:ext']['attrs'];
    } else if (parentExt?.['a:ext']?.['attrs']) {
        ext = parentExt['a:ext']['attrs'];
    } else if (sType?.['a:ext']?.['attrs']) {
        ext = sType['a:ext']['attrs'];
    }

    if (ext === undefined) {
        return '';
    }
    
    w = parseInt(ext['cx']) * slideFactor;
    h = parseInt(ext['cy']) * slideFactor;
    return (isNaN(w) || isNaN(h)) ? '' : `width:${w}px; height:${h}px;`;
}

/**
 * 古老数字格式化（如希伯来数字）
 * @param {Array<Array>} arr - 数字映射数组
 * @returns {Object} 格式化对象
 */
function archaicNumbers(arr) {
    const arrParse = arr.slice().sort((a, b) => b[1].length - a[1].length);
    return {
        format: function (n) {
            let ret = '';
            for (let i = 0; i < arr.length; i++) {
                const item = arr[i];
                const num = item[0];
                if (parseInt(num) > 0) {
                    for (; n >= num; n -= num) ret += item[1];
                } else {
                    ret = ret.replace(num, item[1]);
                }
            }
            return ret;
        }
    };
}

/**
 * 数字编号格式化
 * @param {string} numTyp - 编号类型
 * @param {number} num - 数字
 * @returns {string} 格式化后的编号
 */
function getNumTypeNum(numTyp, num) {
    switch (numTyp) {
        case 'arabicPeriod':
            return `${num}. `;
        case 'arabicParenR':
            return `${num}) `;
        case 'alphaLcParenR':
            return `${alphaNumeric(num, 'lowerCase')}) `;
        case 'alphaLcPeriod':
            return `${alphaNumeric(num, 'lowerCase')}. `;
        case 'alphaUcParenR':
            return `${alphaNumeric(num, 'upperCase')}) `;
        case 'alphaUcPeriod':
            return `${alphaNumeric(num, 'upperCase')}. `;
        case 'romanUcPeriod':
            return `${romanize(num)}. `;
        case 'romanLcParenR':
            return `${romanize(num)}) `;
        case 'hebrew2Minus':
            return `${hebrew2Minus.format(num)}-`;
        default:
            return String(num);
    }
}

/**
 * 罗马数字转换
 * @param {number} num - 数字
 * @returns {string} 罗马数字
 */
function romanize(num) {
    if (!num || num <= 0) return '';
    
    const digits = String(+num).split('');
    const key = [
        "", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM",
        "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC",
        "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"
    ];
    let roman = '';
    let i = 3;
    
    while (i--) {
        roman = (key[+digits.pop() + (i * 10)] || "") + roman;
    }
    
    return 'M'.repeat(+digits.join("") + 1) + roman;
}

/**
 * 字母数字编号生成 (a, b, c, ... or A, B, C, ...)
 * @param {number} num - 数字
 * @param {string} upperLower - 大小写 ('upperCase' | 'lowerCase')
 * @returns {string} 字母编号
 */
function alphaNumeric(num, upperLower) {
    const n = Number(num) - 1;
    const aNum = ((n / 26 >= 1) ? String.fromCharCode(Math.floor(n / 26) + 64) : '') + 
                 String.fromCharCode((n % 26) + 65);
    return upperLower === 'upperCase' ? aNum.toUpperCase() : aNum.toLowerCase();
}

// 希伯来数字编号器
const hebrew2Minus = archaicNumbers([
    [1000, ''], [400, 'ת'], [300, 'ש'], [200, 'ר'], [100, 'ק'],
    [90, 'צ'], [80, 'פ'], [70, 'ע'], [60, 'ס'], [50, 'נ'],
    [40, 'מ'], [30, 'ל'], [20, 'כ'], [10, 'י'], [9, 'ט'],
    [8, 'ח'], [7, 'ז'], [6, 'ו'], [5, 'ה'], [4, 'ד'],
    [3, 'ג'], [2, 'ב'], [1, 'א'],
    [/יה/, 'ט״ו'], [/יו/, 'ט״ז'],
    [/([א-ת])([א-ת])$/, '$1״$2'],
    [/^([א-ת])$/, "$1׳"]
]);

// 公开工具函数（保持向后兼容）
const PPTXUtils = {
    angleToDegrees,
    degreesToRadians,
    getMimeType,
    base64ArrayBuffer,
    arrayBufferToBlobUrl,
    base64ToBlobUrl,
    IsVideoLink,
    resolvePath,
    resolveRelationshipTarget,
    extractFileExtension,
    escapeHtml,
    getTextByPathList,
    setTextByPathList,
    eachElement,
    archaicNumbers,
    getNumTypeNum,
    romanize,
    alphaNumeric,
    hebrew2Minus,
    getPosition,
    getSize,
    getSlideFactor: () => slideFactor,
    setSlideFactor: (factor) => { slideFactor = factor; }
};

export { PPTXUtils };

// ============================================================================
// 文件读写工具函数 (File Reader Utils)
// ============================================================================

/**
 * 读取 File 或 Blob 对象为 ArrayBuffer
 * @param {File|Blob} file - File 或 Blob 对象
 * @param {Function} onLoad - 加载成功回调
 * @param {Function} onError - 加载失败回调
 */
export function readAsArrayBuffer(file, onLoad, onError) {
    const reader = new FileReader();
    reader.onload = (event) => {
        if (onLoad) onLoad(event.target.result);
    };
    reader.onerror = (event) => {
        if (onError) onError(new Error(`Failed to read file: ${event.target.error}`));
    };
    reader.readAsArrayBuffer(file);
}

/**
 * 读取 File 或 Blob 对象为 Text
 * @param {File|Blob} file - File 或 Blob 对象
 * @param {Function} onLoad - 加载成功回调
 * @param {Function} onError - 加载失败回调
 */
export function readAsText(file, onLoad, onError) {
reader = new FileReader();
    reader.onload = (event) => {
        if (onLoad) onLoad(event.target.result);
    };
    reader.onerror = (event) => {
        if (onError) onError(new Error(`Failed to read file: ${event.target.error}`));
    };
    reader.readAsText(file);
}

/**
 * 读取 File 或 Blob 对象为 DataURL
 * @param {File|Blob} file - File 或 Blob 对象
 * @param {Function} onLoad - 加载成功回调
 * @param {Function} onError - 加载失败回调
 */
export function readAsDataURL(file, onLoad, onError) {
reader = new FileReader();
    reader.onload = (event) => {
        if (onLoad) onLoad(event.target.result);
    };
    reader.onerror = (event) => {
        if (onError) onError(new Error(`Failed to read file: ${event.target.error}`));
    };
    reader.readAsDataURL(file);
}

/**
 * 读取 File 或 Blob 对象为 ArrayBuffer (Promise版本)
 * @param {File|Blob} file - File 或 Blob 对象
 * @returns {Promise<ArrayBuffer>} Promise对象
 */
export function readAsArrayBufferAsync(file) {
    return new Promise((resolve, reject) => {
        readAsArrayBuffer(file, resolve, reject);
    });
}

/**
 * 读取 File 或 Blob 对象为 Text (Promise版本)
 * @param {File|Blob} file - File 或 Blob 对象
 * @returns {Promise<string>} Promise对象
 */
export function readAsTextAsync(file) {
    return new Promise((resolve, reject) => {
        readAsText(file, resolve, reject);
    });
}

/**
 * 读取 File 或 Blob 对象为 DataURL (Promise版本)
 * @param {File|Blob} file - File 或 Blob 对象
 * @returns {Promise<string>} Promise对象
 */
export function readAsDataURLAsync(file) {
    return new Promise((resolve, reject) => {
        readAsDataURL(file, resolve, reject);
    });
}

/**
 * File Reader 工具对象 (兼容接口)
 */
export const PPTXFileReader = {
    setupBlob(file, options) {
        if (!file) return;
        const onCallbacks = options?.on || {};
        readAsArrayBuffer(file, onCallbacks.load, onCallbacks.error);
    },
    
    setupBlobAsText(file, options) {
        if (!file) return;
onCallbacks = options?.on || {};
        readAsText(file, onCallbacks.load, onCallbacks.error);
    },
    
    setSync(value) {
    },
    
    getSync() {
        return false;
    }
};
