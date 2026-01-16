
    // PPTX 坐标转换因子 (96 pixels per inch / 914400 EMUs per inch)
var slideFactor = 96 / 914400;


    // 角度转度数（EMU 单位 1/60000 度 -> 标准度）
function angleToDegrees(angle) {
    if (angle == "" || angle == null) return 0;
    return Math.round(angle / 60000);
}

    // 度转弧度
function degreesToRadians(degrees) {
    return degrees * (Math.PI / 180);
}

    // ArrayBuffer 转 Base64（用于缩略图）
function base64ArrayBuffer(arrayBuffer) {
    var base64    = '';
    var encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

    var bytes         = new Uint8Array(arrayBuffer);
    var byteLength    = bytes.byteLength;
    var byteRemainder = byteLength % 3;
    var mainLength    = byteLength - byteRemainder;

    var a, b, c, d;
    var chunk;

    for (var i = 0; i < mainLength; i = i + 3) {
        chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
        a = (chunk & 16515072) >> 18;
        b = (chunk & 258048)   >> 12;
        c = (chunk & 4032)     >> 6;
        d = chunk & 63;
        base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d];
    }

    if (byteRemainder == 1) {
        chunk = bytes[mainLength];
        a = (chunk & 252) >> 2;
        b = (chunk & 3)   << 4;
        base64 += encodings[a] + encodings[b] + '==';
    } else if (byteRemainder == 2) {
        chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];
        a = (chunk & 64512) >> 10;
        b = (chunk & 1008)  >> 4;
        c = (chunk & 15)    << 2;
        base64 += encodings[a] + encodings[b] + encodings[c] + '=';
    }
    return base64;
}

    // Base64 或 ArrayBuffer 转 Blob URL
function arrayBufferToBlobUrl(arrayBuffer, mimeType) {
    var blob = new Blob([arrayBuffer], { type: mimeType });
    return URL.createObjectURL(blob);
}

    // Base64 字符串转 Blob URL
function base64ToBlobUrl(base64Data, mimeType) {
    var binaryString = window.atob(base64Data);
    var bytes = new Uint8Array(binaryString.length);
    for (var i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    var blob = new Blob([bytes], { type: mimeType });
    return URL.createObjectURL(blob);
}

    // 获取 MIME 类型
function getMimeType(imgFileExt) {
    var mimeType = "";
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
        case "emf":
            mimeType = "image/x-emf";
            break;
        case "wmf":
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
        case "tif":
        case "tiff":
            mimeType = "image/tiff";
            break;
    }
    return mimeType;
}

    // 判断是否为视频链接
function IsVideoLink(vdoFile) {
    var urlregex = /^(https?|ftp):\/\/([a-zA-Z0-9.-]+(:[a-zA-Z0-9.&%$-]+)*@)*((25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]?)(\.(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])){3}|([a-zA-Z0-9-]+\.)*[a-zA-Z0-9-]+\.(com|edu|gov|int|mil|net|org|biz|arpa|info|name|pro|aero|coop|museum|[a-zA-Z]{2}))(:[0-9]+)*(\/($|[a-zA-Z0-9.,?'\\+&%$#=~_-]+))*$/;
    return urlregex.test(vdoFile);
}

    // 解析相对路径
function resolvePath(basePath, relativePath) {
    if (relativePath.startsWith("ppt/") || relativePath.startsWith("[Content_Types].xml") || relativePath.startsWith("docProps/")) {
        return relativePath;
    }
    
    var baseDir = basePath.substring(0, basePath.lastIndexOf("/") + 1);
    
    var parts = relativePath.split("/");
    var resultParts = baseDir.split("/").filter(function(part) {
        return part !== "";
    });
    
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i];
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

    // 解析关系文件目标路径
function resolveRelationshipTarget(relFilePath, target) {
    var basePath = relFilePath;
    if (basePath.indexOf("/_rels/") !== -1) {
        basePath = basePath.substring(0, basePath.indexOf("/_rels/")) + "/";
    }
    return resolvePath(basePath, target);
}

    // 提取文件扩展名
function extractFileExtension(filename) {
    return filename.substr((~-filename.lastIndexOf(".") >>> 0) + 2);
}

    // 转义 HTML 特殊字符
function escapeHtml(text) {
    var map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, function (m) { return map[m]; });
}

    // 通过路径列表获取节点值
function getTextByPathList(node, path) {
    if (path.constructor !== Array) {
        throw Error("Error of path type! path is not array.");
    }

    if (node === undefined || node === null) {
        return undefined;
    }

    var l = path.length;
    for (var i = 0; i < l; i++) {
        node = node[path[i]];
        if (node === undefined || node === null) {
            return undefined;
        }
    }

    return node;
}

    // 通过路径列表设置节点值
function setTextByPathList(node, path, value) {
    if (path.constructor !== Array) {
        throw Error("Error of path type! path is not array.");
    }

    if (node === undefined) {
        return undefined;
    }

    Object.prototype.set = function (parts, value) {
        var obj = this;
        var lent = parts.length;
        for (var i = 0; i < lent; i++) {
            var p = parts[i];
            if (obj[p] === undefined) {
                if (i == lent - 1) {
                    obj[p] = value;
                } else {
                    obj[p] = {};
                }
            }
            obj = obj[p];
        }
        return obj;
    };

    node.set(path, value);
}

    // 遍历数组或对象
function eachElement(node, doFunction) {
    if (node === undefined) {
        return;
    }
    var result = "";
    if (node.constructor === Array) {
        var l = node.length;
        for (var i = 0; i < l; i++) {
            result += doFunction(node[i], i);
        }
    } else {
        result += doFunction(node, 0);
    }
    return result;
}

    // 计算元素位置 CSS（top, left）- 支持从多个来源获取位置
function getPosition(xfrmNode, pNode, parentOff, parentExt, sType) {
    // 简单版本：只处理 xfrmNode
    if (xfrmNode && arguments.length <= 5) {
        if (!xfrmNode) return "";
        
        // 检查必需的属性
        var offNode = xfrmNode["a:off"];
        if (!offNode || !offNode["attrs"]) return "";
        var offAttrs = offNode["attrs"];
        var x = parseInt(offAttrs["x"]) * slideFactor;
        var y = parseInt(offAttrs["y"]) * slideFactor;
        if (isNaN(x) || isNaN(y)) return "";

        var css = "";
        if (sType === "group-rotate") {
            css += "top: " + y + "px; left: " + x + "px;";
        } else {
            var chOff = xfrmNode["a:chOff"];
            var chExt = xfrmNode["a:chExt"];
            if (chOff && chExt && chOff["attrs"] && chExt["attrs"]) {
                var chx = parseInt(chOff["attrs"]["x"]) * slideFactor;
                var chy = parseInt(chOff["attrs"]["y"]) * slideFactor;
                if (!isNaN(chx) && !isNaN(chy)) {
                    css += "top: " + (y - chy) + "px; left: " + (x - chx) + "px;";
                } else {
                    css += "top: " + y + "px; left: " + x + "px;";
                }
            } else {
                css += "top: " + y + "px; left: " + x + "px;";
            }
        }
        return css;
    }

    // 复杂版本：支持从多个来源获取位置（slideSpNode, slideLayoutSpNode, slideMasterSpNode）
    var off;
    var x = -1, y = -1;

    if (xfrmNode !== undefined && xfrmNode["a:off"] && xfrmNode["a:off"]["attrs"]) {
        off = xfrmNode["a:off"]["attrs"];
    }

    if (off === undefined && parentOff !== undefined && parentOff["a:off"] && parentOff["a:off"]["attrs"]) {
        off = parentOff["a:off"]["attrs"];
    } else if (off === undefined && parentExt !== undefined && parentExt["a:off"] && parentExt["a:off"]["attrs"]) {
        off = parentExt["a:off"]["attrs"];
    }
    var offX = 0, offY = 0;
    var grpX = 0, grpY = 0;
    if (sType == "group") {

        var grpXfrmNode = getTextByPathList(pNode, ["p:grpSpPr", "a:xfrm"]);
        if (grpXfrmNode !== undefined && grpXfrmNode["a:off"] && grpXfrmNode["a:off"]["attrs"]) {
            var offAttrs = grpXfrmNode["a:off"]["attrs"];
            var tmpX = parseInt(offAttrs["x"]) * slideFactor;
            var tmpY = parseInt(offAttrs["y"]) * slideFactor;
            if (!isNaN(tmpX) && !isNaN(tmpY)) {
                grpX = tmpX;
                grpY = tmpY;
            }
        }
    }
    if (sType == "group-rotate" && pNode && pNode["p:grpSpPr"] !== undefined) {
        var grpXfrmNode2 = pNode["p:grpSpPr"]["a:xfrm"];
        if (grpXfrmNode2 && grpXfrmNode2["a:chOff"] && grpXfrmNode2["a:chOff"]["attrs"]) {
            var chAttrs = grpXfrmNode2["a:chOff"]["attrs"];
            var chx = parseInt(chAttrs["x"]) * slideFactor;
            var chy = parseInt(chAttrs["y"]) * slideFactor;
            if (!isNaN(chx) && !isNaN(chy)) {
                offX = chx;
                offY = chy;
            }
        }
    }
    if (off === undefined) {
        return "";
    } else {
        x = parseInt(off["x"]) * slideFactor;
        y = parseInt(off["y"]) * slideFactor;
        return (isNaN(x) || isNaN(y)) ? "" : "top:" + (y - offY + grpY) + "px; left:" + (x - offX + grpX) + "px;";
    }
}

    // 计算元素尺寸 CSS（width, height）- 支持从多个来源获取尺寸
function getSize(xfrmNode, parentExt, sType) {
    // 简单版本：只处理 xfrmNode
    if (xfrmNode && arguments.length <= 3) {
        if (!xfrmNode) return "";

        var ext = xfrmNode["a:ext"];
        if (!ext || !ext["attrs"]) return "";

        var attrs = ext["attrs"];
        var w = parseInt(attrs["cx"]) * slideFactor;
        var h = parseInt(attrs["cy"]) * slideFactor;
        if (isNaN(w) || isNaN(h)) return "";

        return "width: " + w + "px; height: " + h + "px;";
    }

    // 复杂版本：支持从多个来源获取尺寸（xfrmNode, slideLayoutSpNode, slideMasterSpNode）
    var ext = undefined;
    var w = -1, h = -1;

    if (xfrmNode !== undefined && xfrmNode["a:ext"] && xfrmNode["a:ext"]["attrs"]) {
        ext = xfrmNode["a:ext"]["attrs"];
    } else if (parentExt !== undefined && parentExt["a:ext"] && parentExt["a:ext"]["attrs"]) {
        ext = parentExt["a:ext"]["attrs"];
    } else if (sType !== undefined && sType["a:ext"] && sType["a:ext"]["attrs"]) {
        ext = sType["a:ext"]["attrs"];
    }

    if (ext === undefined) {
        return "";
    } else {
        w = parseInt(ext["cx"]) * slideFactor;
        h = parseInt(ext["cy"]) * slideFactor;
        return (isNaN(w) || isNaN(h)) ? "" : "width:" + w + "px; height:" + h + "px;";
    }
}



    // 古老数字格式化（如希伯来数字）
function archaicNumbers(arr) {
    var arrParse = arr.slice().sort(function (a, b) { return b[1].length - a[1].length });
    return {
        format: function (n) {
            var ret = '';
            for (var i = 0; i < arr.length; i++) {
                var item = arr[i];
                var num = item[0];
                if (parseInt(num) > 0) {
                    for (; n >= num; n -= num) ret += item[1];
                } else {
                    ret = ret.replace(num, item[1]);
                }
            }
            return ret;
        }
    }
}

    // 数字编号格式化
function getNumTypeNum(numTyp, num) {
    var rtrnNum = "";
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
            rtrnNum = hebrew2Minus.format(num) + "-";
            break;
        default:
            rtrnNum = num;
    }
    return rtrnNum;
}

    // 罗马数字转换
function romanize(num) {
    if (!+num)
        return false;
    var digits = String(+num).split(""),
        key = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM",
            "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC",
            "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"],
        roman = "",
        i = 3;
    while (i--)
        roman = (key[+digits.pop() + (i * 10)] || "") + roman;
    return Array(+digits.join("") + 1).join("M") + roman;
}

    // 字母数字编号生成 (a, b, c, ... or A, B, C, ...)
function alphaNumeric(num, upperLower) {
    num = Number(num) - 1;
    var aNum = "";
    if (upperLower == "upperCase") {
        aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toUpperCase();
    } else if (upperLower == "lowerCase") {
        aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toLowerCase();
    }
    return aNum;
}

    // 希伯来数字编号器
var hebrew2Minus = archaicNumbers([
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


    // 公开工具函数
const PPTXUtils = {
    angleToDegrees: angleToDegrees,
    degreesToRadians: degreesToRadians,
    getMimeType: getMimeType,
    base64ArrayBuffer: base64ArrayBuffer,
    arrayBufferToBlobUrl: arrayBufferToBlobUrl,
    base64ToBlobUrl: base64ToBlobUrl,
    IsVideoLink: IsVideoLink,
    resolvePath: resolvePath,
    resolveRelationshipTarget: resolveRelationshipTarget,
    extractFileExtension: extractFileExtension,
    escapeHtml: escapeHtml,
    getTextByPathList: getTextByPathList,
    setTextByPathList: setTextByPathList,
    eachElement: eachElement,
    archaicNumbers: archaicNumbers,
    getNumTypeNum: getNumTypeNum,
    romanize: romanize,
    alphaNumeric: alphaNumeric,
    hebrew2Minus: hebrew2Minus,
    getPosition: getPosition,
    getSize: getSize,
    getSlideFactor: function() { return slideFactor; },
    setSlideFactor: function(factor) { slideFactor = factor; }
};


export { PPTXUtils };

// ============================================================================
// 文件读写工具函数 (File Reader Utils)
// ============================================================================

/**
 * 读取 File 或 Blob 对象为 ArrayBuffer
 * @param {File|Blob} file - File 或 Blob 对象
 * @param {Function} onLoad - 加载成功回调，参数为 ArrayBuffer
 * @param {Function} onError - 加载失败回调，参数为 Error 对象
 */
export function readAsArrayBuffer(file, onLoad, onError) {
    var reader = new FileReader();
    reader.onload = function(event) {
        if (onLoad) {
            onLoad(event.target.result);
        }
    };
    reader.onerror = function(event) {
        if (onError) {
            onError(new Error("Failed to read file: " + event.target.error));
        }
    };
    reader.readAsArrayBuffer(file);
}

/**
 * 读取 File 或 Blob 对象为 Text
 * @param {File|Blob} file - File 或 Blob 对象
 * @param {Function} onLoad - 加载成功回调，参数为文本内容
 * @param {Function} onError - 加载失败回调，参数为 Error 对象
 */
export function readAsText(file, onLoad, onError) {
    var reader = new FileReader();
    reader.onload = function(event) {
        if (onLoad) {
            onLoad(event.target.result);
        }
    };
    reader.onerror = function(event) {
        if (onError) {
            onError(new Error("Failed to read file: " + event.target.error));
        }
    };
    reader.readAsText(file);
}

/**
 * 读取 File 或 Blob 对象为 DataURL (Base64)
 * @param {File|Blob} file - File 或 Blob 对象
 * @param {Function} onLoad - 加载成功回调，参数为 DataURL 字符串
 * @param {Function} onError - 加载失败回调，参数为 Error 对象
 */
export function readAsDataURL(file, onLoad, onError) {
    var reader = new FileReader();
    reader.onload = function(event) {
        if (onLoad) {
            onLoad(event.target.result);
        }
    };
    reader.onerror = function(event) {
        if (onError) {
            onError(new Error("Failed to read file: " + event.target.error));
        }
    };
    reader.readAsDataURL(file);
}

/**
 * 读取 File 或 Blob 对象为 ArrayBuffer (Promise 版本)
 * @param {File|Blob} file - File 或 Blob 对象
 * @returns {Promise<ArrayBuffer>} Promise 对象
 */
export function readAsArrayBufferAsync(file) {
    return new Promise(function(resolve, reject) {
        readAsArrayBuffer(file, resolve, reject);
    });
}

/**
 * 读取 File 或 Blob 对象为 Text (Promise 版本)
 * @param {File|Blob} file - File 或 Blob 对象
 * @returns {Promise<string>} Promise 对象
 */
export function readAsTextAsync(file) {
    return new Promise(function(resolve, reject) {
        readAsText(file, resolve, reject);
    });
}

/**
 * 读取 File 或 Blob 对象为 DataURL (Promise 版本)
 * @param {File|Blob} file - File 或 Blob 对象
 * @returns {Promise<string>} Promise 对象
 */
export function readAsDataURLAsync(file) {
    return new Promise(function(resolve, reject) {
        readAsDataURL(file, resolve, reject);
    });
}

/**
 * File Reader 工具对象 (兼容旧的 FileReaderJS 接口)
 */
export var PPTXFileReader = {
    /**
     * 设置文件读取为 ArrayBuffer
     * @param {File|Blob} file - 文件对象
     * @param {Object} options - 配置选项
     * @param {Function} options.on.load - 加载成功回调
     * @param {Function} options.on.error - 加载失败回调
     */
    setupBlob: function(file, options) {
        if (!file) return;

        options = options || {};
        var onCallbacks = options.on || {};

        readAsArrayBuffer(file, onCallbacks.load, onCallbacks.error);
    },

    /**
     * 设置文件读取为 Text
     * @param {File|Blob} file - 文件对象
     * @param {Object} options - 配置选项
     * @param {Function} options.on.load - 加载成功回调
     * @param {Function} options.on.error - 加载失败回调
     */
    setupBlobAsText: function(file, options) {
        if (!file) return;

        options = options || {};
        var onCallbacks = options.on || {};

        readAsText(file, onCallbacks.load, onCallbacks.error);
    },

    /**
     * 设置同步模式 (兼容接口，当前不支持)
     * @param {boolean} value - 是否同步模式
     */
    setSync: function(value) {
        // 当前实现不支持同步模式
        console.warn("PPTXFileReader: Sync mode is not supported");
    },

    /**
     * 获取同步模式状态 (兼容接口)
     * @returns {boolean} 始终返回 false
     */
    getSync: function() {
        return false;
    }
};

