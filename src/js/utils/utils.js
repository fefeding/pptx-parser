    // PPTX 坐标转换因子 (96 pixels per inch / 914400 EMUs per inch)
import { PPTXColorUtils } from '../core/color-utils.js';
var slideFactor = 96 / 914400;

// 辅助函数：获取 XML 属性，兼容 tXml 和 fast-xml-parser 格式
export function getAttrs(obj) {
    if (!obj) return {};
    // fast-xml-parser 直接在对象上，tXml 在 attrs 属性中
    if (obj.attrs) {
        return obj.attrs;
    }
    // fast-xml-parser 格式：属性直接在对象上
    var result = {};
    for (var key in obj) {
        if (key !== '#text' && key !== '__cdata' && typeof obj[key] !== 'object') {
            result[key] = obj[key];
        }
    }
    return result;
}

// 辅助函数：获取单个属性值
export function getAttr(obj, attrName) {
    var attrs = getAttrs(obj);
    return attrs[attrName];
}

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
        c = (chunk & 4032)     >>  6;
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
        b = (chunk & 1008)  >>  4;
        c = (chunk & 15)    <<  2;
        base64 += encodings[a] + encodings[b] + encodings[c] + '=';
    }

    return base64;
}

    // 从多种输入类型获取 ArrayBuffer
function getArrayBufferFromFileInput(input) {
    return new Promise(function(resolve, reject) {
        // 处理 ArrayBuffer
        if (input instanceof ArrayBuffer) {
            resolve({ arrayBuffer: input, fileName: '' });
            return;
        }

        // 处理 File 或 Blob 对象
        if (input instanceof File || input instanceof Blob) {
            const reader = new FileReader();
            reader.onload = function(e) {
                resolve({ arrayBuffer: e.target.result, fileName: input.name || '' });
            };
            reader.onerror = function(e) {
                reject(new Error('Failed to read file: ' + e));
            };
            reader.readAsArrayBuffer(input);
            return;
        }

        // 处理 URL 字符串
        if (typeof input === 'string') {
            fetch(input)
                .then(function(response) {
                    if (!response.ok) {
                        throw new Error('HTTP error! status: ' + response.status);
                    }
                    return response.arrayBuffer();
                })
                .then(function(arrayBuffer) {
                    // 从 URL 提取文件名
                    var fileName = '';
                    var urlParts = input.split('/');
                    if (urlParts.length > 0) {
                        fileName = urlParts[urlParts.length - 1];
                        var dotIndex = fileName.lastIndexOf('.');
                        if (dotIndex > 0) {
                            fileName = fileName.substring(0, dotIndex);
                        }
                    }
                    resolve({ arrayBuffer: arrayBuffer, fileName: fileName });
                })
                .catch(function(err) {
                    reject(new Error('Failed to fetch URL: ' + err.message));
                });
            return;
        }

        reject(new Error('Invalid input type. Expected string URL, File, Blob, or ArrayBuffer.'));
    });
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
        var key = path[i];

        // 特殊处理 "attrs" 键：兼容 tXml 和 fast-xml-parser
        if (key === "attrs") {
            // fast-xml-parser 格式：属性直接在对象上
            // 如果下一个键在当前节点中直接存在，直接返回
            if (i + 1 < l && node[path[i + 1]] !== undefined) {
                return node[path[i + 1]];
            }
            // 否则尝试访问 attrs 属性（tXml 格式）
            node = node[key];
        } else {
            node = node[key];
        }

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

    // 获取边框样式
function getBorder(node, pNode, isSvgMode, bType, warpObj) {
    var cssText, lineNode;

    if (bType == "shape") {
        cssText = "border: ";
        lineNode = node["p:spPr"]["a:ln"];
    } else if (bType == "text") {
        cssText = "";
        lineNode = node["a:rPr"]["a:ln"];
    }

    var is_noFill = getTextByPathList(lineNode, ["a:noFill"]);
    if (is_noFill !== undefined) {
        return "hidden";
    }

    if (lineNode == undefined) {
        var lnRefNode = getTextByPathList(node, ["p:style", "a:lnRef"]);
        if (lnRefNode !== undefined){
            var lnIdx = getTextByPathList(lnRefNode, ["attrs", "idx"]);
            lineNode = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"][Number(lnIdx) - 1];
        }
    }
    if (lineNode == undefined) {
        cssText = "";
        lineNode = node;
    }

    var borderColor;
    if (lineNode !== undefined) {
        var borderWidth = parseInt(getTextByPathList(lineNode, ["attrs", "w"])) / 12700;
        if (isNaN(borderWidth) || borderWidth < 1) {
            cssText += (4/3) + "px ";
        } else {
            cssText += borderWidth + "px ";
        }

        var borderType = getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
        if (borderType === undefined) {
            borderType = getTextByPathList(lineNode, ["attrs", "cmpd"]);
        }
        var strokeDasharray = "0";
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
            default:
                cssText += "solid";
                strokeDasharray = "0";
        }

        var fillTyp = PPTXColorUtils.getFillType(lineNode);
        if (fillTyp == "NO_FILL") {
            borderColor = isSvgMode ? "none" : "";
        } else if (fillTyp == "SOLID_FILL") {
            borderColor = PPTXColorUtils.getSolidFill(lineNode["a:solidFill"], undefined, undefined, warpObj);
        } else if (fillTyp == "GRADIENT_FILL") {
            borderColor = PPTXColorUtils.getGradientFill(lineNode["a:gradFill"], warpObj);
        } else if (fillTyp == "PATTERN_FILL") {
            borderColor = PPTXColorUtils.getPatternFill(lineNode["a:pattFill"], warpObj);
        }
    }

    if (borderColor === undefined) {
        var lnRefNode = getTextByPathList(node, ["p:style", "a:lnRef"]);
        if (lnRefNode !== undefined) {
            borderColor = PPTXColorUtils.getSolidFill(lnRefNode, undefined, undefined, warpObj);
        }
    }

    if (borderColor === undefined) {
        if (isSvgMode) {
            borderColor = "none";
        } else {
            borderColor = "hidden";
        }
    } else {
        borderColor = "#" + borderColor;
    }
    cssText += " " + borderColor + " ";

    if (isSvgMode) {
        return { "color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray };
    } else {
        return cssText + ";";
    }
}

    // 获取表格边框样式
function getTableBorders(node, warpObj) {
    var borderStyle = "";
    if (node["a:bottom"] !== undefined) {
        var obj = {
            "p:spPr": {
                "a:ln": node["a:bottom"]["a:ln"]
            }
        }
        var borders = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-bottom");
    }
    if (node["a:top"] !== undefined) {
        var obj = {
            "p:spPr": {
                "a:ln": node["a:top"]["a:ln"]
            }
        }
        var borders = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-top");
    }
    if (node["a:right"] !== undefined) {
        var obj = {
            "p:spPr": {
                "a:ln": node["a:right"]["a:ln"]
            }
        }
        var borders = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-right");
    }
    if (node["a:left"] !== undefined) {
        var obj = {
            "p:spPr": {
                "a:ln": node["a:left"]["a:ln"]
            }
        }
        var borders = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-left");
    }

    return borderStyle;
}

    // 公开工具函数
const PPTXUtils = {
    angleToDegrees: angleToDegrees,
    degreesToRadians: degreesToRadians,
    getMimeType: getMimeType,
    base64ArrayBuffer: base64ArrayBuffer,
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
    getBorder: getBorder,
    getTableBorders: getTableBorders,
    getArrayBufferFromFileInput: getArrayBufferFromFileInput,
    getSlideFactor: function() { return slideFactor; },
    setSlideFactor: function(factor) { slideFactor = factor; },
    getAttrs,
    getAttr,
};


export { PPTXUtils };
