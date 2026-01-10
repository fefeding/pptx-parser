/**
 * PPTXUtils - 通用工具函数库
 * 提取自 pptxjs.js
 */

(function () {
    var $ = window.jQuery;

    // 角度转换
    function angleToDegrees(angle) {
        if (angle == "" || angle == null) {
            return 0;
        }
        return Math.round(angle / 60000);
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

    // Base64 编码 ArrayBuffer
    function base64ArrayBuffer(arrayBuffer) {
        var base64 = '';
        var encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
        var bytes = new Uint8Array(arrayBuffer);
        var byteLength = bytes.byteLength;
        var byteRemainder = byteLength % 3;
        var mainLength = byteLength - byteRemainder;

        var a, b, c, d;
        var chunk;

        for (var i = 0; i < mainLength; i = i + 3) {
            chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
            a = (chunk & 16515072) >> 18;
            b = (chunk & 258048) >> 12;
            c = (chunk & 4032) >> 6;
            d = chunk & 63;
            base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d];
        }

        if (byteRemainder == 1) {
            chunk = bytes[mainLength];
            a = (chunk & 252) >> 2;
            b = (chunk & 3) << 4;
            base64 += encodings[a] + encodings[b] + '==';
        } else if (byteRemainder == 2) {
            chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];
            a = (chunk & 64512) >> 10;
            b = (chunk & 1008) >> 4;
            c = (chunk & 15) << 2;
            base64 += encodings[a] + encodings[b] + encodings[c] + '=';
        }

        return base64;
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

        if (node === undefined) {
            return undefined;
        }

        var l = path.length;
        for (var i = 0; i < l; i++) {
            node = node[path[i]];
            if (node === undefined) {
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

    // 颜色处理函数
    function applyShade(rgbStr, shadeValue, isAlpha) {
        var color = tinycolor(rgbStr).toHsl();
        if (shadeValue >= 1) {
            shadeValue = 1;
        }
        var cacl_l = Math.min(color.l * shadeValue, 1);
        if (isAlpha)
            return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
        return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
    }

    function applyTint(rgbStr, tintValue, isAlpha) {
        var color = tinycolor(rgbStr).toHsl();
        if (tintValue >= 1) {
            tintValue = 1;
        }
        var cacl_l = color.l * tintValue + (1 - tintValue);
        if (isAlpha)
            return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
        return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
    }

    function applyLumOff(rgbStr, offset, isAlpha) {
        var color = tinycolor(rgbStr).toHsl();
        var lum = offset + color.l;
        if (lum >= 1) {
            if (isAlpha)
                return tinycolor({ h: color.h, s: color.s, l: 1, a: color.a }).toHex8();
            return tinycolor({ h: color.h, s: color.s, l: 1, a: color.a }).toHex();
        }
        if (isAlpha)
            return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex8();
        return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex();
    }

    function applyLumMod(rgbStr, multiplier, isAlpha) {
        var color = tinycolor(rgbStr).toHsl();
        var cacl_l = color.l * multiplier;
        if (cacl_l >= 1) {
            cacl_l = 1;
        }
        if (isAlpha)
            return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
        return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
    }

    function applyHueMod(rgbStr, multiplier, isAlpha) {
        var color = tinycolor(rgbStr).toHsl();
        var cacl_h = color.h * multiplier;
        if (cacl_h >= 360) {
            cacl_h = cacl_h - 360;
        }
        if (isAlpha)
            return tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex8();
        return tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex();
    }

    function applySatMod(rgbStr, multiplier, isAlpha) {
        var color = tinycolor(rgbStr).toHsl();
        var cacl_s = color.s * multiplier;
        if (cacl_s >= 1) {
            cacl_s = 1;
        }
        if (isAlpha)
            return tinycolor({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex8();
        return tinycolor({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex();
    }

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
            a = 1;
        }
        a = ((a * 255) | 1 << 8).toString(16).slice(1);
        hex = hex + a;

        return hex;
    }

    // 公开工具函数
    window.PPTXUtils = {
        angleToDegrees: angleToDegrees,
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
        applyShade: applyShade,
        applyTint: applyTint,
        applyLumOff: applyLumOff,
        applyLumMod: applyLumMod,
        applyHueMod: applyHueMod,
        applySatMod: applySatMod,
        rgba2hex: rgba2hex
    };

})();