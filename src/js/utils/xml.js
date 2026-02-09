/**
 * XML 工具函数模块
 * 提供XML节点遍历和查询功能
 */

import { SLIDE_FACTOR, FONT_SIZE_FACTOR } from '../core/constants.js';
import tXml from '../core/tXml.js';

export const PPTXXmlUtils = (function() {

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

        var l = path.length;
        for (var i = 0; i < l; i++) {
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
        var map = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        };
        return text.replace(/[&<>"']/g, function (m) { return map[m]; });
    }

    /**
     * readXmlFile - 读取XML文件并解析为对象
     * @param {Object} zip - JSZip实例
     * @param {string} filename - 文件名
     * @param {boolean} isSlideContent - 是否为幻灯片内容
     * @param {number} appVersion - 应用版本
     * @returns {Object} 解析后的XML对象
     */
    function readXmlFile(zip, filename, isSlideContent, appVersion) {
        try {
            var fileContent = zip.file(filename).asText();
            if (isSlideContent && appVersion <= 12) {
                //< office2007
                //remove "<!CDATA[ ... ]]>" tag
                fileContent = fileContent.replace(/<!\[CDATA\[(.*?)\]\]>/g, '$1');
            }
            var xmlData = tXml(fileContent, { simplify: 1 });
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
     * @returns {Object} 包含slides和slideLayouts的对象
     */
    function getContentTypes(zip, appVersion) {
        var ContentTypesJson = PPTXXmlUtils.readXmlFile(zip, "[Content_Types].xml", false, appVersion);
        
        var subObj = ContentTypesJson["Types"]["Override"];
        var slidesLocArray = [];
        var slideLayoutsLocArray = [];
        for (var i = 0; i < subObj.length; i++) {
            switch (subObj[i]["attrs"]["ContentType"]) {
                case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
                    slidesLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
                    break;
                case "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml":
                    slideLayoutsLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
                    break;
                default:
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
     * @returns {Object} 包含width和height的对象
     */
    function getSlideSizeAndSetDefaultTextStyle(zip, settings) {
        //get app version
        var app = PPTXXmlUtils.readXmlFile(zip, "docProps/app.xml");
        var app_verssion_str = app["Properties"]["AppVersion"]
        const app_verssion = Number(app_verssion_str);
        console.log("create by Office PowerPoint app verssion: ", app_verssion_str)

        //get slide dimensions
        var rtenObj = {};
        var content = PPTXXmlUtils.readXmlFile(zip, "ppt/presentation.xml");
        var sldSzAttrs = content["p:presentation"]["p:sldSz"]["attrs"];
        var sldSzWidth = parseInt(sldSzAttrs["cx"]);
        var sldSzHeight = parseInt(sldSzAttrs["cy"]);
        var sldSzType = sldSzAttrs["type"];
        console.log("Presentation size type: ", sldSzType)

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

        const slideWidth = sldSzWidth * SLIDE_FACTOR + settings.incSlide.width|0;// * scaleX;//parseInt(sldSzAttrs["cx"]) * 96 / 914400;
        const slideHeight = sldSzHeight * SLIDE_FACTOR + settings.incSlide.height|0;// * scaleY;//parseInt(sldSzAttrs["cy"]) * 96 / 914400;
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
                console.log(`Media file found at alternative path: ${altPath} (originally: ${originalPath})`);
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

    function extractFileExtension(filename) {
            return filename.substr((~-filename.lastIndexOf(".") >>> 0) + 2);
        }
    function getMimeType(imgFileExt) {
            var mimeType = "";
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
                case "emf":
                    mimeType = "image/emf";
                    break;
                case "wmf":
                    mimeType = "image/wmf";
                case "tif":
                case "tiff":
                    mimeType = "image/tiff";
                    break;
            }
            return mimeType;
        }

    
        function getPosition(slideSpNode, pNode, slideLayoutSpNode, slideMasterSpNode, sType) {
            var off;
            var x = -1, y = -1;

            if (slideSpNode !== undefined) {
                off = slideSpNode["a:off"]["attrs"];
            }

            if (off === undefined && slideLayoutSpNode !== undefined) {
                off = slideLayoutSpNode["a:off"]["attrs"];
            } else if (off === undefined && slideMasterSpNode !== undefined) {
                off = slideMasterSpNode["a:off"]["attrs"];
            }
            var offX = 0, offY = 0;
            var grpX = 0, grpY = 0;
            if (sType == "group") {

                var grpXfrmNode = PPTXXmlUtils.getTextByPathList(pNode, ["p:grpSpPr", "a:xfrm"]);
                if (xfrmNode !== undefined) {
                    grpX = parseInt(grpXfrmNode["a:off"]["attrs"]["x"]) * SLIDE_FACTOR;
                    grpY = parseInt(grpXfrmNode["a:off"]["attrs"]["y"]) * SLIDE_FACTOR;
                    // var chx = parseInt(grpXfrmNode["a:chOff"]["attrs"]["x"]) * SLIDE_FACTOR;
                    // var chy = parseInt(grpXfrmNode["a:chOff"]["attrs"]["y"]) * SLIDE_FACTOR;
                    // var cx = parseInt(grpXfrmNode["a:ext"]["attrs"]["cx"]) * SLIDE_FACTOR;
                    // var cy = parseInt(grpXfrmNode["a:ext"]["attrs"]["cy"]) * SLIDE_FACTOR;
                    // var chcx = parseInt(grpXfrmNode["a:chExt"]["attrs"]["cx"]) * SLIDE_FACTOR;
                    // var chcy = parseInt(grpXfrmNode["a:chExt"]["attrs"]["cy"]) * SLIDE_FACTOR;
                    // var rotate = parseInt(grpXfrmNode["attrs"]["rot"])
                }
            }
            if (sType == "group-rotate" && pNode["p:grpSpPr"] !== undefined) {
                var xfrmNode = pNode["p:grpSpPr"]["a:xfrm"];
                // var ox = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * SLIDE_FACTOR;
                // var oy = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * SLIDE_FACTOR;
                var chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * SLIDE_FACTOR;
                var chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * SLIDE_FACTOR;

                offX = chx;
                offY = chy;
            }
            if (off === undefined) {
                return "";
            } else {
                x = parseInt(off["x"]) * SLIDE_FACTOR;
                y = parseInt(off["y"]) * SLIDE_FACTOR;
                // if (type = "body")
                //     console.log("getPosition: slideSpNode: ", slideSpNode, ", type: ", type, "x: ", x, "offX:", offX, "y:", y, "offY:", offY)
                return (isNaN(x) || isNaN(y)) ? "" : "top:" + (y - offY + grpY) + "px; left:" + (x - offX + grpX) + "px;";
            }

        }

        function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
            var ext = undefined;
            var w = -1, h = -1;

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
                w = parseInt(ext["cx"]) * SLIDE_FACTOR;
                h = parseInt(ext["cy"]) * SLIDE_FACTOR;
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
            var urlregex = /^(https?|ftp):\/\/([a-zA-Z0-9.-]+(:[a-zA-Z0-9.&%$-]+)*@)*((25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]?)(\.(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])){3}|([a-zA-Z0-9-]+\.)*[a-zA-Z0-9-]+\.(com|edu|gov|int|mil|net|org|biz|arpa|info|name|pro|aero|coop|museum|[a-zA-Z]{2}))(:[0-9]+)*(\/($|[a-zA-Z0-9.,?'\\+&%$#=~_-]+))*$/;
            return urlregex.test(vdoFile);
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
    };
})();
