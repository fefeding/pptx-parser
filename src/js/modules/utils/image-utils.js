/**
 * 图片和媒体工具函数模块
 * 提供图片处理、MIME类型识别等功能
 */

/**
 * getMimeType - 根据文件扩展名获取MIME类型
 * @param {string} imgFileExt - 文件扩展名
 * @returns {string} MIME类型字符串
 */

var PPTXImageUtils = (function() {
    function getMimeType(imgFileExt) {
    let mimeType = "";
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

/**
 * IsVideoLink - 检查是否为视频链接
 * @param {string} vdoFile - 视频文件或URL
 * @returns {boolean} 是否为视频链接
 */
    function IsVideoLink(vdoFile) {
    const urlregex = /^(https?|ftp):\/\/([a-zA-Z0-9.-]+(:[a-zA-Z0-9.&%$-]+)*@)*((25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9]?)(\.(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])){3}|([a-zA-Z0-9-]+\.)*[a-zA-Z0-9-]+\.(com|edu|gov|int|mil|net|org|biz|arpa|info|name|pro|aero|coop|museum|[a-zA-Z]{2}))(:[0-9]+)*(\/($|[a-zA-Z0-9.,?'\\+&%$#=~_-]+))*$/;
    return urlregex.test(vdoFile);
}

/**
 * extractFileExtension - 提取文件扩展名
 * @param {string} filename - 文件名
 * @returns {string} 文件扩展名
 */
    function extractFileExtension(filename) {
    return filename.substr((~-filename.lastIndexOf(".") >>> 0) + 2);
}

/**
 * base64ArrayBuffer - 将ArrayBuffer转换为Base64字符串
 * @param {ArrayBuffer} arrayBuffer - ArrayBuffer对象
 * @returns {string} Base64编码的字符串
 */
    function base64ArrayBuffer(arrayBuffer) {
    let base64 = '';
    const encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
    const bytes = new Uint8Array(arrayBuffer);
    const byteLength = bytes.byteLength;
    const byteRemainder = byteLength % 3;
    const mainLength = byteLength - byteRemainder;

    let a, b, c, d;
    let chunk;

    for (let i = 0; i < mainLength; i = i + 3) {
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

/**
 * getBase64ImageDimensions - 获取Base64图片的尺寸
 * @param {string} imgSrc - Base64图片数据
 * @returns {Array} [宽度, 高度]
 */
    function getBase64ImageDimensions(imgSrc) {
    const image = new Image();
    let w, h;
    image.onload = function () {
        w = image.width;
        h = image.height;
    };
    image.src = imgSrc;

    do {
        if (image.width !== undefined) {
            return [image.width, image.height];
        }
    } while (image.width === undefined);
}

/**
 * getSvgImagePattern - 生成SVG图片图案
 * @param {Object} node - XML节点
 * @param {string} fill - 填充内容
 * @param {string} shpId - 形状ID
 * @param {Object} warpObj - 包装对象
 * @returns {string} SVG图案字符串
 */
    function getSvgImagePattern(node, fill, shpId, warpObj) {
    const pic_dim = getBase64ImageDimensions(fill);
    const width = pic_dim[0];
    const height = pic_dim[1];

    const blipFillNode = node["p:spPr"]["a:blipFill"];
    const tileNode = getTextByPathList(blipFillNode, ["a:tile", "attrs"]);
    let sx, sy;
    if (tileNode !== undefined && tileNode["sx"] !== undefined) {
        sx = (parseInt(tileNode["sx"]) / 100000) * width;
        sy = (parseInt(tileNode["sy"]) / 100000) * height;
    }

    const blipNode = node["p:spPr"]["a:blipFill"]["a:blip"];
    const tialphaModFixNode = getTextByPathList(blipNode, ["a:alphaModFix", "attrs"]);
    let imgOpacity = "";
    if (tialphaModFixNode !== undefined && tialphaModFixNode["amt"] !== undefined && tialphaModFixNode["amt"] != "") {
        const amt = parseInt(tialphaModFixNode["amt"]) / 100000;
        const opacity = amt;
        imgOpacity = "opacity='" + opacity + "'";
    }

    let ptrn;
    if (sx !== undefined && sx != 0) {
        ptrn = '<pattern id="imgPtrn_' + shpId + '" x="0" y="0"  width="' + sx + '" height="' + sy + '" patternUnits="userSpaceOnUse">';
    } else {
        ptrn = '<pattern id="imgPtrn_' + shpId + '"  patternContentUnits="objectBoundingBox"  width="1" height="1">';
    }

    const duotoneNode = getTextByPathList(blipNode, ["a:duotone"]);
    let fillterNode = "";
    let filterUrl = "";
    if (duotoneNode !== undefined) {
        const clr_ary = [];
        Object.keys(duotoneNode).forEach(function (clr_type) {
            if (clr_type != "attrs") {
                const obj = {};
                obj[clr_type] = duotoneNode[clr_type];
                const hexClr = getSolidFill(obj, undefined, undefined, warpObj);
                const color = window.tinycolor("#" + hexClr);
                clr_ary.push(color.toRgb());
            }
        });

        if (clr_ary.length == 2) {
            fillterNode = '<filter id="svg_image_duotone"> ' +
                '<feColorMatrix type="matrix" values=".33 .33 .33 0 0' +
                '.33 .33 .33 0 0' +
                '.33 .33 .33 0 0' +
                '0 0 0 1 0">' +
                '</feColorMatrix>' +
                '<feComponentTransfer color-interpolation-filters="sRGB">' +
                '<feFuncR type="table" tableValues="' + clr_ary[0].r / 255 + ' ' + clr_ary[1].r / 255 + '"></feFuncR>' +
                '<feFuncG type="table" tableValues="' + clr_ary[0].g / 255 + ' ' + clr_ary[1].g / 255 + '"></feFuncG>' +
                '<feFuncB type="table" tableValues="' + clr_ary[0].b / 255 + ' ' + clr_ary[1].b / 255 + '"></feFuncB>' +
                '</feComponentTransfer>' +
                ' </filter>';
        }

        filterUrl = 'filter="url(#svg_image_duotone)"';
        ptrn += fillterNode;
    }

    fill = escapeHtml(fill);
    if (sx !== undefined && sx != 0) {
        ptrn += '<image  xlink:href="' + fill + '" x="0" y="0" width="' + sx + '" height="' + sy + '" ' + imgOpacity + ' ' + filterUrl + '></image>';
    } else {
        ptrn += '<image  xlink:href="' + fill + '" preserveAspectRatio="none" width="1" height="1" ' + imgOpacity + ' ' + filterUrl + '></image>';
    }
    ptrn += '</pattern>';

    return ptrn;
}

/**
 * escapeHtml - 转义HTML特殊字符
 * @param {string} text - 原始文本
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
    return text.replace(/[&<>"']/g, function (m) { return map[m]; });
}


    return {
        getMimeType: getMimeType,
        IsVideoLink: IsVideoLink,
        extractFileExtension: extractFileExtension,
        base64ArrayBuffer: base64ArrayBuffer,
        getBase64ImageDimensions: getBase64ImageDimensions,
        getSvgImagePattern: getSvgImagePattern
    };
})();

window.PPTXImageUtils = PPTXImageUtils;