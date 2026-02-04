/**
 * 颜色工具函数模块
 * 提供颜色转换、填充获取等功能
 * 依赖 tinycolor 库
 */

/**
 * toHex - 将数字转换为两位十六进制
 * @param {number} n - 数字
 * @returns {string} 十六进制字符串
 */

var PPTXColorUtils = (function() {
    function toHex(n) {
    const hex = n.toString(16);
    while (hex.length < 2) { hex = "0" + hex; }
    return hex;
}

/**
 * hslToRgb - 将HSL颜色转换为RGB
 * @param {number} hue - 色相 (0-360)
 * @param {number} sat - 饱和度 (0-1)
 * @param {number} light - 亮度 (0-1)
 * @returns {Object} RGB对象 {r, g, b}
 */
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

/**
 * hueToRgb - HSL到RGB的辅助函数
 * @param {number} t1 - 色彩值1
 * @param {number} t2 - 色彩值2
 * @param {number} hue - 色相
 * @returns {number} RGB分量值
 */
function hueToRgb(t1, t2, hue) {
    if (hue < 0) hue += 6;
    if (hue >= 6) hue -= 6;
    if (hue < 1) return (t2 - t1) * hue + t1;
    else if (hue < 3) return t2;
    else if (hue < 4) return (t2 - t1) * (4 - hue) + t1;
    else return t1;
}

/**
 * applyShade - 应用阴影效果（使颜色变暗）
 * @param {string} rgbStr - RGB颜色字符串
 * @param {number} shadeValue - 阴影值 (0-1)
 * @param {boolean} isAlpha - 是否包含透明度
 * @returns {string} 转换后的颜色
 */
    function applyShade(rgbStr, shadeValue, isAlpha) {
    const color = window.tinycolor(rgbStr).toHsl();
    if (shadeValue >= 1) {
        shadeValue = 1;
    }
    const cacl_l = Math.min(color.l * shadeValue, 1);
    if (isAlpha)
        return window.tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return window.tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}

/**
 * applyTint - 应用着色效果（使颜色变亮）
 * @param {string} rgbStr - RGB颜色字符串
 * @param {number} tintValue - 着色值 (0-1)
 * @param {boolean} isAlpha - 是否包含透明度
 * @returns {string} 转换后的颜色
 */
    function applyTint(rgbStr, tintValue, isAlpha) {
    const color = window.tinycolor(rgbStr).toHsl();
    if (tintValue >= 1) {
        tintValue = 1;
    }
    const cacl_l = color.l * tintValue + (1 - tintValue);
    if (isAlpha)
        return window.tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return window.tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}

/**
 * applyLumOff - 应用亮度偏移
 * @param {string} rgbStr - RGB颜色字符串
 * @param {number} offset - 偏移量
 * @param {boolean} isAlpha - 是否包含透明度
 * @returns {string} 转换后的颜色
 */
    function applyLumOff(rgbStr, offset, isAlpha) {
    const color = window.tinycolor(rgbStr).toHsl();
    const lum = offset + color.l;
    if (lum >= 1) {
        if (isAlpha)
            return window.tinycolor({ h: color.h, s: color.s, l: 1, a: color.a }).toHex8();
        return window.tinycolor({ h: color.h, s: color.s, l: 1, a: color.a }).toHex();
    }
    if (isAlpha)
        return window.tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex8();
    return window.tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex();
}

/**
 * applyLumMod - 应用亮度调制
 * @param {string} rgbStr - RGB颜色字符串
 * @param {number} multiplier - 乘数
 * @param {boolean} isAlpha - 是否包含透明度
 * @returns {string} 转换后的颜色
 */
    function applyLumMod(rgbStr, multiplier, isAlpha) {
    const color = window.tinycolor(rgbStr).toHsl();
    let cacl_l = color.l * multiplier;
    if (cacl_l >= 1) {
        cacl_l = 1;
    }
    if (isAlpha)
        return window.tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return window.tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}

/**
 * applyHueMod - 应用色相调制
 * @param {string} rgbStr - RGB颜色字符串
 * @param {number} multiplier - 乘数
 * @param {boolean} isAlpha - 是否包含透明度
 * @returns {string} 转换后的颜色
 */
    function applyHueMod(rgbStr, multiplier, isAlpha) {
    const color = window.tinycolor(rgbStr).toHsl();
    let cacl_h = color.h * multiplier;
    if (cacl_h >= 360) {
        cacl_h = cacl_h - 360;
    }
    if (isAlpha)
        return window.tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex8();
    return window.tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex();
}

/**
 * applySatMod - 应用饱和度调制
 * @param {string} rgbStr - RGB颜色字符串
 * @param {number} multiplier - 乘数
 * @param {boolean} isAlpha - 是否包含透明度
 * @returns {string} 转换后的颜色
 */
    function applySatMod(rgbStr, multiplier, isAlpha) {
    const color = window.tinycolor(rgbStr).toHsl();
    let cacl_s = color.s * multiplier;
    if (cacl_s >= 1) {
        cacl_s = 1;
    }
    if (isAlpha)
        return window.tinycolor({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex8();
    return window.tinycolor({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex();
}

/**
 * rgba2hex - 将RGBA颜色转换为十六进制
 * @param {string} rgbaStr - RGBA颜色字符串
 * @returns {string} 十六进制颜色字符串
 */
    function rgba2hex(rgbaStr) {
    let a,
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
    a = ((a * 255) | 1 << 8).toString(16).slice(1);
    hex = hex + a;

    return hex;
}

/**
 * getColorName2Hex - 将颜色名称转换为十六进制
 * @param {string} name - 颜色名称
 * @returns {string} 十六进制颜色字符串
 */
    function getColorName2Hex(name) {
    const colorName = ['white', 'AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk', 'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenRod', 'DarkGray', 'DarkGrey', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray', 'DarkSlateGrey', 'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DimGrey', 'DodgerBlue', 'FireBrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'GoldenRod', 'Gray', 'Grey', 'Green', 'GreenYellow', 'HoneyDew', 'HotPink', 'IndianRed', 'Indigo', 'Ivory', 'Khaki', 'Lavender', 'LavenderBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenRodYellow', 'LightGray', 'LightGrey', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray', 'LightSlateGrey', 'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquaMarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue', 'MintCream', 'MistyRose', 'Moccasin', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenRod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue', 'Purple', 'RebeccaPurple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'SlateGrey', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Turquoise', 'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen'];
    const colorHex = ['ffffff', 'f0f8ff', 'faebd7', '00ffff', '7fffd4', 'f0ffff', 'f5f5dc', 'ffe4c4', '000000', 'ffebcd', '0000ff', '8a2be2', 'a52a2a', 'deb887', '5f9ea0', '7fff00', 'd2691e', 'ff7f50', '6495ed', 'fff8dc', 'dc143c', '00ffff', '00008b', '008b8b', 'b8860b', 'a9a9a9', 'a9a9a9', '006400', 'bdb76b', '8b008b', '556b2f', 'ff8c00', '9932cc', '8b0000', 'e9967a', '8fbc8f', '483d8b', '2f4f4f', '2f4f4f', '00ced1', '9400d3', 'ff1493', '00bfff', '696969', '696969', '1e90ff', 'b22222', 'fffaf0', '228b22', 'ff00ff', 'dcdcdc', 'f8f8ff', 'ffd700', 'daa520', '808080', '808080', '008000', 'adff2f', 'f0fff0', 'ff69b4', 'cd5c5c', '4b0082', 'fffff0', 'f0e68c', 'e6e6fa', 'fff0f5', '7cfc00', 'fffacd', 'add8e6', 'f08080', 'e0ffff', 'fafad2', 'd3d3d3', 'd3d3d3', '90ee90', 'ffb6c1', 'ffa07a', '20b2aa', '87cefa', '778899', '778899', 'b0c4de', 'ffffe0', '00ff00', '32cd32', 'faf0e6', 'ff00ff', '800000', '66cdaa', '0000cd', 'ba55d3', '9370db', '3cb371', '7b68ee', '00fa9a', '48d1cc', 'c71585', '191970', 'f5fffa', 'ffe4e1', 'ffe4b5', 'ffdead', '000080', 'fdf5e6', '808000', '6b8e23', 'ffa500', 'ff4500', 'da70d6', 'eee8aa', '98fb98', 'afeeee', 'db7093', 'ffefd5', 'ffdab9', 'cd853f', 'ffc0cb', 'dda0dd', 'b0e0e6', '800080', '663399', 'ff0000', 'bc8f8f', '4169e1', '8b4513', 'fa8072', 'f4a460', '2e8b57', 'fff5ee', 'a0522d', 'c0c0c0', '87ceeb', '6a5acd', '708090', '708090', 'fffafa', '00ff7f', '4682b4', 'd2b48c', '008080', 'd8bfd8', 'ff6347', '40e0d0', 'ee82ee', 'f5deb3', 'ffffff', 'f5f5f5', 'ffff00', '9acd32'];
    const findIndx = colorName.indexOf(name);
    let hex;
    if (findIndx != -1) {
        hex = colorHex[findIndx];
    }
    return hex;
}

/**
 * getSchemeColorFromTheme - 从主题中获取方案颜色
 * @param {string} schemeClr - 方案颜色名称
 * @param {Object} clrMap - 颜色映射
 * @param {string} phClr - 占位符颜色
 * @param {Object} warpObj - 包含主题内容的包装对象
 * @returns {string} 十六进制颜色字符串
 */
    function getSchemeColorFromTheme(schemeClr, clrMap, phClr, warpObj) {
    let slideLayoutClrOvride;
    if (clrMap !== undefined) {
        slideLayoutClrOvride = clrMap;
    } else {
        const sldClrMapOvr = getTextByPathList(warpObj["slideContent"], ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
        if (sldClrMapOvr !== undefined) {
            slideLayoutClrOvride = sldClrMapOvr;
        } else {
            const sldClrMapOvr = getTextByPathList(warpObj["slideLayoutContent"], ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                slideLayoutClrOvride = sldClrMapOvr;
            } else {
                slideLayoutClrOvride = getTextByPathList(warpObj["slideMasterContent"], ["p:sldMaster", "p:clrMap", "attrs"]);
            }
        }
    }

    const schmClrName = schemeClr.substr(2);
    let color;
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

        const refNode = getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
        color = getTextByPathList(refNode, ["a:srgbClr", "attrs", "val"]);
        if (color === undefined && refNode !== undefined) {
            color = getTextByPathList(refNode, ["a:sysClr", "attrs", "lastClr"]);
        }
    }
    return color;
}

/**
 * getSvgGradient - 生成SVG渐变
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} angl - 角度
 * @param {Array} color_arry - 颜色数组
 * @param {string} shpId - 形状ID
 * @returns {string} SVG渐变字符串
 */
    function getSvgGradient(w, h, angl, color_arry, shpId) {
    const stopsArray = getMiddleStops(color_arry.length - 2);

    let svgAngle = '',
        svgHeight = h,
        svgWidth = w,
        svg = '',
        xy_ary = SVGangle(angl, svgHeight, svgWidth),
        x1 = xy_ary[0],
        y1 = xy_ary[1],
        x2 = xy_ary[2],
        y2 = xy_ary[3];

    const sal = stopsArray.length,
        sr = sal < 20 ? 100 : 1000;
    svgAngle = ' gradientUnits="userSpaceOnUse" x1="' + x1 + '%" y1="' + y1 + '%" x2="' + x2 + '%" y2="' + y2 + '%"';
    svgAngle = '<linearGradient id="linGrd_' + shpId + '"' + svgAngle + '>\n';
    svg += svgAngle;

    for (let i = 0; i < sal; i++) {
        const tinClr = window.tinycolor("#" + color_arry[i]);
        const alpha = tinClr.getAlpha();
        svg += '<stop offset="' + Math.round(parseFloat(stopsArray[i]) / 100 * sr) / sr + '" style="stop-color:' + tinClr.toHexString() + '; stop-opacity:' + (alpha) + ';"';
        svg += '/>\n';
    }

    svg += '</linearGradient>\n';

    return svg;
}

/**
 * getMiddleStops - 获取中间停止点
 * @param {number} s - 停止点数量
 * @returns {Array} 停止点数组
 */
function getMiddleStops(s) {
    const sArry = ['0%', '100%'];
    if (s == 0) {
        return sArry;
    } else {
        let i = s;
        while (i--) {
            const middleStop = 100 - ((100 / (s + 1)) * (i + 1)),
                middleStopString = middleStop + "%";
            sArry.splice(-1, 0, middleStopString);
        }
    }
    return sArry;
}

/**
 * SVGangle - 计算SVG角度坐标
 * @param {number} deg - 角度
 * @param {number} svgHeight - SVG高度
 * @param {number} svgWidth - SVG宽度
 * @returns {Array} 坐标数组 [x1, y1, x2, y2]
 */
function SVGangle(deg, svgHeight, svgWidth) {
    const w = parseFloat(svgWidth),
        h = parseFloat(svgHeight),
        ang = parseFloat(deg);
    let o = 2,
        n = 2,
        wc = w / 2,
        hc = h / 2,
        tx1 = 2,
        ty1 = 2,
        tx2 = 2,
        ty2 = 2;
    const k = (((ang % 360) + 360) % 360),
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

    const m = o + (n / i),
        tx1Val = tx1 == 2 ? i * (m - l) / (Math.pow(i, 2) + 1) : tx1,
        ty1Val = ty1 == 2 ? i * tx1Val + l : ty1,
        tx2Val = tx2 == 2 ? w - tx1Val : tx2,
        ty2Val = ty2 == 2 ? h - ty1Val : ty2,
        x1 = Math.round(tx2Val / w * 100 * 100) / 100,
        y1 = Math.round(ty2Val / h * 100 * 100) / 100,
        x2 = Math.round(tx1Val / w * 100 * 100) / 100,
        y2 = Math.round(ty1Val / h * 100 * 100) / 100;
    return [x1, y1, x2, y2];
}

/**
 * getFillType - 获取填充类型
 * @param {Object} node - 节点对象
 * @returns {string} 填充类型
 */
function getFillType(node) {
    if (!node) return "NO_FILL";
    
    if (node["a:solidFill"] !== undefined) {
        return "SOLID_FILL";
    } else if (node["a:gradFill"] !== undefined) {
        return "GRADIENT_FILL";
    } else if (node["a:pattFill"] !== undefined) {
        return "PATTERN_FILL";
    } else if (node["a:blipFill"] !== undefined) {
        return "PIC_FILL";
    } else if (node["a:noFill"] !== undefined) {
        return "NO_FILL";
    } else {
        return "NO_FILL";
    }
}

/**
 * getSolidFill - 获取纯色填充
 * @param {Object} node - 节点对象
 * @param {Object} clrMap - 颜色映射
 * @param {string} phClr - 占位符颜色
 * @param {Object} warpObj - 包装对象
 * @returns {string} 十六进制颜色
 */
function getSolidFill(node, clrMap, phClr, warpObj) {
    if (!node) return "000000";
    
    var srgbClr = node["a:srgbClr"];
    if (srgbClr && srgbClr["attrs"] && srgbClr["attrs"]["val"]) {
        return srgbClr["attrs"]["val"];
    }
    
    var schemeClr = node["a:schemeClr"];
    if (schemeClr && schemeClr["attrs"] && schemeClr["attrs"]["val"]) {
        return getSchemeColorFromTheme(schemeClr["attrs"]["val"], clrMap, phClr, warpObj);
    }
    
    var prstClr = node["a:prstClr"];
    if (prstClr && prstClr["attrs"] && prstClr["attrs"]["val"]) {
        return getColorName2Hex(prstClr["attrs"]["val"]);
    }
    
    var sysClr = node["a:sysClr"];
    if (sysClr && sysClr["attrs"] && sysClr["attrs"]["lastClr"]) {
        return sysClr["attrs"]["lastClr"];
    }
    
    return "000000";
}

/**
 * getTextByPathList - 通过路径列表获取节点值
 * @param {Object} node - 节点对象
 * @param {Array} path - 路径数组
 * @returns {*} 节点值
 */
function getTextByPathList(node, path) {
    if (!node || !path || path.constructor !== Array) {
        return undefined;
    }
    
    var current = node;
    for (var i = 0; i < path.length; i++) {
        current = current[path[i]];
        if (current === undefined || current === null) {
            return undefined;
        }
    }
    
    return current;
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

    return {
        toHex: toHex,
        hslToRgb: hslToRgb,
        applyShade: applyShade,
        applyTint: applyTint,
        applyLumOff: applyLumOff,
        applyLumMod: applyLumMod,
        applyHueMod: applyHueMod,
        applySatMod: applySatMod,
        rgba2hex: rgba2hex,
        getColorName2Hex: getColorName2Hex,
        getSchemeColorFromTheme: getSchemeColorFromTheme,
        getSvgGradient: getSvgGradient,
        getFillType: getFillType,
        getSolidFill: getSolidFill,
        getTextByPathList: getTextByPathList,
        angleToDegrees: angleToDegrees
    };
})();

window.PPTXColorUtils = PPTXColorUtils;