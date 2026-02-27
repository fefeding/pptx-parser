// @ts-nocheck
/**
 * TinyColor v1.4.2
 * https://github.com/bgrins/TinyColor
 * Brian Grinstead, MIT License
 */

// =============================================================================
// 工具常量
// =============================================================================

const TRIM_LEFT = /^\s+/;
const TRIM_RIGHT = /\s+$/;
const MATH_ROUND = Math.round;
const MATH_MIN = Math.min;
const MATH_MAX = Math.max;
const MATH_RANDOM = Math.random;

let tinyCounter = 0;

// =============================================================================
// TinyColor 类
// =============================================================================

/**
 * TinyColor 构造函数
 * @constructor
 * @param {string|Object} color - 颜色输入值
 * @param {Object} [opts] - 配置选项
 */
function TinyColor(color, opts) {
    color = color || '';
    opts = opts || {};

    // 如果输入已经是 TinyColor 实例，直接返回
    if (color instanceof TinyColor) {
        return color;
    }

    // 如果作为普通函数调用，转为构造函数调用
    if (!(this instanceof TinyColor)) {
        return new TinyColor(color, opts);
    }

    const rgb = inputToRGB(color);
    this._originalInput = color;
    this._r = rgb.r;
    this._g = rgb.g;
    this._b = rgb.b;
    this._a = rgb.a;
    this._roundA = MATH_ROUND(100 * this._a) / 100;
    this._format = opts.format || rgb.format;
    this._gradientType = opts.gradientType;

    // 确保 [0,255] 范围不会变成 [0,1]
    if (this._r < 1) { this._r = MATH_ROUND(this._r); }
    if (this._g < 1) { this._g = MATH_ROUND(this._g); }
    if (this._b < 1) { this._b = MATH_ROUND(this._b); }

    this._ok = rgb.ok;
    this._tc_id = tinyCounter++;
}

// =============================================================================
// 工厂函数
// =============================================================================

/**
 * TinyColor 工厂函数
 * @param {string|Object} color - 颜色输入值
 * @param {Object} [opts] - 配置选项
 * @returns {TinyColor} TinyColor 实例
 */
export default function tinycolor(color, opts) {
    return new TinyColor(color, opts);
}

// =============================================================================
// 原型方法
// =============================================================================

TinyColor.prototype = {
    /**
     * 判断颜色是否为深色
     * @returns {boolean}
     */
    isDark: function() {
        return this.getBrightness() < 128;
    },

    /**
     * 判断颜色是否为浅色
     * @returns {boolean}
     */
    isLight: function() {
        return !this.isDark();
    },

    /**
     * 判断颜色是否有效
     * @returns {boolean}
     */
    isValid: function() {
        return this._ok;
    },

    /**
     * 获取原始输入值
     * @returns {string|Object}
     */
    getOriginalInput: function() {
        return this._originalInput;
    },

    /**
     * 获取颜色格式
     * @returns {string}
     */
    getFormat: function() {
        return this._format;
    },

    /**
     * 获取 Alpha 通道值
     * @returns {number}
     */
    getAlpha: function() {
        return this._a;
    },

    /**
     * 获取颜色亮度
     * @returns {number}
     */
    getBrightness: function() {
        // http://www.w3.org/TR/AERT#color-contrast
        const rgb = this.toRgb();
        return (rgb.r * 299 + rgb.g * 587 + rgb.b * 114) / 1000;
    },

    /**
     * 获取颜色明度
     * @returns {number}
     */
    getLuminance: function() {
        // http://www.w3.org/TR/2008/REC-WCAG20-20081211/#relativeluminancedef
        const rgb = this.toRgb();
        const RsRGB = rgb.r / 255;
        const GsRGB = rgb.g / 255;
        const BsRGB = rgb.b / 255;

        const R = RsRGB <= 0.03928 ? RsRGB / 12.92 : Math.pow((RsRGB + 0.055) / 1.055, 2.4);
        const G = GsRGB <= 0.03928 ? GsRGB / 12.92 : Math.pow((GsRGB + 0.055) / 1.055, 2.4);
        const B = BsRGB <= 0.03928 ? BsRGB / 12.92 : Math.pow((BsRGB + 0.055) / 1.055, 2.4);

        return 0.2126 * R + 0.7152 * G + 0.0722 * B;
    },

    /**
     * 设置 Alpha 通道值
     * @param {number} value - Alpha 值 (0-1)
     * @returns {TinyColor}
     */
    setAlpha: function(value) {
        this._a = boundAlpha(value);
        this._roundA = MATH_ROUND(100 * this._a) / 100;
        return this;
    },

    /**
     * 转换为 HSV 对象
     * @returns {Object}
     */
    toHsv: function() {
        const hsv = rgbToHsv(this._r, this._g, this._b);
        return { h: hsv.h * 360, s: hsv.s, v: hsv.v, a: this._a };
    },

    /**
     * 转换为 HSV 字符串
     * @returns {string}
     */
    toHsvString: function() {
        const hsv = rgbToHsv(this._r, this._g, this._b);
        const h = MATH_ROUND(hsv.h * 360);
        const s = MATH_ROUND(hsv.s * 100);
        const v = MATH_ROUND(hsv.v * 100);
        return this._a === 1
            ? `hsv(${h}, ${s}%, ${v}%)`
            : `hsva(${h}, ${s}%, ${v}%, ${this._roundA})`;
    },

    /**
     * 转换为 HSL 对象
     * @returns {Object}
     */
    toHsl: function() {
        const hsl = rgbToHsl(this._r, this._g, this._b);
        return { h: hsl.h * 360, s: hsl.s, l: hsl.l, a: this._a };
    },

    /**
     * 转换为 HSL 字符串
     * @returns {string}
     */
    toHslString: function() {
        const hsl = rgbToHsl(this._r, this._g, this._b);
        const h = MATH_ROUND(hsl.h * 360);
        const s = MATH_ROUND(hsl.s * 100);
        const l = MATH_ROUND(hsl.l * 100);
        return this._a === 1
            ? `hsl(${h}, ${s}%, ${l}%)`
            : `hsla(${h}, ${s}%, ${l}%, ${this._roundA})`;
    },

    /**
     * 转换为 Hex
     * @param {boolean} [allow3Char] - 允许 3 字符简写
     * @returns {string}
     */
    toHex: function(allow3Char) {
        return rgbToHex(this._r, this._g, this._b, allow3Char);
    },

    /**
     * 转换为 Hex 字符串
     * @param {boolean} [allow3Char] - 允许 3 字符简写
     * @returns {string}
     */
    toHexString: function(allow3Char) {
        return '#' + this.toHex(allow3Char);
    },

    /**
     * 转换为带 Alpha 的 Hex
     * @param {boolean} [allow4Char] - 允许 4 字符简写
     * @returns {string}
     */
    toHex8: function(allow4Char) {
        return rgbaToHex(this._r, this._g, this._b, this._a, allow4Char);
    },

    /**
     * 转换为带 Alpha 的 Hex 字符串
     * @param {boolean} [allow4Char] - 允许 4 字符简写
     * @returns {string}
     */
    toHex8String: function(allow4Char) {
        return '#' + this.toHex8(allow4Char);
    },

    /**
     * 转换为 RGB 对象
     * @returns {Object}
     */
    toRgb: function() {
        return {
            r: MATH_ROUND(this._r),
            g: MATH_ROUND(this._g),
            b: MATH_ROUND(this._b),
            a: this._a
        };
    },

    /**
     * 转换为 RGB 字符串
     * @returns {string}
     */
    toRgbString: function() {
        const r = MATH_ROUND(this._r);
        const g = MATH_ROUND(this._g);
        const b = MATH_ROUND(this._b);
        return this._a === 1
            ? `rgb(${r}, ${g}, ${b})`
            : `rgba(${r}, ${g}, ${b}, ${this._roundA})`;
    },

    /**
     * 转换为百分比 RGB 对象
     * @returns {Object}
     */
    toPercentageRgb: function() {
        return {
            r: MATH_ROUND(bound01(this._r, 255) * 100) + '%',
            g: MATH_ROUND(bound01(this._g, 255) * 100) + '%',
            b: MATH_ROUND(bound01(this._b, 255) * 100) + '%',
            a: this._a
        };
    },

    /**
     * 转换为百分比 RGB 字符串
     * @returns {string}
     */
    toPercentageRgbString: function() {
        const r = MATH_ROUND(bound01(this._r, 255) * 100);
        const g = MATH_ROUND(bound01(this._g, 255) * 100);
        const b = MATH_ROUND(bound01(this._b, 255) * 100);
        return this._a === 1
            ? `rgb(${r}%, ${g}%, ${b}%)`
            : `rgba(${r}%, ${g}%, ${b}%, ${this._roundA})`;
    },

    /**
     * 获取颜色名称
     * @returns {string|false}
     */
    toName: function() {
        if (this._a === 0) {
            return 'transparent';
        }
        if (this._a < 1) {
            return false;
        }
        return hexNames[rgbToHex(this._r, this._g, this._b, true)] || false;
    },

    /**
     * 加深颜色
     * @param {number} [amount=10] - 加深量 (0-100)
     * @returns {TinyColor}
     */
    darken: function(amount) {
        amount = amount === 0 ? 0 : (amount || 10);
        const hsl = this.toHsl();
        hsl.l -= amount / 100;
        hsl.l = clamp01(hsl.l);
        return new TinyColor(hsl);
    },

    /**
     * 减淡颜色
     * @param {number} [amount=10] - 减淡量 (0-100)
     * @returns {TinyColor}
     */
    lighten: function(amount) {
        amount = amount === 0 ? 0 : (amount || 10);
        const hsl = this.toHsl();
        hsl.l += amount / 100;
        hsl.l = clamp01(hsl.l);
        return new TinyColor(hsl);
    },

    /**
     * 增加饱和度
     * @param {number} [amount=10] - 增加量 (0-100)
     * @returns {TinyColor}
     */
    saturate: function(amount) {
        amount = amount === 0 ? 0 : (amount || 10);
        const hsl = this.toHsl();
        hsl.s += amount / 100;
        hsl.s = clamp01(hsl.s);
        return new TinyColor(hsl);
    },

    /**
     * 降低饱和度
     * @param {number} [amount=10] - 降低量 (0-100)
     * @returns {TinyColor}
     */
    desaturate: function(amount) {
        amount = amount === 0 ? 0 : (amount || 10);
        const hsl = this.toHsl();
        hsl.s -= amount / 100;
        hsl.s = clamp01(hsl.s);
        return new TinyColor(hsl);
    }
};

// =============================================================================
// 辅助函数
// =============================================================================

/**
 * 解析输入颜色
 * @param {string|Object} color - 颜色输入
 * @returns {Object} RGB 对象
 */
function inputToRGB(color) {
    let rgb = { r: 0, g: 0, b: 0 };
    let a = 1;
    let ok = false;
    let format = false;

    if (typeof color === 'string') {
        color = stringInputToObject(color);
    }

    if (typeof color === 'object') {
        if (isValidCSSUnit(color.r) && isValidCSSUnit(color.g) && isValidCSSUnit(color.b)) {
            rgb = rgbToRgb(color.r, color.g, color.b);
            ok = true;
            format = String(color.r).substr(-1) === '%' ? 'prgb' : 'rgb';
        } else if (isValidCSSUnit(color.h) && isValidCSSUnit(color.s) && isValidCSSUnit(color.v)) {
            color.s = convertToPercentage(color.s);
            color.v = convertToPercentage(color.v);
            rgb = hsvToRgb(color.h, color.s, color.v);
            ok = true;
            format = 'hsv';
        } else if (isValidCSSUnit(color.h) && isValidCSSUnit(color.s) && isValidCSSUnit(color.l)) {
            color.s = convertToPercentage(color.s);
            color.l = convertToPercentage(color.l);
            rgb = hslToRgb(color.h, color.s, color.l);
            ok = true;
            format = 'hsl';
        }

        if (color.hasOwnProperty('a')) {
            a = color.a;
        }
    }

    a = boundAlpha(a);

    return {
        ok: ok,
        format: color.format || format,
        r: MATH_MIN(255, MATH_MAX(rgb.r, 0)),
        g: MATH_MIN(255, MATH_MAX(rgb.g, 0)),
        b: MATH_MIN(255, MATH_MAX(rgb.b, 0)),
        a: a
    };
}

/**
 * 标准化 RGB 值
 */
function rgbToRgb(r, g, b) {
    return {
        r: bound01(r, 255) * 255,
        g: bound01(g, 255) * 255,
        b: bound01(b, 255) * 255
    };
}

/**
 * RGB 转 HSL
 */
function rgbToHsl(r, g, b) {
    r = bound01(r, 255);
    g = bound01(g, 255);
    b = bound01(b, 255);

    const max = MATH_MAX(r, g, b);
    const min = MATH_MIN(r, g, b);
    let h, s;
    const l = (max + min) / 2;

    if (max === min) {
        h = s = 0;
    } else {
        const d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        switch (max) {
            case r: h = (g - b) / d + (g < b ? 6 : 0); break;
            case g: h = (b - r) / d + 2; break;
            case b: h = (r - g) / d + 4; break;
        }
        h /= 6;
    }

    return { h: h, s: s, l: l };
}

/**
 * HSL 转 RGB 的辅助函数
 */
function hue2rgb(p, q, t) {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1/6) return p + (q - p) * 6 * t;
    if (t < 1/2) return q;
    if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
    return p;
}

/**
 * HSL 转 RGB
 */
function hslToRgb(h, s, l) {
    let r, g, b;

    h = bound01(h, 360);
    s = bound01(s, 100);
    l = bound01(l, 100);

    if (s === 0) {
        r = g = b = l;
    } else {
        const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        const p = 2 * l - q;
        r = hue2rgb(p, q, h + 1/3);
        g = hue2rgb(p, q, h);
        b = hue2rgb(p, q, h - 1/3);
    }

    return { r: r * 255, g: g * 255, b: b * 255 };
}

/**
 * RGB 转 HSV
 */
function rgbToHsv(r, g, b) {
    r = bound01(r, 255);
    g = bound01(g, 255);
    b = bound01(b, 255);

    const max = MATH_MAX(r, g, b);
    const min = MATH_MIN(r, g, b);
    let h;
    const v = max;

    const d = max - min;
    const s = max === 0 ? 0 : d / max;

    if (max === min) {
        h = 0;
    } else {
        switch (max) {
            case r: h = (g - b) / d + (g < b ? 6 : 0); break;
            case g: h = (b - r) / d + 2; break;
            case b: h = (r - g) / d + 4; break;
        }
        h /= 6;
    }

    return { h: h, s: s, v: v };
}

/**
 * HSV 转 RGB
 */
function hsvToRgb(h, s, v) {
    h = bound01(h, 360) * 6;
    s = bound01(s, 100);
    v = bound01(v, 100);

    const i = Math.floor(h);
    const f = h - i;
    const p = v * (1 - s);
    const q = v * (1 - f * s);
    const t = v * (1 - (1 - f) * s);
    const mod = i % 6;

    const r = [v, q, p, p, t, v][mod];
    const g = [t, v, v, q, p, p][mod];
    const b = [p, p, t, v, v, q][mod];

    return { r: r * 255, g: g * 255, b: b * 255 };
}

/**
 * RGB 转 Hex
 */
function rgbToHex(r, g, b, allow3Char) {
    const hex = [
        pad2(MATH_ROUND(r).toString(16)),
        pad2(MATH_ROUND(g).toString(16)),
        pad2(MATH_ROUND(b).toString(16))
    ];

    if (allow3Char && 
        hex[0].charAt(0) === hex[0].charAt(1) && 
        hex[1].charAt(0) === hex[1].charAt(1) && 
        hex[2].charAt(0) === hex[2].charAt(1)) {
        return hex[0].charAt(0) + hex[1].charAt(0) + hex[2].charAt(0);
    }

    return hex.join('');
}

/**
 * RGBA 转 Hex8
 */
function rgbaToHex(r, g, b, a, allow4Char) {
    const hex = [
        pad2(MATH_ROUND(r).toString(16)),
        pad2(MATH_ROUND(g).toString(16)),
        pad2(MATH_ROUND(b).toString(16)),
        pad2(convertDecimalToHex(a))
    ];

    if (allow4Char && 
        hex[0].charAt(0) === hex[0].charAt(1) && 
        hex[1].charAt(0) === hex[1].charAt(1) && 
        hex[2].charAt(0) === hex[2].charAt(1) && 
        hex[3].charAt(0) === hex[3].charAt(1)) {
        return hex[0].charAt(0) + hex[1].charAt(0) + hex[2].charAt(0) + hex[3].charAt(0);
    }

    return hex.join('');
}

// =============================================================================
// 工具函数
// =============================================================================

function bound01(n, max) {
    if (isOnePointZero(n)) {
        n = '100%';
    }

    const processPercent = isPercentage(n);
    n = MATH_MIN(max, MATH_MAX(0, parseFloat(n)));

    if (processPercent) {
        n = parseInt(n * max, 10) / 100;
    }

    if (Math.abs(n - max) < 0.000001) {
        return 1;
    }

    return (n % max) / parseFloat(max);
}

function boundAlpha(a) {
    a = parseFloat(a);

    if (isNaN(a) || a < 0 || a > 1) {
        return 1;
    }

    return a;
}

function clamp01(val) {
    return MATH_MIN(1, MATH_MAX(0, val));
}

function isOnePointZero(n) {
    return typeof n === 'string' && n.indexOf('.') !== -1 && parseFloat(n) === 1;
}

function isPercentage(n) {
    return typeof n === 'string' && n.indexOf('%') !== -1;
}

function pad2(c) {
    return c.length === 1 ? '0' + c : '' + c;
}

function convertToPercentage(n) {
    if (n <= 1) {
        n = (n * 100) + '%';
    }
    return n;
}

function convertDecimalToHex(d) {
    return Math.round(parseFloat(d) * 255).toString(16);
}

function isValidCSSUnit(color) {
    return !!MATCHERS.CSS_UNIT.exec(color);
}

function stringInputToObject(color) {
    color = color.replace(TRIM_LEFT, '').replace(TRIM_RIGHT, '').toLowerCase();
    
    let named = false;
    if (CSS_NAMES[color]) {
        color = CSS_NAMES[color];
        named = true;
    } else if (color === 'transparent') {
        return { r: 0, g: 0, b: 0, a: 0, format: 'name' };
    }

    let match;
    if ((match = MATCHERS.rgb.exec(color))) {
        return { r: match[1], g: match[2], b: match[3] };
    }
    if ((match = MATCHERS.rgba.exec(color))) {
        return { r: match[1], g: match[2], b: match[3], a: match[4] };
    }
    if ((match = MATCHERS.hsl.exec(color))) {
        return { h: match[1], s: match[2], l: match[3] };
    }
    if ((match = MATCHERS.hsla.exec(color))) {
        return { h: match[1], s: match[2], l: match[3], a: match[4] };
    }
    if ((match = MATCHERS.hsv.exec(color))) {
        return { h: match[1], s: match[2], v: match[3] };
    }
    if ((match = MATCHERS.hsva.exec(color))) {
        return { h: match[1], s: match[2], v: match[3], a: match[4] };
    }
    if ((match = MATCHERS.hex8.exec(color))) {
        return {
            r: parseIntFromHex(match[1]),
            g: parseIntFromHex(match[2]),
            b: parseIntFromHex(match[3]),
            a: convertHexToDecimal(match[4]),
            format: named ? 'name' : 'hex8'
        };
    }
    if ((match = MATCHERS.hex6.exec(color))) {
        return {
            r: parseIntFromHex(match[1]),
            g: parseIntFromHex(match[2]),
            b: parseIntFromHex(match[3]),
            format: named ? 'name' : 'hex'
        };
    }
    if ((match = MATCHERS.hex4.exec(color))) {
        return {
            r: parseIntFromHex(match[1] + match[1]),
            g: parseIntFromHex(match[2] + match[2]),
            b: parseIntFromHex(match[3] + match[3]),
            a: convertHexToDecimal(match[4] + match[4]),
            format: named ? 'name' : 'hex8'
        };
    }
    if ((match = MATCHERS.hex3.exec(color))) {
        return {
            r: parseIntFromHex(match[1] + match[1]),
            g: parseIntFromHex(match[2] + match[2]),
            b: parseIntFromHex(match[3] + match[3]),
            format: named ? 'name' : 'hex'
        };
    }

    return false;
}

function parseIntFromHex(val) {
    return parseInt(val, 16);
}

function convertHexToDecimal(val) {
    return parseIntFromHex(val) / 255;
}

// =============================================================================
// 正则表达式匹配器
// =============================================================================

const MATCHERS = (function() {
    const CSS_INTEGER = '[-\\+]?\\d+%?';
    const CSS_NUMBER = '[-\\+]?\\d*\\.\\d+%?';
    const CSS_UNIT = '(?:' + CSS_NUMBER + ')|(?:' + CSS_INTEGER + ')';
    const PERMISSIVE_MATCH3 = '[\\s|\\(]+(' + CSS_UNIT + ')[,|\\s]+(' + CSS_UNIT + ')[,|\\s]+(' + CSS_UNIT + ')\\s*\\)?';
    const PERMISSIVE_MATCH4 = '[\\s|\\(]+(' + CSS_UNIT + ')[,|\\s]+(' + CSS_UNIT + ')[,|\\s]+(' + CSS_UNIT + ')[,|\\s]+(' + CSS_UNIT + ')\\s*\\)?';

    return {
        CSS_UNIT: new RegExp(CSS_UNIT),
        rgb: new RegExp('rgb' + PERMISSIVE_MATCH3),
        rgba: new RegExp('rgba' + PERMISSIVE_MATCH4),
        hsl: new RegExp('hsl' + PERMISSIVE_MATCH3),
        hsla: new RegExp('hsla' + PERMISSIVE_MATCH4),
        hsv: new RegExp('hsv' + PERMISSIVE_MATCH3),
        hsva: new RegExp('hsva' + PERMISSIVE_MATCH4),
        hex3: /^#?([0-9a-fA-F]{1})([0-9a-fA-F]{1})([0-9a-fA-F]{1})$/,
        hex6: /^#?([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/,
        hex4: /^#?([0-9a-fA-F]{1})([0-9a-fA-F]{1})([0-9a-fA-F]{1})([0-9a-fA-F]{1})$/,
        hex8: /^#?([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/
    };
})();

// =============================================================================
// CSS 颜色名称
// =============================================================================

const CSS_NAMES = {
    aliceblue: 'f0f8ff',
    antiquewhite: 'faebd7',
    aqua: '0ff',
    aquamarine: '7fffd4',
    azure: 'f0ffff',
    beige: 'f5f5dc',
    bisque: 'ffe4c4',
    black: '000',
    blanchedalmond: 'ffebcd',
    blue: '00f',
    blueviolet: '8a2be2',
    brown: 'a52a2a',
    burlywood: 'deb887',
    burntsienna: 'ea7e5d',
    cadetblue: '5f9ea0',
    chartreuse: '7fff00',
    chocolate: 'd2691e',
    coral: 'ff7f50',
    cornflowerblue: '6495ed',
    cornsilk: 'fff8dc',
    crimson: 'dc143c',
    cyan: '0ff',
    darkblue: '00008b',
    darkcyan: '008b8b',
    darkgoldenrod: 'b8860b',
    darkgray: 'a9a9a9',
    darkgreen: '006400',
    darkgrey: 'a9a9a9',
    darkkhaki: 'bdb76b',
    darkmagenta: '8b008b',
    darkolivegreen: '556b2f',
    darkorange: 'ff8c00',
    darkorchid: '9932cc',
    darkred: '8b0000',
    darksalmon: 'e9967a',
    darkseagreen: '8fbc8f',
    darkslateblue: '483d8b',
    darkslategray: '2f4f4f',
    darkslategrey: '2f4f4f',
    darkturquoise: '00ced1',
    darkviolet: '9400d3',
    deeppink: 'ff1493',
    deepskyblue: '00bfff',
    dimgray: '696969',
    dimgrey: '696969',
    dodgerblue: '1e90ff',
    firebrick: 'b22222',
    floralwhite: 'fffaf0',
    forestgreen: '228b22',
    fuchsia: 'f0f',
    gainsboro: 'dcdcdc',
    ghostwhite: 'f8f8ff',
    gold: 'ffd700',
    goldenrod: 'daa520',
    gray: '808080',
    green: '008000',
    greenyellow: 'adff2f',
    grey: '808080',
    honeydew: 'f0fff0',
    hotpink: 'ff69b4',
    indianred: 'cd5c5c',
    indigo: '4b0082',
    ivory: 'fffff0',
    khaki: 'f0e68c',
    lavender: 'e6e6fa',
    lavenderblush: 'fff0f5',
    lawngreen: '7cfc00',
    lemonchiffon: 'fffacd',
    lightblue: 'add8e6',
    lightcoral: 'f08080',
    lightcyan: 'e0ffff',
    lightgoldenrodyellow: 'fafad2',
    lightgray: 'd3d3d3',
    lightgreen: '90ee90',
    lightgrey: 'd3d3d3',
    lightpink: 'ffb6c1',
    lightsalmon: 'ffa07a',
    lightseagreen: '20b2aa',
    lightskyblue: '87cefa',
    lightslategray: '789',
    lightslategrey: '789',
    lightsteelblue: 'b0c4de',
    lightyellow: 'ffffe0',
    lime: '0f0',
    limegreen: '32cd32',
    linen: 'faf0e6',
    magenta: 'f0f',
    maroon: '800000',
    mediumaquamarine: '66cdaa',
    mediumblue: '0000cd',
    mediumorchid: 'ba55d3',
    mediumpurple: '9370db',
    mediumseagreen: '3cb371',
    mediumslateblue: '7b68ee',
    mediumspringgreen: '00fa9a',
    mediumturquoise: '48d1cc',
    mediumvioletred: 'c71585',
    midnightblue: '191970',
    mintcream: 'f5fffa',
    mistyrose: 'ffe4e1',
    moccasin: 'ffe4b5',
    navajowhite: 'ffdead',
    navy: '000080',
    oldlace: 'fdf5e6',
    olive: '808000',
    olivedrab: '6b8e23',
    orange: 'ffa500',
    orangered: 'ff4500',
    orchid: 'da70d6',
    palegoldenrod: 'eee8aa',
    palegreen: '98fb98',
    paleturquoise: 'afeeee',
    palevioletred: 'db7093',
    papayawhip: 'ffefd5',
    peachpuff: 'ffdab9',
    peru: 'cd853f',
    pink: 'ffc0cb',
    plum: 'dda0dd',
    powderblue: 'b0e0e6',
    purple: '800080',
    rebeccapurple: '663399',
    red: 'f00',
    rosybrown: 'bc8f8f',
    royalblue: '4169e1',
    saddlebrown: '8b4513',
    salmon: 'fa8072',
    sandybrown: 'f4a460',
    seagreen: '2e8b57',
    seashell: 'fff5ee',
    sienna: 'a0522d',
    silver: 'c0c0c0',
    skyblue: '87ceeb',
    slateblue: '6a5acd',
    slategray: '708090',
    slategrey: '708090',
    snow: 'fffafa',
    springgreen: '00ff7f',
    steelblue: '4682b4',
    tan: 'd2b48c',
    teal: '008080',
    thistle: 'd8bfd8',
    tomato: 'ff6347',
    turquoise: '40e0d0',
    violet: 'ee82ee',
    wheat: 'f5deb3',
    white: 'fff',
    whitesmoke: 'f5f5f5',
    yellow: 'ff0',
    yellowgreen: '9acd32'
};

// 创建反向映射
const hexNames = {};
for (const key in CSS_NAMES) {
    hexNames[CSS_NAMES[key]] = key;
}
