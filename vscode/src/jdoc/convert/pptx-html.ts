/**
 * PPTX -> HTML（版式还原）
 *
 * 与 PPTX->PDF 的「幻灯片版式直排」是同一套几何语义：
 * 每页一张与幻灯片等大的画布（默认 1280x720），元素按文档顺序绝对定位，
 * 形状走 SVG、文本走 HTML（flex 对齐），因此渲染结果与 PDF / 编辑器预览一致。
 *
 * 与 blocks 路线（pptx.ts）的区别：那条链路只抽文本，天然丢版面。
 *
 * @module convert/pptx-html
 */

import { pptxToStandard, presetShapePath } from '../pptx/index';
import type {
    PptxBullet,
    PptxDocument,
    PptxElement,
    PptxFill,
    PptxGroupElement,
    PptxImageElement,
    PptxParagraph,
    PptxShapeElement,
    PptxSlide,
    PptxTableElement,
    PptxTextElement,
    PptxTextRun
} from '@fefeding/ppt-parser';
import type { DocInput } from '../core/input';
import { wrapError } from '../core/errors';
import {
    DASH_PATTERNS,
    flipOn,
    gradientAt,
    gradientAxis,
    gradientStopsOf,
    normalizeColor,
    rgbOf,
    roundedRectPathLocal,
    shapeCornerRadii,
    solidFillOf,
    withThemeColors,
    type GradientStop
} from './pptx-pdf';

/** PPTX -> HTML 选项 */
export interface PptxToHtmlOptions {
    /** 页面标题（缺省用文件名） */
    title?: string;
}

/**
 * 默认字体栈（与编辑器预览端一致：微软雅黑优先，回退系统无衬线）
 *
 * 字体名一律用**单引号**：这些值会被拼进 HTML 的 style="..." 属性，
 * 双引号会提前闭合属性（背景图那次踩过同样的坑）。
 */
const FONT_STACK =
    "'Microsoft YaHei','微软雅黑','PingFang SC','Hiragino Sans GB','Source Han Sans SC'," +
    "'Noto Sans CJK SC',-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif";

/** HTML 转义 */
function esc(value: string): string {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/** 数字保留两位小数（避免样式里出现长浮点尾数） */
function num(value: number): string {
    return String(Math.round(value * 100) / 100);
}

/** 路径段 -> SVG path 的 d 属性 */
function pathToD(segments: { op: string; args: number[] }[]): string {
    const parts: string[] = [];
    for (const seg of segments) {
        if (seg.op === 'm') parts.push(`M ${num(seg.args[0])} ${num(seg.args[1])}`);
        else if (seg.op === 'l') parts.push(`L ${num(seg.args[0])} ${num(seg.args[1])}`);
        else if (seg.op === 'c')
            parts.push(
                `C ${num(seg.args[0])} ${num(seg.args[1])} ${num(seg.args[2])} ${num(seg.args[3])} ${num(seg.args[4])} ${num(seg.args[5])}`
            );
        else if (seg.op === 'h') parts.push('Z');
    }
    return parts.join(' ');
}

/** 渐变停点 -> SVG linearGradient 的 stop 列表 */
function gradientStopsSvg(stops: GradientStop[]): string {
    return stops
        .map((s) => {
            const { color, alpha } = gradientAt([s], 0);
            const offset = Math.round((s.position ?? 0) * 1000) / 10;
            const op = alpha < 1 ? ` stop-opacity="${num(alpha)}"` : '';
            return `<stop offset="${num(offset)}%" stop-color="${color}"${op}/>`;
        })
        .join('');
}

/** 带箭头信息的线条样式（解析器类型未声明 arrow 字段，运行时存在） */
interface ConnectorLine {
    color?: string;
    width?: number;
    dashType?: string;
    startArrow?: string;
    endArrow?: string;
}

/** 形状的描边样式（线宽单位是 pt，HTML 里换算成 px，与 PDF 端视觉一致） */
function strokeAttrs(el: PptxShapeElement): { color?: string; width: number; dash?: string } {
    const line = el.line && el.line !== 'none' ? (el.line as ConnectorLine) : undefined;
    const color = normalizeColor(line?.color);
    const width = line?.width ? Math.max(0.5, line.width * (96 / 72)) : 0;
    const dashArr = DASH_PATTERNS[line?.dashType ?? ''];
    return {
        color,
        width,
        dash: dashArr ? ` stroke-dasharray="${dashArr.map((d) => num(d)).join(' ')}"` : undefined
    };
}

/** 单个 SVG 片段（自带定位样式） */
function svgBox(x: number, y: number, w: number, h: number, inner: string, rotation?: number): string {
    const rot = rotation ? ` transform:rotate(${num(rotation)}deg);` : '';
    return (
        `<svg style="left:${num(x)}px;top:${num(y)}px;width:${num(Math.max(w, 0.1))}px;` +
        `height:${num(Math.max(h, 0.1))}px;${rot}" viewBox="0 0 ${num(Math.max(w, 0.1))} ${num(
            Math.max(h, 0.1)
        )}" overflow="visible" xmlns="http://www.w3.org/2000/svg">${inner}</svg>`
    );
}

/** 连接线 / 直线 -> SVG（与编辑器渲染端的连接线分支同源） */
function connectorToSvg(el: PptxShapeElement): string {
    const st = (el.shapeType || '').toLowerCase();
    const { x, y, width: w, height: h } = el;
    const flipH = flipOn(el.flipH);
    const flipV = flipOn(el.flipV);
    const line = el.line && el.line !== 'none' ? (el.line as ConnectorLine) : undefined;
    const color = normalizeColor(line?.color) ?? '#000000';
    const lw = Math.max(1, (line?.width ?? 2) * (96 / 72));
    const dashArr = DASH_PATTERNS[line?.dashType ?? ''];
    const dash = dashArr ? ` stroke-dasharray="${dashArr.map((d) => num(d)).join(' ')}"` : '';
    const isPlainLine = st === 'line';
    const hasArrow = !isPlainLine && !!(line?.endArrow || line?.startArrow);

    /** 局部坐标（0,0 ~ w,h）-> 绝对坐标（含镜像） */
    const P = (lx: number, ly: number): [number, number] => [
        x + (flipH ? w - lx : lx),
        y + (flipV ? h - ly : ly)
    ];

    if (st.startsWith('bentconnector')) {
        // M 0,0 L W/2,0 L W/2,H L W,H（开放路径）
        const pts = [P(0, 0), P(w / 2, 0), P(w / 2, h), P(w, h)];
        const d = pts.map((p, i) => `${i ? 'L' : 'M'} ${num(p[0] - x)} ${num(p[1] - y)}`).join(' ');
        return svgBox(
            x,
            y,
            w,
            h,
            `<path d="${d}" fill="none" stroke="${color}" stroke-width="${num(lw)}"${dash}/>`,
            el.rotation
        );
    }

    if (st.startsWith('curvedconnector')) {
        const r = Math.min(Math.max((el.adjust?.adj1 ?? 50000) / 100000, 0), 1);
        const p0 = P(0, 0);
        const c1 = P(w * r, 0);
        const mid = P(w / 2, h / 2);
        const c2 = P(w * (1 - r), h);
        const p1 = P(w, h);
        const q = (a: [number, number], c: [number, number], b: [number, number]) =>
            `Q ${num(c[0] - x)} ${num(c[1] - y)} ${num(b[0] - x)} ${num(b[1] - y)}`;
        const d = `M ${num(p0[0] - x)} ${num(p0[1] - y)} ${q(p0, c1, mid)} ${q(mid, c2, p1)}`;
        return svgBox(
            x,
            y,
            w,
            h,
            `<path d="${d}" fill="none" stroke="${color}" stroke-width="${num(lw)}"${dash}/>`,
            el.rotation
        );
    }

    // straightConnector1 / line：左上 -> 右下
    const a = P(0, 0);
    const b = P(w, h);
    // marker id 必须在同一个 svg 内唯一
    const markerId = `ah${uid()}`;
    const defs =
        hasArrow
            ? `<defs><marker id="${markerId}" markerWidth="5" markerHeight="5" refX="10" refY="5" ` +
              `orient="auto-start-reverse" markerUnits="strokeWidth" viewBox="0 0 10 10">` +
              `<path d="M 0 0 L 10 5 L 0 10 z" fill="${color}"/></marker></defs>`
            : '';
    const endMarker =
        hasArrow && line?.endArrow && line.endArrow !== 'none'
            ? ` marker-end="url(#${markerId})"`
            : '';
    const startMarker =
        hasArrow && line?.startArrow && line.startArrow !== 'none'
            ? ` marker-start="url(#${markerId})"`
            : '';
    const lineEl =
        `<line x1="${num(a[0] - x)}" y1="${num(a[1] - y)}" x2="${num(b[0] - x)}" y2="${num(
            b[1] - y
        )}" stroke="${color}" stroke-width="${num(lw)}"${dash}${endMarker}${startMarker}/>`;
    return svgBox(x, y, w, h, defs + lineEl, el.rotation);
}

let uidCounter = 0;
/** 生成文档内唯一的 id（渐变 / marker 用） */
function uid(): string {
    uidCounter += 1;
    return String(uidCounter);
}

/** 块箭头（rightArrow / downArrow）-> SVG 多边形 */
function blockArrowToSvg(el: PptxShapeElement): string {
    const { x, y, width: w, height: h } = el;
    const ss = Math.min(w, h);
    const adj = el.adjust || {};
    const body = ss * ((adj.adj1 ?? 50000) / 100000);
    const head = ss * ((adj.adj2 ?? 50000) / 100000);
    const fill = solidFillOf(el.fill);
    if (!fill) return '';
    const stroke = strokeAttrs(el);
    let pts: [number, number][];
    if ((el.shapeType || '').toLowerCase() === 'downarrow') {
        const left = (w - body) / 2;
        const right = (w + body) / 2;
        pts = [
            [left, 0],
            [right, 0],
            [right, h - head],
            [w, h - head],
            [w / 2, h],
            [0, h - head],
            [left, h - head]
        ];
    } else {
        const top = (h - body) / 2;
        const bot = (h + body) / 2;
        pts = [
            [0, top],
            [w - head, top],
            [w - head, 0],
            [w, h / 2],
            [w - head, h],
            [w - head, bot],
            [0, bot]
        ];
    }
    const d = pts.map((p, i) => `${i ? 'L' : 'M'} ${num(p[0])} ${num(p[1])}`).join(' ') + ' Z';
    const fillAttr = `fill="${fill.color}"${fill.alpha < 1 ? ` fill-opacity="${num(fill.alpha)}"` : ''}`;
    const strokeAttr = stroke.color
        ? ` stroke="${stroke.color}" stroke-width="${num(stroke.width)}"`
        : '';
    return svgBox(
        x,
        y,
        w,
        h,
        `<path d="${d}" ${fillAttr}${strokeAttr}/>`,
        el.rotation
    );
}

/** 圆柱 / 数据库符号（can / cylinder / flowChartMagneticDisk）-> SVG */
function canToSvg(el: PptxShapeElement): string {
    const { x, y, width: w, height: h } = el;
    const ss = Math.min(w, h);
    const ry = (ss * ((el.adjust?.adj ?? 25000) / 100000)) / 2;
    const fill = solidFillOf(el.fill);
    if (!fill || ry <= 0 || h <= 2 * ry) return '';
    const stroke = strokeAttrs(el);
    const fillAttr = `fill="${fill.color}"${fill.alpha < 1 ? ` fill-opacity="${num(fill.alpha)}"` : ''}`;
    const strokeAttr = stroke.color
        ? ` stroke="${stroke.color}" stroke-width="${num(stroke.width)}"`
        : '';
    const bodyTop = ry;
    const bodyH = h - 2 * ry;
    const inner =
        `<ellipse cx="${num(w / 2)}" cy="${num(h - ry)}" rx="${num(w / 2)}" ry="${num(ry)}" ${fillAttr}${strokeAttr}/>` +
        `<rect x="0" y="${num(bodyTop)}" width="${num(w)}" height="${num(bodyH)}" ${fillAttr}${strokeAttr}/>` +
        `<ellipse cx="${num(w / 2)}" cy="${num(ry)}" rx="${num(w / 2)}" ry="${num(ry)}" ${fillAttr}${strokeAttr}/>`;
    return svgBox(x, y, w, h, inner, el.rotation);
}

/** 常规形状（矩形 / 圆角矩形 / 椭圆）-> SVG */
function shapeToSvg(el: PptxShapeElement): string {
    const { x, y, width: w, height: h } = el;
    if (!w || !h) return '';
    const st = (el.shapeType || '').toLowerCase();
    const isEllipse = st === 'ellipse';
    const radii = shapeCornerRadii(el.shapeType, el.adjust, w, h);
    const uniform = radii[0] === radii[1] && radii[1] === radii[2] && radii[2] === radii[3];
    const rx = uniform ? radii[0] : radii[0];
    const stroke = strokeAttrs(el);
    const strokeAttr = stroke.color
        ? ` stroke="${stroke.color}" stroke-width="${num(stroke.width)}"${stroke.dash ?? ''}`
        : '';
    const grad = gradientStopsOf(el.fill);
    let fillAttr = 'fill="none"';
    let defs = '';
    if (grad) {
        const axis = gradientAxis(grad.direction, grad.angle, w, h);
        const x1 = axis.p1[0];
        const y1 = axis.p1[1];
        const x2 = axis.p1[0] + axis.dir[0] * axis.len;
        const y2 = axis.p1[1] + axis.dir[1] * axis.len;
        const gid = `grad${uid()}`;
        defs =
            `<defs><linearGradient id="${gid}" gradientUnits="userSpaceOnUse" ` +
            `x1="${num(x1)}" y1="${num(y1)}" x2="${num(x2)}" y2="${num(y2)}">` +
            gradientStopsSvg(grad.stops) +
            '</linearGradient></defs>';
        fillAttr = `fill="url(#${gid})"`;
    } else {
        const fill = solidFillOf(el.fill);
        if (fill) {
            fillAttr = `fill="${fill.color}"${fill.alpha < 1 ? ` fill-opacity="${num(fill.alpha)}"` : ''}`;
        }
    }
    // 非等半径圆角用路径表达（SVG rect 只支持统一 rx）
    if (!isEllipse && !uniform && radii.some((r) => r > 0)) {
        const d = pathToD(roundedRectPathLocal(w, h, radii));
        return svgBox(x, y, w, h, `${defs}<path d="${d}" ${fillAttr}${strokeAttr}/>`, el.rotation);
    }
    if (isEllipse) {
        return svgBox(
            x,
            y,
            w,
            h,
            `${defs}<ellipse cx="${num(w / 2)}" cy="${num(h / 2)}" rx="${num(w / 2)}" ry="${num(
                h / 2
            )}" ${fillAttr}${strokeAttr}/>`,
            el.rotation
        );
    }
    return svgBox(
        x,
        y,
        w,
        h,
        `${defs}<rect x="0" y="0" width="${num(w)}" height="${num(h)}"${rx > 0 ? ` rx="${num(rx)}"` : ''} ${fillAttr}${strokeAttr}/>`,
        el.rotation
    );
}

/**
 * 预设几何 -> SVG（与预览端同源公式：直接用解析器的 presetShapePath）
 *
 * 解析器已实现 OOXML 预设形状的真实路径（圆柱 flowChartMagneticDisk、星形、
 * 弧形等），手工近似会画错；解析器未实现的形状返回 null，交给下面的兜底分支。
 */
function presetShapeSvg(el: PptxShapeElement): string {
    const st = el.shapeType || '';
    const { x, y, width: w, height: h } = el;
    if (!st || !w || !h) return '';
    const geo = presetShapePath(st, w, h, el.adjust || {});
    if (!geo?.d) return '';

    // 镜像（a:xfrm/@flipH/@flipV）：OOXML 先 flip 后 rot
    const flipH = flipOn(el.flipH);
    const flipV = flipOn(el.flipV);
    let tf = geo.transform || '';
    if (flipH || flipV) {
        const sx = flipH ? -1 : 1;
        const sy = flipV ? -1 : 1;
        const tx = flipH ? w : 0;
        const ty = flipV ? h : 0;
        tf += (tf ? ' ' : '') + `translate(${tx},${ty}) scale(${sx},${sy})`;
    }
    const tfAttr = tf.trim() ? ` transform="${tf.trim()}"` : '';

    // 填充
    let defs = '';
    let fillAttr = 'fill="none"';
    if (geo.noFill) {
        fillAttr = 'fill="none"';
    } else {
        const grad = gradientStopsOf(el.fill);
        if (grad) {
            const axis = gradientAxis(grad.direction, grad.angle, w, h);
            const gid = `grad${uid()}`;
            defs =
                `<defs><linearGradient id="${gid}" gradientUnits="userSpaceOnUse" ` +
                `x1="${num(axis.p1[0])}" y1="${num(axis.p1[1])}" ` +
                `x2="${num(axis.p1[0] + axis.dir[0] * axis.len)}" ` +
                `y2="${num(axis.p1[1] + axis.dir[1] * axis.len)}">` +
                gradientStopsSvg(grad.stops) +
                '</linearGradient></defs>';
            fillAttr = `fill="url(#${gid})"`;
        } else {
            const fill = solidFillOf(el.fill);
            if (fill) {
                fillAttr = `fill="${fill.color}"${fill.alpha < 1 ? ` fill-opacity="${num(fill.alpha)}"` : ''}`;
            }
        }
    }

    const stroke = strokeAttrs(el);
    const strokeAttr = stroke.color
        ? ` stroke="${stroke.color}" stroke-width="${num(stroke.width)}"${stroke.dash ?? ''}`
        : '';
    const main =
        `<path d="${geo.d}"${tfAttr} ${fillAttr}${strokeAttr}` +
        `${geo.fillRule ? ` fill-rule="${geo.fillRule}"` : ''}/>`;

    // 附加描边路径（callout 引线、笑脸嘴等）
    const accent = stroke.color ?? '#5B9BD5';
    const extra = (geo.strokes || [])
        .map(
            (s) =>
                `<path d="${s.d}"${tfAttr} fill="none" stroke="${accent}" ` +
                `stroke-width="${num(s.width ?? Math.max(1, w * 0.05))}" stroke-linecap="round"/>`
        )
        .join('');

    return svgBox(x, y, w, h, `${defs}${main}${extra}`, el.rotation);
}

/** 形状 -> SVG 片段 */
function shapeElementToSvg(el: PptxShapeElement): string {
    const st = (el.shapeType || '').toLowerCase();
    if (/^(straightconnector|bentconnector|curvedconnector)/.test(st) || st === 'line') {
        return connectorToSvg(el);
    }
    // 预设几何优先（与预览端同一套公式）
    const preset = presetShapeSvg(el);
    if (preset) return preset;
    if (st === 'rightarrow' || st === 'downarrow') return blockArrowToSvg(el);
    if (st === 'can' || st === 'cylinder' || st === 'flowchartmagneticdisk') return canToSvg(el);
    return shapeToSvg(el);
}

/** 字体名加引号（CSS 里含空格/中文的字体名必须加引号；用单引号避免闭合 HTML 属性） */
function quoteFont(face: string): string {
    return /^[A-Za-z0-9 _-]+$/.test(face) ? face : `'${face.replace(/'/g, "\\'")}'`;
}

/**
 * 拼 font-family：西文（latin）在前、东亚（ea）在后，浏览器按字符集自动回退
 * （中文 PPT 常见 latin="Arial" + ea="微软雅黑"，只写 latin 会让中文走系统默认字形）
 */
function fontFamilyOf(...faces: (string | undefined)[]): string | undefined {
    const out: string[] = [];
    for (const face of faces) {
        if (!face) continue;
        const quoted = quoteFont(face);
        if (!out.includes(quoted)) out.push(quoted);
    }
    return out.length ? out.join(', ') : undefined;
}

/** 主题正文色（tx1），取不到时兜底黑色 */
function themeTextColor(): string {
    return normalizeColor('scheme:tx1') ?? '#000000';
}

/** 行距 -> line-height（与编辑器渲染端同一套语义） */
function lineHeightStyle(spacing: PptxParagraph['lineSpacing']): string {
    if (spacing == null) return '';
    if (typeof spacing === 'number') {
        // 两种约定混存：>=10 是 OOXML 百分值（100=100%），<5 是倍数
        return `line-height:${num(spacing >= 10 ? spacing / 100 : spacing)}`;
    }
    if (spacing.type === 'percent') return `line-height:${num(spacing.value / 100)}`;
    if (spacing.type === 'pt') return `line-height:${num(spacing.value)}pt`;
    return '';
}

// ---- 自动编号格式化（ST_TextAutonumberScheme，与编辑器渲染端一致） ----
const CN_DIGITS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
const CN_LEGAL = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
const CIRCLED = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩'];

function cnNumber(n: number, legal: boolean): string {
    const digits = legal ? CN_LEGAL : CN_DIGITS;
    const ten = legal ? '拾' : '十';
    if (n <= 0) return String(n);
    if (n < 10) return digits[n];
    if (n === 10) return ten;
    if (n < 20) return ten + digits[n % 10];
    if (n < 100) return digits[Math.floor(n / 10)] + ten + (n % 10 ? digits[n % 10] : '');
    return String(n);
}

function romanNumber(n: number, upper: boolean): string {
    const map: [number, string][] = [
        [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
        [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']
    ];
    let out = '';
    let rest = n;
    for (const [v, s] of map) {
        while (rest >= v) {
            out += s;
            rest -= v;
        }
    }
    return upper ? out : out.toLowerCase();
}

function alphaNumber(n: number, upper: boolean): string {
    let out = '';
    let rest = n;
    while (rest > 0) {
        rest -= 1;
        out = String.fromCharCode(97 + (rest % 26)) + out;
        rest = Math.floor(rest / 26);
    }
    return upper ? out.toUpperCase() : out;
}

function hebrewNumber(n: number): string {
    const letters = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט', 'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ', 'ק', 'ר', 'ש', 'ת'];
    let out = '';
    let rest = Math.max(0, n - 1);
    do {
        out = letters[rest % letters.length] + out;
        rest = Math.floor(rest / letters.length) - 1;
    } while (rest >= 0);
    return out || letters[0];
}

function formatAutoNum(fmt: string, n: number): string {
    switch (fmt) {
        case 'hebrew1Minus':
        case 'hebrew2Minus':
            return `${hebrewNumber(n)}-`;
        case 'arabicPeriod':
            return `${n}.`;
        case 'arabicParenR':
            return `${n})`;
        case 'arabicParenBoth':
            return `(${n})`;
        case 'arabicPlain':
            return `${n}`;
        case 'alphaLcPeriod':
            return `${alphaNumber(n, false)}.`;
        case 'alphaUcPeriod':
            return `${alphaNumber(n, true)}.`;
        case 'alphaLcParenR':
            return `${alphaNumber(n, false)})`;
        case 'alphaUcParenR':
            return `${alphaNumber(n, true)})`;
        case 'alphaLcParenBoth':
            return `(${alphaNumber(n, false)})`;
        case 'alphaUcParenBoth':
            return `(${alphaNumber(n, true)})`;
        case 'romanLcPeriod':
            return `${romanNumber(n, false)}.`;
        case 'romanUcPeriod':
            return `${romanNumber(n, true)}.`;
        case 'romanLcParenR':
            return `${romanNumber(n, false)})`;
        case 'romanUcParenR':
            return `${romanNumber(n, true)})`;
        case 'romanLcParenBoth':
            return `(${romanNumber(n, false)})`;
        case 'romanUcParenBoth':
            return `(${romanNumber(n, true)})`;
        case 'chineseCounting':
        case 'chineseCountingThousand':
        case 'ea1ChsPeriod':
        case 'ea1ChtPeriod':
            return `${cnNumber(n, false)}、`;
        case 'chineseLegalSimplified':
            return `${cnNumber(n, true)}、`;
        case 'ea1ChsPlain':
        case 'ea1ChtPlain':
            return cnNumber(n, false);
        case 'ideographDigital':
        case 'circleNumDbPlain':
        case 'circleNumWdWhitePlain':
        case 'circleNumWdBlackPlain':
            return CIRCLED[(n - 1) % CIRCLED.length];
        default:
            return `${n}.`;
    }
}

/** run -> 行内样式 */
function runStyle(run: PptxTextRun, el: PptxTextElement, fallbackSize: number, fontScale: number): string {
    const size = (run.fontSize ?? fallbackSize) * fontScale;
    const styles: string[] = [];
    if (size) styles.push(`font-size:${num(size)}pt`);
    const family = fontFamilyOf(run.fontFace, run.fontFaceEa, el.fontFace, el.fontFaceEa);
    if (family) styles.push(`font-family:${family},${FONT_STACK}`);
    const color =
        normalizeColor(run.color) ?? normalizeColor(el.color) ?? themeTextColor();
    if (color) styles.push(`color:${color}`);
    if (run.bold ?? el.bold) styles.push('font-weight:700');
    if (run.italic ?? el.italic) styles.push('font-style:italic');
    const deco: string[] = [];
    if (run.underline ?? el.underline) deco.push('underline');
    if (run.strike) deco.push('line-through');
    if (deco.length) styles.push(`text-decoration:${deco.join(' ')}`);
    // 上下标（a:rPr@baseline）
    if (run.baseline) styles.push(`vertical-align:${run.baseline > 0 ? 'super' : 'sub'}`);
    // 字符间距（a:rPr/a:spc）
    if (run.spacing) styles.push(`letter-spacing:${num(run.spacing)}px`);
    // 文本高亮（a:rPr/a:highlight）
    if (run.highlight) styles.push(`background-color:${normalizeColor(run.highlight) ?? 'transparent'}`);
    if (run.smallCaps) styles.push('font-variant:small-caps');
    // 着重号（a:rPr/a:em）
    if (run.emphasisMark) {
        const em = run.emphasisMark;
        styles.push(`text-emphasis-style:${em === 'circle' ? 'circle' : em === 'comma' ? 'comma' : 'dot'}`);
        styles.push('text-emphasis-position:under');
    }
    // 文字描边 + 外阴影（用 text-shadow 模拟）
    const shadows: string[] = [];
    if (run.outline && run.outline !== 'none' && run.outline.color) {
        const oc = normalizeColor(run.outline.color);
        if (oc) {
            const w = Math.max(1, Math.floor(run.outline.width && run.outline.width >= 1 ? run.outline.width : 4 / 3));
            shadows.push(`-${w}px 0 ${oc}, 0 ${w}px ${oc}, ${w}px 0 ${oc}, 0 -${w}px ${oc}`);
        }
    }
    if (run.shadow?.color) {
        const sc = normalizeColor(run.shadow.color);
        if (sc) {
            const alpha = run.shadow.alpha != null ? run.shadow.alpha : 1;
            const blur = run.shadow.blur != null ? run.shadow.blur : 0;
            shadows.push(
                `${num((run.shadow.x ?? 0) * (96 / 72))}px ${num((run.shadow.y ?? 0) * (96 / 72))}px ` +
                    `${num(blur * (96 / 72))}px rgba(${rgbOf(sc)?.join(',') ?? '0,0,0'},${num(alpha)})`
            );
        }
    }
    if (shadows.length) styles.push(`text-shadow:${shadows.join(',')}`);
    return styles.join(';');
}

/**
 * 项目符号 / 编号 -> HTML 片段
 *
 * counter 为跨段落的编号状态（连续编号段落共用计数器，遇到非编号段落重置）。
 */
function bulletToHtml(
    bullet: boolean | 'number' | PptxBullet,
    el: PptxTextElement,
    firstRun: PptxTextRun,
    fontSize: number,
    counter: { n: number | null }
): string {
    const isNumber = bullet === 'number' || (typeof bullet === 'object' && bullet.type === 'number');
    if (isNumber) {
        const obj = typeof bullet === 'object' ? bullet : undefined;
        const fmt = obj?.fmt || 'arabicPeriod';
        if (counter.n == null) counter.n = obj?.start || 1;
        const text = formatAutoNum(fmt, counter.n);
        counter.n += 1;
        const color =
            normalizeColor(obj?.color) ?? normalizeColor(firstRun.color) ?? normalizeColor(el.color) ?? themeTextColor();
        const family = fontFamilyOf(firstRun.fontFace, firstRun.fontFaceEa, el.fontFace, el.fontFaceEa, '微软雅黑');
        return (
            `<span class="jt-bullet" style="font-size:${num(fontSize)}pt;color:${color};` +
            `font-family:${family ?? FONT_STACK};margin-right:0.3em">${esc(text)}</span>`
        );
    }
    const obj = typeof bullet === 'object' ? bullet : undefined;
    const color =
        normalizeColor(obj?.color) ?? normalizeColor(firstRun.color) ?? normalizeColor(el.color) ?? themeTextColor();
    const size = obj?.sizePct ? (fontSize * obj.sizePct) / 100 : fontSize;
    if (obj?.type === 'picture' && obj.data) {
        const px = num((size * 96) / 72);
        return (
            `<span class="jt-bullet jt-bullet-img" style="margin-right:0.3em">` +
            `<img src='${obj.data}' style="width:${px}px;height:${px}px;vertical-align:-0.15em"/></span>`
        );
    }
    const char = obj?.char || '•';
    // 符号字体（a:buFont，如 Wingdings）：缺了它，私用区字符会退化成豆腐块
    const family = obj?.font
        ? quoteFont(obj.font)
        : fontFamilyOf(firstRun.fontFace, firstRun.fontFaceEa, el.fontFace, el.fontFaceEa, '微软雅黑');
    return (
        `<span class="jt-bullet" style="font-size:${num(size)}pt;color:${color};` +
        `font-family:${family ?? FONT_STACK};margin-right:0.3em">${esc(char)}</span>`
    );
}

/** 段落 -> HTML */
function paragraphToHtml(
    p: PptxParagraph,
    el: PptxTextElement,
    fallbackSize: number,
    fontScale: number,
    counter: { n: number | null }
): string {
    const runs: PptxTextRun[] = p.runs?.length
        ? p.runs
        : p.text != null
          ? [{ text: p.text } as PptxTextRun]
          : [];
    const firstRun = runs[0] || ({} as PptxTextRun);
    const effSize = (r: PptxTextRun) => (r.fontSize ?? fallbackSize) * fontScale;

    // 空段落不画项目符号（PowerPoint 放映时空段只占一行空白）
    const hasContent = runs.some((r) => r.text || r.break);
    const bulletRaw = p.bullet ?? el.bullet;
    const bullet = hasContent ? bulletRaw : null;

    const styles: string[] = [`text-align:${p.align ?? el.align ?? 'left'}`];
    const lh = lineHeightStyle(p.lineSpacing ?? el.lineSpacing);
    if (lh) styles.push(lh);
    // 悬挂缩进：marL = 正文左边界（折行也停在这里），indent = 首行相对 marL 的偏移
    const marLSrc = p.indentLeft ?? el.indentLeft ?? el.indent;
    const indentSrc = p.indent;
    if (marLSrc != null || indentSrc != null) {
        const marL = Math.max(0, marLSrc || 0);
        const firstLine = marL + (indentSrc || 0) < 0 ? -marL : indentSrc || 0;
        styles.push(`padding-left:${num(marL)}px`);
        styles.push(`text-indent:${num(firstLine)}px`);
    }
    if (p.spaceBefore) styles.push(`margin-top:${num(p.spaceBefore)}px`);
    if (p.spaceAfter) styles.push(`margin-bottom:${num(p.spaceAfter)}px`);

    let bulletHtml = '';
    if (bullet) {
        bulletHtml = bulletToHtml(bullet, el, firstRun, effSize(firstRun), counter);
    } else {
        counter.n = null;
    }

    const body = runs
        .map((r) => {
            if (r.break) return '<br/>';
            const style = runStyle(r, el, fallbackSize, fontScale);
            const text = esc(r.text ?? '');
            if (r.href) {
                return (
                    `<a href="${esc(r.href)}" target="_blank" rel="noopener noreferrer" ` +
                    `style="${style};text-decoration:underline">${text}</a>`
                );
            }
            return `<span style="${style}">${text}</span>`;
        })
        .join('');

    return `<div class="jt-p" style="${styles.join(';')}">${bulletHtml}${body || '<br/>'}</div>`;
}

/** 文本框 -> HTML（flex 垂直对齐，与编辑器渲染端一致） */
function textToHtml(el: PptxTextElement): string {
    const out: string[] = [];
    // 文本框自带的形状底（先画底再画字）
    if (el.fill && el.fill !== 'none') {
        if (el.shapeType) {
            const box = {
                type: 'shape' as const,
                shapeType: el.shapeType,
                x: el.x,
                y: el.y,
                width: el.width,
                height: el.height,
                fill: el.fill,
                line: el.line,
                adjust: el.adjust,
                rotation: el.rotation
            } as unknown as PptxShapeElement;
            const svg = shapeElementToSvg(box);
            if (svg) out.push(svg);
        } else {
            const fill = solidFillOf(el.fill as PptxFill);
            if (fill) {
                out.push(
                    `<div style="left:${num(el.x)}px;top:${num(el.y)}px;width:${num(el.width)}px;` +
                        `height:${num(el.height)}px;background:${fill.color};` +
                        (fill.alpha < 1 ? `opacity:${num(fill.alpha)};` : '') +
                        `${el.rotation ? `transform:rotate(${num(el.rotation)}deg);` : ''}"></div>`
                );
            }
        }
    }

    // 「溢出则缩字」（normAutofit fontScale）：PPT/WPS 按此比例整体缩小字号
    const fontScale = (el.fontScale ?? 100) / 100;
    const paragraphs = el.paragraphs?.length
        ? el.paragraphs
        : el.runs?.length || el.text
          ? [{ runs: el.runs, text: el.text } as unknown as PptxParagraph]
          : [];
    if (!paragraphs.length) return out.join('');

    const fallbackSize = el.fontSize ?? 18;
    const justify =
        el.valign === 'middle' ? 'center' : el.valign === 'bottom' ? 'flex-end' : 'flex-start';
    const counter: { n: number | null } = { n: null };
    const body = paragraphs
        .map((p) => paragraphToHtml(p, el, fallbackSize, fontScale, counter))
        .join('');

    // 文本框内边距（a:bodyPr lIns/tIns，单位 px）
    const inset = el.inset;
    const padding = inset
        ? `padding:${num(inset.t ?? 0)}px ${num(inset.r ?? 0)}px ${num(inset.b ?? 0)}px ${num(inset.l ?? 0)}px;`
        : '';
    const whiteSpace = el.noWrap ? 'nowrap' : 'pre-wrap';
    const dirStyle =
        el.textDirection && el.textDirection !== 'horz'
            ? 'writing-mode:vertical-rl;text-orientation:upright;'
            : '';
    const rot = el.rotation ? ` transform:rotate(${num(el.rotation)}deg);` : '';
    out.push(
        `<div class="jt-text" style="left:${num(el.x)}px;top:${num(el.y)}px;width:${num(el.width)}px;` +
            `height:${num(el.height)}px;justify-content:${justify};${padding}white-space:${whiteSpace};` +
            `${dirStyle}${rot}">${body}</div>`
    );
    return out.join('');
}

/** 表格 -> HTML */
function tableToHtml(el: PptxTableElement): string {
    const rows = el.rows || [];
    if (!rows.length) return '';
    const body = rows
        .map((row, ri) => {
            const cells = (row.cells || [])
                .map((cell) => {
                    const text = esc(
                        cell.text ?? (cell.paragraphs || []).map((p) => p.text ?? '').join('\n')
                    );
                    const fill = cell.fill ? `background:${normalizeColor(cell.fill) ?? 'transparent'};` : '';
                    const tag = ri === 0 ? 'th' : 'td';
                    return `<${tag} style="${fill}">${text.replace(/\n/g, '<br/>')}</${tag}>`;
                })
                .join('');
            return `<tr>${cells}</tr>`;
        })
        .join('');
    return (
        `<table class="jt-table" style="left:${num(el.x)}px;top:${num(el.y)}px;width:${num(el.width)}px">` +
        `<tbody>${body}</tbody></table>`
    );
}

/** 图片 -> HTML（裁剪 / 亮度对比度 / 透明度与编辑器渲染端一致） */
function imageToHtml(el: PptxImageElement): string {
    const src = el.data || el.src;
    if (!src) return '';
    const rot = el.rotation ? ` transform:rotate(${num(el.rotation)}deg);` : '';
    const adj = el.imageAdjust || {};
    const filters: string[] = [];
    if (adj.brightness) filters.push(`brightness(${num(1 + adj.brightness / 100)})`);
    if (adj.contrast) filters.push(`contrast(${num(1 + adj.contrast / 100)})`);
    const filter = filters.length ? `filter:${filters.join(' ')};` : '';
    const opacity = adj.transparency ? `opacity:${num(Math.max(0, Math.min(1, 1 - adj.transparency / 100)))};` : '';

    // 裁剪（a:srcRect，各边为 0~1 比例）：容器裁掉四周后图片放大填满
    const crop = el.crop;
    const hasCrop = crop && (crop.l || crop.t || crop.r || crop.b);
    if (hasCrop) {
        const cw = Math.max(0.01, 1 - (crop.l || 0) - (crop.r || 0));
        const ch = Math.max(0.01, 1 - (crop.t || 0) - (crop.b || 0));
        const imgStyle =
            `position:absolute;left:${num((-(crop.l || 0) / cw) * 100)}%;top:${num((-(crop.t || 0) / ch) * 100)}%;` +
            `width:${num(100 / cw)}%;height:${num(100 / ch)}%;max-width:none;${filter}${opacity}`;
        return (
            `<div class="jt-img-box" style="left:${num(el.x)}px;top:${num(el.y)}px;width:${num(el.width)}px;` +
            `height:${num(el.height)}px;overflow:hidden;${rot}">` +
            `<img class="jt-img" src='${src}' style="${imgStyle}"/></div>`
        );
    }
    return (
        `<img class="jt-img" src='${src}' style="left:${num(el.x)}px;top:${num(el.y)}px;` +
        `width:${num(el.width)}px;height:${num(el.height)}px;${filter}${opacity}${rot}"/>`
    );
}

/**
 * 递归收集元素 -> HTML 片段（保持文档顺序 = 层叠顺序）
 *
 * offset：当前层级元素坐标的平移量
 * base：当前容器原点的绝对坐标（元素按 `el.x + offset - base` 定位）
 */
function collectHtml(
    list: PptxElement[],
    out: string[],
    offset?: { x: number; y: number },
    base?: { x: number; y: number }
): void {
    const dx = offset?.x ?? 0;
    const dy = offset?.y ?? 0;
    const bx = base?.x ?? 0;
    const by = base?.y ?? 0;
    for (const el of list || []) {
        const ex = el.x + dx - bx;
        const ey = el.y + dy - by;
        switch (el.type) {
            case 'text': {
                const t = el as PptxTextElement;
                out.push(textToHtml({ ...t, x: ex, y: ey }));
                break;
            }
            case 'shape': {
                const s = el as PptxShapeElement;
                const svg = shapeElementToSvg({ ...s, x: ex, y: ey });
                if (svg) out.push(svg);
                break;
            }
            case 'image': {
                const img = el as PptxImageElement;
                out.push(imageToHtml({ ...img, x: ex, y: ey }));
                break;
            }
            case 'table': {
                const t = el as PptxTableElement;
                out.push(tableToHtml({ ...t, x: ex, y: ey }));
                break;
            }
            case 'group': {
                const g = el as PptxGroupElement;
                // relative：子元素坐标相对组合原点；其余（page/绝对坐标）已是页面坐标
                const relKids = g.childrenCoordinates === 'relative';
                const gx = g.x + dx;
                const gy = g.y + dy;
                const tf: string[] = [];
                if (g.rotation) tf.push(`rotate(${num(g.rotation)}deg)`);
                if (flipOn(g.flipH)) tf.push('scaleX(-1)');
                if (flipOn(g.flipV)) tf.push('scaleY(-1)');
                if (!tf.length) {
                    collectHtml(
                        g.children || [],
                        out,
                        relKids ? { x: gx, y: gy } : { x: 0, y: 0 },
                        { x: bx, y: by }
                    );
                    break;
                }
                // 组级旋转/翻转：套一层组容器，子元素相对容器定位
                // （旋转中心为组边界框中心；OOXML 先 flip 后 rot）
                const inner: string[] = [];
                collectHtml(
                    g.children || [],
                    inner,
                    relKids ? { x: gx, y: gy } : { x: 0, y: 0 },
                    { x: gx, y: gy }
                );
                out.push(
                    `<div class="jt-group" style="left:${num(gx - bx)}px;top:${num(gy - by)}px;` +
                        `width:${num(g.width || 0)}px;height:${num(g.height || 0)}px;` +
                        `transform:${tf.join(' ')};transform-origin:center center">${inner.join('')}</div>`
                );
                break;
            }
            default:
                break;
        }
    }
}

/** 幻灯片背景色（纯色 / 渐变取首停色） */
function slideBackgroundColor(slide: PptxSlide): string {
    const bg = slide.background;
    if (typeof bg === 'string') return normalizeColor(bg) ?? '#ffffff';
    if (bg) {
        const b = bg as { type?: string; color?: string; stops?: GradientStop[] };
        const color =
            b.type === 'gradient' ? normalizeColor(b.stops?.[0]?.color) : normalizeColor(b.color);
        if (color) return color;
    }
    return '#ffffff';
}

/** 幻灯片背景图（data URL） */
function slideBackgroundImage(slide: PptxSlide): string | undefined {
    const bg = slide.background;
    if (bg && typeof bg === 'object') {
        const b = bg as { type?: string; data?: string; src?: string };
        if (b.type === 'image') return b.data || b.src;
    }
    return undefined;
}

/** 幻灯片 -> HTML */
function slideToHtml(slide: PptxSlide, index: number, pageW: number, pageH: number): string {
    const bg = slideBackgroundColor(slide);
    const bgImage = slideBackgroundImage(slide);
    const bgStyle = bgImage
        ? // 注意：URL 用单引号 —— 双引号会提前闭合 style 属性
          `background:#${bg.replace('#', '')} url('${bgImage}') center/cover no-repeat;`
        : `background:${bg};`;
    const parts: string[] = [];
    collectHtml(slide.elements || [], parts);
    return (
        `<section class="jt-slide" data-slide="${index}" style="width:${num(pageW)}px;` +
        `height:${num(pageH)}px;${bgStyle}">${parts.join('')}</section>`
    );
}

/**
 * PPTX -> HTML（版式还原）
 *
 * 每页一张与幻灯片等大的画布，元素按文档顺序绝对定位；
 * 形状走 SVG、文本走 HTML（flex 垂直对齐），与 PDF / 编辑器预览保持一致。
 */
export async function pptxToSlideHtml(
    input: DocInput,
    options: PptxToHtmlOptions = {}
): Promise<string> {
    try {
        const doc = (await pptxToStandard(input, { assets: 'both' })) as PptxDocument;
        const pageW = doc.slideSize.width;
        const pageH = doc.slideSize.height;
        const slides = doc.slides.filter((s) => !s.hidden);

        // 取色依赖主题色板（scheme:bg1 等），在同步作用域内切换
        const body = withThemeColors(doc, () =>
            slides
                .map((slide, i) => slideToHtml(slide, i + 1, pageW, pageH))
                .join('\n')
        );

        const title = options.title ?? 'PPTX';
        return `<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>${esc(title)}</title>
<style>
html,body{margin:0;padding:0;background:#eef0f3}
body{padding:24px 12px;font-family:${FONT_STACK}}
.jt-slides{display:flex;flex-direction:column;align-items:center;gap:24px}
.jt-slide{position:relative;overflow:hidden;background:#fff;box-shadow:0 1px 6px rgba(0,0,0,.2);flex:none}
.jt-slide>svg,.jt-slide>.jt-text,.jt-slide>.jt-img,.jt-slide>.jt-img-box,.jt-slide>.jt-table,.jt-slide>.jt-group,.jt-slide>div{position:absolute}
.jt-group>svg,.jt-group>.jt-text,.jt-group>.jt-img,.jt-group>.jt-img-box,.jt-group>.jt-table,.jt-group>.jt-group,.jt-group>div{position:absolute}
.jt-text{display:flex;flex-direction:column;box-sizing:border-box;overflow:visible;word-break:break-word}
.jt-text>.jt-p{margin:0;min-height:1em;flex-shrink:0}
.jt-bullet{display:inline-block;min-width:1.2em;text-align:left;white-space:pre}
.jt-bullet-img img{display:inline-block;max-width:none;object-fit:contain}
.jt-img{object-fit:fill}
.jt-table{border-collapse:collapse;table-layout:fixed}
.jt-table td,.jt-table th{border:1px solid #d0d0d0;padding:2px 6px;font-size:12pt;text-align:left;vertical-align:top}
@media print{html,body{background:#fff;padding:0}.jt-slides{gap:0}.jt-slide{box-shadow:none;page-break-after:always;margin:0}}
</style>
</head>
<body>
<div class="jt-slides">
${body}
</div>
</body>
</html>
`;
    } catch (error) {
        throw wrapError('pptx', '转 HTML', error);
    }
}
