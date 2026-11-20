/**
 * PPTX -> PDF（幻灯片版式直排）
 *
 * 幻灯片是「固定画布 + 绝对定位元素」，与 docx/md 的流式文档不同。经 blocks
 * 中转的流式排版会把 16:9 幻灯片重排成 A4 竖版：绝对定位的文本框互相叠压、
 * 大图溢出成多页、页尾大面积留白。这里按幻灯片原稿直排：
 *
 *   - 每张幻灯片一页，页面尺寸 = 幻灯片尺寸（解析器坐标为 96dpi px，转 pt ×0.75）
 *   - 文本框 / 图片 / 表格 / 形状按原始位置与尺寸绝对定位（依赖 pages 模式
 *     对显式 y 元素的绝对定位语义，见 generate.ts renderPage）
 *   - 渐变填充用多段纯色条带近似（PDF 生成端只支持单色填充），按停点位置
 *     插值并把透明度混合到幻灯片背景色
 *   - 形状几何还原圆角（roundRect / round2SameRect）、椭圆与旋转
 *   - 文本框自带的形状填充（如「汇报人」的蓝色胶囊底）单独绘制在文字下层
 *   - 文字按框内垂直对齐（top / middle / bottom）定位，行高按 1.4（CJK）/ 1.2（西文）
 *
 * @module convert/pptx-pdf
 */

import type { DocInput } from '../core/input';
import { pptxToStandard } from '../pptx/index';
import { jsonToPdf } from '../pdf/index';
import { detectSystemFonts } from '../pdf/fonts';
import type {
    PdfElement,
    PdfFontSet,
    PdfPageSpec,
    PdfPathSegment,
    PdfRun,
    PdfRunsElement,
    PdfTableCell,
    PdfTableCellObject
} from '../pdf/types';
import type {
    PptxDocument,
    PptxElement,
    PptxFill,
    PptxGroupElement,
    PptxImageElement,
    PptxShapeElement,
    PptxSlide,
    PptxTableElement,
    PptxTableCell,
    PptxTextElement,
    PptxTextRun
} from '@fefeding/ppt-parser';

/** 解析器导出的坐标单位是 px（96dpi，16:9 幻灯片为 1280×720），PDF 用 pt */
const PX_TO_PT = 0.75;

/** 幻灯片上未显式给字号的文本默认 18pt（与 PowerPoint / WPS 默认一致） */
const DEFAULT_FONT_SIZE = 18;

/** 行高系数：CJK 字体行高约 1.4，西文约 1.2（与编辑器渲染一致） */
const LINE_HEIGHT_CJK = 1.4;
const LINE_HEIGHT_LATIN = 1.2;

/** 渐变条带数：足够密以看不出阶梯 */
const GRADIENT_BANDS = 48;

/** pptxToPdf 选项 */
export interface PptxToPdfOptions {
    /** 内嵌字体（中文必须）；未提供且内容含中文时自动探测系统字体 */
    fonts?: PdfFontSet;
    /** 是否自动探测系统字体（默认 true，仅 Node 端有效） */
    autoFont?: boolean;
}

// ---- 颜色 ----

/**
 * 主题色解析器（scheme:bg1 / scheme:accent1 等 -> 具体色值）
 *
 * 解析器输出里幻灯片背景等保留的是 `scheme:bg1` 这类主题色引用，
 * 需要用 theme1.xml 的 clrScheme + 母版 clrMap 才能落到具体色值。
 * bg1/tx1/bg2/tx2 是 clrMap 的别名槽（默认 bg1=lt1、tx1=dk1）。
 *
 * 下列纯函数同时被 PPTX→HTML 的版式还原链路复用（见 pptx-html.ts）。
 */
export interface SchemeResolver {
    (value: string | undefined | null): string | undefined;
}

/** 解析 theme xml 的 clrScheme 色板与母版 clrMap，返回 scheme 颜色解析函数 */
export function buildSchemeResolver(themeXml?: string, masterXml?: string): SchemeResolver {
    const palette = new Map<string, string>();
    const pick = (xml: string, tag: string): string | undefined => {
        const m = xml.match(new RegExp(`<a:${tag}>[\\s\\S]*?</a:${tag}>`));
        if (!m) return undefined;
        const srgb = m[0].match(/srgbClr val="([0-9A-Fa-f]{6})"/);
        if (srgb) return `#${srgb[1].toUpperCase()}`;
        // sysClr（windowText/window 等）取 lastClr
        const sys = m[0].match(/sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"/);
        if (sys) return `#${sys[1].toUpperCase()}`;
        return undefined;
    };
    if (themeXml) {
        const scheme = themeXml.match(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/);
        if (scheme) {
            for (const slot of [
                'dk1', 'lt1', 'dk2', 'lt2',
                'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
                'hlink', 'folHlink'
            ]) {
                const color = pick(scheme[0], slot);
                if (color) palette.set(slot, color);
            }
        }
    }
    // 母版 clrMap：bg1/tx1/bg2/tx2 -> 实际色槽（OOXML 默认 bg1=lt1、tx1=dk1）
    const colorMap: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
    if (masterXml) {
        const map = masterXml.match(/<p:clrMap[^>]*\/>/);
        if (map) {
            for (const [k, v] of Object.entries(colorMap)) {
                const m = map[0].match(new RegExp(`${k}="(\\w+)"`));
                if (m) colorMap[k] = m[1];
            }
        }
    }
    return (value) => {
        if (!value) return undefined;
        const name = value.replace(/^scheme:/i, '').trim();
        const slot = colorMap[name] ?? name;
        return palette.get(slot);
    };
}

/**
 * 当前转换调用的 scheme 解析器。
 *
 * 形状 / 文本 / 渐变的取色函数众多且都是同步纯函数，逐层传参改动过大；
 * 这里以模块级变量在一次 pptxToPdf 调用内设置 —— 取色过程同步执行、
 * 期间没有 await，事件循环无法插队，因此并发转换不会串色。
 */
let resolveScheme: SchemeResolver = () => undefined;

/**
 * 在指定文档的主题色板下执行同步回调
 *
 * PPTX→HTML 的版式还原链路复用本模块的取色函数（normalizeColor 等），
 * 通过该函数在同步作用域内切换主题色板；回调必须同步执行。
 */
export function withThemeColors<T>(doc: PptxDocument, fn: () => T): T {
    const prev = resolveScheme;
    resolveScheme = buildSchemeResolver(
        doc.themeXmls?.[0] ?? doc.themeXml,
        (doc.masters?.[0] as { __rawXml?: string } | undefined)?.__rawXml
    );
    try {
        return fn();
    } finally {
        resolveScheme = prev;
    }
}

/** 颜色归一：解析器的颜色可能不带 # 前缀；非法色值视为未指定 */
export function normalizeColor(value: string | undefined | null): string | undefined {
    if (!value || value === 'none' || typeof value !== 'string') return undefined;
    if (/^scheme:/i.test(value)) value = resolveScheme(value) ?? undefined;
    if (!value) return undefined;
    if (!value.startsWith('#')) value = `#${value}`;
    // 非法色值一律视为未指定（此前会生成 "#scheme:bg1" 这类垃圾值，
    // 进而被 PDF 端兜底成黑色，出现整页黑底）
    return /^#[0-9a-fA-F]{6}$/.test(value) ? value : undefined;
}

/** #RRGGBB -> [r, g, b]（非法输入返回 null） */
export function rgbOf(color: string | undefined): [number, number, number] | null {
    const hex = normalizeColor(color)?.slice(1);
    if (!hex || !/^[0-9a-fA-F]{6}$/.test(hex)) return null;
    return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}

/** [r, g, b] -> #RRGGBB */
export function toHex(rgb: [number, number, number]): string {
    return `#${rgb.map((v) => Math.round(v).toString(16).padStart(2, '0')).join('')}`;
}

// ---- 渐变 ----

/** 渐变停点 */
export interface GradientStop {
    color: string;
    position: number;
    transparency?: number;
}

/** 取填充里的渐变停点（非渐变返回 null） */
export function gradientStopsOf(fill: PptxFill | undefined): { stops: GradientStop[]; direction: string; angle?: number } | null {
    if (!fill || typeof fill === 'string') return null;
    const f = fill as { type?: string; stops?: GradientStop[]; direction?: string; angle?: number };
    if (f.type !== 'gradient' || !f.stops?.length) return null;
    const stops = [...f.stops]
        .map((s) => ({ color: normalizeColor(s.color)!, position: s.position ?? 0, transparency: s.transparency }))
        .filter((s) => !!s.color)
        .sort((a, b) => a.position - b.position);
    if (!stops.length) return null;
    return { stops, direction: f.direction || 'vertical', angle: f.angle };
}

/** 渐变在 t（0~1）处的颜色与不透明度 */
export function gradientAt(stops: GradientStop[], t: number): { color: string; alpha: number } {
    const clamped = Math.max(0, Math.min(1, t));
    let lo = stops[0];
    let hi = stops[stops.length - 1];
    // 与 jdoc 上游的差异：原条件 `i < stops.length - i` 在单停点数组（HTML 渐变停点
    // 逐个转 SVG stop 时会传 [s]）里仍会进入循环并读 stops[1].position 而抛 TypeError。
    // 正确的区间遍历上界是 length - 1。待上游修正后可去掉这处本地补丁。
    for (let i = 0; i < stops.length - 1; i++) {
        const a = stops[i];
        const b = stops[i + 1];
        if (clamped >= a.position && clamped <= b.position) {
            lo = a;
            hi = b;
            break;
        }
    }
    const span = hi.position - lo.position;
    const k = span > 0 ? (clamped - lo.position) / span : 0;
    const c1 = rgbOf(lo.color);
    const c2 = rgbOf(hi.color);
    const color = c1 && c2 ? toHex([0, 1, 2].map((i) => c1[i] + (c2[i] - c1[i]) * k) as [number, number, number]) : lo.color;
    const a1 = 1 - (lo.transparency ?? 0) / 100;
    const a2 = 1 - (hi.transparency ?? 0) / 100;
    return { color, alpha: a1 + (a2 - a1) * k };
}

/**
 * 渐变轴（形状本地坐标，px）
 *
 * 方向语义与 PPT 的 a:lin 一致：ang 以 60000 分之一度表示、自正东方向顺时针；
 * 解析器给的是度。渐变起点是沿反方向的那个角（stop 0），终点是沿方向的角。
 */
export function gradientAxis(
    direction: string,
    angle: number | undefined,
    w: number,
    h: number
): { dir: [number, number]; p1: [number, number]; len: number; coverage: number; angle: number } {
    let dx: number;
    let dy: number;
    if (angle != null) {
        const rad = (angle * Math.PI) / 180;
        dx = Math.cos(rad);
        dy = Math.sin(rad);
    } else if (direction === 'horizontal') {
        dx = 1;
        dy = 0;
    } else if (direction === 'diagonal') {
        dx = Math.SQRT1_2;
        dy = -Math.SQRT1_2;
    } else {
        dx = 0;
        dy = 1;
    }
    const len = Math.hypot(dx, dy) || 1;
    dx /= len;
    dy /= len;
    // 矩形四角在方向上的投影，取最小 / 最大者为起点 / 终点
    const corners: [number, number][] = [
        [0, 0],
        [w, 0],
        [0, h],
        [w, h]
    ];
    let p1 = corners[0];
    let p2 = corners[0];
    let sMin = Infinity;
    let sMax = -Infinity;
    for (const c of corners) {
        const s = c[0] * dx + c[1] * dy;
        if (s < sMin) {
            sMin = s;
            p1 = c;
        }
        if (s > sMax) {
            sMax = s;
            p2 = c;
        }
    }
    // 垂直方向的覆盖长度（条带要盖住整个形状）
    const coverage = Math.abs(w * -dy) + Math.abs(h * dx);
    return { dir: [dx, dy], p1, len: sMax - sMin, coverage, angle: (Math.atan2(dy, dx) * 180) / Math.PI };
}

// ---- 形状 ----

/** 贝塞尔圆弧控制点系数（4/3·(√2−1)） */
const BEZIER_K = 0.5523;

/**
 * 四角圆角半径（px，顺序：左上 / 右上 / 右下 / 左下）
 *
 * 与编辑器渲染端（render.ts shapeGeometry）同源：
 *   - roundRect：adj（默认 16667）占短边比例，四角同半径
 *   - round1Rect：仅左上角（adj1，默认 33333）
 *   - round2SameRect：上方两角（adj1）
 *   - round2DiagRect：对角两角（adj1 / adj2）
 */
export function shapeCornerRadii(
    shapeType: string | undefined,
    adjust: Record<string, number> | undefined,
    width: number,
    height: number
): [number, number, number, number] {
    const st = (shapeType || '').toLowerCase();
    const min = Math.min(width, height);
    if (min <= 0) return [0, 0, 0, 0];
    const adj = adjust || {};
    if (st === 'roundrect') {
        const r = ((adj.adj != null ? adj.adj : 16667) / 100000) * min;
        return [r, r, r, r];
    }
    if (st === 'round1rect' || st === 'round2samerect' || st === 'round2diagrect') {
        const r1 = ((adj.adj1 != null ? adj.adj1 : 33333) / 100000) * min;
        const r2 = ((adj.adj2 != null ? adj.adj2 : 0) / 100000) * min;
        if (st === 'round1rect') return [r1, 0, 0, 0];
        if (st === 'round2samerect') return [r1, r1, 0, 0];
        return [r1, 0, r1, 0];
    }
    return [0, 0, 0, 0];
}

/** 形状填充 -> 颜色 + 不透明度（透明度过高视为不可见） */
export function solidFillOf(fill: PptxFill | undefined): { color: string; alpha: number } | undefined {
    if (!fill || fill === 'none') return undefined;
    if (typeof fill === 'string') {
        const color = normalizeColor(fill);
        return color ? { color, alpha: 1 } : undefined;
    }
    const f = fill as { type?: string; color?: string; transparency?: number; fg?: string };
    if (f.type === 'gradient' || f.type === 'image') return undefined;
    const color = normalizeColor(f.type === 'pattern' ? f.fg : f.color);
    if (!color) return undefined;
    const alpha = 1 - (f.transparency ?? 0) / 100;
    if (alpha < 0.03) return undefined;
    return { color, alpha };
}

/**
 * PPT 虚线类型 -> 虚线间隔（px）
 *
 * 与编辑器渲染端的 SVG_DASH_MAP 保持一致（dasharray 单位是 px）：
 * solid 无、dash 5、dot 1,5、lgDash 10,5、sysDash 5,2 等。
 */
export const DASH_PATTERNS: Record<string, number[]> = {
    dash: [5, 5],
    dashDot: [5, 5, 1, 5],
    dot: [1, 5],
    lgDash: [10, 5],
    lgDashDot: [10, 5, 1, 5],
    lgDashDotDot: [10, 5, 1, 5, 1, 5],
    sysDash: [5, 2],
    sysDot: [2, 5],
    sysDashDot: [5, 2, 1, 5],
    sysDashDotDot: [5, 2, 1, 5, 1, 5]
};

/** 带箭头信息的线条样式（解析器类型未声明 arrow 字段，运行时存在） */
interface ConnectorLine {
    color?: string;
    width?: number;
    transparency?: number;
    dashType?: string;
    startArrow?: string;
    endArrow?: string;
}

/** OOXML 的 flip 属性既可能是布尔也可能是 '1' / 'true' 字符串 */
export function flipOn(value: unknown): boolean {
    return value === true || value === '1' || value === 'true';
}

/**
 * 折线点列 -> 路径段
 *
 * closed=false 时保持开放（连接线是开放路径，闭合会多出一段回首点的斜线）；
 * closed=true 用于多边形（如块箭头）。
 */
function pathFromPoints(points: [number, number][], closed = false): PdfPathSegment[] {
    const segs: PdfPathSegment[] = points.map(([x, y], i) => ({
        op: i === 0 ? 'm' : 'l',
        args: [x, y]
    }));
    if (closed) segs.push({ op: 'h', args: [] });
    return segs;
}

/** 二次贝塞尔 -> 三次贝塞尔控制点（PDF 只支持三次） */
function quadToCubic(
    p0: [number, number],
    c: [number, number],
    p1: [number, number]
): [number, number, number, number, number, number] {
    return [
        p0[0] + (2 / 3) * (c[0] - p0[0]),
        p0[1] + (2 / 3) * (c[1] - p0[1]),
        p1[0] + (2 / 3) * (c[0] - p1[0]),
        p1[1] + (2 / 3) * (c[1] - p1[1]),
        p1[0],
        p1[1]
    ];
}

/**
 * 连接线 / 直线 -> PDF 元素
 *
 * 几何与编辑器渲染端（render.ts 的连接线分支）同源：
 *   - straightConnector1：左上 -> 右下 的直线
 *   - bentConnector*：M 0,0 L W/2,0 L W/2,H L W,H
 *   - curvedConnector*：两段二次贝塞尔 M 0,0 Q W*r,0 W/2,H/2 Q W*(1-r),H W,H
 *   - line：同 straightConnector1（编辑器同样画成对角线）
 * flipH / flipV 以形状中心镜像（与编辑器的 scaleX(-1) 一致）。
 */
function connectorToElements(el: PptxShapeElement): PdfElement[] {
    const line: ConnectorLine | undefined =
        el.line && el.line !== 'none' ? (el.line as unknown as ConnectorLine) : undefined;
    const color = normalizeColor(line?.color);
    if (!color) return [];
    const st = (el.shapeType || '').toLowerCase();
    const { x, y, width: w, height: h } = el;
    const flipH = flipOn(el.flipH);
    const flipV = flipOn(el.flipV);
    const dash = DASH_PATTERNS[line?.dashType || '']?.map((d) => d * PX_TO_PT);
    // 线宽单位是 pt（解析器用 emuToPt 换算），编辑器端下限为 1px
    const lineWidth = Math.max(0.75, line?.width ?? 2);
    // 直线类形状（shapeType='line'）编辑器端不画箭头
    const isPlainLine = st === 'line';
    const startArrow = !isPlainLine && line?.startArrow && line.startArrow !== 'none' ? true : undefined;
    const endArrow = !isPlainLine && line?.endArrow && line.endArrow !== 'none' ? true : undefined;

    /** 局部坐标（0,0 ~ w,h）-> 页面绝对坐标（含镜像） */
    const P = (lx: number, ly: number): [number, number] => [
        (x + (flipH ? w - lx : lx)) * PX_TO_PT,
        (y + (flipV ? h - ly : ly)) * PX_TO_PT
    ];

    if (st.startsWith('bentconnector')) {
        // M 0,0 L W/2,0 L W/2,H L W,H
        return [
            {
                type: 'path',
                segments: pathFromPoints([P(0, 0), P(w / 2, 0), P(w / 2, h), P(w, h)]),
                stroke: color,
                lineWidth,
                dash,
                startArrow,
                endArrow
            }
        ];
    }

    if (st.startsWith('curvedconnector')) {
        // 两段二次贝塞尔（r = adj1/100000，默认 0.5）
        const r = Math.min(Math.max((el.adjust?.adj1 ?? 50000) / 100000, 0), 1);
        const p0 = P(0, 0);
        const c1 = P(w * r, 0);
        const mid = P(w / 2, h / 2);
        const c2 = P(w * (1 - r), h);
        const p1 = P(w, h);
        return [
            {
                type: 'path',
                segments: [
                    { op: 'm', args: [p0[0], p0[1]] },
                    { op: 'c', args: quadToCubic(p0, c1, mid) },
                    { op: 'c', args: quadToCubic(mid, c2, p1) }
                ],
                stroke: color,
                lineWidth,
                dash,
                startArrow,
                endArrow
            }
        ];
    }

    // straightConnector1 / line：左上 -> 右下
    const a = P(0, 0);
    const b = P(w, h);
    return [
        {
            type: 'line',
            x1: a[0],
            y1: a[1],
            x2: b[0],
            y2: b[1],
            lineWidth,
            color,
            dash,
            startArrow,
            endArrow
        }
    ];
}

/** 块箭头（rightArrow / downArrow）-> 多边形路径（adj1 箭身厚度 / adj2 箭头长度，占短边） */
function blockArrowToElements(el: PptxShapeElement): PdfElement[] {
    const fill = solidFillOf(el.fill);
    if (!fill) return [];
    const { x, y, width: w, height: h } = el;
    const ss = Math.min(w, h);
    const adj = el.adjust || {};
    const body = ss * ((adj.adj1 ?? 50000) / 100000);
    const head = ss * ((adj.adj2 ?? 50000) / 100000);
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
    const segments: PdfPathSegment[] = [];
    pts.forEach(([px, py], i) => {
        segments.push({
            op: i === 0 ? 'm' : 'l',
            args: [(x + px) * PX_TO_PT, (y + py) * PX_TO_PT]
        });
    });
    segments.push({ op: 'h', args: [] });
    const stroke = normalizeColor(
        el.line && el.line !== 'none' ? (el.line as { color?: string }).color : undefined
    );
    return [
        {
            type: 'path',
            segments,
            fill: fill.color,
            opacity: fill.alpha < 1 ? fill.alpha : undefined,
            stroke,
            lineWidth: stroke ? Math.max(0.5, (el.line as { width?: number }).width ?? 1) : undefined
        }
    ];
}

/** 圆柱（can / cylinder）-> 底椭圆 + 柱身 + 顶椭圆（adj 为椭圆高占比，默认 25000） */
function canToElements(el: PptxShapeElement): PdfElement[] {
    const fill = solidFillOf(el.fill);
    if (!fill) return [];
    const { x, y, width: w, height: h } = el;
    const ss = Math.min(w, h);
    const ry = (ss * ((el.adjust?.adj ?? 25000) / 100000)) / 2;
    if (ry <= 0 || h <= 2 * ry) return [];
    const opacity = fill.alpha < 1 ? fill.alpha : undefined;
    return [
        // 底（下半外露）
        {
            type: 'rect',
            x: x * PX_TO_PT,
            y: (y + h - 2 * ry) * PX_TO_PT,
            width: w * PX_TO_PT,
            height: 2 * ry * PX_TO_PT,
            fill: fill.color,
            opacity,
            ellipse: true
        },
        // 柱身
        {
            type: 'rect',
            x: x * PX_TO_PT,
            y: (y + ry) * PX_TO_PT,
            width: w * PX_TO_PT,
            height: (h - 2 * ry) * PX_TO_PT,
            fill: fill.color,
            opacity
        },
        // 顶
        {
            type: 'rect',
            x: x * PX_TO_PT,
            y: y * PX_TO_PT,
            width: w * PX_TO_PT,
            height: 2 * ry * PX_TO_PT,
            fill: fill.color,
            opacity,
            ellipse: true
        }
    ];
}

/** 圆角矩形路径（局部坐标：左上为原点、y 向下、宽高为 w/h） */
export function roundedRectPathLocal(w: number, h: number, radii: [number, number, number, number]): PdfPathSegment[] {
    const [tl, tr, br, bl] = radii;
    const k = (v: number) => BEZIER_K * v;
    const segs: PdfPathSegment[] = [{ op: 'm', args: [tl, 0] }];
    segs.push({ op: 'l', args: [w - tr, 0] });
    if (tr > 0) segs.push({ op: 'c', args: [w - tr + k(tr), 0, w, tr - k(tr), w, tr] });
    segs.push({ op: 'l', args: [w, h - br] });
    if (br > 0) segs.push({ op: 'c', args: [w, h - br + k(br), w - br + k(br), h, w - br, h] });
    segs.push({ op: 'l', args: [bl, h] });
    if (bl > 0) segs.push({ op: 'c', args: [bl - k(bl), h, 0, h - bl + k(bl), 0, h - bl] });
    segs.push({ op: 'l', args: [0, tl] });
    if (tl > 0) segs.push({ op: 'c', args: [0, tl - k(tl), tl - k(tl), 0, tl, 0] });
    segs.push({ op: 'h', args: [] });
    return segs;
}

/** 椭圆路径（局部坐标） */
export function ellipsePathLocal(w: number, h: number): PdfPathSegment[] {
    const cx = w / 2;
    const cy = h / 2;
    const kx = BEZIER_K * (w / 2);
    const ky = BEZIER_K * (h / 2);
    return [
        { op: 'm', args: [cx + w / 2, cy] },
        { op: 'c', args: [cx + w / 2, cy + ky, cx + kx, cy + h / 2, cx, cy + h / 2] },
        { op: 'c', args: [cx - kx, cy + h / 2, cx - w / 2, cy + ky, cx - w / 2, cy] },
        { op: 'c', args: [cx - w / 2, cy - ky, cx - kx, cy - h / 2, cx, cy - h / 2] },
        { op: 'c', args: [cx + kx, cy - h / 2, cx + w / 2, cy - ky, cx + w / 2, cy] },
        { op: 'h', args: [] }
    ];
}

/** 绕中心旋转一个点（屏幕坐标顺时针为正，与 PPTX 的 rot 一致） */
export function rotatePoint(p: [number, number], center: [number, number], deg: number): [number, number] {
    if (!deg) return p;
    const rad = (deg * Math.PI) / 180;
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    const dx = p[0] - center[0];
    const dy = p[1] - center[1];
    return [center[0] + dx * cos - dy * sin, center[1] + dx * sin + dy * cos];
}

/** 路径段整体仿射（旋转 + 平移），坐标从局部 px 转到页面绝对 pt */
function pathToPage(
    segs: PdfPathSegment[],
    originX: number,
    originY: number,
    center: [number, number],
    theta: number
): PdfPathSegment[] {
    const out: PdfPathSegment[] = [];
    for (const seg of segs) {
        if (seg.op === 'h') {
            out.push({ op: 'h', args: [] });
            continue;
        }
        // 'm' / 'l' 为 1 个点，'c' 为 3 个控制点
        const pointCount = seg.op === 'c' ? 3 : 1;
        const args: number[] = [];
        for (let i = 0; i < pointCount; i++) {
            const [rx, ry] = rotatePoint(
                [seg.args[i * 2], seg.args[i * 2 + 1]],
                center,
                theta
            );
            args.push((originX + rx) * PX_TO_PT, (originY + ry) * PX_TO_PT);
        }
        out.push({ op: seg.op, args });
    }
    return out;
}

/**
 * 形状 -> PDF 元素（渐变展开成多条带 + 轮廓裁剪）
 *
 * 渐变沿轴切成若干条带（每条带是一个旋转矩形），位置 0 在轴的起点角、
 * 位置 1 在终点角，与 PPT 的 a:lin 语义一致。圆角 / 椭圆形状的轮廓用裁剪
 * 路径限制条带范围（PDF 无原生渐变，也不能直接画渐变填充的圆角形状）。
 *
 * 旋转：PPTX 的 rot 为顺时针为正，而 PDF 生成端的 rotation 为逆时针为正，
 * 两者相差一个符号（这里统一换算）。
 */
function shapeToElements(el: PptxShapeElement): PdfElement[] {
    if (!el.width && !el.height) return [];

    // 连接线：straightConnector / bentConnector / curvedConnector / line 都有各自的
    // 几何（与编辑器渲染端同源），不能按矩形轮廓画。
    const st = (el.shapeType || '').toLowerCase();
    if (/^(straightconnector|bentconnector|curvedconnector)/.test(st) || st === 'line') {
        return connectorToElements(el);
    }

    // 块箭头（rightArrow / downArrow）：多边形路径
    if (st === 'rightarrow' || st === 'downarrow') return blockArrowToElements(el);

    // 圆柱 / 数据库符号（can / cylinder / flowChartMagneticDisk）
    if (st === 'can' || st === 'cylinder' || st === 'flowchartmagneticdisk') {
        return canToElements(el);
    }

    const x = el.x * PX_TO_PT;
    const y = el.y * PX_TO_PT;
    const w = el.width * PX_TO_PT;
    const h = el.height * PX_TO_PT;
    const line = el.line && el.line !== 'none' ? el.line : undefined;
    const stroke = normalizeColor(line?.color);
    const lineWidth = line?.width ? Math.max(0.5, line.width) : undefined;
    const radii = shapeCornerRadii(el.shapeType, el.adjust, el.width, el.height);
    const uniformRadius = radii[0] === radii[1] && radii[1] === radii[2] && radii[2] === radii[3];
    const ellipse = (el.shapeType || '').toLowerCase() === 'ellipse';
    // PPTX 旋转（顺时针）-> PDF 生成端旋转（逆时针）
    const theta = el.rotation || 0;
    const rotation = theta ? -theta : undefined;

    const gradient = gradientStopsOf(el.fill);
    if (gradient) {
        const axis = gradientAxis(gradient.direction, gradient.angle, el.width, el.height);
        const n = GRADIENT_BANDS;
        const bands: PdfElement[] = [];
        for (let i = 0; i < n; i++) {
            const t = (i + 0.5) / n;
            const { color, alpha } = gradientAt(gradient.stops, t);
            if (alpha < 0.02) continue;
            // 条带中心：以形状中心为基准，沿渐变轴推进（垂直于轴的方向必须以形状
            // 中心对齐，否则窄形状会被切掉一半）
            const s = t * axis.len - axis.len / 2;
            const cx = el.width / 2 + axis.dir[0] * s;
            const cy = el.height / 2 + axis.dir[1] * s;
            const bandW = axis.len / n;
            const bandH = axis.coverage;
            // 形状整体旋转：条带中心绕形状中心旋转
            const [rx, ry] = rotatePoint([cx, cy], [el.width / 2, el.height / 2], theta);
            bands.push({
                type: 'rect',
                x: (el.x + rx - bandW / 2) * PX_TO_PT,
                y: (el.y + ry - bandH / 2) * PX_TO_PT,
                width: bandW * PX_TO_PT,
                height: bandH * PX_TO_PT,
                fill: color,
                opacity: alpha,
                // 条带方向即渐变轴（屏幕坐标顺时针角度，与 PPT 的 rot/ang 同向）
                rotation: -(axis.angle + theta)
            });
        }
        if (!bands.length) return [];

        // 轮廓裁剪：圆角 / 椭圆形状的条带必须限制在轮廓内，否则会溢出成矩形
        const hasOutline = ellipse || radii.some((r) => r > 0);
        if (!hasOutline) return bands;
        const local = ellipse
            ? ellipsePathLocal(el.width, el.height)
            : roundedRectPathLocal(el.width, el.height, radii);
        const clip = pathToPage(local, el.x, el.y, [el.width / 2, el.height / 2], theta);
        const out: PdfElement[] = [{ type: 'group', clip, children: bands }];
        // 描边：四角同半径 / 椭圆可精确还原，其他情况跳过（形状无边框的素材居多）
        if (stroke && (ellipse || uniformRadius)) {
            out.push({
                type: 'rect',
                x,
                y,
                width: w,
                height: h,
                stroke,
                lineWidth,
                radius: ellipse ? undefined : radii[0] * PX_TO_PT,
                ellipse: ellipse || undefined,
                rotation
            });
        }
        return out;
    }

    const fill = solidFillOf(el.fill);
    if (!fill && !stroke) return [];
    return [
        {
            type: 'rect',
            x,
            y,
            width: w,
            height: h,
            fill: fill?.color,
            opacity: fill && fill.alpha < 1 ? fill.alpha : undefined,
            stroke,
            lineWidth,
            radius: uniformRadius && radii[0] > 0 ? radii[0] * PX_TO_PT : undefined,
            ellipse: ellipse || undefined,
            rotation
        }
    ];
}

// ---- 文本 ----

/** PPTX 文本片段 -> PDF run（解析器字号已是 pt，不需要缩放） */
function runToPdfRun(run: PptxTextRun, fallbackSize?: number, fontScale = 1): PdfRun {
    const r: PdfRun = { text: run.text ?? '' };
    if (run.bold) r.bold = true;
    if (run.italic) r.italic = true;
    if (run.strike) r.strike = true;
    if (run.color) r.color = normalizeColor(run.color);
    const size = (run.fontSize ?? fallbackSize) * fontScale;
    if (size) r.size = size;
    if (run.href) r.link = run.href;
    return r;
}

/** 字符宽度估算系数（CJK 全宽 1.0，西文数字 0.6，其余标点 0.5） */
function charWidthFactor(ch: string): number {
    if (/[　-〿㐀-鿿豈-﫿＀-￯]/.test(ch)) return 1;
    if (/[A-Za-z0-9@%]/.test(ch)) return 0.6;
    return 0.5;
}

/** 估算一段文本的渲染宽度（pt） */
function estimateWidth(text: string, size: number): number {
    let w = 0;
    for (const ch of Array.from(text)) w += charWidthFactor(ch) * size;
    return w;
}

/** 文本是否含 CJK（决定行高系数） */
function hasCjk(text: string): boolean {
    return /[　-〿㐀-鿿豈-﫿＀-￯]/.test(text);
}

/**
 * 文本框 -> PDF 元素（先画框内形状底，再画文字）
 *
 * 垂直对齐按编辑器渲染语义（flex-start / center / flex-end）：bottom 把文字
 * 底边贴到框底，middle 居中，top 贴框顶。文字块高度按「行数 × 字号 × 行高」
 * 估算（行数由估算宽度与框宽比较得出）。
 */
function textToElements(el: PptxTextElement): PdfElement[] {
    const out: PdfElement[] = [];

    // 文本框自带的形状底（如「汇报人」的蓝色胶囊、正文里的彩色底板）。
    // 无形状类型的纯文本框填充按矩形绘制（PPT 的 TextBox 也常带底色）。
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
            out.push(...shapeToElements(box));
        } else {
            const fill = solidFillOf(el.fill as PptxFill);
            if (fill) {
                out.push({
                    type: 'rect',
                    x: el.x * PX_TO_PT,
                    y: el.y * PX_TO_PT,
                    width: el.width * PX_TO_PT,
                    height: el.height * PX_TO_PT,
                    fill: fill.color,
                    opacity: fill.alpha < 1 ? fill.alpha : undefined,
                    rotation: el.rotation ? -el.rotation : undefined
                });
            }
        }
    }

    // 「溢出则缩字」（normAutofit fontScale）：PPT/WPS 按此比例整体缩小字号
    const fontScale = (el.fontScale ?? 100) / 100;
    const paragraphs = el.paragraphs?.length
        ? el.paragraphs
        : el.runs?.length || el.text
          ? [{ runs: el.runs, text: el.text }]
          : [];
    const boxW = el.width * PX_TO_PT;
    const boxH = el.height * PX_TO_PT;

    const lines: PdfRun[][] = [];
    let align: PdfRunsElement['align'];
    let textHeight = 0;
    for (const p of paragraphs) {
        const runs = p.runs?.length
            ? p.runs.map((r) => runToPdfRun(r, el.fontSize, fontScale))
            : p.text
              ? [{ text: p.text, size: (el.fontSize ?? DEFAULT_FONT_SIZE) * fontScale }]
              : [];
        const plain = runs.map((r) => r.text).join('');
        const size = Math.max(...runs.map((r) => r.size ?? 0), el.fontSize ?? 0, 1);
        const lh = size * (hasCjk(plain) ? LINE_HEIGHT_CJK : LINE_HEIGHT_LATIN);
        // 估算折行数：与生成端的逐字断行近似
        const wrapped = Math.max(1, Math.ceil(estimateWidth(plain, size) / Math.max(1, boxW) - 0.02));
        textHeight += wrapped * lh;
        if (!plain.trim()) {
            lines.push([{ text: '' }]);
            continue;
        }
        lines.push(runs);
        align = align ?? (p.align as PdfRunsElement['align']);
    }
    if (!lines.length) return out;

    // 垂直对齐：文字块在框内的偏移
    let offsetY = 0;
    if (el.valign === 'middle') offsetY = (boxH - textHeight) / 2;
    else if (el.valign === 'bottom') offsetY = boxH - textHeight;

    const runs: PdfRun[] = [];
    lines.forEach((line, i) => {
        if (i > 0) runs.push({ text: '\n' });
        runs.push(...line);
    });
    out.push({
        type: 'runs',
        x: el.x * PX_TO_PT,
        y: el.y * PX_TO_PT + offsetY,
        width: boxW,
        height: boxH,
        runs,
        align,
        lineHeight: LINE_HEIGHT_CJK,
        spacingAfter: 0,
        // 文本框旋转：PPTX 顺时针 -> PDF 生成端逆时针
        rotation: el.rotation ? -el.rotation : undefined
    });
    return out;
}

// ---- 表格 ----

/** 单元格 -> PDF 单元格（保留合并信息，否则合并表头会错位） */
function cellToPdfCell(cell: PptxTableCell): PdfTableCell {
    const text =
        cell.text ??
        cell.paragraphs
            ?.map((p) => p.text ?? p.runs?.map((r) => r.text ?? '').join('') ?? '')
            .join('\n') ??
        '';
    if (!cell.colSpan && !cell.rowSpan && !cell.fill) return text;
    const obj: PdfTableCellObject = { content: text };
    if (cell.colSpan && cell.colSpan > 1) obj.colspan = cell.colSpan;
    if (cell.rowSpan && cell.rowSpan > 1) obj.rowspan = cell.rowSpan;
    const fill = normalizeColor(cell.fill);
    if (fill) obj.fill = fill;
    return obj;
}

/** 表格字号：取首个显式给出的 run 字号，缺省 14pt */
function tableFontSize(el: PptxTableElement): number {
    for (const row of el.rows) {
        for (const cell of row.cells) {
            for (const p of cell.paragraphs || []) {
                for (const r of p.runs || []) {
                    if (r.fontSize) return r.fontSize;
                }
            }
        }
    }
    return 14;
}

/** 表格 -> 绝对定位表格元素 */
function tableToElement(el: PptxTableElement): PdfElement {
    return {
        type: 'table',
        x: el.x * PX_TO_PT,
        y: el.y * PX_TO_PT,
        rows: el.rows.map((row) =>
            row.cells.flatMap((cell) => (cell.hMerge || cell.vMerge ? [] : [cellToPdfCell(cell)]))
        ),
        header: false,
        fontSize: tableFontSize(el),
        columnWidths: el.colWidths?.map((w) => w * PX_TO_PT)
    };
}

// ---- 元素收集 ----

/**
 * 递归收集元素（展平组合）
 *
 * 组合子元素坐标语义与编辑器渲染端一致：`relative` 表示相对组合原点（需加上
 * 组合位置），其余（page / 绝对坐标）已是页面绝对坐标。
 */
function collectElements(
    list: PptxElement[],
    out: PdfElement[],
    offset?: { x: number; y: number }
): void {
    const dx = offset?.x ?? 0;
    const dy = offset?.y ?? 0;
    for (const el of list) {
        switch (el.type) {
            case 'text': {
                const t = el as PptxTextElement;
                out.push(...textToElements({ ...t, x: t.x + dx, y: t.y + dy }));
                break;
            }
            case 'shape': {
                const s = el as PptxShapeElement;
                out.push(...shapeToElements({ ...s, x: s.x + dx, y: s.y + dy }));
                break;
            }
            case 'image': {
                const img = el as PptxImageElement;
                const src = img.data || img.src;
                if (!src || typeof src !== 'string') break;
                out.push({
                    type: 'image',
                    src,
                    x: (img.x + dx) * PX_TO_PT,
                    y: (img.y + dy) * PX_TO_PT,
                    width: img.width * PX_TO_PT,
                    height: img.height * PX_TO_PT,
                    crop: img.crop
                });
                break;
            }
            case 'table': {
                const t = el as PptxTableElement;
                out.push(tableToElement({ ...t, x: t.x + dx, y: t.y + dy }));
                break;
            }
            case 'group': {
                const g = el as PptxGroupElement;
                if (g.childrenCoordinates === 'relative') {
                    collectElements(g.children || [], out, { x: (g.x || 0) + dx, y: (g.y || 0) + dy });
                } else {
                    collectElements(g.children || [], out);
                }
                break;
            }
            default:
                // chart / diagram / video 等：PDF 直排暂不绘制（图表数据在
                // pptx -> md/html/docx 路径仍有呈现）
                break;
        }
    }
}

/** 幻灯片背景色（纯色 / 渐变取首停色），无背景返回白色兜底 */
function slideBackgroundColor(slide: PptxSlide): string {
    const bg = slide.background;
    if (typeof bg === 'string') return normalizeColor(bg) ?? '#ffffff';
    if (bg) {
        const b = bg as { type?: string; color?: string; stops?: { color: string }[] };
        const color =
            b.type === 'gradient' ? normalizeColor(b.stops?.[0]?.color) : normalizeColor(b.color);
        if (color) return color;
    }
    return '#ffffff';
}

/** 幻灯片背景图片（data URL），无图片背景返回 undefined */
function slideBackgroundImage(slide: PptxSlide): string | undefined {
    const bg = slide.background;
    if (bg && typeof bg === 'object') {
        const b = bg as { type?: string; data?: string; src?: string };
        if (b.type === 'image') return b.data || b.src;
    }
    return undefined;
}

/** 幻灯片是否含中文（决定是否探测系统字体） */
function slidesHaveCjk(slides: PptxSlide[]): boolean {
    let found = false;
    const walk = (els?: PptxElement[]) => {
        for (const el of els || []) {
            if (found) return;
            if (el.type === 'text') {
                const t = el as PptxTextElement;
                const text =
                    t.text ||
                    (t.paragraphs || []).map((p) => p.text || p.runs?.map((r) => r.text ?? '').join('') || '').join('');
                if (/[一-鿿]/.test(text)) found = true;
            } else if (el.type === 'group') {
                walk((el as PptxGroupElement).children);
            }
        }
    };
    for (const s of slides) walk(s.elements);
    return found;
}

/**
 * 把 PPTX 按幻灯片版式直排为 PDF（每页 = 一张幻灯片）
 *
 * @param input - PPTX 字节 / base64 / data URL
 * @param options - 字体选项
 * @returns PDF 字节数据
 */
export async function pptxToPdf(input: DocInput, options: PptxToPdfOptions = {}): Promise<Uint8Array> {
    const doc = (await pptxToStandard(input, { assets: 'both' })) as PptxDocument;
    const pageW = doc.slideSize.width * PX_TO_PT;
    const pageH = doc.slideSize.height * PX_TO_PT;

    // 自动字体：含中文且未提供字体时探测系统字体（仅 Node 端）
    let fonts = options.fonts;
    if (!fonts && options.autoFont !== false && slidesHaveCjk(doc.slides)) {
        try {
            fonts = await detectSystemFonts();
        } catch {
            // 浏览器端无系统字体：交给 jsonToPdf 抛友好的中文提示
        }
    }

    // 主题色：scheme:bg1 / scheme:accent1 等引用需要 clrScheme + clrMap 落到具体色值。
    // 必须在所有 await 之后再设置：下面的取色过程是同步的，事件循环无法插队，
    // 否则并发转换两个 PPTX 时会串用对方的主题色板。
    resolveScheme = buildSchemeResolver(
        doc.themeXmls?.[0] ?? doc.themeXml,
        (doc.masters?.[0] as { __rawXml?: string } | undefined)?.__rawXml
    );

    const pages: PdfPageSpec[] = [];
    try {
        for (const slide of doc.slides) {
            if (slide.hidden) continue;
            const bg = slideBackgroundColor(slide);
            const elements: PdfElement[] = [{ type: 'rect', x: 0, y: 0, width: pageW, height: pageH, fill: bg }];
            // 背景图片（封面整图背景很常见）：铺满整页，画在底色之上、内容之下
            const bgImage = slideBackgroundImage(slide);
            if (bgImage) {
                elements.push({ type: 'image', src: bgImage, x: 0, y: 0, width: pageW, height: pageH });
            }
            collectElements(slide.elements || [], elements);
            pages.push({ size: [pageW, pageH], margin: 0, elements });
        }
        if (!pages.length) pages.push({ size: [pageW, pageH], margin: 0, elements: [] });
    } finally {
        resolveScheme = () => undefined;
    }

    return jsonToPdf({ pages, fonts, defaultSize: DEFAULT_FONT_SIZE }, {});
}
