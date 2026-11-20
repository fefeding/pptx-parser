/**
 * 共享图表渲染（SVG）
 *
 * PPTX 与 XLSX 共用同一套绘图内核：各格式负责把 OOXML 解析成 `ChartSpec`，
 * 画法统一在这里，避免两套 SVG 渲染器各画各的、能力越差越多。
 *
 * 纯字符串输出、不触碰 DOM，因此 Node（库）与浏览器（webview）都能用，
 * 也便于做快照测试。
 *
 * 支持：
 * - 类型：柱（column）/ 条（bar）/ 折线（line）/ 面积（area）/ 饼（pie）/
 *   环形（doughnut）/ 散点（scatter）/ 雷达（radar）
 * - 分组：簇状（clustered）/ 堆积（stacked）/ 百分比堆积（percentStacked）
 * - 图例：none / 左 / 右 / 上 / 下（饼图的图例项是分类，其余是系列）
 * - 数据标签：数值 / 百分比 / 分类名
 * - 数字格式：Excel 数字格式串（如 `¥#,##0`），轴标签与数据标签通用
 * - 其他：标题、平滑折线、数据点标记、环形内径、逐点配色
 *
 * @module chart/svg
 */

import { formatByPattern } from '../core/numfmt';

/** 图表类型（与格式无关） */
export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar';

/** 图例位置 */
export type ChartLegendPos = 'none' | 'l' | 'r' | 't' | 'b';

/** 数据系列 */
export interface ChartSeriesSpec {
    /** 系列名 */
    name?: string;
    /** 数值（null 表示缺值） */
    values: (number | null)[];
    /** 散点图的 x 值；缺省用序号 */
    xValues?: (number | null)[];
    /** 系列色 #RRGGBB */
    color?: string;
    /** 逐点配色（饼图扇区） */
    pointColors?: string[];
}

/** 数据标签开关 */
export interface ChartLabels {
    /** 显示数值 */
    value?: boolean;
    /** 显示百分比 */
    percent?: boolean;
    /** 显示分类名 */
    category?: boolean;
}

/** 图表描述（与格式无关） */
export interface ChartSpec {
    /** 图表类型 */
    type: ChartType;
    /** 标题 */
    title?: string;
    /** 分类轴 */
    categories?: (string | number | null)[];
    /** 数据系列 */
    series?: ChartSeriesSpec[];
    /** 图例位置 */
    legend?: ChartLegendPos;
    /** 数据标签 */
    labels?: ChartLabels | boolean;
    /** 分组方式 */
    grouping?: 'clustered' | 'stacked' | 'percentStacked';
    /** 环形图内径百分比 */
    holeSize?: number;
    /** 折线图数据点标记 */
    marker?: boolean;
    /** 折线图平滑 */
    smooth?: boolean;
    /** 饼图逐点配色 */
    varyColors?: boolean;
    /** 数值格式（Excel 数字格式串） */
    numFmt?: string;
    /** 调色板 */
    palette?: string[];
    /** 文本色 */
    textColor?: string;
    /** 标题色 */
    titleColor?: string;
    /** 背景色 */
    background?: string;
    /** 边框色 */
    borderColor?: string;
    /** 字体 */
    fontFamily?: string;
}

/** 渲染尺寸 */
export interface ChartRenderOptions {
    width: number;
    height: number;
}

/** Office 默认配色（Excel / WPS 内置图表同款） */
export const DEFAULT_CHART_PALETTE = [
    '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E'
];

const DEFAULT_FONT = "'Microsoft YaHei','PingFang SC','Segoe UI',sans-serif";

/** XML/HTML 转义 */
function esc(value: unknown): string {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/** 数值转显示文本（优先套用 Excel 数字格式） */
function numText(v: number, fmt?: string): string {
    if (!isFinite(v)) return '';
    if (fmt) {
        const text = formatByPattern(v, fmt);
        if (text) return text;
    }
    const abs = Math.abs(v);
    if (abs >= 10000) return v.toLocaleString('zh-CN', { maximumFractionDigits: 0 });
    if (Number.isInteger(v)) return String(v);
    return String(Math.round(v * 100) / 100);
}

/** 轴刻度取整（1/2/5 × 10^n） */
function niceCeil(v: number): number {
    if (v <= 0) return 1;
    const exp = Math.floor(Math.log10(v));
    const base = Math.pow(10, exp);
    const n = v / base;
    const step = n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10;
    return step * base;
}

/** 平滑折线路径（Catmull-Rom 转二次贝塞尔） */
function smoothPath(points: { x: number; y: number }[]): string {
    if (!points.length) return '';
    if (points.length < 3) {
        return points.map((p, i) => `${i ? 'L' : 'M'}${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
    }
    let d = `M${points[0].x.toFixed(1)},${points[0].y.toFixed(1)}`;
    for (let i = 0; i < points.length - 1; i++) {
        const p0 = points[i - 1] || points[i];
        const p1 = points[i];
        const p2 = points[i + 1];
        const p3 = points[i + 2] || p2;
        const c1x = p1.x + (p2.x - p0.x) / 6;
        const c1y = p1.y + (p2.y - p0.y) / 6;
        const c2x = p2.x - (p3.x - p1.x) / 6;
        const c2y = p2.y - (p3.y - p1.y) / 6;
        d += ` C${c1x.toFixed(1)},${c1y.toFixed(1)} ${c2x.toFixed(1)},${c2y.toFixed(1)} ${p2.x.toFixed(1)},${p2.y.toFixed(1)}`;
    }
    return d;
}

/** 折线（不平滑）路径 */
function polyPath(points: { x: number; y: number }[]): string {
    if (!points.length) return '';
    return (
        `M${points[0].x.toFixed(1)},${points[0].y.toFixed(1)}` +
        points
            .slice(1)
            .map((p) => `L${p.x.toFixed(1)},${p.y.toFixed(1)}`)
            .join('')
    );
}

/** 分类标签（过长时截断） */
function catLabel(cat: unknown, max = 8): string {
    const text = String(cat ?? '');
    return text.length > max ? text.slice(0, max - 1) + '…' : text;
}

/** 图例项 */
interface LegendItem {
    label: string;
    color: string;
    extra?: string;
}

/** 归一化数据标签开关 */
function labelFlags(labels: ChartSpec['labels']): Required<ChartLabels> {
    if (labels === true) return { value: true, percent: false, category: false };
    if (!labels) return { value: false, percent: false, category: false };
    return { value: !!labels.value, percent: !!labels.percent, category: !!labels.category };
}

/** 是否饼/环类（无直角坐标轴） */
function isPieLike(type: ChartType): boolean {
    return type === 'pie' || type === 'doughnut';
}

/**
 * 渲染图表为 SVG 字符串
 *
 * @param spec - 图表描述
 * @param opts - 尺寸
 */
export function renderChartSvg(spec: ChartSpec, opts: ChartRenderOptions): string {
    const W = Math.max(80, Math.round(opts.width || 640));
    const H = Math.max(60, Math.round(opts.height || 380));
    const palette = spec.palette && spec.palette.length ? spec.palette : DEFAULT_CHART_PALETTE;
    const textColor = spec.textColor || '#6b7280';
    const titleColor = spec.titleColor || '#1f2328';
    const font = spec.fontFamily || DEFAULT_FONT;
    const labels = labelFlags(spec.labels);
    const series = (spec.series || []).filter(Boolean);
    const categories = spec.categories || [];
    const legendPos: ChartLegendPos = spec.legend || 'none';

    const parts: string[] = [];
    parts.push(
        `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${font}">`
    );
    if (spec.background !== 'none') {
        parts.push(
            `<rect x="0.5" y="0.5" width="${W - 1}" height="${H - 1}" fill="${spec.background || '#ffffff'}" stroke="${spec.borderColor || '#d8dce3'}"/>`
        );
    }

    // 标题
    let top = 4;
    if (spec.title) {
        parts.push(
            `<text x="${(W / 2).toFixed(1)}" y="17" text-anchor="middle" font-size="13" font-weight="600" fill="${titleColor}">${esc(spec.title)}</text>`
        );
        top = 26;
    }

    // 图例内容：饼/环用分类，其余用系列
    const legendItems: LegendItem[] = [];
    if (legendPos !== 'none') {
        if (isPieLike(spec.type)) {
            const vals = series[0]?.values || [];
            const total = vals.reduce<number>((s, v) => s + (typeof v === 'number' ? v : 0), 0);
            categories.forEach((cat, i) => {
                const v = typeof vals[i] === 'number' ? (vals[i] as number) : 0;
                const color = series[0]?.pointColors?.[i] || palette[i % palette.length];
                const extra = [
                    labels.value ? numText(v, spec.numFmt) : '',
                    labels.percent && total ? `${((v / total) * 100).toFixed(0)}%` : ''
                ]
                    .filter(Boolean)
                    .join(' ');
                legendItems.push({ label: String(cat ?? ''), color, extra: extra || undefined });
            });
        } else {
            series.forEach((s, i) => {
                legendItems.push({ label: s.name || `系列${i + 1}`, color: s.color || palette[i % palette.length] });
            });
        }
    }

    // 图例占位
    let legendW = 0;
    let legendH = 0;
    if (legendItems.length) {
        if (legendPos === 'r' || legendPos === 'l') {
            const widest = legendItems.reduce(
                (m, it) => Math.max(m, it.label.length * 7 + (it.extra ? it.extra.length * 7 : 0) + 26),
                60
            );
            legendW = Math.min(200, Math.max(80, widest));
        } else if (legendPos === 'b' || legendPos === 't') {
            legendH = 20;
        }
    }
    const legendLeft = legendPos === 'l' ? legendW : 0;

    const pie = isPieLike(spec.type);
    const plotLeft = legendLeft + (pie ? 8 : 64);
    const plotRight = W - (legendPos === 'r' ? legendW : 0) - 8;
    const plotTop = top + (legendPos === 't' ? legendH : 0);
    const plotBottom = H - legendH - (pie ? 8 : 26);
    const plotW = Math.max(20, plotRight - plotLeft);
    const plotH = Math.max(20, plotBottom - plotTop);

    if (pie) {
        parts.push(...pieParts(spec, series, categories, plotLeft, plotTop, plotW, plotH, palette, labels, textColor));
    } else if (spec.type === 'scatter') {
        parts.push(...scatterParts(spec, series, plotLeft, plotTop, plotW, plotH, palette, labels, textColor));
    } else if (spec.type === 'radar') {
        parts.push(...radarParts(spec, series, categories, plotLeft, plotTop, plotW, plotH, palette, textColor));
    } else if (spec.type === 'line' || spec.type === 'area') {
        parts.push(...lineParts(spec, series, categories, plotLeft, plotTop, plotW, plotH, palette, labels, textColor));
    } else {
        parts.push(...barParts(spec, series, categories, plotLeft, plotTop, plotW, plotH, palette, labels, textColor));
    }

    // 图例
    if (legendItems.length && legendPos !== 'none') {
        parts.push(
            ...legendParts(legendItems, legendPos, W, H, plotLeft, plotRight, legendLeft, legendH, textColor)
        );
    }

    parts.push('</svg>');
    return parts.join('');
}

/** 图例绘制 */
function legendParts(
    items: LegendItem[],
    pos: ChartLegendPos,
    W: number,
    H: number,
    plotLeft: number,
    plotRight: number,
    legendLeft: number,
    legendH: number,
    textColor: string
): string[] {
    const parts: string[] = [];
    if (pos === 'r' || pos === 'l') {
        const x = pos === 'r' ? plotRight + 8 : 6;
        let y = 16;
        for (const item of items) {
            parts.push(`<rect x="${x}" y="${y - 8}" width="9" height="9" fill="${item.color}"/>`);
            parts.push(
                `<text x="${x + 13}" y="${y}" font-size="11" fill="${textColor}">${esc(item.label)}${item.extra ? ' ' + esc(item.extra) : ''}</text>`
            );
            y += 15;
        }
        return parts;
    }
    // 上 / 下：横向排布
    const y = pos === 't' ? 14 : H - legendH + 16;
    let x = pos === 't' ? legendLeft + 4 : plotLeft;
    for (const item of items) {
        const w = 20 + item.label.length * 7 + (item.extra ? item.extra.length * 7 : 0);
        if (x + w > W - 6) break;
        parts.push(`<rect x="${x}" y="${y - 8}" width="9" height="9" fill="${item.color}"/>`);
        parts.push(
            `<text x="${x + 13}" y="${y}" font-size="11" fill="${textColor}">${esc(item.label)}${item.extra ? ' ' + esc(item.extra) : ''}</text>`
        );
        x += w;
    }
    return parts;
}

/** 数值范围（含堆积） */
function valueRange(
    spec: ChartSpec,
    series: ChartSeriesSpec[],
    categories: unknown[]
): { min: number; max: number } {
    const stacked = spec.grouping === 'stacked' || spec.grouping === 'percentStacked';
    let min = Infinity;
    let max = -Infinity;
    if (stacked) {
        for (let i = 0; i < categories.length; i++) {
            let sum = 0;
            for (const s of series) sum += typeof s.values[i] === 'number' ? (s.values[i] as number) : 0;
            min = Math.min(min, Math.min(0, sum));
            max = Math.max(max, sum);
        }
    } else {
        for (const s of series) {
            for (const v of s.values) {
                if (typeof v !== 'number') continue;
                min = Math.min(min, v);
                max = Math.max(max, v);
            }
        }
    }
    if (!isFinite(min)) {
        min = 0;
        max = 1;
    }
    if (min > 0) min = 0;
    // 百分比堆积：每根柱/每个点的高度是占比，上限恒为 100%
    if (spec.grouping === 'percentStacked') return { min: 0, max: 100 };
    return { min, max: max || 1 };
}

/** 堆积/百分比堆积时的每类合计 */
function stackTotals(series: ChartSeriesSpec[], count: number): number[] {
    const totals: number[] = new Array(count).fill(0);
    for (const s of series) {
        for (let i = 0; i < count; i++) {
            totals[i] += typeof s.values[i] === 'number' ? (s.values[i] as number) : 0;
        }
    }
    return totals;
}

/** 柱状 / 条形图 */
function barParts(
    spec: ChartSpec,
    series: ChartSeriesSpec[],
    categories: unknown[],
    left: number,
    top: number,
    w: number,
    h: number,
    palette: string[],
    labels: Required<ChartLabels>,
    textColor: string
): string[] {
    const parts: string[] = [];
    const horizontal = spec.type === 'bar';
    const count = Math.max(1, categories.length);
    const stacked = spec.grouping === 'stacked' || spec.grouping === 'percentStacked';
    const { min, max } = valueRange(spec, series, categories);
    const axisMax = stacked ? max : niceCeil(max);
    const axisMin = min < 0 ? -niceCeil(-min) : 0;
    const span = axisMax - axisMin || 1;
    const totals = stackTotals(series, count);

    // 网格线与刻度
    const ticks = 4;
    for (let i = 0; i <= ticks; i++) {
        const ratio = i / ticks;
        const value = axisMin + span * ratio;
        if (horizontal) {
            const x = left + w * ratio;
            parts.push(`<line x1="${x.toFixed(1)}" y1="${top}" x2="${x.toFixed(1)}" y2="${top + h}" stroke="#e6e8ec"/>`);
            parts.push(
                `<text x="${(x - 4).toFixed(1)}" y="${top + h + 14}" font-size="10" text-anchor="end" fill="${textColor}">${esc(numText(value, spec.numFmt))}</text>`
            );
        } else {
            const y = top + h - h * ratio;
            parts.push(`<line x1="${left}" y1="${y.toFixed(1)}" x2="${left + w}" y2="${y.toFixed(1)}" stroke="#e6e8ec"/>`);
            parts.push(
                `<text x="${(left - 6).toFixed(1)}" y="${(y + 3.5).toFixed(1)}" font-size="10" text-anchor="end" fill="${textColor}">${esc(numText(value, spec.numFmt))}</text>`
            );
        }
    }

    // 柱体
    const group = (horizontal ? h : w) / count;
    const barW = Math.max(2, (group * 0.62) / Math.max(1, stacked ? 1 : series.length));
    series.forEach((s, si) => {
        const color = s.color || palette[si % palette.length];
        let acc = 0;
        s.values.forEach((v, i) => {
            if (typeof v !== 'number') return;
            let value = v;
            if (spec.grouping === 'percentStacked') {
                const total = totals[i] || 1;
                value = (v / total) * 100;
            } else if (stacked) {
                value = acc + v;
            }
            const base = axisMin >= 0 ? 0 : axisMin;
            if (horizontal) {
                const len = ((value - base) / span) * w;
                const x = value >= base ? left : left + len;
                const y =
                    top +
                    group * i +
                    (group - barW * (stacked ? 1 : series.length)) / 2 +
                    barW * (stacked ? 0 : si);
                parts.push(
                    `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${Math.abs(len).toFixed(1)}" height="${barW.toFixed(1)}" fill="${color}"/>`
                );
            } else {
                const len = ((value - base) / span) * h;
                const y = value >= base ? top + h - len : top + h;
                const x =
                    left +
                    group * i +
                    (group - barW * (stacked ? 1 : series.length)) / 2 +
                    barW * (stacked ? 0 : si);
                parts.push(
                    `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${Math.abs(len).toFixed(1)}" fill="${color}"/>`
                );
            }
            if (stacked) acc += v;
        });
    });

    // 分类标签
    categories.forEach((cat, i) => {
        const label = catLabel(cat);
        if (horizontal) {
            const y = top + group * (i + 0.5) + 4;
            parts.push(
                `<text x="${left - 6}" y="${y.toFixed(1)}" font-size="10" text-anchor="end" fill="${textColor}">${esc(label)}</text>`
            );
        } else {
            const x = left + group * (i + 0.5);
            parts.push(
                `<text x="${x.toFixed(1)}" y="${top + h + 14}" font-size="10" text-anchor="middle" fill="${textColor}">${esc(label)}</text>`
            );
        }
    });

    // 轴线
    parts.push(
        horizontal
            ? `<line x1="${left}" y1="${top}" x2="${left}" y2="${top + h}" stroke="#c9ced6"/>`
            : `<line x1="${left}" y1="${top + h}" x2="${left + w}" y2="${top + h}" stroke="#c9ced6"/>`
    );
    return parts;
}

/** 折线 / 面积图 */
function lineParts(
    spec: ChartSpec,
    series: ChartSeriesSpec[],
    categories: unknown[],
    left: number,
    top: number,
    w: number,
    h: number,
    palette: string[],
    labels: Required<ChartLabels>,
    textColor: string
): string[] {
    const parts: string[] = [];
    const count = Math.max(1, categories.length);
    const stacked = spec.grouping === 'stacked' || spec.grouping === 'percentStacked';
    const { min, max } = valueRange(spec, series, categories);
    const axisMax = stacked ? max : niceCeil(max);
    const axisMin = min < 0 ? -niceCeil(-min) : 0;
    const span = axisMax - axisMin || 1;

    // 网格线与刻度
    const ticks = 4;
    for (let i = 0; i <= ticks; i++) {
        const ratio = i / ticks;
        const value = axisMin + span * ratio;
        const y = top + h - h * ratio;
        parts.push(`<line x1="${left}" y1="${y.toFixed(1)}" x2="${left + w}" y2="${y.toFixed(1)}" stroke="#e6e8ec"/>`);
        parts.push(
            `<text x="${(left - 6).toFixed(1)}" y="${(y + 3.5).toFixed(1)}" font-size="10" text-anchor="end" fill="${textColor}">${esc(numText(value, spec.numFmt))}</text>`
        );
    }

    const step = count > 1 ? w / (count - 1) : w;
    const totals = stackTotals(series, count);
    series.forEach((s, si) => {
        const color = s.color || palette[si % palette.length];
        const points: { x: number; y: number }[] = [];
        let acc = 0;
        s.values.forEach((v, i) => {
            if (typeof v !== 'number') return;
            let value = v;
            if (spec.grouping === 'percentStacked') {
                const total = totals[i] || 1;
                value = (v / total) * 100;
            } else if (stacked) {
                value = acc + v;
            }
            points.push({ x: left + step * i, y: top + h - ((value - axisMin) / span) * h });
            if (stacked) acc += v;
        });
        if (!points.length) return;
        const path = spec.smooth ? smoothPath(points) : polyPath(points);
        if (spec.type === 'area') {
            const base = top + h;
            parts.push(
                `<path d="${path} L${points[points.length - 1].x.toFixed(1)},${base} L${points[0].x.toFixed(1)},${base} Z" fill="${color}" fill-opacity="0.18"/>`
            );
        }
        parts.push(
            `<path d="${path}" fill="none" stroke="${color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>`
        );
        if (spec.marker !== false) {
            for (const p of points) {
                parts.push(
                    `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="2.6" fill="#fff" stroke="${color}" stroke-width="1.6"/>`
                );
            }
        }
        if (labels.value) {
            points.forEach((p, i) => {
                const raw = s.values[i];
                if (typeof raw !== 'number') return;
                parts.push(
                    `<text x="${p.x.toFixed(1)}" y="${(p.y - 8).toFixed(1)}" font-size="10" text-anchor="middle" fill="${textColor}">${esc(numText(raw, spec.numFmt))}</text>`
                );
            });
        }
    });

    categories.forEach((cat, i) => {
        const x = left + step * i;
        parts.push(
            `<text x="${x.toFixed(1)}" y="${top + h + 14}" font-size="10" text-anchor="middle" fill="${textColor}">${esc(catLabel(cat))}</text>`
        );
    });

    parts.push(`<line x1="${left}" y1="${top + h}" x2="${left + w}" y2="${top + h}" stroke="#c9ced6"/>`);
    return parts;
}

/** 散点图（x 用系列的 xValues，缺省用序号） */
function scatterParts(
    spec: ChartSpec,
    series: ChartSeriesSpec[],
    left: number,
    top: number,
    w: number,
    h: number,
    palette: string[],
    labels: Required<ChartLabels>,
    textColor: string
): string[] {
    const parts: string[] = [];
    const allX: number[] = [];
    const allY: number[] = [];
    for (const s of series) {
        s.values.forEach((v, i) => {
            if (typeof v !== 'number') return;
            allY.push(v);
            const x = s.xValues && typeof s.xValues[i] === 'number' ? (s.xValues[i] as number) : i + 1;
            allX.push(x);
        });
    }
    if (!allX.length) {
        parts.push(
            `<text x="${(left + w / 2).toFixed(1)}" y="${(top + h / 2).toFixed(1)}" font-size="11" text-anchor="middle" fill="${textColor}">无数据</text>`
        );
        return parts;
    }
    const xMinRaw = Math.min(...allX);
    const xMaxRaw = Math.max(...allX);
    const yMinRaw = Math.min(...allY);
    const yMaxRaw = Math.max(...allY);
    const xMin = xMinRaw;
    const xMax = xMaxRaw === xMinRaw ? xMinRaw + 1 : xMaxRaw;
    const yMin = yMinRaw > 0 ? 0 : yMinRaw;
    const yMax = niceCeil(yMaxRaw);
    const ySpan = yMax - yMin || 1;
    const xSpan = xMax - xMin || 1;

    // 网格线与刻度
    for (let i = 0; i <= 4; i++) {
        const ratio = i / 4;
        const y = top + h - h * ratio;
        parts.push(`<line x1="${left}" y1="${y.toFixed(1)}" x2="${left + w}" y2="${y.toFixed(1)}" stroke="#e6e8ec"/>`);
        parts.push(
            `<text x="${(left - 6).toFixed(1)}" y="${(y + 3.5).toFixed(1)}" font-size="10" text-anchor="end" fill="${textColor}">${esc(numText(yMin + ySpan * ratio, spec.numFmt))}</text>`
        );
    }

    series.forEach((s, si) => {
        const color = s.color || palette[si % palette.length];
        s.values.forEach((v, i) => {
            if (typeof v !== 'number') return;
            const xv = s.xValues && typeof s.xValues[i] === 'number' ? (s.xValues[i] as number) : i + 1;
            const x = left + ((xv - xMin) / xSpan) * w;
            const y = top + h - ((v - yMin) / ySpan) * h;
            parts.push(`<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="3.5" fill="${color}" fill-opacity="0.85"/>`);
            if (labels.value) {
                parts.push(
                    `<text x="${x.toFixed(1)}" y="${(y - 8).toFixed(1)}" font-size="10" text-anchor="middle" fill="${textColor}">${esc(numText(v, spec.numFmt))}</text>`
                );
            }
        });
    });

    parts.push(`<line x1="${left}" y1="${top + h}" x2="${left + w}" y2="${top + h}" stroke="#c9ced6"/>`);
    return parts;
}

/** 雷达图 */
function radarParts(
    spec: ChartSpec,
    series: ChartSeriesSpec[],
    categories: unknown[],
    left: number,
    top: number,
    w: number,
    h: number,
    palette: string[],
    textColor: string
): string[] {
    const parts: string[] = [];
    const n = Math.max(1, categories.length);
    const { min, max } = valueRange(spec, series, categories);
    const span = max - min || 1;
    const cx = left + w / 2;
    const cy = top + h / 2;
    const rad = Math.min(w, h) / 2 - 22;
    const angle = (i: number): number => -Math.PI / 2 + (i / n) * Math.PI * 2;

    for (let ring = 1; ring <= 4; ring++) {
        const rr = (rad * ring) / 4;
        const p = categories
            .map(
                (_: unknown, i: number) =>
                    `${(cx + rr * Math.cos(angle(i))).toFixed(1)},${(cy + rr * Math.sin(angle(i))).toFixed(1)}`
            )
            .join(' ');
        parts.push(`<polygon points="${p}" fill="none" stroke="#e6e8ec"/>`);
    }
    categories.forEach((cat, i) => {
        const x = cx + rad * Math.cos(angle(i));
        const y = cy + rad * Math.sin(angle(i));
        parts.push(
            `<line x1="${cx.toFixed(1)}" y1="${cy.toFixed(1)}" x2="${x.toFixed(1)}" y2="${y.toFixed(1)}" stroke="#e6e8ec"/>`
        );
        parts.push(
            `<text x="${(cx + (rad + 12) * Math.cos(angle(i))).toFixed(1)}" y="${(cy + (rad + 12) * Math.sin(angle(i)) + 3).toFixed(1)}" font-size="10" text-anchor="middle" fill="${textColor}">${esc(catLabel(cat))}</text>`
        );
    });
    series.forEach((s, si) => {
        const color = s.color || palette[si % palette.length];
        const p = categories
            .map((_: unknown, i: number) => {
                const v = typeof s.values[i] === 'number' ? (s.values[i] as number) : 0;
                const rr = rad * Math.max(0, Math.min(1, (v - min) / span));
                return `${(cx + rr * Math.cos(angle(i))).toFixed(1)},${(cy + rr * Math.sin(angle(i))).toFixed(1)}`;
            })
            .join(' ');
        parts.push(`<polygon points="${p}" fill="${color}" fill-opacity="0.22" stroke="${color}" stroke-width="2"/>`);
    });
    return parts;
}

/** 饼图 / 环形图 */
function pieParts(
    spec: ChartSpec,
    series: ChartSeriesSpec[],
    categories: unknown[],
    left: number,
    top: number,
    w: number,
    h: number,
    palette: string[],
    labels: Required<ChartLabels>,
    textColor: string
): string[] {
    const parts: string[] = [];
    const s = series[0];
    const vals = (s?.values || []).map((v) => (typeof v === 'number' ? v : 0));
    const total = vals.reduce((a, b) => a + b, 0);
    const cx = left + w / 2;
    const cy = top + h / 2;
    const radius = Math.max(10, Math.min(w, h) / 2 - (labels.percent ? 24 : 10));

    if (total <= 0) {
        parts.push(
            `<text x="${cx.toFixed(1)}" y="${cy.toFixed(1)}" font-size="11" text-anchor="middle" fill="#9aa1ac">无数据</text>`
        );
        return parts;
    }

    const donut = spec.type === 'doughnut';
    const inner = donut ? radius * Math.max(0.1, Math.min(0.9, (spec.holeSize ?? 50) / 100)) : 0;

    let angle = -Math.PI / 2;
    vals.forEach((v, i) => {
        const frac = v / total;
        const end = angle + frac * Math.PI * 2;
        const color =
            s?.pointColors?.[i] ||
            (spec.varyColors === false ? s?.color || palette[0] : palette[i % palette.length]);
        const large = frac > 0.5 ? 1 : 0;
        const x1 = cx + radius * Math.cos(angle);
        const y1 = cy + radius * Math.sin(angle);
        const x2 = cx + radius * Math.cos(end);
        const y2 = cy + radius * Math.sin(end);
        if (frac >= 0.9999) {
            parts.push(`<circle cx="${cx.toFixed(1)}" cy="${cy.toFixed(1)}" r="${radius.toFixed(1)}" fill="${color}"/>`);
        } else {
            let d =
                `M${cx.toFixed(1)},${cy.toFixed(1)} L${x1.toFixed(1)},${y1.toFixed(1)} ` +
                `A${radius.toFixed(1)},${radius.toFixed(1)} 0 ${large} 1 ${x2.toFixed(1)},${y2.toFixed(1)} Z`;
            if (donut) {
                const ix1 = cx + inner * Math.cos(end);
                const iy1 = cy + inner * Math.sin(end);
                const ix2 = cx + inner * Math.cos(angle);
                const iy2 = cy + inner * Math.sin(angle);
                d +=
                    ` M${ix1.toFixed(1)},${iy1.toFixed(1)} ` +
                    `A${inner.toFixed(1)},${inner.toFixed(1)} 0 ${large} 0 ${ix2.toFixed(1)},${iy2.toFixed(1)} Z`;
                parts.push(`<path d="${d}" fill="${color}" fill-rule="evenodd"/>`);
            } else {
                parts.push(`<path d="${d}" fill="${color}"/>`);
            }
        }

        // 数据标签：只显示百分比时放在扇区外，其余放扇区内
        const mid = (angle + end) / 2;
        if (labels.percent || labels.value || labels.category) {
            const bits: string[] = [];
            if (labels.category) bits.push(catLabel(categories[i], 6));
            if (labels.value) bits.push(numText(v, spec.numFmt));
            if (labels.percent) bits.push(`${(frac * 100).toFixed(0)}%`);
            if (bits.length) {
                const text = bits.join(' ');
                // 只有百分比时放扇区内（放得下），带分类名/数值的放扇区外（Excel 的做法）
                const percentOnly = labels.percent && !labels.value && !labels.category;
                if (percentOnly) {
                    const lr = inner > 0 ? (radius + inner) / 2 : radius * 0.68;
                    parts.push(
                        `<text x="${(cx + lr * Math.cos(mid)).toFixed(1)}" y="${(cy + lr * Math.sin(mid) + 3).toFixed(1)}" font-size="10" text-anchor="middle" fill="#ffffff">${esc(text)}</text>`
                    );
                } else {
                    const outR = radius + 12;
                    const lx = cx + outR * Math.cos(mid);
                    const ly = cy + outR * Math.sin(mid);
                    parts.push(
                        `<text x="${lx.toFixed(1)}" y="${(ly + 3).toFixed(1)}" font-size="10" text-anchor="middle" fill="${textColor}">${esc(text)}</text>`
                    );
                }
            }
        }
        angle = end;
    });

    return parts;
}
