/**
 * PPTX 图表 → 共享渲染模型
 *
 * PPTX 与 XLSX 的图表画法统一交给 `chart/svg`，本模块只负责把
 * ppt-parser 的图表元素映射成格式无关的 `ChartSpec`。
 *
 * @module pptx/chart
 */

import type { ChartSeriesSpec, ChartSpec, ChartType } from '../chart/svg';

/** PPTX 图表元素（ppt-parser 的 PptxChartElement，字段按需取用） */
interface PptxChartLike {
    chartType?: string;
    categories?: string[];
    series?: { name?: string; values?: unknown[]; color?: string; pointColors?: string[] }[];
    title?: string;
    legend?: boolean;
    legendPosition?: string;
    dataLabels?: boolean;
    grouping?: string;
    holeSize?: number;
    smooth?: boolean;
    marker?: boolean;
    varyColors?: boolean;
}

/** 主题色覆盖项 */
export interface PptxChartTheme {
    /** 调色板 */
    palette?: string[];
    /** 文本色 */
    textColor?: string;
    /** 标题色 */
    titleColor?: string;
}

/** PPTX chartType → 共享渲染类型 */
const TYPE_MAP: Record<string, ChartType> = {
    barChart: 'column',
    lineChart: 'line',
    areaChart: 'area',
    pieChart: 'pie',
    pie3DChart: 'pie',
    ofPieChart: 'pie',
    doughnutChart: 'doughnut',
    scatterChart: 'scatter',
    bubbleChart: 'scatter',
    radarChart: 'radar',
    line3DChart: 'line',
    area3DChart: 'area',
    bar3DChart: 'column',
    stockChart: 'line',
    surfaceChart: 'line',
    surface3DChart: 'line'
};

/** 系列取值：兼容 {y} / {x,y} / {x,y,size} / 裸数字 */
function seriesValues(values: unknown[] | undefined): {
    values: (number | null)[];
    xValues?: (number | null)[];
} {
    const out: (number | null)[] = [];
    const xs: (number | null)[] = [];
    let hasX = false;
    for (const v of values || []) {
        if (v === null || v === undefined) {
            out.push(null);
            xs.push(null);
            continue;
        }
        if (typeof v === 'number') {
            out.push(v);
            xs.push(null);
            continue;
        }
        const pt = v as { x?: number; y?: number };
        if (typeof pt.x === 'number') {
            hasX = true;
            xs.push(pt.x);
        } else {
            xs.push(null);
        }
        out.push(typeof pt.y === 'number' ? pt.y : null);
    }
    return hasX ? { values: out, xValues: xs } : { values: out };
}

/**
 * 把 PPTX 图表元素转成共享渲染模型
 *
 * @param el - 图表元素
 * @param theme - 主题色（调色板 / 文本色 / 标题色）
 */
export function pptxChartToSpec(
    el: PptxChartLike | undefined | null,
    theme?: PptxChartTheme
): ChartSpec {
    const raw = String(el?.chartType || 'barChart');
    const [base, variant] = raw.split('|');
    let type: ChartType = TYPE_MAP[base] || 'column';
    if (type === 'column' && variant === 'bar') type = 'bar';

    const series: ChartSeriesSpec[] = (el?.series || []).map((s) => {
        const { values, xValues } = seriesValues(s.values);
        const out: ChartSeriesSpec = { name: s.name, values };
        if (xValues) out.xValues = xValues;
        if (s.color) out.color = s.color;
        if (s.pointColors?.length) out.pointColors = s.pointColors;
        return out;
    });

    const spec: ChartSpec = {
        type,
        categories: el?.categories || [],
        series
    };
    if (el?.title) spec.title = el.title;
    if (el?.legend) {
        const pos = String(el.legendPosition || 'r').toLowerCase();
        spec.legend = (['l', 'r', 't', 'b'] as const).includes(pos as 'l')
            ? (pos as 'l' | 'r' | 't' | 'b')
            : 'r';
    }
    if (el?.dataLabels) spec.labels = { value: true };
    const grouping = String(el?.grouping || '').toLowerCase();
    if (grouping === 'stacked') spec.grouping = 'stacked';
    else if (grouping === 'percentstacked') spec.grouping = 'percentStacked';
    else if (grouping === 'clustered') spec.grouping = 'clustered';
    if (el?.holeSize !== undefined) spec.holeSize = el.holeSize;
    if (el?.smooth !== undefined) spec.smooth = el.smooth;
    if (el?.marker !== undefined) spec.marker = el.marker;
    if (el?.varyColors !== undefined) spec.varyColors = el.varyColors;
    if (theme?.palette?.length) spec.palette = theme.palette;
    if (theme?.textColor) spec.textColor = theme.textColor;
    if (theme?.titleColor) spec.titleColor = theme.titleColor;
    return spec;
}
