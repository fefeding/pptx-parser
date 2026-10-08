// 组合图 multi-plot 解析：c:plotArea 下所有 c:*Chart 节点都应被提取到 `plots`，
// 而非像旧实现那样仅取首个（extractChart 原 `break` 丢失后续绘图区）。
// 主 plot 的属性仍映射到顶层 chartType/series，保持单图表场景向后兼容。
import { describe, it, expect } from 'vitest';
import { extractChart } from '../src/serializer/json-from-pptx';

/** 构造一个带 name/cat/val 的 series */
function ser(name: string, cats: string[], vals: number[]) {
    return {
        'c:tx': { 'c:strRef': { 'c:strCache': { 'c:pt': [{ attrs: { idx: '0' }, 'c:v': name }] } } },
        'c:cat': {
            'c:strRef': {
                'c:strCache': { 'c:pt': cats.map((c, i) => ({ attrs: { idx: String(i) }, 'c:v': c })) }
            }
        },
        'c:val': {
            'c:numRef': {
                'c:numCache': { 'c:pt': vals.map((v, i) => ({ attrs: { idx: String(i) }, 'c:v': String(v) })) }
            }
        }
    };
}

function chartSpace(plotArea: any) {
    return { 'c:chartSpace': { 'c:chart': { 'c:plotArea': plotArea } } };
}

describe('组合图 multi-plot 解析', () => {
    it('提取 plotArea 下全部图表节点到 plots（组合图不再丢数据）', () => {
        const xml = chartSpace({
            'c:barChart': [{ 'c:ser': [ser('Bar', ['Q1', 'Q2'], [10, 20])] }],
            'c:lineChart': [{ 'c:ser': [ser('Line', ['Q1', 'Q2'], [30, 40])] }],
            'c:pieChart': [{ 'c:ser': [ser('Pie', ['A', 'B'], [5, 6])] }],
            'c:valAx': [{ 'c:numFmt': { attrs: { formatCode: '0.00' } } }]
        });

        const r = extractChart(xml)!;
        expect(r).toBeDefined();
        expect(r!.plots).toBeDefined();
        // 三个绘图区全部保留（顺序按 CHART_PLOT_TYPES：bar → line → pie）
        expect(r!.plots!.map((p) => p.chartType)).toEqual(['barChart', 'lineChart', 'pieChart']);
        // 主 plot（第一个）属性映射到顶层，保持单图表兼容
        expect(r!.chartType).toBe('barChart');
        expect(r!.series!.length).toBe(1);
        expect(r!.series![0].name).toBe('Bar');
        expect(r!.series![0].values).toEqual([10, 20]);
        // 每个 plot 独立保留自己的系列与类别
        expect(r!.plots![0].series[0].values).toEqual([10, 20]);
        expect(r!.plots![0].categories).toEqual(['Q1', 'Q2']);
        expect(r!.plots![1].series[0].values).toEqual([30, 40]);
        expect(r!.plots![2].series[0].values).toEqual([5, 6]);
        // 数值轴数字格式合并到顶层
        expect(r!.numberFormat).toBe('0.00');
    });

    it('单图表场景向后兼容：plots 长度为 1 且顶层字段正常', () => {
        const xml = chartSpace({
            'c:barChart': [{ 'c:ser': [ser('Only', ['X', 'Y'], [1, 2])] }]
        });
        const r = extractChart(xml)!;
        expect(r!.plots!.length).toBe(1);
        expect(r!.chartType).toBe('barChart');
        expect(r!.series![0].values).toEqual([1, 2]);
        expect(r!.plots![0].chartType).toBe('barChart');
    });

    it('散点/气泡/股票 plot 按各自结构解析', () => {
        const scatter = {
            'c:ser': [{
                'c:tx': { 'c:strRef': { 'c:strCache': { 'c:pt': [{ attrs: { idx: '0' }, 'c:v': 'S' }] } } },
                'c:xVal': { 'c:numRef': { 'c:numCache': { 'c:pt': [{ attrs: { idx: '0' }, 'c:v': '1' }, { attrs: { idx: '1' }, 'c:v': '2' }] } } },
                'c:yVal': { 'c:numRef': { 'c:numCache': { 'c:pt': [{ attrs: { idx: '0' }, 'c:v': '10' }, { attrs: { idx: '1' }, 'c:v': '20' }] } } }
            }]
        };
        const line = { 'c:ser': [ser('L', ['Q1', 'Q2'], [3, 4])] };
        const xml = chartSpace({ 'c:scatterChart': [scatter], 'c:lineChart': [line] });
        const r = extractChart(xml)!;
        // 按 CHART_PLOT_TYPES 规范顺序提取：lineChart 排在 scatterChart 之前
        expect(r!.plots!.map((p) => p.chartType)).toEqual(['lineChart', 'scatterChart']);
        // 折线图的系列结构（values）
        expect(r!.plots![0].series[0].values).toEqual([3, 4]);
        // 散点图的系列结构（x/y 而非 values）
        expect(r!.plots![1].series[0].x).toEqual([1, 2]);
        expect(r!.plots![1].series[0].y).toEqual([10, 20]);
    });
});
