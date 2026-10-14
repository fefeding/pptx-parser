/**
 * 组合图 multi-plot 生成端（el.plots 长度 >= 2）：
 * - 结构级：多个 c:*Chart 节点、共享轴 id、series idx/order 跨 plot 全局连续、
 *   数据引用列按 plot 顺序分配（不串列）
 * - round-trip：与解析端 plots 提取（chart-multi-plot.test.ts）形成闭环
 * - 约束 fail-fast：scatter/bubble/stock/surface 混入、与 secondaryValueAxis 并用
 * - 向后兼容：无 plots / plots 长度为 1 时行为不变
 */
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx, pptxToJson } from '../src/index.ts';
import type { PptxDocument } from '../src/types/pptx-document';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

/** 生成单图表幻灯片文档，返回 chart1.xml 文本 */
async function genChartXml(el: Record<string, unknown>): Promise<string> {
    const doc: PptxDocument = {
        version: '1.0',
        slideSize: { width: 1280, height: 720 },
        slides: [{ elements: [{ type: 'chart', x: 0, y: 0, width: 600, height: 400, ...el } as any] }]
    };
    const pptx = await jsonToPptx(doc, { outputType: 'uint8array' });
    const zip = await JSZip.loadAsync(toArrayBuffer(pptx as Uint8Array));
    const chartFile = zip.file('ppt/charts/chart1.xml');
    expect(chartFile, 'chart1.xml 应存在').toBeTruthy();
    return await chartFile!.async('string');
}

/** 提取全部 c:ser 的 idx（验证跨 plot 连续性） */
function serIdxs(xml: string): string[] {
    return [...xml.matchAll(/<c:ser><c:idx val="(\d+)"\/>/g)].map((m) => m[1]);
}

describe('组合图 multi-plot 生成端', () => {
    it('barChart + lineChart：两个图表节点共享轴 id，series idx 跨 plot 连续', async () => {
        const xml = await genChartXml({
            chartType: 'barChart',
            categories: ['Q1', 'Q2', 'Q3', 'Q4'],
            plots: [
                { chartType: 'barChart', series: [{ name: '营收', values: [10, 20, 30, 40] }] },
                { chartType: 'lineChart', series: [{ name: '利润', values: [5, 8, 12, 18] }] }
            ]
        });
        // 两个图表节点各出现一次
        expect(xml.match(/<c:barChart>/g)?.length).toBe(1);
        expect(xml.match(/<c:lineChart>/g)?.length).toBe(1);
        // 共享轴 id：两个图表节点各引用一次 + catAx/valAx 定义各一次 = 3
        // （ECMA-376 允许同 plotArea 多图表节点引用相同 axId，即组合图同轴）
        expect(xml.match(/<c:axId val="111"\/>/g)?.length).toBe(3);
        expect(xml.match(/<c:axId val="112"\/>/g)?.length).toBe(3);
        // series idx/order 跨 plot 全局连续（0=营收，1=利润）
        expect(serIdxs(xml)).toEqual(['0', '1']);
        // 数据引用列按 plot 顺序分配：bar 系列 → B 列，line 系列 → C 列（不串列）
        expect(xml).toContain('<c:f>Sheet1!$B$2:$B$5</c:f>');
        expect(xml).toContain('<c:f>Sheet1!$C$2:$C$5</c:f>');
        // 内联数据缓存保真（渲染端以 cache 为准）
        expect(xml).toContain('<c:v>40</c:v>');
        expect(xml).toContain('<c:v>18</c:v>');
    });

    it('三 plot 组合：列分配连续（B/C/D），idx 0..2', async () => {
        const xml = await genChartXml({
            chartType: 'barChart',
            categories: ['A', 'B'],
            plots: [
                { chartType: 'barChart', series: [{ name: 'S1', values: [1, 2] }] },
                { chartType: 'lineChart', series: [{ name: 'S2', values: [3, 4] }] },
                { chartType: 'areaChart', series: [{ name: 'S3', values: [5, 6] }] }
            ]
        });
        expect(xml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>');
        expect(xml).toContain('<c:f>Sheet1!$C$2:$C$3</c:f>');
        expect(xml).toContain('<c:f>Sheet1!$D$2:$D$3</c:f>');
        expect(serIdxs(xml)).toEqual(['0', '1', '2']);
    });

    it('pieChart + doughnutChart：饼类组合无坐标轴，holeSize 按 plot 生效，varyColors 各自默认开', async () => {
        const xml = await genChartXml({
            chartType: 'pieChart',
            plots: [
                { chartType: 'pieChart', series: [{ name: '渠道', values: [60, 40] }] },
                { chartType: 'doughnutChart', holeSize: 65, series: [{ name: '成本', values: [30, 70] }] }
            ]
        });
        expect(xml).toContain('<c:pieChart>');
        expect(xml).toContain('<c:doughnutChart>');
        expect(xml).toContain('<c:holeSize val="65"/>');
        // 全饼类组合不生成坐标轴
        expect(xml).not.toContain('<c:axId');
        expect(xml).not.toContain('<c:catAx');
        // varyColors 默认开（两个节点各一个）
        expect(xml.match(/<c:varyColors val="1"\/>/g)?.length).toBe(2);
    });

    it('per-plot 配置：bar 默认 clustered / line 默认 standard，dLbls 按 plot 独立生效', async () => {
        const xml = await genChartXml({
            chartType: 'barChart',
            categories: ['A', 'B'],
            plots: [
                { chartType: 'barChart', series: [{ name: 'S1', values: [1, 2] }] },
                { chartType: 'lineChart', dataLabels: true, series: [{ name: 'S2', values: [3, 4] }] }
            ]
        });
        expect(xml).toContain('<c:grouping val="clustered"/>');
        expect(xml).toContain('<c:grouping val="standard"/>');
        // dLbls 仅出现在 line 节点内（CT_Chart 本身不含 dLbls，必须位于图表类型节点内）
        const barPart = xml.slice(xml.indexOf('<c:barChart>'), xml.indexOf('</c:barChart>'));
        const linePart = xml.slice(xml.indexOf('<c:lineChart>'), xml.indexOf('</c:lineChart>'));
        expect(linePart).toContain('<c:dLbls>');
        expect(linePart).toContain('<c:showVal val="1"/>');
        expect(barPart).not.toContain('<c:dLbls>');
    });

    it('round-trip：生成 → 解析回读 plots（类型/系列/类目保真）', async () => {
        const doc: PptxDocument = {
            version: '1.0',
            slideSize: { width: 1280, height: 720 },
            slides: [{
                elements: [{
                    type: 'chart', x: 0, y: 0, width: 600, height: 400,
                    chartType: 'barChart',
                    categories: ['Q1', 'Q2'],
                    plots: [
                        { chartType: 'barChart', series: [{ name: '营收', values: [10, 20] }] },
                        { chartType: 'lineChart', series: [{ name: '利润', values: [5, 8] }] }
                    ]
                } as any]
            }]
        };
        const pptx = await jsonToPptx(doc, { outputType: 'uint8array' });
        const back = await pptxToJson(toArrayBuffer(pptx as Uint8Array), { mode: 'semantic' }) as any;
        const chart = back.document.slides[0].elements.find((e: any) => e.type === 'chart');
        expect(chart.plots).toBeDefined();
        expect(chart.plots.map((p: any) => p.chartType)).toEqual(['barChart', 'lineChart']);
        expect(chart.plots[0].series[0].values).toEqual([10, 20]);
        expect(chart.plots[1].series[0].values).toEqual([5, 8]);
        // 主 plot 映射顶层，保持单图表场景兼容
        expect(chart.chartType).toBe('barChart');
    });

    it('fail-fast：组合图混入 scatter 抛错（独立坐标系不支持共享轴）', async () => {
        await expect(genChartXml({
            chartType: 'barChart',
            plots: [
                { chartType: 'barChart', series: [{ name: 'S1', values: [1, 2] }] },
                { chartType: 'scatterChart', series: [{ name: 'S2', x: [1, 2], y: [3, 4] }] }
            ]
        })).rejects.toThrow(/组合图不支持混入 scatterChart/);
    });

    it('fail-fast：组合图与 secondaryValueAxis 并用抛错', async () => {
        await expect(genChartXml({
            chartType: 'barChart',
            secondaryValueAxis: true,
            plots: [
                { chartType: 'barChart', series: [{ name: 'S1', values: [1, 2] }] },
                { chartType: 'lineChart', series: [{ name: 'S2', values: [3, 4] }] }
            ]
        })).rejects.toThrow(/组合图（plots）不支持 secondaryValueAxis/);
    });

    it('向后兼容：plots 长度 1 时退化为顶层单图表路径', async () => {
        const xml = await genChartXml({
            chartType: 'barChart',
            categories: ['A', 'B'],
            series: [{ name: 'S', values: [1, 2] }],
            plots: [{ chartType: 'pieChart', series: [{ name: 'P', values: [9, 9] }] }]
        });
        // 走顶层字段：单 barChart 节点 + 轴，plots 被忽略
        expect(xml.match(/<c:barChart>/g)?.length).toBe(1);
        expect(xml).not.toContain('<c:pieChart>');
        expect(xml).toContain('<c:axId val="111"/>');
        expect(xml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>');
    });

    it('向后兼容：无 plots 时单图表输出不变', async () => {
        const xml = await genChartXml({
            chartType: 'barChart',
            categories: ['A', 'B'],
            series: [{ name: 'S', values: [1, 2] }]
        });
        expect(xml.match(/<c:barChart>/g)?.length).toBe(1);
        expect(xml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>');
        expect(serIdxs(xml)).toEqual(['0']);
    });
});
