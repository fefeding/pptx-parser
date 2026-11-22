import { describe, it, expect } from 'vitest';
import { ChartRenderer } from '../examples/chart-lib/chart-renderer.js';

describe('chart-renderer 选项构建', () => {
  const r = new ChartRenderer();
  it('scatter 生成散点 series', () => {
    const info = {
      chartId: 'c1', type: 'scatterChart', title: '',
      data: [
        { key: 'S1', xlabels: [], values: [{ x: 1, y: 10 }, { x: 2, y: 20 }] },
        { key: 'S2', xlabels: [], values: [{ x: 1, y: 5 }, { x: 2, y: 15 }] }
      ],
      style: { legend: { position: 'right' } }
    };
    const opt = r.prepareEChartsOption(info as any);
    console.log('SCATTER OPT', JSON.stringify(opt?.series));
    expect(opt?.series?.[0]?.type).toBe('scatter');
    expect(opt?.xAxis?.type).toBe('value');
    expect(JSON.stringify(opt?.series?.[0]?.data)).toContain('[1,10]');
  });
  it('趋势线：linear 追加虚线 series 且线性拟合正确', () => {
    const info = {
      chartId: 'c3', type: 'lineChart', title: '',
      data: [{ key: 'S', xlabels: ['a', 'b', 'c', 'd'], values: [{ y: 2 }, { y: 4 }, { y: 6 }, { y: 8 }] }],
      style: { legend: { position: 'right' } },
      trendlines: [[{ type: 'linear' }]]
    };
    const opt = r.prepareEChartsOption(info as any);
    const tl = opt?.series?.[1];
    expect(tl?.type).toBe('line');
    expect(tl?.showSymbol).toBe(false);
    expect(tl?.lineStyle?.type).toBe('dashed');
    expect(tl?.data?.[0]).toBeCloseTo(2);
    expect(tl?.data?.[3]).toBeCloseTo(8);
  });
  it('趋势线：forward/backward 扩展类目轴并补齐原 series', () => {
    const info = {
      chartId: 'c4', type: 'lineChart', title: '',
      data: [{ key: 'S', xlabels: ['a', 'b'], values: [{ y: 1 }, { y: 3 }] }],
      style: { legend: { position: 'right' } },
      trendlines: [[{ type: 'linear', backward: 1, forward: 2 }]]
    };
    const opt = r.prepareEChartsOption(info as any);
    expect(opt?.xAxis?.data?.length).toBe(5);
    expect(opt?.series?.[0]?.data?.length).toBe(5);
    expect(opt?.series?.[1]?.data?.length).toBe(5);
    expect(opt?.series?.[0]?.data?.[0]).toBeNull();
    expect(opt?.series?.[1]?.data?.[0]).toBeCloseTo(-1);
    expect(opt?.series?.[1]?.data?.[4]).toBeCloseTo(7);
  });
  it('趋势线：poly / movingAvg 可拟合，非法数据不生成 series', () => {
    const info = {
      chartId: 'c5', type: 'lineChart', title: '',
      data: [{ key: 'S', xlabels: ['a', 'b', 'c'], values: [{ y: 1 }, { y: 4 }, { y: 9 }] }],
      style: { legend: { position: 'right' } },
      trendlines: [[{ type: 'poly', order: 2 }, { type: 'movingAvg', period: 2 }]]
    };
    const opt = r.prepareEChartsOption(info as any);
    expect(opt?.series).toHaveLength(3);
    expect(opt?.series?.[1]?.smooth).toBe(true);
    expect(opt?.series?.[2]?.data?.[0]).toBe(1);
    expect(opt?.series?.[2]?.data?.[1]).toBeCloseTo(2.5);

    const bad = r.prepareEChartsOption({ ...info, chartId: 'c6', data: [{ key: 'S', xlabels: [], values: [] }] } as any);
    expect(bad?.series).toHaveLength(1); // 只有原 series，点数不足不生成趋势线
    expect(bad?.series?.[1]).toBeUndefined();
  });
  it('stock 生成蜡烛 series', () => {
    const info = {
      chartId: 'c2', type: 'stockChart', title: '',
      data: [
        { key: 'K', xlabels: ['d1', 'd2'], values: [[1, 1.5, 0.5, 2], [2, 2.5, 1.5, 3]] }
      ],
      style: { legend: { position: 'right' } }
    };
    const opt = r.prepareEChartsOption(info as any);
    console.log('STOCK OPT', JSON.stringify(opt?.series));
    expect(opt?.series?.[0]?.type).toBe('candlestick');
    expect(JSON.stringify(opt?.series?.[0]?.data)).toContain('[1,1.5,0.5,2]');
  });
});
