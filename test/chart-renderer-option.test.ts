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
