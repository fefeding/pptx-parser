import { describe, it, expect } from 'vitest';
import { chartRenderer } from '../examples/chart-lib/chart-renderer.js';

describe('chart-renderer 散点/K线 option 构建', () => {
  it('普通散点（size 为 undefined）不应产生 symbolSize，避免 canvas 点不可见', () => {
    const option = chartRenderer.prepareEChartsOption({
      chartId: 'c1', type: 'scatterChart',
      data: [{ key: 'A', values: [{ x: 1, y: 2, size: undefined }, { x: 2, y: 4, size: undefined }] }],
      style: {}, title: '', theme: {}
    });
    const series = option?.series?.[0];
    expect(series?.type).toBe('scatter');
    expect(series?.data?.[0]).toEqual([1, 2]);
    expect(series?.data?.[1]).toEqual([2, 4]);
  });

  it('股票图无类目时仍使用 category X 轴，蜡烛数据保持 [open,close,low,high]', () => {
    const option = chartRenderer.prepareEChartsOption({
      chartId: 'c2', type: 'stockChart',
      data: [{ key: 'OHLC', values: [[10, 12, 8, 15], [12, 11, 10, 16]] }],
      style: {}, title: '', theme: {}
    });
    expect(option?.series?.[0]?.type).toBe('candlestick');
    expect(option?.xAxis?.type).toBe('category');
    expect(option?.series?.[0]?.data).toEqual([[10, 12, 8, 15], [12, 11, 10, 16]]);
  });
});
