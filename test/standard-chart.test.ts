import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToStandard } from '../src/index.ts';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
  return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

describe('pptxToStandard 图表', () => {
  it('scatter/stock 在 standard 文档里保留 x/y/open', async () => {
    const doc = {
      slides: [{
        elements: [
          { type: 'chart' as const, x: 0, y: 0, width: 500, height: 300, chartType: 'scatterChart' as const,
            series: [{ name: 'S1', x: [1, 2, 3], y: [10, 20, 30] }] },
          { type: 'chart' as const, x: 0, y: 0, width: 500, height: 300, chartType: 'stockChart' as const,
            categories: ['d1', 'd2', 'd3'],
            series: [{ name: 'K', open: [1, 2, 3], high: [2, 3, 4], low: [0.5, 1.5, 2.5], close: [1.5, 2.5, 3.5] }] }
        ]
      }]
    };
    const data = await jsonToPptx(doc);
    const res = await pptxToStandard(toArrayBuffer(data)) as any;
    const els = res.slides[0].elements;
    console.log('STD SCATTER', JSON.stringify(els[0].series));
    console.log('STD STOCK', JSON.stringify(els[1].series));
    expect(els[0].series[0].x).toEqual([1, 2, 3]);
    expect(els[1].series[0].open).toEqual([1, 2, 3]);
  });
});
