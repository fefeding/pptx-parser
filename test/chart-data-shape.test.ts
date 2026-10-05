import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToJson } from '../src/index.ts';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
  return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

describe('图表数据形态', () => {
  it('scatter/bubble/stock 解析保留 x/y/size/OHLC', async () => {
    const doc = {
      slides: [{
        elements: [
          {
            type: 'chart' as const, x: 0, y: 0, width: 500, height: 300,
            chartType: 'scatterChart' as const,
            series: [
              { name: 'S1', x: [1, 2, 3], y: [10, 20, 30] },
              { name: 'S2', x: [1, 2, 3], y: [5, 15, 25] }
            ]
          },
          {
            type: 'chart' as const, x: 0, y: 0, width: 500, height: 300,
            chartType: 'bubbleChart' as const,
            series: [{ name: 'B1', x: [1, 2, 3], y: [10, 20, 30], values: [5, 10, 15] }]
          },
          {
            type: 'chart' as const, x: 0, y: 0, width: 500, height: 300,
            chartType: 'stockChart' as const,
            categories: ['d1', 'd2', 'd3'],
            series: [{
              name: 'K', open: [1, 2, 3], high: [2, 3, 4],
              low: [0.5, 1.5, 2.5], close: [1.5, 2.5, 3.5]
            }]
          }
        ]
      }]
    };
    const data = await jsonToPptx(doc);
    const res = await pptxToJson(toArrayBuffer(data), { mode: 'semantic' });
    const els = res.document.slides[0].elements as any[];
    console.log('SCATTER', JSON.stringify(els[0].series));
    console.log('BUBBLE', JSON.stringify(els[1].series));
    console.log('STOCK', JSON.stringify(els[2].series));
    expect(els[0].series[0]).toHaveProperty('x');
    expect(els[0].series[0]).toHaveProperty('y');
    expect(els[1].series[0]).toHaveProperty('values');
    expect(els[2].series[0]).toHaveProperty('open');
  });
});
