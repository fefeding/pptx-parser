// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToJson } from '../src/index.ts';
import { docFromPptx } from '../examples/editor/src/model.js';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
  return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

describe('编辑器导入散点/K线', () => {
  it('docFromPptx 后 chart 系列保留散点结构', async () => {
    const doc = {
      slides: [{
        elements: [
          {
            type: 'chart' as const, x: 0, y: 0, width: 500, height: 300,
            chartType: 'scatterChart' as const,
            series: [{ name: 'S1', x: [1, 2, 3], y: [10, 20, 30] }]
          },
          {
            type: 'chart' as const, x: 0, y: 0, width: 500, height: 300,
            chartType: 'stockChart' as const,
            categories: ['d1', 'd2', 'd3'],
            series: [{ name: 'K', open: [1, 2, 3], high: [2, 3, 4], low: [0.5, 1.5, 2.5], close: [1.5, 2.5, 3.5] }]
          }
        ]
      }]
    };
    const data = await jsonToPptx(doc);
    const res = await pptxToJson(toArrayBuffer(data), { mode: 'semantic' });
    const edoc = docFromPptx(res.document);
    const els = edoc.slides[0].elements as any[];
    console.log('SCATTER EL', JSON.stringify(els[0].series));
    console.log('STOCK EL', JSON.stringify(els[1].series));
    expect(els[0].series[0].values[0]).toMatchObject({ x: 1, y: 10 });
    expect(Array.isArray(els[1].series[0].values[0])).toBe(true);
  });
});
