import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx } from '../src/serializer/json-to-pptx';
import { pptxToStandard } from '../src/index';
import type { PptxDocument } from '../src/types/pptx-document';

const PNG_1x1 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC';

const doc: PptxDocument = {
    version: '1.0',
    slideSize: { width: 1280, height: 720 },
    metadata: { title: 'RT', author: 'tester' },
    slides: [
        {
            background: 'FFFFFF',
            elements: [
                { type: 'text', x: 100, y: 50, width: 400, height: 80, text: '标题', align: 'center', fontSize: 24, bold: true, color: 'FF0000' },
                { type: 'shape', shapeType: 'roundRect', x: 100, y: 200, width: 200, height: 100, fill: { color: '0000FF' }, line: { color: '333333', width: 2 } },
                { type: 'image', x: 400, y: 200, width: 120, height: 120, data: PNG_1x1 },
                {
                    type: 'chart', chartType: 'barChart', x: 600, y: 100, width: 500, height: 400,
                    title: '销量', legend: true, categories: ['Q1', 'Q2'], series: [{ name: 'A', values: [1, 2] }]
                }
            ]
        }
    ]
};

describe('pptx round-trip (jsonToPptx -> pptxToStandard)', () => {
    it('还原的元素数量与类型一致', async () => {
        const pptx = await jsonToPptx(doc, { outputType: 'uint8array' });
        expect(pptx).toBeInstanceOf(Uint8Array);

        const back = await pptxToStandard(pptx as Uint8Array);
        expect(back.version).toBe('1.0');
        expect(back.metadata?.title).toBe('RT');
        expect(back.slides.length).toBe(1);

        const els = back.slides[0].elements;
        const types = els.map((e) => e.type).sort();
        expect(types).toEqual(['chart', 'image', 'shape', 'text']);

        const text = els.find((e) => e.type === 'text') as any;
        expect(text.text ?? text.paragraphs?.[0]?.runs?.[0]?.text).toMatch(/标题/);
        expect(text.x).toBeCloseTo(100, 0);
        expect(text.fontSize).toBe(24);

        const shape = els.find((e) => e.type === 'shape') as any;
        expect(shape.shapeType).toBe('roundRect');
        expect(shape.fill).toBe('0000FF');

        const img = els.find((e) => e.type === 'image') as any;
        expect(img.data).toMatch(/^data:image\/png;base64,/);

        const chart = els.find((e) => e.type === 'chart') as any;
        expect(chart.chartType).toBe('barChart');
        expect(chart.categories).toEqual(['Q1', 'Q2']);
        expect(chart.series[0].values).toEqual([1, 2]);
    });
});
