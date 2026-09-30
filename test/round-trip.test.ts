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
            background: { type: 'gradient', direction: 'vertical', stops: [{ color: 'FF0000', position: 0 }, { color: '0000FF', position: 1 }] },
            transition: { type: 'fade', duration: 1000 },
            notes: '这是演讲者备注',
            elements: [
                { type: 'text', x: 100, y: 50, width: 400, height: 80, text: '标题', align: 'center', fontSize: 24, bold: true, color: 'FF0000' },
                { type: 'shape', shapeType: 'roundRect', x: 100, y: 200, width: 200, height: 100, fill: { color: '0000FF' }, line: { color: '333333', width: 2 } },
                { type: 'image', x: 400, y: 200, width: 120, height: 120, data: PNG_1x1 },
                {
                    type: 'chart', chartType: 'barChart', x: 600, y: 100, width: 500, height: 400,
                    title: '销量', legend: true, categories: ['Q1', 'Q2'], series: [{ name: 'A', values: [1, 2] }]
                },
                {
                    type: 'table', x: 100, y: 350, width: 400, height: 120,
                    colWidths: [200, 200],
                    rows: [
                        { height: 60, cells: [{ text: '姓名', fill: 'DDEEFF' }, { text: '分数', fill: 'DDEEFF' }] },
                        { height: 60, cells: [{ text: '张三' }, { text: '95' }] }
                    ]
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
        expect(types).toEqual(['chart', 'image', 'shape', 'table', 'text']);

        // 背景 round-trip（渐变）
        expect(back.slides[0].background).toMatchObject({ type: 'gradient', direction: 'vertical' });
        expect((back.slides[0].background as any).stops.map((s: any) => s.color)).toEqual(['FF0000', '0000FF']);

        // 过渡 round-trip
        expect(back.slides[0].transition).toEqual({ type: 'fade', duration: 1000 });

        // 备注 round-trip
        expect(back.slides[0].notes).toContain('这是演讲者备注');

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

        // 表格 round-trip
        const table = els.find((e) => e.type === 'table') as any;
        expect(table.rows.length).toBe(2);
        expect(table.rows[0].cells.map((c: any) => c.text)).toEqual(['姓名', '分数']);
        expect(table.rows[1].cells.map((c: any) => c.text)).toEqual(['张三', '95']);
        expect(table.rows[0].cells[0].fill).toBe('DDEEFF');
        expect(table.colWidths?.length).toBe(2);
        expect(table.rows[0].height).toBeCloseTo(60, 0);
    });
});
