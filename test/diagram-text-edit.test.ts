// SmartArt（图示）文字编辑：改写后应同步落到数据部件与缓存绘图部件，并可被重新解析
import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { pptxToStandard, jsonToPptx, updateDiagramTexts, docFromPptx, docToPptx } from '../src/index.ts';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SAMPLE = path.resolve(__dirname, '../examples/Sample_12.pptx');

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteLength);
}

describe('SmartArt 文字编辑', () => {
    it('改写文字后数据部件与绘图部件同步更新，且能重新解析出来', async () => {
        const buf = fs.readFileSync(SAMPLE);
        const std = await pptxToStandard(toArrayBuffer(buf));
        const slide = std.slides[10];
        const diagram = slide.elements.find((e: any) => e.type === 'diagram');
        expect(diagram).toBeTruthy();
        const before = [...(diagram as any).texts];
        expect(before.length).toBeGreaterThan(1);

        const next = before.map((_t: string, i: number) => `节点${i + 1}`);
        const res = updateDiagramTexts(diagram, next);
        expect(res.changed).toBe(true);
        expect(res.texts).toEqual(next);
        expect((diagram as any).texts).toEqual(next);

        // 形状缓存同步（画布渲染读的是 el.shapes[].text）
        const shapeTexts = (diagram as any).shapes.filter((s: any) => s.text).map((s: any) => s.text);
        expect(shapeTexts.some((t: string) => t.includes('节点1'))).toBe(true);

        // 导出后再解析，文字应保持修改后的内容
        const out = await jsonToPptx(std as any);
        const back = await pptxToStandard(toArrayBuffer(out));
        const backDiagram = back.slides[10].elements.find((e: any) => e.type === 'diagram');
        expect((backDiagram as any).texts).toEqual(next);
    }, 60000);

    it('经编辑器模型（docFromPptx/docToPptx）后仍能改写并落盘', async () => {
        const buf = fs.readFileSync(SAMPLE);
        const std = await pptxToStandard(toArrayBuffer(buf));
        const edoc = docFromPptx(std);
        const diagram = (edoc as any).slides[10].elements.find((e: any) => e.type === 'diagram');
        expect(diagram).toBeTruthy();

        const next = (diagram.texts as string[]).map((_t, i) => `改${i}`);
        updateDiagramTexts(diagram, next);

        const out = await jsonToPptx(docToPptx(edoc));
        const back = await pptxToStandard(toArrayBuffer(out));
        const backDiagram = back.slides[10].elements.find((e: any) => e.type === 'diagram');
        expect((backDiagram as any).texts).toEqual(next);
    }, 60000);

    it('非图示元素调用时安全返回', () => {
        const res = updateDiagramTexts({ type: 'text', texts: ['a'] }, ['b']);
        expect(res.changed).toBe(false);
        expect(res.texts).toEqual([]);
    });
});
