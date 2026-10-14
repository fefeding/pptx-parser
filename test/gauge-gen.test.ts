/**
 * gauge 仪表盘生成端（type:'gauge'，OOXML 无原生类型，用 blockArc 形状组合模拟）：
 * - 结构级：两个 blockArc（背景环 + 值环）+ 中心数值文本
 * - 值环扫掠比例 = clamp((value-min)/(max-min), 0, 1)（OOXML 角度单位 1/60000 度）
 * - round-trip：pptxToStandard 读回不崩，拆为 2 个 blockArc 形状 + 文本
 */
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx, pptxToStandard } from '../src/index.ts';
import type { PptxDocument } from '../src/types/pptx-document';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

async function genSlideXml(el: Record<string, unknown>): Promise<string> {
    const doc: PptxDocument = {
        version: '1.0',
        slideSize: { width: 1280, height: 720 },
        slides: [{ elements: [el as any] }]
    };
    const pptx = await jsonToPptx(doc, { outputType: 'uint8array' });
    const zip = await JSZip.loadAsync(toArrayBuffer(pptx as Uint8Array));
    const slide = zip.file('ppt/slides/slide1.xml');
    expect(slide, 'slide1.xml 应存在').toBeTruthy();
    return await slide!.async('string');
}

describe('gauge 仪表盘生成端', () => {
    it('写出两个 blockArc（背景+值）与中心数值文本', async () => {
        const xml = await genSlideXml({
            type: 'gauge', x: 100, y: 100, width: 300, height: 300,
            value: 75, max: 100, unit: '%', label: '完成率', color: '#2E7D32'
        });
        // 两个 blockArc 形状（背景环 + 值环）
        expect(xml.match(/prst="blockArc"/g)?.length).toBe(2);
        // 值环扫掠 0.75 → 0.75*359*60000 = 16155000
        expect(xml).toContain('<a:gd name="adj2" fmla="val 16155000"/>');
        // 背景环扫掠 359° → 21540000
        expect(xml).toContain('<a:gd name="adj2" fmla="val 21540000"/>');
        // 背景环与值环均有 adj1=0（起点）/ adj3（厚度）
        expect(xml).toContain('<a:gd name="adj1" fmla="val 0"/>');
        // 中心数值文本含 "75%"
        expect(xml).toContain('75%');
        // 标签
        expect(xml).toContain('完成率');
    });

    it('value 超过 max 时 clamp 到满环（sweep=21540000）', async () => {
        const xml = await genSlideXml({
            type: 'gauge', x: 0, y: 0, width: 200, height: 200, value: 150, max: 100
        });
        expect(xml).toContain('<a:gd name="adj2" fmla="val 21540000"/>');
        expect(xml).toContain('150');
    });

    it('自定义 min/max 与跨零范围', async () => {
        // -50..50, value=25 → ratio 0.75 → sweep 16155000
        const xml = await genSlideXml({
            type: 'gauge', x: 0, y: 0, width: 200, height: 200,
            value: 25, min: -50, max: 50
        });
        expect(xml).toContain('<a:gd name="adj2" fmla="val 16155000"/>');
    });

    it('showValue:false 时不写中心文本', async () => {
        const xml = await genSlideXml({
            type: 'gauge', x: 0, y: 0, width: 200, height: 200,
            value: 40, max: 100, showValue: false
        });
        expect(xml.match(/prst="blockArc"/g)?.length).toBe(2);
        // 无中心数值文本块（blockArc 形状无 txBody）
        expect(xml).not.toContain('>40<');
    });

    it('round-trip：pptxToStandard 读回含 2 个 blockArc 形状与中心文本', async () => {
        const doc: PptxDocument = {
            version: '1.0',
            slideSize: { width: 1280, height: 720 },
            slides: [{ elements: [{ type: 'gauge', x: 100, y: 100, width: 300, height: 300, value: 40, max: 100, unit: '%' } as any] }]
        };
        const pptx = await jsonToPptx(doc, { outputType: 'uint8array' });
        const std = await pptxToStandard(toArrayBuffer(pptx as Uint8Array)) as any;
        const els = std.slides[0].elements;
        const blocks = els.filter((e: any) => e.shapeType === 'blockArc');
        const texts = els.filter((e: any) => e.type === 'text');
        expect(blocks.length).toBe(2);
        expect(texts.length).toBeGreaterThanOrEqual(1);
        expect(texts.some((t: any) => JSON.stringify(t).includes('40%'))).toBe(true);
    });
});
