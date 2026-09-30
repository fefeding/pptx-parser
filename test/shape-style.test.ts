import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx, pptxToStandard } from '../src/index.ts';

/** 1pt = 12700 EMU */
const PT = 12700;

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

function shapeEl(overrides: any) {
    return { type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 100, height: 100, ...overrides };
}
function shapeDoc(el: any) {
    return { slides: [{ elements: [el] }] };
}

describe('T2 形状渐变填充 / 透明度 / 阴影 / 发光（生成端）', () => {
    it('渐变填充生成 a:gradFill + a:gsLst + a:lin（vertical → ang=5400000）', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({
            fill: { type: 'gradient', direction: 'vertical', stops: [{ color: '#FF0000', position: 0 }, { color: '#00FF00', position: 1 }] }
        })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:gradFill');
        expect(xml).toContain('<a:gsLst');
        expect(xml).toContain('pos="0"');
        expect(xml).toContain('pos="100000"');
        expect(xml).toContain('val="FF0000"');
        expect(xml).toContain('val="00FF00"');
        expect(xml).toContain('ang="5400000"');
    });

    it('纯色透明度生成 a:solidFill/a:srgbClr/a:alpha（transparency=50 → val=50000）', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({ fill: { type: 'solid', color: '#FF0000', transparency: 50 } })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:solidFill');
        expect(xml).toContain('val="FF0000"');
        expect(xml).toContain('<a:alpha val="50000"');
    });

    it('边框透明度 + 虚线生成 a:alpha 与 a:prstDash', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({ line: { color: '#000000', width: 2, transparency: 30, dashType: 'dash' } })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('w="25400"');        // 2pt
        expect(xml).toContain('<a:prstDash val="dash"');
        expect(xml).toContain('<a:alpha val="70000"'); // transparency 30 → 70000
    });

    it('外阴影生成 a:effectLst/a:outerShdw + blur/dist/dir + alpha', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({
            effects: { shadow: { type: 'outer', color: '#000000', blur: 4, distance: 3, angle: 90 } }
        })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:effectLst');
        expect(xml).toContain('<a:outerShdw');
        expect(xml).toContain('blur="50800"');   // 4pt
        expect(xml).toContain('dist="38100"');   // 3pt
        expect(xml).toContain('dir="5400000"');  // 90°
    });

    it('发光生成 a:glow + blur + 颜色', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({ effects: { glow: { color: '#FFFF00', blur: 5 } } })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:glow');
        expect(xml).toContain('blur="63500"');   // 5pt
        expect(xml).toContain('val="FFFF00"');
    });

    it('默认阴影 effects.shadow=true 生成 a:outerShdw', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({ effects: { shadow: true } })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('<a:outerShdw');
    });

    it('兼容：字符串 fill 仍生成 solidFill（无回归）', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({ fill: '#4f46e5' })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('<a:solidFill');
        expect(xml).toContain('val="4F46E5"');
        expect(xml).not.toContain('<a:gradFill');
    });
});

describe('T2 形状样式 round-trip 回读', () => {
    it('pptxToStandard 回读渐变与阴影', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({
            fill: { type: 'gradient', direction: 'vertical', stops: [{ color: '#FF0000', position: 0 }, { color: '#00FF00', position: 1 }] },
            effects: { shadow: { type: 'outer', color: '#000000', blur: 4, distance: 3, angle: 90, transparency: 40 } }
        })));
        const std: any = await pptxToStandard(toArrayBuffer(data));
        const shape = std.slides[0].elements.find((e: any) => e.type === 'shape');

        expect(shape.fill.type).toBe('gradient');
        expect(shape.fill.direction).toBe('vertical');
        expect(shape.fill.stops[0]).toEqual({ color: 'FF0000', position: 0 });
        expect(shape.fill.stops[1]).toEqual({ color: '00FF00', position: 1 });
        expect(shape.effects.shadow.type).toBe('outer');
        expect(shape.effects.shadow.transparency).toBe(40);
    });

    it('pptxToStandard 回读透明度与虚线', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({
            fill: { type: 'solid', color: '#FF0000', transparency: 50 },
            line: { color: '#000000', width: 2, transparency: 30, dashType: 'dash' }
        })));
        const std: any = await pptxToStandard(toArrayBuffer(data));
        const shape = std.slides[0].elements.find((e: any) => e.type === 'shape');

        expect(shape.fill).toEqual({ type: 'solid', color: 'FF0000', transparency: 50 });
        expect(shape.line).toMatchObject({ color: '000000', width: 2, transparency: 30, dashType: 'dash' });
    });
});
