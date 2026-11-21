import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToStandard } from '../src/index';

const presWithGroupFill = () => ({
    slides: [{
        elements: [{
            type: 'group', x: 50, y: 50, width: 300, height: 200,
            fill: { type: 'solid', color: '#4f46e5' },
            children: [
                { type: 'shape', shapeType: 'rect', x: 10, y: 10, width: 100, height: 60, fill: { type: 'group' } },
                { type: 'shape', shapeType: 'rect', x: 140, y: 10, width: 100, height: 60, fill: '#ef4444' }
            ]
        }]
    }]
});

describe('组合填充 a:grpFill 语义往返', () => {
    it('生成端：grpSpPr 写组合填充，子形状写 a:grpFill', async () => {
        const buf: any = await jsonToPptx(presWithGroupFill());
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(buf);
        const xml = await zip.file('ppt/slides/slide1.xml').async('text');
        expect(xml).toContain('<p:grpSp');
        expect(xml).toContain('<p:grpSpPr>');
        // 组合级填充写成 srgbClr 4f46e5
        expect(xml.toLowerCase()).toContain('val="4f46e5"');
        // 子元素跟随组填充
        expect(xml).toContain('<a:grpFill');
    });

    it('解析端：group.fill 与子元素 {type:group} 均可回读', async () => {
        const buf: any = await jsonToPptx(presWithGroupFill());
        const std: any = await pptxToStandard(buf, { mediaProcess: false } as any);
        const group = std.slides[0].elements.find((e: any) => e.type === 'group');
        expect(group).toBeDefined();
        const gf = group.fill;
        expect(gf).toBeTruthy();
        const color = typeof gf === 'string' ? gf : gf.color;
        expect(String(color).toLowerCase()).toContain('4f46e5');
        expect(group.children[0].fill).toEqual({ type: 'group' });
    });
});
