import { describe, it, expect } from 'vitest';
import pptxToHtml, { jsonToPptx } from '../src/index';
import { PPTXStyleUtils } from '../src/utils/style';

function presWithRadial() {
    return {
        slides: [{
            background: '#ffffff',
            elements: [{
                type: 'shape', shapeType: 'ellipse', x: 100, y: 100, width: 200, height: 200,
                fill: {
                    type: 'gradient', gradientType: 'radial', gradientPath: 'circle',
                    stops: [
                        { position: 0, color: '#ffffff' },
                        { position: 1, color: '#ff0000' }
                    ]
                }
            }]
        }]
    };
}

describe('路径（径向）渐变', () => {
    it('生成端写出 a:path + a:fillToRect（无 a:lin）', async () => {
        const buf: any = await jsonToPptx(presWithRadial());
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(buf);
        const xml = await zip.file('ppt/slides/slide1.xml').async('text');
        expect(xml).toContain('<a:gradFill');
        expect(xml).toContain('<a:path');
        expect(xml).toContain('fillToRect');
        expect(xml).not.toContain('<a:lin');
    });

    it('getGradientFill 标记 radial=true 并解析 fillToRect', async () => {
        const buf: any = await jsonToPptx(presWithRadial());
        const res: any = await pptxToHtml(buf, { mediaProcess: false } as any);
        const svg = res.slides[0].html as string;
        // 径向渐变在 SVG 中应为 radialGradient，且仍用 linGrd_ id 供形状 fill 引用
        expect(svg).toContain('<radialGradient');
        expect(svg).toContain('id="linGrd_');
        expect(svg).toContain('gradientTransform=');
        expect(svg).toContain('url(#linGrd_');
    });
});
