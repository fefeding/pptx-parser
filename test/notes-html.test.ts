import { describe, it, expect } from 'vitest';
import pptxToHtml, { jsonToPptx, pptxToStandard } from '../src/index';

function presWithNotes() {
    return {
        slides: [
            {
                background: '#ffffff',
                elements: [
                    { type: 'text', x: 60, y: 30, width: 1160, height: 50, text: '备注测试', fontSize: 30, bold: true }
                ],
                notes: '第一行备注\n第二行备注'
            }
        ]
    };
}

describe('演讲者备注：解析 + HTML 渲染', () => {
    it('语义层仍可读回 notes', async () => {
        const buf: any = await jsonToPptx(presWithNotes());
        const std: any = await pptxToStandard(buf, { mediaProcess: false } as any);
        expect(std.slides[0].notes).toContain('第一行备注');
        expect(std.slides[0].notes).toContain('第二行备注');
    });

    it('pptxToHtml 输出 pptx-notes 容器与 data-notes，且默认隐藏', async () => {
        const buf: any = await jsonToPptx(presWithNotes());
        const res: any = await pptxToHtml(buf, { mediaProcess: false } as any);
        const html = res.slides[0].html as string;
        expect(html).toContain('pptx-notes');
        expect(html).toContain('data-notes=');
        expect(html).toContain('第一行备注');
        expect(html).toContain('第二行备注');
        expect(html).toContain('display:none');
    });

    it('无备注时不输出 pptx-notes 容器', async () => {
        const buf: any = await jsonToPptx({ slides: [{ elements: [] }] });
        const res: any = await pptxToHtml(buf, { mediaProcess: false } as any);
        expect(res.slides[0].html).not.toContain('pptx-notes');
    });
});
