import { describe, it, expect } from 'vitest';
import { jsonToPptx } from '../src/index';
import { convert } from '../src/convert/index';

function presWithNotes() {
    return {
        slides: [{
            background: '#ffffff',
            elements: [{ type: 'text', x: 60, y: 30, width: 600, height: 50, text: '备注导出测试' }],
            notes: '转换后的第一行备注\n转换后的第二行备注'
        }]
    };
}

describe('convert(html) 保留演讲者备注', () => {
    it('pptxToSlideHtml 输出隐藏的 pptx-notes 容器与 data-notes', async () => {
        const buf: any = await jsonToPptx(presWithNotes());
        const html = await convert(buf, 'html') as string;
        expect(html).toContain('pptx-notes');
        expect(html).toContain('data-notes=');
        expect(html).toContain('转换后的第一行备注');
        expect(html).toContain('转换后的第二行备注');
        expect(html).toContain('display:none');
    });

    it('无备注时不输出 pptx-notes 容器', async () => {
        const buf: any = await jsonToPptx({ slides: [{ elements: [] }] });
        const html = await convert(buf, 'html') as string;
        expect(html).not.toContain('pptx-notes');
    });
});
