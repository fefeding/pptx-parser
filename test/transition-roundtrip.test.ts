import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToStandard } from '../src/index';

describe('切换时序（T14）生成 + 解析往返', () => {
    it('advTm / advClick / dir 生成规范属性且可回读', async () => {
        const buf: any = await jsonToPptx({
            slides: [{
                background: '#ffffff',
                elements: [],
                transition: { type: 'wipe', duration: 1000, advanceAfterTime: 5000, advanceOnClick: false, direction: 'l' }
            }]
        });
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(buf);
        const xml = await zip.file('ppt/slides/slide1.xml').async('text');
        expect(xml).toContain('advTm="5000"');
        expect(xml).toContain('advClick="0"');
        expect(xml).not.toContain('advanceOnClick=');

        const std: any = await pptxToStandard(buf, { mediaProcess: false } as any);
        const tr = std.slides[0].transition;
        expect(tr).toBeTruthy();
        expect(tr.type).toBe('wipe');
        expect(tr.advanceAfterTime).toBe(5000);
        expect(tr.advanceOnClick).toBe(false);
        expect(tr.direction).toBe('l');
    });
});
