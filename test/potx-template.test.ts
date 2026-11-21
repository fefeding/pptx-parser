import { describe, it, expect } from 'vitest';
import { jsonToPptx } from '../src/index';

const pres = () => ({
    slides: [{ elements: [{ type: 'text', x: 60, y: 30, width: 600, height: 50, text: '模板标题' }] }]
});

async function readContentTypes(buf: any): Promise<string> {
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buf);
    return zip.file('[Content_Types].xml').async('text');
}

describe('.potx 模板输出', () => {
    it('asTemplate:false 时 presentation.xml 为 presentation.main+xml（默认，向后兼容）', async () => {
        const buf: any = await jsonToPptx(pres());
        const ct = await readContentTypes(buf);
        expect(ct).toContain('presentationml.presentation.main+xml');
        expect(ct).not.toContain('presentationml.template.main+xml');
    });

    it('asTemplate:true 时 presentation.xml 为 template.main+xml', async () => {
        const buf: any = await jsonToPptx(pres(), { asTemplate: true } as any);
        const ct = await readContentTypes(buf);
        expect(ct).toContain('presentationml.template.main+xml');
        expect(ct).not.toContain('presentationml.presentation.main+xml');
    });
});
