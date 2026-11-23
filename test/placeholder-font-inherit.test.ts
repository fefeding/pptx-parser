import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToStandard } from '../src/index';

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

/** 构造一页只含占位符的幻灯片；run 不写 sz，字号必须靠版式/母版继承链补出。 */
function slideWithPlaceholder(phXml: string, runXml = '<a:rPr lang="en-US"/>'): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
        + `<p:sld ${NS}><p:cSld><p:spTree>`
        + `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`
        + `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Placeholder"/>`
        + `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phXml}</p:nvPr></p:nvSpPr>`
        + `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>`
        + `<a:p><a:r>${runXml}<a:t>Hello</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>`
        + `</p:txBody></p:sp></p:spTree></p:cSld>`
        + `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

async function parseWithSlide(slideXml: string, masterXml?: string): Promise<any> {
    const buf: any = await jsonToPptx({ slides: [{ elements: [] }] });
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buf);
    zip.file('ppt/slides/slide1.xml', slideXml);
    if (masterXml) zip.file('ppt/slideMasters/slideMaster1.xml', masterXml);
    const patched: any = await zip.generateAsync({ type: 'nodebuffer' });
    return pptxToStandard(patched, { mediaProcess: false } as any);
}

describe('占位符文本缺省字号继承链', () => {
    it('标题占位符 run 无 sz 时继承母版 p:titleStyle 的 44pt', async () => {
        const std = await parseWithSlide(slideWithPlaceholder('<p:ph type="title"/>'));
        const el = std.slides[0].elements[0];
        expect(el.paragraphs[0].runs[0].fontSize).toBe(44);
        expect(el.fontSize).toBe(44);
    });

    it('正文占位符继承母版 p:bodyStyle 的 32pt', async () => {
        const std = await parseWithSlide(slideWithPlaceholder('<p:ph type="body" idx="1"/>'));
        expect(std.slides[0].elements[0].paragraphs[0].runs[0].fontSize).toBe(32);
    });

    it('整条继承链都无字号时回退 18pt（编辑器不再用 24pt 顶替）', async () => {
        // 母版去掉 txStyles，占位符类型走 otherStyle 也无字号可取
        const buf: any = await jsonToPptx({ slides: [{ elements: [] }] });
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(buf);
        const master = await zip.file('ppt/slideMasters/slideMaster1.xml')!.async('text');
        zip.file('ppt/slideMasters/slideMaster1.xml', master.replace(/<p:txStyles>[\s\S]*?<\/p:txStyles>/, ''));
        zip.file('ppt/slides/slide1.xml', slideWithPlaceholder('<p:ph type="body" idx="1"/>'));
        const patched: any = await zip.generateAsync({ type: 'nodebuffer' });
        const std: any = await pptxToStandard(patched, { mediaProcess: false } as any);
        expect(std.slides[0].elements[0].paragraphs[0].runs[0].fontSize).toBe(18);
    });

    it('母版 p:txStyles 用 a:defPPr（无 a:lvl1pPr）时字号仍生效', async () => {
        const buf: any = await jsonToPptx({ slides: [{ elements: [] }] });
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(buf);
        const master = await zip.file('ppt/slideMasters/slideMaster1.xml')!.async('text');
        // 把 bodyStyle 的 a:lvl1pPr/a:defRPr 改成顶层 a:defPPr/a:defRPr
        const rewritten = master.replace(
            /<p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"\/><\/a:lvl1pPr><\/p:bodyStyle>/,
            '<p:bodyStyle><a:defPPr><a:defRPr sz="2800"/></a:defPPr></p:bodyStyle>'
        );
        expect(rewritten).not.toBe(master);
        zip.file('ppt/slideMasters/slideMaster1.xml', rewritten);
        zip.file('ppt/slides/slide1.xml', slideWithPlaceholder('<p:ph type="body" idx="1"/>'));
        const patched: any = await zip.generateAsync({ type: 'nodebuffer' });
        const std: any = await pptxToStandard(patched, { mediaProcess: false } as any);
        expect(std.slides[0].elements[0].paragraphs[0].runs[0].fontSize).toBe(28);
    });
});
