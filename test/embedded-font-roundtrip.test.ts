import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToStandard } from '../src/index';

// 一段最小的「伪 TTF」字节（仅用于验证 XOR 混淆/解混淆对称与往返，不做真实字体渲染）
const FONT_BYTES = Uint8Array.from([0x00, 0x01, 0x00, 0x00, 0x4c, 0x05, 0xEC, 0x03, 0x4b, 0x00, 0xab, 0xcd, 0xff, 0x10, 0x20, 0x30, 0x40, 0x50, 0x05, 0x00]);
const FONT_B64 = Buffer.from(FONT_BYTES).toString('base64');

function presWithFont() {
    return {
        slides: [{ elements: [{ type: 'text', x: 60, y: 30, width: 600, height: 50, text: '字体嵌入' }] }],
        fonts: [
            { name: 'Source Han Sans', data: FONT_B64, panose: '020B0604020202020204', bold: false, italic: true, embedType: 'subset' },
            { name: 'Test Serif', data: Buffer.from([1, 2, 3, 4]).toString('base64'), bold: true }
        ]
    };
}

describe('嵌入字体（T19）生成 + 解析往返', () => {
    it('生成端写出混淆的 fntdata / fontTable.xml / rels', async () => {
        const buf: any = await jsonToPptx(presWithFont());
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(buf);
        expect(zip.file('ppt/fonts/font1.fntdata')).toBeTruthy();
        expect(zip.file('ppt/fonts/font2.fntdata')).toBeTruthy();
        const ft = await zip.file('ppt/fontTable.xml').async('text');
        expect(ft).toContain('Source Han Sans');
        expect(ft).toContain('embeddedFont');
        expect(ft).toContain('subsetted="1"');
        const rels = await zip.file('ppt/_rels/fontTable.xml.rels').async('text');
        expect(rels).toContain('font1.fntdata');
    });

    it('解析端回读并解混淆，字体数据与原字节一致', async () => {
        const buf: any = await jsonToPptx(presWithFont());
        const std: any = await pptxToStandard(buf, { mediaProcess: false } as any);
        expect(Array.isArray(std.fonts)).toBe(true);
        expect(std.fonts).toHaveLength(2);
        const f1 = std.fonts.find((f: any) => f.name === 'Source Han Sans');
        expect(f1).toBeTruthy();
        expect(Buffer.from(f1.data, 'base64')).toEqual(Buffer.from(FONT_BYTES));
        expect(f1.panose).toBe('020B0604020202020204');
        expect(f1.italic).toBe(true);
        expect(f1.embedType).toBe('subset');
        const f2 = std.fonts.find((f: any) => f.name === 'Test Serif');
        expect(f2.bold).toBe(true);
        expect(Buffer.from(f2.data, 'base64')).toEqual(Buffer.from([1, 2, 3, 4]));
    });

    it('无嵌入字体时 fonts 不挂到文档', async () => {
        const buf: any = await jsonToPptx({ slides: [{ elements: [] }] });
        const std: any = await pptxToStandard(buf, { mediaProcess: false } as any);
        expect(std.fonts).toBeUndefined();
    });

    it('能读取标准 OOXML 的 a:regular/a:bold 变体结构（第三方 PPTX）', async () => {
        const JSZip = (await import('jszip')).default;
        const { obfuscateFontData } = await import('../src/serializer/templates');
        const { bytesToBase64 } = await import('../src/index');
        const obf = bytesToBase64(obfuscateFontData(new Uint8Array([9, 8, 7, 6])));
        const zip = new JSZip();
        zip.file('ppt/fontTable.xml',
            '<?xml version="1.0"?><p:fontTable xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<a:font typeface="Third Party Font"><a:embeddedFont>' +
            '<a:regular typeface="Third Party Font"><a:fontData r:id="rId1"/></a:regular>' +
            '<a:bold typeface="Third Party Font"><a:fontData r:id="rId2"/></a:bold>' +
            '</a:embeddedFont></a:font></p:fontTable>');
        zip.file('ppt/_rels/fontTable.xml.rels',
            '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Target="fonts/f1.fntdata" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"/>' +
            '<Relationship Id="rId2" Target="../ppt/fonts/f2.fntdata" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"/>' +
            '</Relationships>');
        zip.file('ppt/fonts/f1.fntdata', obf, { base64: true });
        zip.file('ppt/fonts/f2.fntdata', obf, { base64: true });
        const buf: any = await zip.generateAsync({ type: 'uint8array' });
        // 需要一个可解析的最小 PPTX：复用 jsonToPptx 产物，替换其 fontTable 部件
        const base: any = await jsonToPptx({ slides: [{ elements: [] }] });
        const bz = await JSZip.loadAsync(base);
        bz.file('ppt/fontTable.xml', await zip.file('ppt/fontTable.xml').async('text'));
        bz.file('ppt/_rels/fontTable.xml.rels', await zip.file('ppt/_rels/fontTable.xml.rels').async('text'));
        bz.file('ppt/fonts/f1.fntdata', obf, { base64: true });
        bz.file('ppt/fonts/f2.fntdata', obf, { base64: true });
        const merged: any = await bz.generateAsync({ type: 'uint8array' });
        const std: any = await pptxToStandard(merged, { mediaProcess: false } as any);
        expect(std.fonts).toHaveLength(2);
        const reg = std.fonts.find((f: any) => f.name === 'Third Party Font' && !f.bold);
        const bold = std.fonts.find((f: any) => f.name === 'Third Party Font' && f.bold);
        expect(reg).toBeTruthy();
        expect(bold).toBeTruthy();
        expect(Buffer.from(reg.data, 'base64')).toEqual(Buffer.from([9, 8, 7, 6]));
        expect(Buffer.from(bold.data, 'base64')).toEqual(Buffer.from([9, 8, 7, 6]));
    });
});
