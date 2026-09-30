import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx, pptxToStandard, PPTXComposer } from '../src/index.ts';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

function textEl(overrides: any) {
    return { type: 'text', x: 0, y: 0, width: 300, height: 100, ...overrides };
}
function textDoc(el: any) {
    return { slides: [{ elements: [el] }] };
}

describe('T3 列表/行距/段间距/缩进（生成端）', () => {
    it('自动编号 bullet:"number" 生成 a:buAutoNum', async () => {
        const data = await jsonToPptx(textDoc(textEl({ paragraphs: [{ text: 'A', bullet: 'number' }] })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('<a:buAutoNum type="arabicPeriod" startAt="1"');
    });

    it('自定义编号 fmt/start 生成 a:buAutoNum type/startAt', async () => {
        const data = await jsonToPptx(textDoc(textEl({ paragraphs: [{ text: 'A', bullet: { type: 'number', fmt: 'alphaLc', start: 3 } }] })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('<a:buAutoNum type="alphaLcPeriod" startAt="3"');
    });

    it('项目符号 bullet:true 生成 a:buChar char="•"', async () => {
        const data = await jsonToPptx(textDoc(textEl({ paragraphs: [{ text: 'A', bullet: true }] })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('<a:buChar char="•"');
    });

    it('自定义项目符号字符 bullet:{type:"bullet",char:"-"}', async () => {
        const data = await jsonToPptx(textDoc(textEl({ paragraphs: [{ text: 'A', bullet: { type: 'bullet', char: '-' } }] })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('<a:buChar char="-"');
    });

    it('行距百分比 lineSpacing:200 生成 a:lnSpc/a:spcPct val=200000', async () => {
        const data = await jsonToPptx(textDoc(textEl({ paragraphs: [{ text: 'A', lineSpacing: 200 }] })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('<a:lnSpc');
        expect(xml).toContain('<a:spcPct val="200000"');
    });

    it('行距 pt lineSpacing:{type:"pt",value:24} 生成 a:spcPts val=2400', async () => {
        const data = await jsonToPptx(textDoc(textEl({ paragraphs: [{ text: 'A', lineSpacing: { type: 'pt', value: 24 } }] })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('<a:spcPts val="2400"');
    });

    it('段间距 spaceBefore/spaceAfter 生成 a:spcBef/a:spcAft', async () => {
        const data = await jsonToPptx(textDoc(textEl({ paragraphs: [{ text: 'A', spaceBefore: 10, spaceAfter: 6 }] })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('<a:spcBef');
        expect(xml).toContain('val="1000"');
        expect(xml).toContain('<a:spcAft');
        expect(xml).toContain('val="600"');
    });

    it('缩进 indentLeft/indentRight/indent 生成 marL/marR/indent', async () => {
        const data = await jsonToPptx(textDoc(textEl({ paragraphs: [{ text: 'A', indentLeft: 20, indentRight: 15, indent: 10 }] })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('marL="254000"');   // 20pt
        expect(xml).toContain('marR="190500"');   // 15pt
        expect(xml).toContain('indent="127000"'); // 10pt
    });

    it('纯 text 模式（无 paragraphs）元素级段落默认透传：编号 + 行距', async () => {
        const data = await jsonToPptx(textDoc(textEl({ text: 'Hello', bullet: 'number', lineSpacing: 200 })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('<a:buAutoNum type="arabicPeriod" startAt="1"');
        expect(xml).toContain('<a:spcPct val="200000"');
    });
});

describe('T3 composer 纯文本模式段落级 setter', () => {
    it('addText 链式 .bullet/.lineSpacing 生成编号与行距', async () => {
        const comp = new PPTXComposer({ width: 960, height: 540 });
        comp.addSlide((slide: any) => { slide.addText((b: any) => { b.value('Item'); b.bullet('number'); b.lineSpacing(150); }); });
        const data = await comp.save();
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
        expect(xml).toContain('<a:buAutoNum type="arabicPeriod" startAt="1"');
        expect(xml).toContain('<a:spcPct val="150000"');
    });
});

describe('T3 段落样式 round-trip 回读', () => {
    it('pptxToStandard 回读编号/行距/缩进', async () => {
        const data = await jsonToPptx(textDoc(textEl({
            paragraphs: [{
                text: 'A',
                bullet: { type: 'number', fmt: 'alphaLc', start: 2 },
                lineSpacing: 200,
                indentLeft: 20,
                spaceBefore: 10
            }]
        })));
        const std: any = await pptxToStandard(toArrayBuffer(data));
        const para = std.slides[0].elements[0].paragraphs[0];

        // 生成端把友好写法 'alphaLc' 归一为合法值 'alphaLcPeriod'，回读即该合法值
        expect(para.bullet).toEqual({ type: 'number', fmt: 'alphaLcPeriod', start: 2 });
        expect(para.lineSpacing).toEqual({ type: 'percent', value: 200 });
        expect(para.indentLeft).toBe(20);
        expect(para.spaceBefore).toBe(10);
    });

    it('pptxToStandard 回读项目符号字符', async () => {
        const data = await jsonToPptx(textDoc(textEl({ paragraphs: [{ text: 'A', bullet: { type: 'bullet', char: '-' } }] })));
        const std: any = await pptxToStandard(toArrayBuffer(data));
        const para = std.slides[0].elements[0].paragraphs[0];
        expect(para.bullet).toEqual({ type: 'bullet', char: '-' });
    });
});
