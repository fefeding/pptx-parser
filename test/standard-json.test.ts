import { describe, it, expect } from 'vitest';
import { extractSlideToStandard } from '../src/serializer/json-from-pptx';

// 合成一个 slide tree（tXml simplify 形态），验证 text/shape 提取
function makeSlide() {
    return {
        slideContent: {
            'p:sld': {
                'p:cSld': {
                    'p:bg': { 'p:bgPr': { 'a:solidFill': { 'a:srgbClr': { attrs: { val: 'FF0000' } } } } },
                    'p:spTree': {
                        'p:sp': [
                            {
                                'p:nvSpPr': { 'p:cNvPr': { attrs: { id: 2, name: 'TextBox 1' } }, 'p:cNvSpPr': { attrs: { txBox: '1' } }, 'p:nvPr': {} },
                                'p:spPr': {
                                    'a:xfrm': { attrs: { rot: '60000' }, 'a:off': { attrs: { x: '1828800', y: '914400' } }, 'a:ext': { attrs: { cx: '3657600', cy: '914400' } } },
                                    'a:prstGeom': { attrs: { prst: 'rect' }, 'a:avLst': {} }
                                },
                                'p:txBody': {
                                    'a:bodyPr': { attrs: { anchor: 'ctr' } },
                                    'a:lstStyle': {},
                                    'a:p': [{
                                        'a:pPr': { attrs: { algn: 'ctr' } },
                                        'a:r': [{
                                            'a:rPr': { attrs: { sz: '2400', b: '1' }, 'a:latin': { attrs: { typeface: 'Arial' } }, 'a:solidFill': { 'a:srgbClr': { attrs: { val: '00FF00' } } } },
                                            'a:t': 'Hello'
                                        }, {
                                            'a:rPr': { attrs: { sz: '1800' } },
                                            'a:t': ' World'
                                        }]
                                    }]
                                }
                            },
                            {
                                'p:nvSpPr': { 'p:cNvPr': { attrs: { id: 3, name: 'Rectangle 2' } }, 'p:cNvSpPr': {}, 'p:nvPr': {} },
                                'p:spPr': {
                                    'a:xfrm': { 'a:off': { attrs: { x: '0', y: '0' } }, 'a:ext': { attrs: { cx: '1000000', cy: '500000' } } },
                                    'a:prstGeom': { attrs: { prst: 'roundRect' }, 'a:avLst': {} },
                                    'a:solidFill': { 'a:srgbClr': { attrs: { val: '0000FF' } } },
                                    'a:ln': { attrs: { w: '25400' }, 'a:solidFill': { 'a:srgbClr': { attrs: { val: '333333' } } } }
                                }
                            }
                        ]
                    }
                },
                'p:transition': { attrs: { spd: '2' }, 'p:fade': {} }
            }
        },
        notesContent: undefined,
        slideResObj: {}
    };
}

describe('extractSlideToStandard', () => {
    it('提取 text 与 shape，并解析背景/过渡', async () => {
        const zip: any = { file: () => null };
        const slide = await extractSlideToStandard(makeSlide(), zip);

        expect(slide.background).toBe('FF0000');
        expect(slide.transition).toEqual({ type: 'fade', duration: 1000 });
        expect(slide.elements.length).toBe(2);

        const textEl = slide.elements[0] as any;
        expect(textEl.type).toBe('text');
        expect(textEl.x).toBeCloseTo(1828800 * (96 / 914400), 1); // px
        expect(textEl.rotation).toBe(1);
        expect(textEl.align).toBe('center');
        expect(textEl.valign).toBe('middle');
        expect(textEl.paragraphs[0].runs[0]).toMatchObject({ text: 'Hello', fontSize: 24, bold: true, color: '00FF00', fontFace: 'Arial' });
        expect(textEl.paragraphs[0].runs[1].text).toBe(' World');
        expect(textEl.__raw).toBeDefined();

        const shapeEl = slide.elements[1] as any;
        expect(shapeEl.type).toBe('shape');
        expect(shapeEl.shapeType).toBe('roundRect');
        expect(shapeEl.fill).toBe('0000FF');
        expect(shapeEl.line).toEqual({ color: '333333', width: 2 }); // 25400 EMU / 12700 = 2pt
    });
});
