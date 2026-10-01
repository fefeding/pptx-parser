import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx, pptxToHtml } from '../src/index.ts';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

function shapeDoc(el: any) {
    return { slides: [{ elements: [el] }] };
}

function shapeEl(overrides: any) {
    return { type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 240, height: 160, ...overrides };
}

async function renderFirstSlide(data: Uint8Array): Promise<string> {
    const res: any = await pptxToHtml(toArrayBuffer(data), { themeProcess: false });
    return res.slides[0].html as string;
}

/** 1x1 PNG，用于图片填充用例 */
const ONE_PX_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

describe('T6 形状图案填充（生成端）', () => {
    it('pattFill 写入 prst + fgClr/bgClr', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({
            fill: { type: 'pattern', prst: 'diagCross', fg: '#EF4444', bg: '#FEF3C7' }
        })));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:pattFill prst="diagCross"');
        expect(xml).toContain('<a:fgClr>');
        expect(xml).toContain('val="EF4444"');
        expect(xml).toContain('<a:bgClr>');
        expect(xml).toContain('val="FEF3C7"');
    });
});

describe('T6 形状图案填充（渲染端回归）', () => {
    it('pattFill 生成 SVG <pattern> 并由形状 fill=url(#pattPtrn_xx) 引用', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({
            shapeType: 'roundRect',
            fill: { type: 'pattern', prst: 'diagCross', fg: '#EF4444', bg: '#FEF3C7' }
        })));
        const html = await renderFirstSlide(data);

        expect(html).toContain('<pattern id="pattPtrn_');
        expect(html).toMatch(/fill='url\(#pattPtrn_\d+\)'/);
        // 底色 + 前景交叉线落到 pattern 内
        expect(html).toContain('fill="#FEF3C7"');
        expect(html).toContain('stroke="#EF4444"');
        // 旧实现把图案塞进 styleTable 且以图案文本为 key，会产出 `background: 0` 之类的无效 CSS
        expect(html).not.toContain('background: 0');
    });

    it.each([
        'horz', 'vert', 'ltHorz', 'dkVert', 'narHorz', 'dashHorz',
        'upDiag', 'dnDiag', 'ltUpDiag', 'wdDnDiag', 'dashUpDiag',
        'cross', 'smGrid', 'lgGrid', 'dotGrid', 'smCheck', 'lgCheck',
        'dotDmnd', 'openDmnd', 'solidDmnd', 'smConfetti', 'lgConfetti',
        'horzBrick', 'diagBrick', 'weave', 'plaid', 'trellis', 'shingle', 'wave', 'zigZag', 'sphere', 'divot',
        'pct5', 'pct10', 'pct25', 'pct50', 'pct75', 'pct90',
        'notARealPrst' // 未收录图案走兜底，也必须可见
    ])('prst=%s 生成可被引用的图案定义', async (prst) => {
        const data = await jsonToPptx(shapeDoc(shapeEl({
            fill: { type: 'pattern', prst, fg: '#EF4444', bg: '#FEF3C7' }
        })));
        const html = await renderFirstSlide(data);

        expect(html).toContain('<pattern id="pattPtrn_');
        expect(html).toMatch(/fill='url\(#pattPtrn_\d+\)'/);
        // 图案 tile 必须画出内容（有背景色 rect，且不是空 body）
        expect(html).toMatch(/<pattern id="pattPtrn_\d+" width="\d+" height="\d+" patternUnits="userSpaceOnUse"><rect/);
    });

    it('图案填充不再污染 styleTable（同一图案的多个形状互不覆盖）', async () => {
        const data = await jsonToPptx({
            slides: [{
                elements: [
                    shapeEl({ fill: { type: 'pattern', prst: 'diagCross', fg: '#EF4444', bg: '#FEF3C7' } }),
                    shapeEl({ y: 200, fill: { type: 'pattern', prst: 'diagCross', fg: '#EF4444', bg: '#FEF3C7' } })
                ]
            }]
        });
        const html = await renderFirstSlide(data);
        const ids = [...html.matchAll(/fill='url\(#(pattPtrn_\d+)\)'/g)].map(m => m[1]);

        expect(ids).toHaveLength(2);
        expect(new Set(ids).size).toBe(2); // 两个形状各引用各自的 pattern
    });
});

describe('T6 形状图片填充（渲染端回归）', () => {
    it('blipFill 生成 SVG <pattern> 图片定义并由形状引用', async () => {
        const data = await jsonToPptx(shapeDoc(shapeEl({
            fill: { type: 'image', data: ONE_PX_PNG }
        })));
        const html = await renderFirstSlide(data);

        expect(html).toContain('<pattern id="imgPtrn_');
        expect(html).toMatch(/fill='url\(#imgPtrn_\d+\)'/);
        expect(html).toContain('<image');
        expect(html).toContain('preserveAspectRatio="none"');
    });
});
