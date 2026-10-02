import { describe, it, expect } from 'vitest';
import pptxToHtml from '../src/index';
import { jsonToPptx, pptxToStandard } from '../src/index';

function presWith(overrides: any = {}) {
    return {
        slides: [
            {
                background: '#ffffff',
                elements: [
                    { type: 'text', x: 60, y: 30, width: 1160, height: 50, text: 'T18 · 批注 comments', fontSize: 30, bold: true },
                    { type: 'shape', shapeType: 'foldedCorner', x: 200, y: 200, width: 300, height: 120, fill: { color: '#fde68a' } }
                ],
                comments: [
                    { author: 'Alice', text: '这是一条批注', dt: '2026-01-01T00:00:00Z' },
                    { author: 'Bob', text: '第二条批注' }
                ]
            }
        ],
        customProps: { '部门': '研发', '版本': '1.0', '项目': 'pptx-parser' },
        ...overrides
    };
}

describe('批注 + 自定义属性 解析与渲染', () => {
    it('生成端：非法 shapeType 回退 rect（不写出非标准 prst）', async () => {
        const pres = presWith();
        // 注入一个非法形状名，验证被回退为 rect
        pres.slides[0].elements.push({ type: 'shape', shapeType: 'note', x: 10, y: 10, width: 50, height: 50 });
        const buf: any = await jsonToPptx(pres);
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(buf);
        const slideXml = await zip.file('ppt/slides/slide1.xml').async('text');
        // 不应出现非标准 prst="note"
        expect(slideXml).not.toContain('prst="note"');
        // 非法名应被回退为 rect
        expect(slideXml).toContain('prst="rect"');
    });

    it('解析端：customProps 与 comments 回读并在 HTML 渲染', async () => {
        const buf: any = await jsonToPptx(presWith());
        const res: any = await pptxToHtml(buf, { mediaProcess: false } as any);

        // 文档自定义属性回读
        expect(res.customProps).toEqual({ '部门': '研发', '版本': '1.0', '项目': 'pptx-parser' });

        const html = res.slides[0].html as string;

        // 批注渲染：便签 + 作者 + 文本
        expect(html).toContain('pptx-comment');
        expect(html).toContain('Alice');
        expect(html).toContain('这是一条批注');
        expect(html).toContain('Bob');
        expect(html).toContain('第二条批注');

        // 自定义属性渲染
        expect(html).toContain('pptx-custom-props');
        expect(html).toContain('部门');
        expect(html).toContain('研发');

        // 便签形状（foldedCorner）按几何渲染出路径（非空白）
        expect(html).toContain('<path');
    });

    it('语义层（pptxToStandard）：comments 与 customProps 回读', async () => {
        const buf: any = await jsonToPptx(presWith());
        const std: any = await pptxToStandard(buf, { mediaProcess: false } as any);

        // 文档级自定义属性
        expect(std.customProps).toEqual({ '部门': '研发', '版本': '1.0', '项目': 'pptx-parser' });

        // 幻灯片级批注
        const comments = std.slides[0].comments;
        expect(Array.isArray(comments)).toBe(true);
        expect(comments.length).toBe(2);
        expect(comments[0].author).toBe('Alice');
        expect(comments[0].text).toBe('这是一条批注');
        expect(comments[1].author).toBe('Bob');
        expect(comments[1].text).toBe('第二条批注');
        // 锚点位置（EMU）回读
        expect(comments[0].pos).toEqual({ x: 914400, y: 914400 });
    });

    it('解析端：无 customProps / 无 comments 时不渲染附加块', async () => {
        const buf: any = await jsonToPptx(presWith({ customProps: undefined }));
        // 移除 comments
        (buf as any);
        const pres2 = presWith({ customProps: undefined });
        pres2.slides[0].comments = undefined;
        const buf2: any = await jsonToPptx(pres2);
        const res: any = await pptxToHtml(buf2, { mediaProcess: false } as any);
        const html = res.slides[0].html as string;
        expect(res.customProps).toEqual({});
        expect(html).not.toContain('pptx-custom-props');
        expect(html).not.toContain('pptx-comment');
    });
});
