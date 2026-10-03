import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToStandard } from '../src/index';

/**
 * 语义链路 round-trip：生成端产出 → 解析端读回，验证此前「解析端静默丢弃」的项
 * 能完整还原（group 层级 / video / audio）。
 */

async function roundtrip(pres: any) {
    const buf: any = await jsonToPptx(pres);
    const doc = await pptxToStandard(buf);
    return doc.slides[0];
}

describe('语义链路 group 保留（此前被扁平化为叶子）', () => {
    it('group 连同 children 被还原为 PptxGroupElement', async () => {
        const slide = await roundtrip({
            slides: [{
                elements: [
                    {
                        type: 'group', x: 100, y: 100, width: 300, height: 200,
                        childrenCoordinates: 'relative',
                        children: [
                            { type: 'rect', x: 0, y: 0, width: 100, height: 50, fill: { color: 'FF0000' } },
                            { type: 'text', x: 10, y: 10, width: 200, height: 40, text: 'hello' }
                        ]
                    }
                ]
            }]
        });

        const g = slide.elements[0];
        expect(g.type).toBe('group');
        expect((g as any).children).toHaveLength(2);
        // 解析端形状统一输出 type:'shape' + shapeType（既有约定）
        expect((g as any).children[0].type).toBe('shape');
        expect((g as any).children[0].shapeType).toBe('rect');
        expect((g as any).children[1].type).toBe('text');
    });

    it('嵌套 group 也保留层级', async () => {
        const slide = await roundtrip({
            slides: [{
                elements: [
                    {
                        type: 'group', x: 0, y: 0, width: 400, height: 300,
                        childrenCoordinates: 'relative',
                        children: [
                            {
                                type: 'group', x: 50, y: 50, width: 200, height: 150,
                                childrenCoordinates: 'relative',
                                children: [
                                    { type: 'ellipse', x: 0, y: 0, width: 80, height: 80, fill: { color: '00FF00' } }
                                ]
                            }
                        ]
                    }
                ]
            }]
        });

        const outer = slide.elements[0];
        expect(outer.type).toBe('group');
        const inner = (outer as any).children[0];
        expect(inner.type).toBe('group');
        expect(inner.children[0].type).toBe('shape');
        expect(inner.children[0].shapeType).toBe('ellipse');
    });
});

describe('语义链路 video / audio 保留（此前一律退化 image）', () => {
    const media = 'AAAA'; // 占位 base64（测试不校验真实解码）

    it('video 元素往返后仍识别为 video', async () => {
        const slide = await roundtrip({
            slides: [{
                elements: [{
                    type: 'video', x: 50, y: 50, width: 320, height: 180,
                    extension: 'mp4', data: media,
                    poster: { data: media, extension: 'png' }
                }]
            }]
        });
        const v = slide.elements[0];
        expect(v.type).toBe('video');
        expect((v as any).extension).toBe('mp4');
        expect((v as any).data).toBeDefined();
        expect((v as any).poster).toBeDefined();
    });

    it('audio 元素往返后仍识别为 audio', async () => {
        const slide = await roundtrip({
            slides: [{
                elements: [{
                    type: 'audio', x: 20, y: 20, width: 60, height: 60,
                    extension: 'mp3', data: media
                }]
            }]
        });
        const a = slide.elements[0];
        expect(a.type).toBe('audio');
        expect((a as any).extension).toBe('mp3');
    });
});

describe('解析端动画 preset 透传', () => {
    it('非标准 preset 不再退化为 fade', async () => {
        const slide = await roundtrip({
            slides: [{
                elements: [
                    { type: 'rect', x: 0, y: 0, width: 100, height: 100, fill: { color: 'FF0000' } },
                    { type: 'text', x: 0, y: 120, width: 100, height: 100, text: 'b' }
                ],
                animations: [
                    { target: 0, type: 'fade', duration: 0.5 },
                    { target: 1, type: 'swivel', duration: 1 }
                ]
            }]
        });
        const anims = slide.animations || [];
        expect(anims).toHaveLength(2);
        const swivel = anims.find((a: any) => a.type === 'swivel');
        expect(swivel).toBeDefined();
        expect(swivel.presetClass).toBeDefined(); // 透传类别
    });
});
