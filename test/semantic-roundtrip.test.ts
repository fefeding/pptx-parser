import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToStandard, docFromPptx, docToPptx } from '../src/index';

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
        // 解析端 target 现在是 spid 字符串（cNvPr id），不再是数组索引
        expect(typeof anims[0].target).toBe('string');
        const swivel = anims.find((a: any) => a.type === 'swivel');
        expect(swivel).toBeDefined();
        expect(swivel.presetClass).toBeDefined(); // 透传类别
    });
});

describe('group 子元素动画 round-trip（此前 target 索引偏移导致丢失）', () => {
    it('group 内子元素的动画能正确往返', async () => {
        const slide = await roundtrip({
            slides: [{
                elements: [
                    { type: 'rect', x: 0, y: 0, width: 100, height: 100, fill: { color: 'FF0000' } },
                    {
                        type: 'group', x: 200, y: 0, width: 300, height: 200,
                        childrenCoordinates: 'relative',
                        children: [
                            { type: 'ellipse', x: 0, y: 0, width: 80, height: 80, fill: { color: '00FF00' } },
                            { type: 'text', x: 10, y: 10, width: 200, height: 40, text: 'child' }
                        ]
                    },
                    { type: 'text', x: 0, y: 300, width: 100, height: 50, text: 'after-group' }
                ],
                animations: [
                    { target: 0, type: 'fade', duration: 0.5 },
                    { target: 1, type: 'flyIn', duration: 1, presetClass: 'entr' },
                    { target: 2, type: 'wipe', duration: 0.8 }
                ]
            }]
        });

        const anims = slide.animations || [];
        expect(anims).toHaveLength(3);
        // 所有 target 都应为字符串 spid（不再依赖扁平索引）
        for (const a of anims) {
            expect(typeof a.target).toBe('string');
        }
        // 确认三种 preset 都保留。往返后解析端统一回编辑器友好名，
        // 故 fade / wipe 会以 fadeIn / wipeIn 形式回来（语义不变）
        const types = anims.map((a: any) => a.type);
        expect(types).toContain('fadeIn');
        expect(types).toContain('flyIn');
        expect(types).toContain('wipeIn');
    });
});

describe('编辑器动画往返（按真实 spid 绑定，不再错挂到别的元素）', () => {
    const mk = (name: string, y: number) => ({
        type: 'text', x: 20, y, width: 240, height: 60, name, text: name, fontSize: 24
    });
    const size = { slideSize: { width: 1280, height: 720 } };

    it('只给第 3 个元素设置动画，保存再打开仍作用于该元素且参数不丢', async () => {
        const doc = docFromPptx({ ...size, slides: [{ elements: [mk('A', 20), mk('B', 120), mk('C', 220), mk('D', 320)] }] });
        const target = doc.slides[0].elements[2];
        doc.slides[0].animations = [{
            target: target.id, type: 'fadeIn', presetClass: 'entr',
            direction: 'l', duration: 500, trigger: 'onClick'
        }];
        const buf: any = await jsonToPptx(docToPptx(doc), { outputType: 'uint8array' });
        const back = await pptxToStandard(buf);
        const anims = back.slides[0].animations || [];
        expect(anims).toHaveLength(1);
        expect(typeof anims[0].target).toBe('string');
        expect(anims[0].duration).toBeCloseTo(0.5, 3);
        expect((anims[0].trigger as any).type).toBe('onClick');

        // 目标 spid 必须真的指向形状 C，而不是「下标 + 2」推算出来的别的形状
        const namesById: Record<string, string> = {};
        for (const el of back.slides[0].elements as any[]) {
            const raw = el.__raw && el.__raw.node;
            const nv = raw && (raw['p:nvSpPr'] || raw['p:nvPicPr'] || raw['p:nvGrpSpPr']);
            const c = nv && (Array.isArray(nv['p:cNvPr']) ? nv['p:cNvPr'][0] : nv['p:cNvPr']);
            if (c && c.attrs) namesById[String(c.attrs.id)] = String(c.attrs.name);
        }
        expect(namesById[String(anims[0].target)]).toBe('C');

        const reopened = docFromPptx(back);
        const a = reopened.slides[0].animations[0];
        const bound = (reopened.slides[0].elements as any[]).find((e: any) => e.id === a.target);
        expect(bound && bound.name).toBe('C');
        expect(reopened.slides[0].animations).toHaveLength(1);
        expect(a.duration).toBe(500);          // 毫秒，没有被当成 0.5ms / 500 秒
        expect(a.trigger).toBe('onClick');      // 字符串形态，播放端可直接比较
        expect(a.direction).toBe('l');
        expect(a.type).toBe('fadeIn');
    });

    it('group 之后的元素动画不因数字 id 偏移而挂错', async () => {
        const doc = docFromPptx({ ...size, slides: [{ elements: [
            { type: 'rect', x: 0, y: 0, width: 80, height: 60, name: 'R0', fill: { color: 'FF0000' } },
            { type: 'group', x: 200, y: 0, width: 300, height: 200, name: 'G', childrenCoordinates: 'relative',
              children: [
                  { type: 'rect', x: 0, y: 0, width: 80, height: 80, name: 'G1', fill: { color: '00FF00' } },
                  mk('G2', 10)
              ] },
            mk('TARGET', 300), mk('AFTER', 400)
        ] }] });
        doc.slides[0].animations = [{
            target: doc.slides[0].elements[2].id, type: 'flyIn', presetClass: 'entr',
            duration: 500, trigger: 'afterPrev'
        }];
        const buf: any = await jsonToPptx(docToPptx(doc), { outputType: 'uint8array' });
        const reopened = docFromPptx(await pptxToStandard(buf));
        const a = reopened.slides[0].animations[0];
        const bound = (reopened.slides[0].elements as any[]).find((e: any) => e.id === a.target);
        expect(bound && bound.name).toBe('TARGET');
    });

    it('group 元素（没有 __raw）的动画重新打开后不会丢失', async () => {
        const source: any = await jsonToPptx({ ...size, slides: [{ elements: [
            mk('A', 20),
            { type: 'group', x: 200, y: 0, width: 300, height: 200, name: 'G', childrenCoordinates: 'relative',
              children: [mk('G1', 10)] },
            mk('B', 300)
        ] }] }, { outputType: 'uint8array' });
        // 语义元素不挂 __raw，动画只能靠 spid 回绑
        const std = await pptxToStandard(source);
        const parsedGroup: any = (std.slides[0].elements as any[]).find((e: any) => e.type === 'group');
        expect(parsedGroup).toBeTruthy();
        expect(parsedGroup.__raw).toBeUndefined();
        expect(String(parsedGroup.spid)).toMatch(/^\d+$/);

        const doc = docFromPptx(std);
        const edGroup: any = doc.slides[0].elements.find((e: any) => e.type === 'group');
        doc.slides[0].animations = [{
            target: edGroup.id, type: 'fadeIn', presetClass: 'entr',
            duration: 500, trigger: 'onClick'
        }];
        const buf: any = await jsonToPptx(docToPptx(doc), { outputType: 'uint8array' });
        const back = await pptxToStandard(buf);
        expect(back.slides[0].animations).toHaveLength(1);
        const reopened = docFromPptx(back);
        expect(reopened.slides[0].animations).toHaveLength(1);
        const bound: any = (reopened.slides[0].elements as any[])
            .find((e: any) => e.id === reopened.slides[0].animations[0].target);
        expect(bound && bound.type).toBe('group');
    });

    it('导出的 p:timing 是 PowerPoint 可识别的结构，时长不再放大 1000 倍', async () => {
        const doc = docFromPptx({ ...size, slides: [{ elements: [mk('A', 20), mk('B', 120)] }] });
        doc.slides[0].animations = [{
            target: doc.slides[0].elements[1].id, type: 'fadeIn', presetClass: 'entr',
            duration: 500, trigger: 'onClick'
        }];
        const buf: any = await jsonToPptx(docToPptx(doc), { outputType: 'uint8array' });
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(buf);
        const xml = await zip.file('ppt/slides/slide1.xml')!.async('text');
        expect(xml).toContain('nodeType="tmRoot"');
        expect(xml).toContain('nodeType="mainSeq"');
        expect(xml).toContain('nodeType="clickEffect"');
        expect(xml).toContain('<p:animEffect');
        expect(xml).toContain('presetID="10"');
        expect(xml).toContain('dur="500"');
        expect(xml).not.toContain('dur="500000"');
        const spids = [...xml.matchAll(/<p:spTgt spid="(\d+)"/g)].map(m => m[1]);
        expect(new Set(spids).size).toBe(1);
        expect(spids[0]).toBe('3');
    });

    it('改过类型的动画按新类型写 presetID/presetSubtype，不残留旧效果的编号', async () => {
        const doc = docFromPptx({ ...size, slides: [{ elements: [mk('A', 20), mk('B', 120)] }] });
        doc.slides[0].animations = [{
            target: doc.slides[0].elements[1].id, type: 'flyIn', presetClass: 'entr',
            duration: 500, trigger: 'afterPrev', presetId: 2, presetSubtype: 'left'
        }];
        // 模拟 UI 改类型/方向：旧 presetId(2=flyIn) 与 presetSubtype(left) 仍留在模型里
        doc.slides[0].animations[0].type = 'zoomIn';
        doc.slides[0].animations[0].direction = 'r';
        const buf: any = await jsonToPptx(docToPptx(doc), { outputType: 'uint8array' });
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(buf);
        const xml = await zip.file('ppt/slides/slide1.xml')!.async('text');
        const ids = [...xml.matchAll(/presetID="(\d+)"/g)].map(m => m[1]);
        expect(ids).toContain('21');          // zoom
        expect(ids).not.toContain('2');       // 旧 flyIn
        expect(xml).toContain('presetSubtype="right"');
        expect(xml).not.toContain('presetSubtype="left"');

        // 重新解析回来，类型与 presetId 也保持一致
        const back = await pptxToStandard(buf);
        expect(back.slides[0].animations[0].type).toBe('zoomIn');
        expect(back.slides[0].animations[0].presetId).toBe(21);
    });
});
