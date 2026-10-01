import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx, pptxToHtml } from '../src/index.ts';
import { PPTXStyleUtils } from '../src/utils/style.ts';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

function shapeDoc(el: any) {
    return { slides: [{ elements: [el] }] };
}

/** 64x64 纯色 PNG（真机/浏览器可正常解码） */
const PNG_64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAZUlEQVR42u3QQREAAAQAML20E9qXHM4eK7DI6vksBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAgPsWcJEihvVdy3EAAAAASUVORK5CYII=';

async function renderFirstSlide(data: Uint8Array): Promise<string> {
    const res: any = await pptxToHtml(toArrayBuffer(data), { themeProcess: false });
    return res.slides[0].html as string;
}

/** 把序列化出的 a:stretch 替换成 a:tile（生成端暂不支持 tile，改造已有产物模拟 WPS/PowerPoint 的产物） */
async function toTileBlipFill(data: Uint8Array, attrs: string): Promise<Uint8Array> {
    return patchBlipFill(data, `<a:tile ${attrs}/>`);
}

/** 插入 a:srcRect（生成端暂不支持裁剪） */
async function toSrcRectBlipFill(data: Uint8Array, attrs: string): Promise<Uint8Array> {
    return patchBlipFill(data, `<a:srcRect ${attrs}/><a:stretch><a:fillRect/></a:stretch>`);
}

/** 用 a:stretch/a:fillRect 表达同一裁剪（WPS 常见写法） */
async function toFillRectBlipFill(data: Uint8Array, attrs: string): Promise<Uint8Array> {
    return patchBlipFill(data, `<a:stretch><a:fillRect ${attrs}/></a:stretch>`);
}

async function patchBlipFill(data: Uint8Array, replacement: string): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(data);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const patched = xml.replace(/<a:stretch>[\s\S]*?<\/a:stretch>/, replacement);
    expect(patched).not.toBe(xml); // 断言替换确实发生，避免测试空转
    zip.file('ppt/slides/slide1.xml', patched);
    return zip.generateAsync({ type: 'uint8array' });
}

function imageShape() {
    return { type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 240, height: 160, fill: { type: 'image', data: PNG_64 } };
}

describe('生成端：写入 a:tile / a:srcRect', () => {
    it('未指定时仍写 a:stretch/a:fillRect（默认拉伸铺满，无回归）', async () => {
        const data = await jsonToPptx(shapeDoc(imageShape()));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:stretch><a:fillRect/></a:stretch>');
        expect(xml).not.toContain('<a:tile');
        expect(xml).not.toContain('<a:srcRect');
    });

    it('tile：sx/sy/tx/ty 以千分比写入，且不再写 a:stretch', async () => {
        const data = await jsonToPptx(shapeDoc({
            ...imageShape(),
            fill: { type: 'image', data: PNG_64, tile: { sx: 0.5, sy: 0.5 } }
        }));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:tile sx="50000" sy="50000" tx="0" ty="0"/>');
        expect(xml).not.toContain('<a:stretch');
    });

    it('srcRect：l/t/r/b 以千分比写入，并保留拉伸铺满', async () => {
        const data = await jsonToPptx(shapeDoc({
            ...imageShape(),
            fill: { type: 'image', data: PNG_64, srcRect: { l: 0.25, t: 0.25, r: 0.25, b: 0.25 } }
        }));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:srcRect l="25000" t="25000" r="25000" b="25000"/>');
        expect(xml).toContain('<a:stretch><a:fillRect/></a:stretch>');
    });

    it('同时指定：srcRect 在前、tile 在后（符合 CT_BlipFillProperties 顺序）', async () => {
        const data = await jsonToPptx(shapeDoc({
            ...imageShape(),
            fill: {
                type: 'image', data: PNG_64,
                srcRect: { l: 0, t: 0, r: 0.5, b: 0.5 },
                tile: { sx: 0.25, sy: 0.25, tx: 0.1, ty: 0 }
            }
        }));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toMatch(/<a:srcRect[^>]*\/><a:tile[^>]*\/>/);
        expect(xml).toContain('<a:tile sx="25000" sy="25000" tx="10000" ty="0"/>');
    });

    it('越界值被收敛到 [0,1]', async () => {
        const data = await jsonToPptx(shapeDoc({
            ...imageShape(),
            fill: { type: 'image', data: PNG_64, tile: { sx: 2, sy: -1 }, srcRect: { r: 5 } }
        }));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:tile sx="100000" sy="0" tx="0" ty="0"/>');
        expect(xml).toContain('<a:srcRect l="0" t="0" r="100000" b="0"/>');
    });

    it('幻灯片背景图片同样支持 tile / srcRect', async () => {
        const data = await jsonToPptx({
            slideSize: { width: 640, height: 360 },
            slides: [{
                background: {
                    type: 'image', data: PNG_64,
                    tile: { sx: 0.5, sy: 0.5 },
                    srcRect: { l: 0.25, t: 0, r: 0, b: 0 }
                },
                elements: []
            }]
        });
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:srcRect l="25000" t="0" r="0" b="0"/>');
        expect(xml).toContain('<a:tile sx="50000" sy="50000" tx="0" ty="0"/>');
    });
});

describe('图片原始尺寸解析（a:tile 平铺的前置条件）', () => {
    it('PNG：从 IHDR 读出 64x64', () => {
        const buf = Uint8Array.from(Buffer.from(PNG_64.split(',')[1], 'base64'));
        expect(PPTXStyleUtils.getImageSizeFromBuffer(toArrayBuffer(buf))).toEqual({ width: 64, height: 64 });
    });

    it('JPEG：从 SOF0 段读出宽高（height 在前）', () => {
        // SOI + SOF0(len=17, precision=8, height=16, width=32, 3 个分量) + EOI
        const bytes = [
            0xFF, 0xD8,
            0xFF, 0xC0, 0x00, 0x11, 0x08, 0x00, 0x10, 0x00, 0x20, 0x03,
            0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01,
            0xFF, 0xD9
        ];
        expect(PPTXStyleUtils.getImageSizeFromBuffer(toArrayBuffer(Uint8Array.from(bytes))))
            .toEqual({ width: 32, height: 16 });
    });

    it('GIF / BMP：分别按各自头部解析', () => {
        const gif = Uint8Array.from([...Buffer.from('GIF89a'), 0x28, 0x00, 0x14, 0x00, 0x00, 0x00]);
        expect(PPTXStyleUtils.getImageSizeFromBuffer(toArrayBuffer(gif))).toEqual({ width: 40, height: 20 });

        const bmp = new Uint8Array(30);
        bmp[0] = 0x42; bmp[1] = 0x4D;
        new DataView(bmp.buffer).setInt32(18, 24, true);   // width
        new DataView(bmp.buffer).setInt32(22, -12, true);  // height（自下而上，负值）
        expect(PPTXStyleUtils.getImageSizeFromBuffer(toArrayBuffer(bmp))).toEqual({ width: 24, height: 12 });
    });

    it('无法识别的格式返回 0x0（不抛异常）', () => {
        const junk = Uint8Array.from([0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08]);
        expect(PPTXStyleUtils.getImageSizeFromBuffer(toArrayBuffer(junk))).toEqual({ width: 0, height: 0 });
    });
});

describe('形状图片填充 a:tile 平铺（渲染端）', () => {
    it('stretch 模式：整图铺满仍走 objectBoundingBox（无回归）', async () => {
        const data = await jsonToPptx(shapeDoc(imageShape()));
        const html = await renderFirstSlide(data);

        expect(html).toContain('<pattern id="imgPtrn_');
        expect(html).toContain('patternContentUnits="objectBoundingBox"');
    });

    it('tile 模式：按图片原始尺寸算出 tile 像素（64px × 50% = 32px）', async () => {
        const origin = await jsonToPptx(shapeDoc(imageShape()));
        const tiled = await toTileBlipFill(origin, 'sx="50000" sy="50000"');
        const html = await renderFirstSlide(tiled);

        // 旧实现用同步 new Image() 探测尺寸，浏览器/Node 下恒为 0 → tile 尺寸为 0 → 退回铺满
        expect(html).toMatch(/<pattern id="imgPtrn_\d+" x="0" y="0" width="32" height="32" patternUnits="userSpaceOnUse">/);
        expect(html).toContain('width="32" height="32"');
    });

    it('tile 偏移 tx/ty 写入 pattern 的 x/y', async () => {
        const origin = await jsonToPptx(shapeDoc(imageShape()));
        const tiled = await toTileBlipFill(origin, 'sx="100000" sy="100000" tx="25000" ty="0"');
        const html = await renderFirstSlide(tiled);

        expect(html).toMatch(/<pattern id="imgPtrn_\d+" x="16" y="0" width="64" height="64" patternUnits="userSpaceOnUse">/);
    });
});

describe('形状图片填充 a:srcRect 源图裁剪（渲染端）', () => {
    it('无裁剪时不产生 viewBox（无回归）', async () => {
        const html = await renderFirstSlide(await jsonToPptx(shapeDoc(imageShape())));
        expect(html).not.toContain('viewBox');
    });

    it('a:srcRect 四边各裁 25% → 裁剪窗口 viewBox="16 16 32 32"（64px 图）', async () => {
        const origin = await jsonToPptx(shapeDoc(imageShape()));
        const cropped = await toSrcRectBlipFill(origin, 'l="25000" t="25000" r="25000" b="25000"');
        const html = await renderFirstSlide(cropped);

        expect(html).toContain('<svg width="1" height="1" viewBox="16 16 32 32" preserveAspectRatio="none">');
        expect(html).toContain('width="64" height="64"');
    });

    it('a:stretch/a:fillRect 表达同一裁剪（WPS 写法）→ viewBox 一致', async () => {
        const origin = await jsonToPptx(shapeDoc(imageShape()));
        const cropped = await toFillRectBlipFill(origin, 'l="0" t="0" r="50000" b="50000"');
        const html = await renderFirstSlide(cropped);

        // 保留左上 1/4：起点 (0,0)、窗口 32x32
        expect(html).toContain('viewBox="0 0 32 32"');
    });

    it('平铺 + 裁剪：裁剪窗口映射到每一格', async () => {
        const origin = await jsonToPptx(shapeDoc(imageShape()));
        const tiled = await patchBlipFill(origin, '<a:srcRect l="25000" t="25000" r="25000" b="25000"/><a:tile sx="50000" sy="50000"/>');
        const html = await renderFirstSlide(tiled);

        expect(html).toContain('<svg width="32" height="32" viewBox="16 16 32 32" preserveAspectRatio="none">');
    });
});

describe('图片裁剪的 CSS 背景换算（非 SVG 路径）', () => {
    it('铺满 + 四边各裁 25% → background-size 200% / position 50%', async () => {
        const data = await jsonToPptx(shapeDoc(imageShape()));
        const zip = await JSZip.loadAsync(data);
        const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
        const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map(m => m[0]).find(s => s.includes('/image"'));
        const rid = /Id="([^"]+)"/.exec(rel!)?.[1];
        const target = /Target="([^"]+)"/.exec(rel!)?.[1];

        const warpObj: any = {
            zip,
            slideResObj: { [rid!]: { target } },
            'loaded-images': {},
            'loaded-image-sizes': {}
        };
        const blipFill: any = {
            'a:blip': { attrs: { 'r:embed': rid } },
            'a:srcRect': { attrs: { l: '25000', t: '25000', r: '25000', b: '25000' } }
        };
        const result: any = await PPTXStyleUtils.getPicFill('slide', blipFill, warpObj);

        expect(result.srcRect).toEqual({ l: 0.25, t: 0.25, r: 0.25, b: 0.25 });
        // 裁剪窗口占原图 50%，因此放大到 200%；
        // background-position 的百分比是「图片 p% 对齐容器 p%」，对应值为 l/(1-crop)
        expect(result.backgroundSize).toBe('200% 200%');
        expect(result.backgroundPosition).toBe('50% 50%');
    });
});
