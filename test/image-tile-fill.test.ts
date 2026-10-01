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
    const zip = await JSZip.loadAsync(data);
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const patched = xml.replace(/<a:stretch>[\s\S]*?<\/a:stretch>/, `<a:tile ${attrs}/>`);
    expect(patched).not.toBe(xml); // 断言替换确实发生，避免测试空转
    zip.file('ppt/slides/slide1.xml', patched);
    return zip.generateAsync({ type: 'uint8array' });
}

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
        const data = await jsonToPptx(shapeDoc({
            type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 240, height: 160,
            fill: { type: 'image', data: PNG_64 }
        }));
        const html = await renderFirstSlide(data);

        expect(html).toContain('<pattern id="imgPtrn_');
        expect(html).toContain('patternContentUnits="objectBoundingBox"');
    });

    it('tile 模式：按图片原始尺寸算出 tile 像素（64px × 50% = 32px）', async () => {
        const origin = await jsonToPptx(shapeDoc({
            type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 240, height: 160,
            fill: { type: 'image', data: PNG_64 }
        }));
        const tiled = await toTileBlipFill(origin, 'sx="50000" sy="50000"');
        const html = await renderFirstSlide(tiled);

        // 旧实现用同步 new Image() 探测尺寸，浏览器/Node 下恒为 0 → tile 尺寸为 0 → 退回铺满
        expect(html).toMatch(/<pattern id="imgPtrn_\d+" x="0" y="0" width="32" height="32" patternUnits="userSpaceOnUse">/);
        expect(html).toContain('width="32" height="32"');
    });

    it('tile 偏移 tx/ty 写入 pattern 的 x/y', async () => {
        const origin = await jsonToPptx(shapeDoc({
            type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 240, height: 160,
            fill: { type: 'image', data: PNG_64 }
        }));
        const tiled = await toTileBlipFill(origin, 'sx="100000" sy="100000" tx="25000" ty="0"');
        const html = await renderFirstSlide(tiled);

        expect(html).toMatch(/<pattern id="imgPtrn_\d+" x="16" y="0" width="64" height="64" patternUnits="userSpaceOnUse">/);
    });
});
