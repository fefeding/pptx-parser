import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToHtml } from '../src/index.ts';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

/** 生成单个箭头形状并取回渲染出的 polygon points */
async function arrowPoints(shapeType: string, width: number, height: number, adjust?: Record<string, number>): Promise<string> {
    const el: any = { type: 'shape', shapeType, x: 0, y: 0, width, height, fill: { color: '#10b981' } };
    if (adjust) el.adjust = adjust;
    const data = await jsonToPptx({ slides: [{ elements: [el] }] });
    const res: any = await pptxToHtml(toArrayBuffer(data), { themeProcess: false });
    const html: string = res.slides[0].html;
    const m = /<polygon points='([^']+)'/.exec(html);
    return m ? m[1] : '';
}

describe('箭头 avLst 调整值（a:gd adj1/adj2）', () => {
    // w=240, h=120 → ss=120；默认 adj1=adj2=50000
    // dx1 = ss*adj2/100000 = 60（箭头长度），dy1 = h*adj1/200000 = 30（半箭身厚度）
    // 肩部 x1 = w - dx1 = 180，箭身上下边 y1=30 / y2=90
    const RIGHT_DEFAULT = '0 30,180 30,180 0,240 60,180 120,180 90,0 90';

    it('rightArrow 默认（无 avLst）：箭头长度为 min(w,h)*50% = 60，肩部在 x=180', async () => {
        expect(await arrowPoints('rightArrow', 240, 120)).toBe(RIGHT_DEFAULT);
    });

    it('adj2=40000 → 箭头长度 48、肩部 x=192', async () => {
        const points = await arrowPoints('rightArrow', 240, 120, { adj1: 50000, adj2: 40000 });
        expect(points).toBe('0 30,192 30,192 0,240 60,192 120,192 90,0 90');
    });

    it('adj2=100000 → 箭头长度 120、肩部 x=120', async () => {
        const points = await arrowPoints('rightArrow', 240, 120, { adj1: 50000, adj2: 100000 });
        expect(points).toBe('0 30,120 30,120 0,240 60,120 120,120 90,0 90');
    });

    it('adj1=30000 → 箭身更细（y=42/78）', async () => {
        const points = await arrowPoints('rightArrow', 240, 120, { adj1: 30000, adj2: 40000 });
        expect(points).toBe('0 42,192 42,192 0,240 60,192 120,192 78,0 78');
    });

    it('adj2 超过预设上限 maxAdj2=100000*w/ss 时被收敛', async () => {
        // maxAdj2 = 100000*240/120 = 200000 → dx1 = 240，肩部贴到左边缘
        const points = await arrowPoints('rightArrow', 240, 120, { adj1: 50000, adj2: 300000 });
        expect(points).toBe('0 30,0 30,0 0,240 60,0 120,0 90,0 90');
    });

    it('非法的 gd 名会被忽略并回退默认值（与 WPS/PowerPoint 行为一致）', async () => {
        const points = await arrowPoints('rightArrow', 240, 120, { arrowWidth: 50000, arrowLength: 40000 });
        expect(points).toBe(RIGHT_DEFAULT);
    });

    it('leftArrow：箭头朝左，肩部 x = dx1', async () => {
        const points = await arrowPoints('leftArrow', 240, 120, { adj1: 50000, adj2: 50000 });
        expect(points).toBe('240 30,60 30,60 0,0 60,60 120,60 90,240 90');
    });

    it('upArrow：主轴为高度，箭头长度 = min(w,h)*adj2/100000', async () => {
        // w=120,h=240 → ss=120；dy1 = 60，dx1 = w*adj1/200000 = 30
        const points = await arrowPoints('upArrow', 120, 240, { adj1: 50000, adj2: 50000 });
        expect(points).toBe('30 240,30 60,0 60,60 0,120 60,90 60,90 240');
    });

    it('downArrow：箭头朝下', async () => {
        const points = await arrowPoints('downArrow', 120, 240, { adj1: 50000, adj2: 50000 });
        expect(points).toBe('30 0,30 180,0 180,60 240,120 180,90 180,90 0');
    });

    it('leftRightArrow：两端各一段箭头', async () => {
        const points = await arrowPoints('leftRightArrow', 240, 120, { adj1: 50000, adj2: 50000 });
        expect(points).toBe('0 30,60 30,60 0,0 60,60 120,60 90,180 90,180 120,240 60,180 0,180 30');
    });

    it('upDownArrow：上下各一段箭头', async () => {
        // w=120,h=240 → ss=120；dy=60 → y1=60,y2=180；dx1 = w*adj1/200000 = 30
        const points = await arrowPoints('upDownArrow', 120, 240, { adj1: 50000, adj2: 50000 });
        expect(points).toBe('30 240,30 180,0 180,60 240,120 180,90 180,90 60,120 60,60 0,0 60,30 60');
    });
});

describe('箭头 avLst 写出（生成端）', () => {
    it('adjust 的 key 原样写为 a:gd@name', async () => {
        const data = await jsonToPptx({
            slides: [{
                elements: [{
                    type: 'shape', shapeType: 'rightArrow', x: 0, y: 0, width: 240, height: 120,
                    fill: { color: '#10b981' }, adjust: { adj1: 50000, adj2: 40000 }
                }]
            }]
        });
        const JSZip = (await import('jszip')).default;
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:gd name="adj1" fmla="val 50000"/>');
        expect(xml).toContain('<a:gd name="adj2" fmla="val 40000"/>');
    });
});
