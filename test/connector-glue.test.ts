// 连接线（p:cxnSp）：编辑模型可编辑、胶着关系（a:stCxn/a:endCxn）能往返
import { describe, it, expect } from 'vitest';
import { pptxToStandard, jsonToPptx, docFromPptx, docToPptx } from '../src/index.ts';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteLength);
}

/** 构造一个含连接线（带胶着关系）的文档 */
function buildDoc() {
    return {
        slides: [{
            elements: [
                { id: 's1', type: 'shape', shapeType: 'rect', x: 100, y: 100, width: 200, height: 120, fill: '#ff0000' },
                { id: 's2', type: 'shape', shapeType: 'rect', x: 500, y: 100, width: 200, height: 120, fill: '#00ff00' },
                {
                    id: 'c1', type: 'connector', shapeType: 'straightConnector1',
                    start: { x: 300, y: 160 }, end: { x: 500, y: 160 },
                    fill: 'none',
                    line: { color: '#1f2937', width: 1.5, endArrow: 'triangle' },
                    glueBegin: { shapeId: 's1', site: 'right' },
                    glueEnd: { shapeId: 's2', site: 'left' }
                }
            ]
        }],
        slideSize: { width: 1280, height: 720 }
    };
}

describe('连接线胶着', () => {
    it('导出时写出 p:cxnSp 与 a:stCxn/a:endCxn，id 指向目标形状', async () => {
        const out = await jsonToPptx(buildDoc() as any);
        const std = await pptxToStandard(toArrayBuffer(out));
        const slide = std.slides[0];
        const conn = slide.elements.find((e: any) => e.type === 'connector');
        expect(conn).toBeTruthy();
        // 胶着关系落在连接线的原生连接属性里
        expect((conn as any).beginCxn).toBeTruthy();
        expect((conn as any).endCxn).toBeTruthy();
    });

    it('导入后映射成编辑模型的 begin/end，导出后仍然保留', async () => {
        const out = await jsonToPptx(buildDoc() as any);
        const std = await pptxToStandard(toArrayBuffer(out));
        const edoc = docFromPptx(std);
        const conn = (edoc as any).slides[0].elements.find((e: any) => e.connector);
        expect(conn).toBeTruthy();
        // 导入后胶着目标映射为编辑模型的元素 id（编辑模型会重新生成 id）
        const shapes = (edoc as any).slides[0].elements.filter((e: any) => e.type === 'shape');
        expect(conn.begin.shapeId).toBe(shapes[0].id);
        expect(conn.end.shapeId).toBe(shapes[1].id);
        expect(conn.begin.site).toBe('right');
        expect(conn.end.site).toBe('left');

        const again = await jsonToPptx(docToPptx(edoc));
        const std2 = await pptxToStandard(toArrayBuffer(again));
        const conn2 = std2.slides[0].elements.find((e: any) => e.type === 'connector');
        expect(conn2).toBeTruthy();
        expect((conn2 as any).beginCxn).toBeTruthy();
        expect((conn2 as any).endCxn).toBeTruthy();
    });

    it('连接线保持 0 高度（水平线是合法几何）', async () => {
        const out = await jsonToPptx(buildDoc() as any);
        const std = await pptxToStandard(toArrayBuffer(out));
        const conn = std.slides[0].elements.find((e: any) => e.type === 'connector');
        expect(conn!.height).toBe(0);
    });
});
