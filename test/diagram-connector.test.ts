import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx } from '../src/serializer/json-to-pptx';
import pptxToHtml from '../src/index';

/** 构造一个层次结构 SmartArt 演示文稿 JSON */
function hierarchyPres(): any {
    return {
        slides: [{
            elements: [{
                type: 'diagram',
                diagramType: 'hierarchy',
                x: 0, y: 0, width: 400, height: 300,
                nodes: [{
                    text: '根',
                    children: [
                        { text: '子A' },
                        { text: '子B', children: [{ text: '孙' }] }
                    ]
                }]
            }]
        }]
    };
}

describe('SmartArt 连线标准规范（cxnSp 连接器）', () => {
    it('序列化端：hierarchy 生成标准 dsp:cxnSp 直线连接器，无矩形连线、无箭头', async () => {
        const buf = await jsonToPptx(hierarchyPres());
        const zip = await JSZip.loadAsync(buf);
        const drawing = await zip.file('ppt/diagrams/drawing1.xml')!.async('string');

        // 标准连接器元素与几何
        expect(drawing).toContain('cxnSp');
        expect(drawing).toContain('straightConnector1');
        // 线条颜色/宽度由 a:ln 表达（无填充）
        expect(drawing).toContain('<a:ln');
        // 无箭头端点（符合标准层次结构 SmartArt）
        expect(drawing).not.toContain('headEnd');
        expect(drawing).not.toContain('tailEnd');
        // 旧的矩形连线（prst="rect"）不应再出现
        expect(drawing).not.toContain('prst="rect"');
        // 节点仍用 roundRect 形状
        expect(drawing).toContain('roundRect');
        // 连线数量：根→子A、根→子B 各 3 段；子B→孙 同轴对齐省略横段为 2 段；共 8 段 cxnSp
        const cxnCount = (drawing.match(/cxnSp>/g) || []).length;
        expect(cxnCount).toBeGreaterThanOrEqual(8);
    });

    it('解析端：round-trip 渲染出连接线且无箭头 marker', async () => {
        const buf = await jsonToPptx(hierarchyPres());
        const result = (await pptxToHtml(buf as any)) as any;
        const html = (result.slides && result.slides[0] && result.slides[0].html) || '';

        // cxnSp 的 straightConnector1 渲染为 <line .../>
        expect(html).toContain('<line');
        // 不应出现箭头 marker 引用
        expect(html).not.toContain('marker-start');
        expect(html).not.toContain('marker-end');
    });
});
