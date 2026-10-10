import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToHtml } from '../src/index.ts';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

async function render(el: any): Promise<string> {
    const data = await jsonToPptx({ slides: [{ elements: [el] }] });
    const res: any = await pptxToHtml(toArrayBuffer(data), { themeProcess: false });
    return res.slides[0].html;
}

function connector(line: any, shapeType = 'straightConnector1') {
    return { type: 'connector', shapeType, x: 0, y: 0, width: 400, height: 0, line };
}

describe('线条端点箭头渲染（a:headEnd / a:tailEnd → SVG marker）', () => {
    it('粗线箭头尺寸受限，不会随线宽无限膨胀（旧实现为 5×线宽）', async () => {
        const html = await render(connector({ color: '#c00', width: 100, endArrow: 'triangle' }));
        const m = /<marker[^>]*markerWidth='([\d.]+)'[^>]*markerHeight='([\d.]+)'/.exec(html);
        expect(m).not.toBeNull();
        expect(parseFloat(m![1])).toBeLessThanOrEqual(140);
        expect(parseFloat(m![2])).toBeLessThanOrEqual(120);
    });

    it('箭头始终明显粗于线条（markerHeight ≥ min(3×线宽, 120)）', async () => {
        for (const w of [4, 20, 40]) {
            const html = await render(connector({ color: '#c00', width: w, endArrow: 'triangle' }));
            const m = /markerHeight='([\d.]+)'/.exec(html);
            expect(m).not.toBeNull();
            expect(parseFloat(m![1])).toBeGreaterThanOrEqual(Math.min(w * 3, 120) - 0.01);
        }
    });

    it('细线箭头不小于最小尺寸（避免小到看不见）', async () => {
        const html = await render(connector({ color: '#c00', width: 1, endArrow: 'triangle' }));
        const m = /<marker[^>]*markerWidth='([\d.]+)'[^>]*markerHeight='([\d.]+)'/.exec(html);
        expect(parseFloat(m![1])).toBeGreaterThanOrEqual(10);
        expect(parseFloat(m![2])).toBeGreaterThanOrEqual(8);
    });

    it('head/tail 类型不同时分别生成 marker 并正确引用', async () => {
        const html = await render(connector({ color: '#c00', width: 6, startArrow: 'diamond', endArrow: 'triangle' }));
        expect(html).toMatch(/id='markerHead_[^']*'/);
        expect(html).toMatch(/id='markerTail_[^']*'/);
        expect(html).toMatch(/marker-start='url\(#markerHead_/);
        expect(html).toMatch(/marker-end='url\(#markerTail_/);
    });

    it('支持 triangle/stealth/arrow/diamond/oval 多种箭头类型', async () => {
        for (const t of ['triangle', 'stealth', 'arrow', 'diamond', 'oval']) {
            const html = await render(connector({ color: '#c00', width: 6, endArrow: t }));
            expect(html).toMatch(/marker-end='url\(#markerTail_/);
            expect(html).toMatch(/<marker[^>]*>\s*<[^>]+>\s*<\/marker>/);
        }
    });

    it('无箭头时（none/缺省）不生成 marker', async () => {
        const html = await render(connector({ color: '#c00', width: 4 }));
        expect(html).not.toContain('<marker');
        expect(html).not.toContain('marker-end=');
    });
});

describe('line 形状零尺寸渲染（水平/垂直线）', () => {
    it('shapeType:line 水平线（height=0）也能渲染，SVG 容器高度不为 0', async () => {
        const html = await render({
            type: 'shape', shapeType: 'line', x: 0, y: 0, width: 400, height: 0,
            line: { color: '#090', width: 2, endArrow: 'triangle' }
        });
        expect(html).toContain('<line');
        const svg = /<svg[^>]*>/.exec(html)?.[0] ?? '';
        const hMatch = /height:([\d.]+)px/.exec(svg);
        expect(hMatch).not.toBeNull();
        expect(parseFloat(hMatch![1])).toBeGreaterThan(0);
    });
});
