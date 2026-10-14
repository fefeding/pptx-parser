import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { pptxToStandard, jsonToPptx } from '../src/index.js';
import { relayoutDiagram, inferDiagramType, extractDiagramNodes } from '../src/core/diagram-relayout.js';

const SAMPLE = '/Users/jiamao/project/github/pptx-parser/examples/test-sample.pptx';

async function toArrayBuffer(path: string): Promise<ArrayBuffer> {
    const buf = readFileSync(path);
    return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

async function findSampleDiagram(): Promise<{ doc: any; el: any }> {
    const doc = await pptxToStandard(await toArrayBuffer(SAMPLE));
    for (const slide of doc.slides) {
        const el = (slide.elements || []).find((e: any) => e.type === 'diagram');
        if (el) return { doc, el };
    }
    throw new Error('test-sample.pptx 中未找到 diagram 元素');
}

describe('SmartArt 重布局引擎（P2-9）', () => {
    it('extractDiagramNodes：从 __raw 数据部件提取节点树（兼容 dsdgm: 前缀与 str@val 文本）', async () => {
        const { el } = await findSampleDiagram();
        const nodes = extractDiagramNodes(el);
        expect(nodes).toBeTruthy();
        expect(nodes!.length).toBe(1);
        const root = nodes![0];
        expect(root.text).toBe('根');
        expect((root.children || []).map((c) => c.text)).toEqual(['子A', '子B']);
        const childB = root.children![1];
        expect(childB.children?.length).toBe(1);
        expect(childB.children![0].text).toBe('孙');
    });

    it('inferDiagramType：layoutDef cat@type 优先，其次几何启发，兜底 list', async () => {
        const { el } = await findSampleDiagram();
        // 样例 layoutDef：cat type="hier" + alg hierChild
        expect(inferDiagramType(el)).toBe('hierarchy');
        // cat@type 启发
        const fake: any = {
            type: 'diagram',
            __raw: { parts: [{ path: 'ppt/diagrams/layout1.xml', content: '<dgm:layoutDef><dgm:catLst><dgm:cat type="process" priority="1"/></dgm:catLst></dgm:layoutDef>' }] }
        };
        expect(inferDiagramType(fake)).toBe('process');
        // 形状几何启发：垂直排列 → list；水平排列 → process；环形分布 → cycle
        const v: any = { type: 'diagram', shapes: [0, 1, 2].map((i) => ({ x: 10, y: i * 50, width: 100, height: 30 })) };
        expect(inferDiagramType(v)).toBe('list');
        const h: any = { type: 'diagram', shapes: [0, 1, 2].map((i) => ({ x: i * 60, y: 10, width: 50, height: 30 })) };
        expect(inferDiagramType(h)).toBe('process');
        const ring: any = {
            type: 'diagram',
            shapes: [0, 90, 180, 270].map((deg) => ({
                x: 100 + 60 * Math.cos((deg * Math.PI) / 180) - 15,
                y: 100 + 60 * Math.sin((deg * Math.PI) / 180) - 10,
                width: 30,
                height: 20
            }))
        };
        expect(inferDiagramType(ring)).toBe('cycle');
        expect(inferDiagramType({ type: 'diagram' } as any)).toBe('list');
    });

    it('relayoutDiagram：hierarchy 重排（新增节点 + 样式模板沿用 + 坐标入框 + 标记生成路径）', async () => {
        const { el } = await findSampleDiagram();
        const origFills = ((el.shapes || []) as any[])
            .filter((s: any) => !s.connector)
            .map((s: any) => String(s.fill).toUpperCase());
        const nodes = extractDiagramNodes(el)!;
        nodes[0].children!.push({ text: '子C' });
        relayoutDiagram(el, { nodes });
        expect(el.relayout).toBe(true);
        expect(el.layoutType).toBe('hierarchy');
        expect(el.texts).toContain('子C');
        expect(el.nodes!.length).toBe(1);
        expect(el.nodes![0].children!.length).toBe(3);
        const shapes = el.shapes!;
        const nodeShapes = shapes.filter((s: any) => !s.connector);
        const conns = shapes.filter((s: any) => s.connector);
        // 5 个节点框：根 + 子A + 子B + 孙 + 子C
        expect(nodeShapes.length).toBe(5);
        // 肘形连接：根 3 子 = 竖线+汇流线+3 竖线；子B 1 子 = 竖线+竖线
        expect(conns.length).toBeGreaterThanOrEqual(7);
        for (const s of shapes) {
            expect(s.x).toBeGreaterThanOrEqual(-1);
            expect(s.y).toBeGreaterThanOrEqual(-1);
            expect(s.x + s.width).toBeLessThanOrEqual(el.width + 2);
            expect(s.y + s.height).toBeLessThanOrEqual(el.height + 2);
        }
        // 样式模板：新节点填充沿用原节点色板
        expect(origFills.length).toBeGreaterThan(0);
        for (const s of nodeShapes) expect(origFills).toContain(String(s.fill).toUpperCase());
        // 新增节点形状存在且带文本
        const added = nodeShapes.find((s: any) => s.text === '子C');
        expect(added).toBeTruthy();
        expect(added.fontSize).toBeGreaterThan(0);
    });

    it('relayoutDiagram：类别覆盖（matrix/process/cycle/venn/pyramid/list）', () => {
        const base = (texts: string[]) => ({ type: 'diagram' as const, x: 0, y: 0, width: 400, height: 300, texts });
        const relayout = (el: any, type: any) => {
            relayoutDiagram(el, { type });
            return el;
        };

        const m = relayout(base(['A', 'B', 'C', 'D', 'E']), 'matrix');
        expect(m.shapes.filter((s: any) => !s.connector).length).toBe(5);
        expect(m.shapes.some((s: any) => s.connector)).toBe(false);
        expect(m.shapes.every((s: any) => s.prst === 'rect')).toBe(true);

        const p = relayout(base(['A', 'B', 'C', 'D']), 'process');
        const pNodes = p.shapes.filter((s: any) => !s.connector);
        const pConns = p.shapes.filter((s: any) => s.connector);
        expect(pNodes.length).toBe(4);
        expect(pConns.length).toBe(3);
        expect(pConns.every((c: any) => c.endArrow === 'triangle')).toBe(true);
        for (let i = 1; i < pNodes.length; i++) expect(pNodes[i].x).toBeGreaterThan(pNodes[i - 1].x);

        const cy = relayout(base(['A', 'B', 'C', 'D']), 'cycle');
        expect(cy.shapes.filter((s: any) => !s.connector).length).toBe(4);
        expect(cy.shapes.filter((s: any) => s.connector).length).toBe(4);

        const vn = relayout(base(['A', 'B', 'C']), 'venn');
        expect(vn.shapes.length).toBe(3);
        expect(vn.shapes.every((s: any) => s.prst === 'ellipse')).toBe(true);

        const py = relayout(base(['A', 'B', 'C']), 'pyramid');
        expect(py.shapes.length).toBe(3);
        expect(py.shapes.every((s: any) => s.prst === 'trapezoid')).toBe(true);
        expect(py.shapes[2].width).toBeGreaterThan(py.shapes[0].width);

        const li = relayout(base(['A', 'B', 'C']), 'list');
        expect(li.shapes.filter((s: any) => !s.connector).length).toBe(3);
        expect(li.shapes.filter((s: any) => s.connector).length).toBe(2);
    });

    it('extractDiagramNodes：destOrd 从 0 起合法（不因 falsy 回退），乱序 cxnLst 仍按 destOrd 排序', () => {
        const pt = (id: string, text: string) =>
            `<dgm:pt modelId="${id}" type="node"><dgm:t><a:p><a:r><a:t>${text}</a:t></a:r></a:p></dgm:t></dgm:pt>`;
        const dataXml =
            `<dgm:dataModel xmlns:dgm="x"><dgm:ptLst>` +
            `<dgm:pt modelId="1" type="doc"/>` + pt('2', '根') + pt('3', '甲') + pt('4', '乙') +
            `</dgm:ptLst><dgm:cxnLst>` +
            // cxnLst 故意乱序：乙（destOrd=1）列在甲（destOrd=0）之前
            `<dgm:cxn type="parOf" srcId="1" destId="2" srcOrd="0" destOrd="0"/>` +
            `<dgm:cxn type="parOf" srcId="2" destId="4" srcOrd="0" destOrd="1"/>` +
            `<dgm:cxn type="parOf" srcId="2" destId="3" srcOrd="0" destOrd="0"/>` +
            `</dgm:cxnLst></dgm:dataModel>`;
        const el: any = { type: 'diagram', __raw: { parts: [{ path: 'ppt/diagrams/data1.xml', content: dataXml }] } };
        const nodes = extractDiagramNodes(el);
        expect(nodes).toBeTruthy();
        expect(nodes!.length).toBe(1);
        expect((nodes![0].children || []).map((c) => c.text)).toEqual(['甲', '乙']);
    });

    it('cycle 斜向连线 flip 编码：反斜线段由 flipH/flipV 表达（与编辑器 lineEndpoints 同语义）', () => {
        const el: any = { type: 'diagram', x: 0, y: 0, width: 400, height: 300, texts: ['A', 'B', 'C', 'D'] };
        relayoutDiagram(el, { type: 'cycle' });
        const conns = (el.shapes || []).filter((s: any) => s.connector);
        expect(conns.length).toBe(4);
        const pat = (c: any) => `${c.flipH ? 'H' : ''}${c.flipV ? 'V' : ''}`;
        expect(conns.map(pat).sort()).toEqual(['', 'H', 'HV', 'V']);
    });

    it('process 箭头全链路：生成 drawingN.xml 含 a:tailEnd，重解析回读 endArrow', async () => {
        const { doc, el } = await findSampleDiagram();
        relayoutDiagram(el, { type: 'process' });
        const conns = (el.shapes || []).filter((s: any) => s.connector);
        expect(conns.length).toBeGreaterThan(0);
        expect(conns.every((c: any) => c.endArrow === 'triangle')).toBe(true);

        const buf = await jsonToPptx(doc);
        const genZip = await JSZip.loadAsync(buf);
        const drawingName = Object.keys(genZip.files).find((n) => /ppt\/diagrams\/drawing\d+\.xml$/.test(n))!;
        const drawingXml = await genZip.file(drawingName)!.async('string');
        expect(drawingXml).toContain('<a:tailEnd type="triangle"');

        const doc2 = await pptxToStandard(buf);
        let el2: any = null;
        for (const slide of doc2.slides) {
            const f = (slide.elements || []).find((e: any) => e.type === 'diagram');
            if (f) { el2 = f; break; }
        }
        const conns2 = ((el2?.shapes || []) as any[]).filter((s: any) => s.connector);
        expect(conns2.length).toBe(conns.length);
        expect(conns2.every((c: any) => c.endArrow === 'triangle')).toBe(true);
    });

    it('全链路 round-trip：relayout 后导出重生成 data/drawing，layout/colors/quickStyle 原件保真，重解析反映新节点', async () => {
        const { doc, el } = await findSampleDiagram();
        const nodes = extractDiagramNodes(el)!;
        nodes[0].children!.push({ text: '子C' });
        relayoutDiagram(el, { nodes });

        const buf = await jsonToPptx(doc);

        // 部件保真：生成的 layout/colors/quickStyle 与样例包原件逐字节一致（复用 __raw）
        const origZip = await JSZip.loadAsync(await toArrayBuffer(SAMPLE));
        const genZip = await JSZip.loadAsync(buf);
        for (const kind of ['layout', 'colors', 'quickStyle'] as const) {
            const genName = Object.keys(genZip.files).find((n) => new RegExp(`^ppt/diagrams/${kind}\\d+\\.xml$`).test(n));
            expect(genName).toBeTruthy();
            const orig = await origZip.file(`ppt/diagrams/${kind}1.xml`)!.async('string');
            const gen = await genZip.file(genName!)!.async('string');
            expect(gen).toBe(orig);
        }
        // dataN.xml 重生成：包含新节点文本
        const dataName = Object.keys(genZip.files).find((n) => /ppt\/diagrams\/data\d+\.xml$/.test(n));
        expect(dataName).toBeTruthy();
        const dataXml = await genZip.file(dataName!)!.async('string');
        expect(dataXml).toContain('子C');
        // drawingN.xml 重生成：重布局缓存形状含新节点
        const drawingName = Object.keys(genZip.files).find((n) => /ppt\/diagrams\/drawing\d+\.xml$/.test(n));
        expect(drawingName).toBeTruthy();
        const drawingXml = await genZip.file(drawingName!)!.async('string');
        expect(drawingXml).toContain('子C');

        // 重解析：texts/shapes 反映新节点；relayout 为运行期标记，不落盘
        const doc2 = await pptxToStandard(buf);
        let el2: any = null;
        for (const slide of doc2.slides) {
            const f = (slide.elements || []).find((e: any) => e.type === 'diagram');
            if (f) { el2 = f; break; }
        }
        expect(el2).toBeTruthy();
        expect(el2.texts).toContain('子C');
        expect((el2.shapes || []).filter((s: any) => !s.connector).length).toBe(5);
        expect(el2.relayout).toBeUndefined();
    });
});
