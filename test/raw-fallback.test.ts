import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx } from '../src/serializer/json-to-pptx';
import { pptxToStandard } from '../src/index';
import type { PptxDocument } from '../src/types/pptx-document';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** SmartArt（图示）graphicFrame，引用 rId7~rId10 */
const DIAGRAM_FRAME = `<p:graphicFrame>\
<p:nvGraphicFramePr><p:cNvPr id="9" name="SmartArt 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>\
<p:xfrm><a:off x="914400" y="2743200"/><a:ext cx="5486400" cy="2743200"/></p:xfrm>\
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">\
<dgm:rel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId7" r:lo="rId8" r:qs="rId9" r:cs="rId10"/>\
</a:graphicData></a:graphic></p:graphicFrame>`;

/** 图示数据部件：含两个节点文本 */
const DATA_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <dgm:ptLst>
    <dgm:pt modelId="{A}" type="node"><dgm:t><a:p><a:r><a:rPr lang="zh-CN"/><a:t>需求</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="{B}" type="node"><dgm:t><a:p><a:r><a:rPr lang="zh-CN"/><a:t>上线</a:t></a:r></a:p></dgm:t></dgm:pt>
  </dgm:ptLst>
  <dgm:cxnLst/>
</dgm:dataModel>`;

const PART_FILES: { path: string; rel: string; content: string }[] = [
    { path: 'ppt/diagrams/data1.xml', rel: 'diagramData', content: DATA_XML },
    { path: 'ppt/diagrams/layout1.xml', rel: 'diagramLayout', content: '<?xml version="1.0"?><dgm:layout xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:clrMap/></dgm:layout>' },
    { path: 'ppt/diagrams/quickStyle1.xml', rel: 'diagramQuickStyle', content: '<?xml version="1.0"?><dgm:quickStyleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>' },
    { path: 'ppt/diagrams/colors1.xml', rel: 'diagramColors', content: '<?xml version="1.0"?><dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>' }
];

/** 在 jsonToPptx 产出的基准 PPTX 中注入一个 SmartArt，构造含图示的输入文件 */
async function buildPptxWithSmartArt(): Promise<Uint8Array> {
    const base: PptxDocument = {
        version: '1.0',
        slideSize: { width: 1280, height: 720 },
        slides: [{ elements: [{ type: 'text', x: 50, y: 40, width: 400, height: 60, text: '含 SmartArt' }] }]
    };
    const bytes = (await jsonToPptx(base, { outputType: 'uint8array' })) as Uint8Array;
    const zip = await JSZip.loadAsync(bytes);

    // 1) slide1.xml 注入 graphicFrame
    const slideText = await zip.file('ppt/slides/slide1.xml')!.async('string');
    zip.file('ppt/slides/slide1.xml', slideText.replace('</p:spTree>', `${DIAGRAM_FRAME}</p:spTree>`));

    // 2) slide rels 追加图示关系（rId7~rId10）
    const relsText = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
    const newRels = PART_FILES
        .map((p, i) => `<Relationship Id="rId${7 + i}" Type="${REL}/${p.rel}" Target="../${p.path.replace('ppt/', '')}"/>`)
        .join('');
    zip.file('ppt/slides/_rels/slide1.xml.rels', relsText.replace('</Relationships>', `${newRels}</Relationships>`));

    // 3) 写入图示部件
    for (const p of PART_FILES) zip.file(p.path, p.content);

    // 4) Content-Types 覆盖
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    const overrides = PART_FILES
        .map(p => `<Override PartName="/${p.path}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.${p.rel}+xml"/>`)
        .join('');
    zip.file('[Content_Types].xml', ct.replace('</Types>', `${overrides}</Types>`));

    return zip.generateAsync({ type: 'uint8array' });
}

describe('__raw 无损回退序列化（SmartArt）', () => {
    it('解析端为图示元素附带关系与部件依赖', async () => {
        const pptx = await buildPptxWithSmartArt();
        const doc = await pptxToStandard(pptx);

        const diagram: any = doc.slides[0].elements.find((e: any) => e.type === 'diagram');
        expect(diagram).toBeTruthy();
        expect(diagram.texts).toEqual(['需求', '上线']);

        // __raw 载荷：标签 + 节点 + 依赖
        expect(diagram.__raw.tag).toBe('p:graphicFrame');
        expect(Object.keys(diagram.__raw.rels).sort()).toEqual(['rId10', 'rId7', 'rId8', 'rId9']);
        expect(diagram.__raw.rels.rId7.type).toBe(`${REL}/diagramData`);
        expect(diagram.__raw.parts.map((p: any) => p.path).sort())
            .toEqual(PART_FILES.map(p => p.path).sort());
        // 文本型部件以字符串内联，保证回写自包含
        const dataPart = diagram.__raw.parts.find((p: any) => p.path === 'ppt/diagrams/data1.xml');
        expect(dataPart.content).toContain('需求');
    });

    it('生成端回退写回图示：部件落盘、关系重映射、文本再次可解析', async () => {
        const pptx = await buildPptxWithSmartArt();
        const doc = await pptxToStandard(pptx);

        const out = (await jsonToPptx(doc, { outputType: 'uint8array' })) as Uint8Array;
        const zip = await JSZip.loadAsync(out);

        // 图示部件被回写
        for (const p of PART_FILES) {
            expect(zip.file(p.path)).toBeTruthy();
        }

        // slide1.xml 含图示节点，且 r:dm 指向的关系在 rels 中存在
        const slideText = await zip.file('ppt/slides/slide1.xml')!.async('string');
        expect(slideText).toContain('dgm:rel');
        const dmMatch = slideText.match(/r:dm="(rId\d+)"/);
        expect(dmMatch).toBeTruthy();

        const relsText = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
        // 关系目标为相对路径 ../diagrams/data1.xml（而非绝对 ppt/ 路径）
        expect(relsText).toContain(`Id="${dmMatch![1]}"`);
        expect(relsText).toContain('Target="../diagrams/data1.xml"');
        expect(relsText).toContain(`${REL}/diagramData`);

        // Content-Types 已声明
        const ct = await zip.file('[Content_Types].xml')!.async('string');
        expect(ct).toContain('/ppt/diagrams/data1.xml');

        // 再次解析：图示文本仍可还原（闭环）
        const back = await pptxToStandard(out);
        const diagram: any = back.slides[0].elements.find((e: any) => e.type === 'diagram');
        expect(diagram).toBeTruthy();
        expect(diagram.texts).toEqual(['需求', '上线']);
    });

    it('多页同名原始部件去重：各页关系指向各自部件', async () => {
        // 两页各自携带一个 path 相同的 __raw 部件（模拟多页各自的 diagrams/data1.xml）
        const mkPart = (tag: string) => ({
            tag: 'p:graphicFrame',
            node: { 'dgm:rel': { attrs: { 'r:dm': 'rId7' } } },
            rels: {
                rId7: {
                    type: `${REL}/diagramData`,
                    target: 'ppt/diagrams/data1.xml'
                }
            },
            parts: [{
                path: 'ppt/diagrams/data1.xml',
                content: `<data>${tag}</data>`,
                contentType: 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml'
            }]
        });
        const doc: PptxDocument = {
            version: '1.0',
            slideSize: { width: 1280, height: 720 },
            slides: [1, 2].map(n => ({
                elements: [{
                    type: 'diagram', x: 0, y: 0, width: 100, height: 100,
                    __raw: mkPart(`PAGE${n}`), rawFallback: true
                }]
            })) as unknown as PptxDocument['slides']
        };

        const out = (await jsonToPptx(doc, { outputType: 'uint8array' })) as Uint8Array;
        const zip = await JSZip.loadAsync(out);

        // 两个部件都被写入（后者重命名，不覆盖前者）
        expect(await zip.file('ppt/diagrams/data1.xml')!.async('string')).toBe('<data>PAGE1</data>');
        expect(await zip.file('ppt/diagrams/data1__2.xml')!.async('string')).toBe('<data>PAGE2</data>');

        // 各页关系分别指向自己的部件（去重后必须同步改写关系目标）
        const rels1 = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
        const rels2 = await zip.file('ppt/slides/_rels/slide2.xml.rels')!.async('string');
        expect(rels1).toContain('Target="../diagrams/data1.xml"');
        expect(rels2).toContain('Target="../diagrams/data1__2.xml"');
        expect(rels2).not.toContain('Target="../diagrams/data1.xml"');

        // 重命名后的部件也需在 Content-Types 中声明
        const ct = await zip.file('[Content_Types].xml')!.async('string');
        expect(ct).toContain('/ppt/diagrams/data1__2.xml');
    });

    it('跨页媒体与图表编号连续，不相互覆盖', async () => {
        const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8AAAAMBAQAY3Y2wAAAAAElFTkSuQmCC';
        const doc: PptxDocument = {
            version: '1.0',
            slideSize: { width: 1280, height: 720 },
            slides: [1, 2].map(n => ({
                elements: [
                    { type: 'image', x: 0, y: 0, width: 50, height: 50, data: PNG, extension: 'png' },
                    {
                        type: 'chart', chartType: 'barChart', x: 60, y: 0, width: 200, height: 120,
                        categories: ['A', 'B'],
                        series: [{ name: `S${n}`, values: [1, 2] }]
                    }
                ]
            })) as unknown as PptxDocument['slides']
        };

        const out = (await jsonToPptx(doc, { outputType: 'uint8array' })) as Uint8Array;
        const zip = await JSZip.loadAsync(out);

        const media = Object.keys(zip.files).filter(p => p.startsWith('ppt/media/') && !zip.files[p].dir);
        const charts = Object.keys(zip.files).filter(p => p.startsWith('ppt/charts/') && !zip.files[p].dir);
        expect(media.sort()).toEqual(['ppt/media/image1.png', 'ppt/media/image2.png']);
        expect(charts.sort()).toEqual(['ppt/charts/chart1.xml', 'ppt/charts/chart2.xml']);

        // 两页的图表关系不得指向同一部件
        const rels1 = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
        const rels2 = await zip.file('ppt/slides/_rels/slide2.xml.rels')!.async('string');
        expect(rels1).toContain('../charts/chart1.xml');
        expect(rels2).toContain('../charts/chart2.xml');
        expect(rels1).not.toContain('../charts/chart2.xml');

        // 图表内容随页区分（未被后页覆盖）
        expect(await zip.file('ppt/charts/chart1.xml')!.async('string')).toContain('S1');
        expect(await zip.file('ppt/charts/chart2.xml')!.async('string')).toContain('S2');
    });

    it("rawDeps:'all' 时为语义类型也附带依赖，默认不带", async () => {
        // 构造含图片（p:pic 带 r:embed 引用）的输入文件
        const base: PptxDocument = {
            version: '1.0',
            slideSize: { width: 1280, height: 720 },
            slides: [{
                elements: [{
                    type: 'image', x: 0, y: 0, width: 50, height: 50,
                    data: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8AAAAMBAQAY3Y2wAAAAAElFTkSuQmCC',
                    extension: 'png'
                }]
            }] as unknown as PptxDocument['slides']
        };
        const pptx = (await jsonToPptx(base, { outputType: 'uint8array' })) as Uint8Array;

        // 默认：语义类型（image）只有 { tag, node }，不带依赖
        const auto = await pptxToStandard(pptx);
        const imgAuto: any = auto.slides[0].elements.find((e: any) => e.type === 'image');
        expect(imgAuto.__raw.tag).toBe('p:pic');
        expect(imgAuto.__raw.rels).toBeUndefined();
        expect(imgAuto.__raw.parts).toBeUndefined();

        // 'all'：语义类型也附带 r:embed 关系与媒体部件，rawFallback 才可自包含回写
        const all = await pptxToStandard(pptx, { rawDeps: 'all' });
        const imgAll: any = all.slides[0].elements.find((e: any) => e.type === 'image');
        expect(Object.keys(imgAll.__raw.rels).length).toBeGreaterThan(0);
        const rel = Object.values(imgAll.__raw.rels)[0] as any;
        expect(rel.type).toBe(`${REL}/image`);
        expect(imgAll.__raw.parts.some((p: any) => p.media === true && p.base64)).toBe(true);
    });

    it('rawFallback:true 时已支持的类型也按原始节点回写', async () => {
        const doc: PptxDocument = {
            version: '1.0',
            slideSize: { width: 1280, height: 720 },
            slides: [{
                elements: [{
                    type: 'text', x: 0, y: 0, width: 100, height: 40, text: '语义文本',
                    rawFallback: true,
                    __raw: {
                        tag: 'p:sp',
                        node: {
                            attrs: { id: 'x' },
                            'p:nvSpPr': { 'p:cNvPr': { attrs: { id: 5, name: 'RawSp' } } },
                            'p:txBody': { 'a:p': { 'a:r': { 'a:t': '原始文本' } } }
                        }
                    }
                } as any]
            }]
        };
        const out = (await jsonToPptx(doc, { outputType: 'uint8array' })) as Uint8Array;
        const zip = await JSZip.loadAsync(out);
        const slideText = await zip.file('ppt/slides/slide1.xml')!.async('string');

        // 输出的是原始节点内容（RawSp / 原始文本），而非语义重建的文本
        expect(slideText).toContain('RawSp');
        expect(slideText).toContain('原始文本');
        expect(slideText).not.toContain('语义文本');
    });
});
