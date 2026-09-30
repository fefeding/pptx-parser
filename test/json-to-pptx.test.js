import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { pptxToJson, pptxToStandard, jsonToPptx, editPptx, PPTXComposer } from '../src/index.ts';

// 1x1 红色 PNG 的 base64
const TINY_PNG_BASE64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

/**
 * 构建一个两页的测试演示文稿（覆盖文本/形状/图片/多段落/超链接）
 * @returns {PPTXComposer} Composer 实例
 */
function buildTestComposer() {
    const composer = new PPTXComposer();
    composer
        .title('测试演示文稿')
        .author('pptx-parser')
        .addSlide(slide => {
            slide.background('#ffffff');
            slide.addText(t => t
                .value('第一页标题')
                .x(100).y(80).width(600).height(60)
                .fontSize(28).color('#1e293b').bold());
            slide.addText(t => t
                .value('第一行\n第二行')
                .x(100).y(200).width(400).height(120)
                .fontSize(18).color('#64748b'));
            slide.addShape(s => s
                .shapeType('roundRect')
                .x(100).y(400).width(200).height(80)
                .fill({ color: '#4f46e5' })
                .line({ color: '#000000', width: 1 }));
            slide.addImage({
                data: `data:image/png;base64,${TINY_PNG_BASE64}`,
                x: 500, y: 400, width: 100, height: 100
            });
        })
        .addSlide(slide => {
            slide.addText(t => t
                .runs([
                    { text: '外部链接', options: { href: 'https://example.com', color: '#0563C1', underline: true } },
                    { text: ' 与 ', options: {} },
                    { text: '内部跳转', options: { href: '#1', color: '#0563C1' } }
                ])
                .x(100).y(100).width(500).height(50));
            slide.addShape({ shapeType: 'ellipse', x: 600, y: 300, width: 150, height: 150, fill: { color: 'rgb(237, 125, 49)' } });
        });
    return composer;
}

describe('PPTXComposer / jsonToPptx', () => {
    it('生成的 zip 包含完整的 OOXML 部件结构', async () => {
        const composer = buildTestComposer();
        const data = await composer.save();
        expect(data).toBeInstanceOf(Uint8Array);
        expect(data.byteLength).toBeGreaterThan(0);

        const zip = await JSZip.loadAsync(data);
        const requiredFiles = [
            '[Content_Types].xml',
            '_rels/.rels',
            'docProps/core.xml',
            'docProps/app.xml',
            'ppt/presentation.xml',
            'ppt/_rels/presentation.xml.rels',
            'ppt/slides/slide1.xml',
            'ppt/slides/_rels/slide1.xml.rels',
            'ppt/slides/slide2.xml',
            'ppt/slideMasters/slideMaster1.xml',
            'ppt/slideLayouts/slideLayout1.xml',
            'ppt/theme/theme1.xml',
            'ppt/media/image1.png'
        ];
        for (const name of requiredFiles) {
            expect(zip.file(name), `缺少部件 ${name}`).not.toBeNull();
        }
    });

    it('幻灯片 XML 包含正确的文本、形状与样式', async () => {
        const data = await jsonToPptx(buildTestComposer().toJSON());
        const zip = await JSZip.loadAsync(data);
        const slide1 = await zip.file('ppt/slides/slide1.xml').async('string');

        expect(slide1).toContain('第一页标题');
        expect(slide1).toContain('sz="2800"');           // 28pt
        expect(slide1).toContain('val="1E293B"');        // 颜色 HEX
        expect(slide1).toContain('prst="roundRect"');
        expect(slide1).toContain('val="4F46E5"');
        expect(slide1).toContain('<a:buNone/>');         // 无项目符号
        // 两个段落（\n 分段）
        expect(slide1.match(/<a:p>/g).length).toBeGreaterThanOrEqual(3); // 标题1 + 正文2
    });

    it('元数据写入 docProps/core.xml', async () => {
        const data = await buildTestComposer().save();
        const zip = await JSZip.loadAsync(data);
        const core = await zip.file('docProps/core.xml').async('string');
        expect(core).toContain('<dc:title>测试演示文稿</dc:title>');
        expect(core).toContain('<dc:creator>pptx-parser</dc:creator>');
    });

    it('超链接生成 slide 关系（外部 + 内部跳转）', async () => {
        const data = await buildTestComposer().save();
        const zip = await JSZip.loadAsync(data);
        const rels = await zip.file('ppt/slides/_rels/slide2.xml.rels').async('string');

        expect(rels).toContain('https://example.com');
        expect(rels).toContain('TargetMode="External"');
        expect(rels).toContain('slide1.xml');             // 内部跳转指向第 1 页
        const slide2 = await zip.file('ppt/slides/slide2.xml').async('string');
        expect(slide2).toContain('ppaction://hlinksldjump');
    });

    it('round-trip：生成文件可被 pptxToJson 解析', async () => {
        const data = await buildTestComposer().save();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const result = await pptxToJson(buffer);

        expect(result).not.toBeNull();
        expect(result.slides.length).toBe(2);
        expect(result.slideSize.width).toBeCloseTo(1280, 0);
        expect(result.slideSize.height).toBeCloseTo(720, 0);
        expect(result.metadata.title).toBe('测试演示文稿');
        expect(result.metadata.author).toBe('pptx-parser');
    });

    it('空 slides 数组抛出错误', async () => {
        await expect(jsonToPptx({ slides: [] })).rejects.toThrow(/至少需要一页幻灯片/);
    });
});

describe('editPptx', () => {
    async function makeTestFile() {
        const composer = new PPTXComposer();
        for (let i = 1; i <= 3; i++) {
            composer.addSlide(slide => {
                slide.addText(t => t.value(`第${i}页`).x(50).y(50).fontSize(24));
            });
        }
        return composer.save();
    }

    it('deleteSlide 减少页数并更新包结构', async () => {
        const data = await makeTestFile();
        const buffer = data.buffer.slice(data.byteOffset, data.byteLength + data.byteOffset);
        const editor = await editPptx(buffer);

        expect(await editor.getSlideCount()).toBe(3);
        await editor.deleteSlide(2);
        const newData = await editor.save();

        const newBuffer = newData.buffer.slice(newData.byteOffset, newData.byteOffset + newData.byteLength);
        const result = await pptxToJson(newBuffer);
        expect(result.slides.length).toBe(2);

        // 删除的 slide2.xml 部件应不存在
        const zip = await JSZip.loadAsync(newData);
        expect(zip.file('ppt/slides/slide2.xml')).toBeNull();
        const ct = await zip.file('[Content_Types].xml').async('string');
        expect(ct).not.toContain('/ppt/slides/slide2.xml');
    });

    it('moveSlide 重排幻灯片顺序', async () => {
        const data = await makeTestFile();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const editor = await editPptx(buffer);

        await editor.moveSlide(1, 3); // 第1页移到末尾 → 顺序变为 2,3,1
        const newData = await editor.save();
        const newBuffer = newData.buffer.slice(newData.byteOffset, newData.byteOffset + newData.byteLength);
        const result = await pptxToJson(newBuffer);

        expect(result.slides.length).toBe(3);
        // 验证顺序：第1个应是原第2页
        const firstSlideJson = JSON.stringify(result.slides[0]);
        expect(firstSlideJson).toContain('第2页');
    });

    it('addSlide 追加新页并可被解析', async () => {
        const data = await makeTestFile();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const editor = await editPptx(buffer);

        await editor.addSlide({
            elements: [
                { type: 'text', x: 10, y: 10, width: 300, height: 50, text: '追加的页面', fontSize: 20 },
                { type: 'image', data: `data:image/png;base64,${TINY_PNG_BASE64}`, x: 10, y: 100, width: 50, height: 50 }
            ]
        });
        const newData = await editor.save();
        const newBuffer = newData.buffer.slice(newData.byteOffset, newData.byteOffset + newData.byteLength);
        const result = await pptxToJson(newBuffer);

        expect(result.slides.length).toBe(4);
        const lastSlideJson = JSON.stringify(result.slides[3]);
        expect(lastSlideJson).toContain('追加的页面');

        // 新增媒体文件不应与已有冲突
        const zip = await JSZip.loadAsync(newData);
        expect(zip.file('ppt/media/image1.png')).not.toBeNull();
    });

    it('setMetadata 写回元数据', async () => {
        const data = await makeTestFile();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const editor = await editPptx(buffer);

        await editor.setMetadata({ title: '新标题', author: '新作者' });
        const newData = await editor.save();
        const newBuffer = newData.buffer.slice(newData.byteOffset, newData.byteOffset + newData.byteLength);
        const result = await pptxToJson(newBuffer);

        expect(result.metadata.title).toBe('新标题');
        expect(result.metadata.author).toBe('新作者');
    });

    it('getSlide 使用逻辑顺序（moveSlide 后仍按 sldIdLst 顺序读取）', async () => {
        const data = await makeTestFile();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const editor = await editPptx(buffer);

        await editor.moveSlide(1, 3); // 顺序变为 2,3,1
        const slideRoot = await editor.getSlide(1);
        const json = JSON.stringify(slideRoot);
        // 逻辑首张应为原第 2 页（而非文件名 slide1.xml 的旧内容）
        expect(json).toContain('第2页');
    });

    it('deleteSlide 禁止删除最后一页', async () => {
        const composer = new PPTXComposer();
        composer.addSlide(slide => {
            slide.addText(t => t.value('唯一页').x(10).y(10).fontSize(20));
        });
        const data = await composer.save();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const editor = await editPptx(buffer);

        expect(await editor.getSlideCount()).toBe(1);
        await expect(editor.deleteSlide(1)).rejects.toThrow(/至少需保留一页/);
    });

    it('deleteSlide / addSlide 同步更新 docProps/app.xml 的 Slides 计数', async () => {
        const data = await makeTestFile();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const editor = await editPptx(buffer);

        await editor.deleteSlide(2);
        let zip = await JSZip.loadAsync(await editor.save());
        expect(await zip.file('docProps/app.xml').async('string')).toContain('<Slides>2</Slides>');

        await editor.addSlide({ elements: [{ type: 'text', x: 0, y: 0, width: 10, height: 10, text: 'x' }] });
        zip = await JSZip.loadAsync(await editor.save());
        expect(await zip.file('docProps/app.xml').async('string')).toContain('<Slides>3</Slides>');
    });
});

describe('图表生成（原生 OOXML chart 部件）', () => {
    it('addChart 产出 chart 部件、graphicFrame 与 Content-Types 覆盖', async () => {
        const composer = new PPTXComposer();
        composer.addSlide(slide => {
            slide.addChart({
                chartType: 'barChart',
                title: '测试柱状图',
                x: 40, y: 40, width: 600, height: 360,
                categories: ['Q1', 'Q2', 'Q3'],
                series: [
                    { name: '销售额', values: [120, 200, 150] },
                    { name: '利润', values: [30, 55, 40] }
                ]
            });
        });
        const data = await composer.save();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const zip = await JSZip.loadAsync(buffer);

        // 图表部件存在且为合法 chartSpace
        const chartXml = await zip.file('ppt/charts/chart1.xml').async('string');
        expect(chartXml).toContain('<c:chartSpace');
        expect(chartXml).toContain('<c:barChart');
        expect(chartXml).toContain('<c:varyColors');
        // 系列名与数值落入数据缓存
        expect(chartXml).toContain('销售额');
        expect(chartXml).toContain('<c:v>120</c:v>');

        // 幻灯片含 graphicFrame 与 chart 关系
        const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
        expect(slideXml).toContain('<p:graphicFrame');
        const slideRels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
        expect(slideRels).toContain('relationships/chart');

        // Content-Types 含 chart 覆盖
        const ct = await zip.file('[Content_Types].xml').async('string');
        expect(ct).toContain('drawingml.chart+xml');
    });

    it('散点图使用 xVal/yVal 而非 cat/val', async () => {
        const composer = new PPTXComposer();
        composer.addSlide(slide => {
            slide.addChart({
                chartType: 'scatterChart',
                x: 0, y: 0, width: 400, height: 300,
                series: [{ name: '样本', x: [1, 2, 3], y: [2, 4, 1] }]
            });
        });
        const data = await composer.save();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const zip = await JSZip.loadAsync(buffer);
        const chartXml = await zip.file('ppt/charts/chart1.xml').async('string');
        expect(chartXml).toContain('<c:scatterChart');
        expect(chartXml).toContain('<c:xVal>');
        expect(chartXml).toContain('<c:yVal>');
        expect(chartXml).not.toContain('<c:cat>');
    });

    it('股票图生成 open/high/low/close 四引用与 hiLowLines/serLines', async () => {
        const composer = new PPTXComposer();
        composer.addSlide(slide => {
            slide.addChart({
                chartType: 'stockChart', x: 0, y: 0, width: 400, height: 300,
                series: [{ name: '股价',
                    open: [10, 11, 12], high: [15, 16, 17], low: [8, 9, 10], close: [13, 14, 15] }]
            });
        });
        const data = await composer.save();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const zip = await JSZip.loadAsync(buffer);
        const chartXml = await zip.file('ppt/charts/chart1.xml').async('string');
        expect(chartXml).toContain('<c:stockChart');
        expect(chartXml).toContain('<c:openVal>');
        expect(chartXml).toContain('<c:highVal>');
        expect(chartXml).toContain('<c:lowVal>');
        expect(chartXml).toContain('<c:closeVal>');
        expect(chartXml).toContain('<c:hiLowLines/>');
        expect(chartXml).toContain('<c:serLines/>');
    });

    it('雷达图含 radarStyle、曲面图含 bandFmts', async () => {
        const composer = new PPTXComposer();
        composer.addSlide(slide => {
            slide.addChart({ chartType: 'radarChart', x: 0, y: 0, width: 400, height: 300,
                series: [{ name: 'R', values: [1, 2, 3] }] });
        });
        const d1 = await composer.save();
        const z1 = await JSZip.loadAsync(d1.buffer.slice(d1.byteOffset, d1.byteOffset + d1.byteLength));
        expect(await z1.file('ppt/charts/chart1.xml').async('string')).toContain('<c:radarStyle');

        const composer2 = new PPTXComposer();
        composer2.addSlide(slide => {
            slide.addChart({ chartType: 'surfaceChart', x: 0, y: 0, width: 400, height: 300,
                series: [{ name: 'S', values: [1, 2, 3] }] });
        });
        const d2 = await composer2.save();
        const z2 = await JSZip.loadAsync(d2.buffer.slice(d2.byteOffset, d2.byteOffset + d2.byteLength));
        expect(await z2.file('ppt/charts/chart1.xml').async('string')).toContain('<c:bandFmts/>');
    });
});

describe('批注 comments (T18)', () => {
    it('生成 commentsN.xml、commentAuthors.xml 及其关系', async () => {
        const data = await jsonToPptx({
            slides: [
                { elements: [], comments: [
                    { author: 'Alice', text: '第一条批注', dt: '2026-01-01T00:00:00Z' },
                    { author: 'Bob', text: '第二条批注' }
                ] },
                { elements: [] }
            ]
        });
        const zip = await JSZip.loadAsync(data);

        // 评论部件：cmLst / p:cm / authorId / 正文
        const cm = await zip.file('ppt/comments/comments1.xml').async('string');
        expect(cm).toContain('<p:cmLst');
        expect(cm).toContain('<p:cm ');
        expect(cm).toContain('authorId="0"');
        expect(cm).toContain('authorId="1"');
        expect(cm).toContain('第一条批注');
        expect(cm).toContain('<p:pos');
        expect(cm).toContain('<p:text>');

        // 作者表：commentAuthors / 作者名
        const authors = await zip.file('ppt/commentAuthors.xml').async('string');
        expect(authors).toContain('<p:commentAuthors');
        expect(authors).toContain('name="Alice"');
        expect(authors).toContain('name="Bob"');

        // 幻灯片关系 → comments 部件
        const slideRels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
        expect(slideRels).toContain('relationships/comments');
        expect(slideRels).toContain('../comments/comments1.xml');

        // comments 部件关系 → commentAuthors
        const cmRels = await zip.file('ppt/comments/_rels/comments1.xml.rels').async('string');
        expect(cmRels).toContain('relationships/commentAuthors');

        // presentation 关系含 commentAuthors
        const presRels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
        expect(presRels).toContain('relationships/commentAuthors');

        // Content-Types 覆盖
        const ct = await zip.file('[Content_Types].xml').async('string');
        expect(ct).toContain('presentationml.comments+xml');
        expect(ct).toContain('presentationml.commentAuthors+xml');
    });

    it('无批注时不生成评论部件', async () => {
        const data = await jsonToPptx({ slides: [{ elements: [] }] });
        const zip = await JSZip.loadAsync(data);
        expect(zip.file('ppt/comments/comments1.xml')).toBeNull();
        expect(zip.file('ppt/commentAuthors.xml')).toBeNull();
    });
});

describe('组合 group (T4)', () => {
    it('生成 p:grpSp 并内嵌子元素（chOff/chExt 正确）', async () => {
        const data = await jsonToPptx({
            slides: [{
                elements: [{
                    type: 'group', x: 50, y: 50, width: 400, height: 300,
                    children: [
                        { type: 'text', x: 60, y: 60, width: 200, height: 40, text: '组内文本' },
                        { type: 'shape', shapeType: 'rect', x: 60, y: 120, width: 100, height: 50, fill: { color: '#ff0000' } }
                    ]
                }]
            }]
        });
        const zip = await JSZip.loadAsync(data);
        const slide = await zip.file('ppt/slides/slide1.xml').async('string');

        expect(slide).toContain('<p:grpSp>');
        expect(slide).toContain('<a:chOff');
        expect(slide).toContain('<a:chExt');
        // 子元素被递归内嵌
        expect(slide).toContain('组内文本');
        expect(slide).toContain('<p:sp>');
    });

    it('composer.addGroup 写入分组', async () => {
        const composer = new PPTXComposer();
        composer.addSlide(slide => {
            slide.addGroup(g => g
                .x(10).y(10).width(300).height(200)
                .children([
                    { type: 'text', x: 20, y: 20, width: 100, height: 30, text: 'A' },
                    { type: 'text', x: 20, y: 60, width: 100, height: 30, text: 'B' }
                ]));
        });
        const data = await composer.save();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const zip = await JSZip.loadAsync(buffer);
        const slide = await zip.file('ppt/slides/slide1.xml').async('string');
        expect(slide).toContain('<p:grpSp>');
        expect(slide).toContain('A');
        expect(slide).toContain('B');
    });

    it("'page' 坐标模式将子元素相对化（减 group 偏移），'local' 保持原值", async () => {
        const EMU = 914400 / 96; // pxToEmu 常量：1px = 9525 EMU
        const build = async (coords) => {
            const data = await jsonToPptx({ slides: [{ elements: [{
                type: 'group', x: 50, y: 50, width: 400, height: 300,
                childrenCoordinates: coords,
                children: [{ type: 'text', x: 110, y: 110, width: 100, height: 30, text: 'T' }]
            }] }] });
            const zip = await JSZip.loadAsync(data);
            const slide = await zip.file('ppt/slides/slide1.xml').async('string');
            return [...slide.matchAll(/<a:off x="(\d+)"/g)].map(m => m[1]);
        };
        const local = await build('local');
        const page = await build('page');
        // page 模式：子元素 110px 减 group 偏移 50 → 局部 60px (54866400)
        expect(page).toContain(String(60 * EMU));
        expect(page).not.toContain(String(110 * EMU));
        // local 模式：子元素保持 110px (100584000)
        expect(local).toContain(String(110 * EMU));
    });
});

describe('SmartArt 图示 diagram (T16)', () => {
    function buildDiagramComposer() {
        const composer = new PPTXComposer();
        composer.addSlide(slide => {
            slide.addDiagram(d => d
                .x(40).y(40).width(500).height(400)
                .diagramType('hierarchy')
                .nodes([
                    { text: '根', children: [
                        { text: '子A' },
                        { text: '子B', children: [ { text: '孙' } ] }
                    ] }
                ]));
        });
        return composer;
    }

    it('生成 p:graphicFrame + diagrams 四件套及关系', async () => {
        const data = await buildDiagramComposer().save();
        const zip = await JSZip.loadAsync(data);

        // 幻灯片含 graphicFrame 引用 diagramData
        const slide = await zip.file('ppt/slides/slide1.xml').async('string');
        expect(slide).toContain('<p:graphicFrame>');
        expect(slide).toContain('drawingml/2006/diagram');
        expect(slide).toContain('<dgm:rel');

        // 四个图示部件存在
        const dataXml = await zip.file('ppt/diagrams/data1.xml').async('string');
        expect(dataXml).toContain('<dsdgm:dataModel');
        expect(dataXml).toContain('根');
        expect(dataXml).toContain('子A');
        expect(dataXml).toContain('子B');
        expect(dataXml).toContain('孙');
        expect(dataXml).toContain('<dsdgm:pt');
        expect(dataXml).toContain('<dsdgm:cxn');
        expect(await zip.file('ppt/diagrams/layout1.xml').async('string')).toContain('<dgm:layoutDef');
        expect(await zip.file('ppt/diagrams/colors1.xml').async('string')).toContain('<dgm:colorsDef');
        expect(await zip.file('ppt/diagrams/quickStyle1.xml').async('string')).toContain('<dgm:quickStyleDef');

        // data 部件关系 → layout/colors/quickStyle
        const dataRels = await zip.file('ppt/diagrams/_rels/data1.xml.rels').async('string');
        expect(dataRels).toContain('relationships/diagramLayout');
        expect(dataRels).toContain('relationships/diagramColors');
        expect(dataRels).toContain('relationships/diagramQuickStyle');

        // 幻灯片关系 → diagramData
        const slideRels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
        expect(slideRels).toContain('relationships/diagramData');
        expect(slideRels).toContain('../diagrams/data1.xml');

        // Content-Types 覆盖
        const ct = await zip.file('[Content_Types].xml').async('string');
        expect(ct).toContain('drawingml.diagramData+xml');
        expect(ct).toContain('drawingml.diagramLayout+xml');
        expect(ct).toContain('drawingml.diagramColors+xml');
        expect(ct).toContain('drawingml.diagramQuickStyle+xml');
    });

    it('round-trip：图示可被 pptxToStandard 解析出文本', async () => {
        const data = await buildDiagramComposer().save();
        const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        const result = await pptxToStandard(buffer);
        const diagramEl = result.slides[0].elements.find((e) => e.type === 'diagram');
        expect(diagramEl).toBeTruthy();
        expect(JSON.stringify(diagramEl.texts || diagramEl)).toContain('根');
    });

    it('editPptx.addSlide 支持图示并写入部件', async () => {
        const composer = new PPTXComposer();
        composer.addSlide(slide => slide.addText(t => t.value('原页').x(10).y(10).fontSize(20)));
        const base = await composer.save();
        const buffer = base.buffer.slice(base.byteOffset, base.byteOffset + base.byteLength);
        const editor = await editPptx(buffer);
        await editor.addSlide({
            elements: [{
                type: 'diagram', x: 0, y: 0, width: 300, height: 200,
                diagramType: 'list',
                nodes: [{ text: '列表项' }]
            }]
        });
        const newData = await editor.save();
        const newBuffer = newData.buffer.slice(newData.byteOffset, newData.byteOffset + newData.byteLength);
        const zip = await JSZip.loadAsync(newBuffer);
        expect(await zip.file('ppt/diagrams/data1.xml').async('string')).toContain('列表项');
    });
});
