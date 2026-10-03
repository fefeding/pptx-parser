import { describe, it, expect } from 'vitest';
import { jsonToPptx } from '../src/index';

/**
 * 多母版 / 多版式 / 占位符 / 节 / 语义主题
 *
 * 对应 PPTX_GEN_GAP_PLAN 的 T17（标注完成但实际未实现）与 T5（仅整串 XML 覆盖）。
 */

async function load(pres: any) {
    const buf: any = await jsonToPptx(pres);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buf);
    const text = async (name: string) => await zip.file(name)!.async('text');
    return { zip, text };
}

const masters = [
    {
        name: 'Main Master',
        placeholders: [
            { type: 'title', x: 48, y: 40, width: 900, height: 96, fontSize: 40, prompt: '单击此处添加标题' }
        ],
        layouts: [
            {
                name: 'Title Only',
                placeholders: [
                    { type: 'title', idx: 0, x: 48, y: 40, width: 900, height: 96 }
                ]
            },
            {
                name: 'Title and Content',
                placeholders: [
                    { type: 'title', idx: 0, x: 48, y: 40, width: 900, height: 96 },
                    { type: 'body', idx: 1, x: 48, y: 160, width: 900, height: 400 },
                    { type: 'ftr', x: 48, y: 640, width: 400, height: 32 },
                    { type: 'sldNum', x: 800, y: 640, width: 100, height: 32 },
                    { type: 'dt', x: 600, y: 640, width: 180, height: 32 }
                ]
            }
        ]
    }
];

describe('多母版 / 版式 / 占位符（T17）', () => {
    it('按 masters 生成多个版式部件，母版引用全部版式', async () => {
        const { text } = await load({
            masters,
            slides: [
                { layout: 0, elements: [] },
                { layout: 1, elements: [] }
            ]
        });

        const master = await text('ppt/slideMasters/slideMaster1.xml');
        // 母版应引用两个版式
        expect(master.match(/<p:sldLayoutId /g)?.length).toBe(2);
        // 母版占位符
        expect(master).toContain('p:ph type="title"');

        const layout1 = await text('ppt/slideLayouts/slideLayout1.xml');
        const layout2 = await text('ppt/slideLayouts/slideLayout2.xml');
        expect(layout1).toContain('name="Title Only"');
        expect(layout2).toContain('name="Title and Content"');

        // 页脚 / 页码 / 日期占位符（此前完全无法生成）
        expect(layout2).toContain('type="ftr"');
        expect(layout2).toContain('type="sldNum"');
        expect(layout2).toContain('type="dt"');
        // 占位符索引
        expect(layout2).toContain('idx="1"');
    });

    it('幻灯片按 slide.layout 关联对应版式', async () => {
        const { text } = await load({
            masters,
            slides: [
                { layout: 0, elements: [] },
                { layout: 1, elements: [] }
            ]
        });

        const rels1 = await text('ppt/slides/_rels/slide1.xml.rels');
        const rels2 = await text('ppt/slides/_rels/slide2.xml.rels');
        expect(rels1).toContain('slideLayout1.xml');
        expect(rels2).toContain('slideLayout2.xml');
        expect(rels1).not.toContain('slideLayout2.xml');
    });

    it('版式关系指回所属母版，母版关系含版式与主题', async () => {
        const { text } = await load({ masters, slides: [{ layout: 0, elements: [] }] });

        const layoutRels = await text('ppt/slideLayouts/_rels/slideLayout1.xml.rels');
        expect(layoutRels).toContain('slideMaster1.xml');

        const masterRels = await text('ppt/slideMasters/_rels/slideMaster1.xml.rels');
        expect(masterRels).toContain('slideLayout1.xml');
        expect(masterRels).toContain('slideLayout2.xml');
        expect(masterRels).toContain('theme1.xml');
    });

    it('Content-Types 为每个版式补 Override（否则 PowerPoint 判为损坏包）', async () => {
        const { text } = await load({ masters, slides: [{ layout: 0, elements: [] }] });
        const ct = await text('[Content_Types].xml');
        expect(ct).toContain('/ppt/slideLayouts/slideLayout1.xml');
        expect(ct).toContain('/ppt/slideLayouts/slideLayout2.xml');
        expect(ct).toContain('/ppt/slideMasters/slideMaster1.xml');
    });

    it('多母版场景：生成两个母版且各挂自身版式', async () => {
        const { text } = await load({
            masters: [
                { name: 'M1', layouts: [{ name: 'L1' }] },
                { name: 'M2', layouts: [{ name: 'L2' }, { name: 'L3' }] }
            ],
            slides: [{ layout: 0, elements: [] }, { layout: 2, elements: [] }]
        });

        const pres = await text('ppt/presentation.xml');
        expect(pres.match(/<p:sldMasterId /g)?.length).toBe(2);

        const m2 = await text('ppt/slideMasters/slideMaster2.xml');
        expect(m2.match(/<p:sldLayoutId /g)?.length).toBe(2);
        // 第二母版的第二版式（全局第 3 个）
        expect(await text('ppt/slideLayouts/slideLayout3.xml')).toContain('name="L3"');
        // 版式 3 应指回母版 2
        expect(await text('ppt/slideLayouts/_rels/slideLayout3.xml.rels')).toContain('slideMaster2.xml');
    });

    it('未使用 masters 时保持单母版单版式（不回归）', async () => {
        const { text } = await load({ slides: [{ elements: [] }] });
        expect(await text('ppt/slideLayouts/slideLayout1.xml')).toContain('p:sldLayout');
        const pres = await text('ppt/presentation.xml');
        expect(pres.match(/<p:sldMasterId /g)?.length).toBe(1);
    });
});

describe('文档节（sections）', () => {
    it('写出 p:sectionLst 并按索引引用幻灯片', async () => {
        const { text } = await load({
            slides: [{ elements: [] }, { elements: [] }, { elements: [] }],
            sections: [
                { name: '第一章', slides: [0, 1] },
                { name: '第二章', slides: [2] }
            ]
        });

        const pres = await text('ppt/presentation.xml');
        expect(pres).toContain('p:sectionLst');
        expect(pres).toContain('第一章');
        expect(pres).toContain('第二章');
        // sldId 的数值 id 从 256 起
        expect(pres).toContain('<p:sldId id="256"/>');
        expect(pres).toContain('<p:sldId id="258"/>');
    });

    it('未声明 sections 时不输出 sectionLst', async () => {
        const { text } = await load({ slides: [{ elements: [] }] });
        expect(await text('ppt/presentation.xml')).not.toContain('p:sectionLst');
    });
});

describe('语义级主题（T5）', () => {
    it('colors / fonts 覆盖写入 theme1.xml', async () => {
        const { text } = await load({
            theme: {
                name: 'Brand Theme',
                colors: { accent1: '#FF0000', dk1: '112233' },
                fonts: { major: { latin: 'Arial', ea: '微软雅黑' }, minor: { latin: 'Calibri' } }
            },
            slides: [{ elements: [] }]
        });

        const theme = await text('ppt/theme/theme1.xml');
        expect(theme).toContain('name="Brand Theme"');
        // 自定义色槽（#FF0000 → FF0000）
        expect(theme).toContain('val="FF0000"');
        expect(theme).toContain('val="112233"');
        // 字体
        expect(theme).toContain('typeface="Arial"');
        expect(theme).toContain('typeface="微软雅黑"');
        // 未指定的色槽保留默认值（accent2 默认 ED7D31）
        expect(theme).toContain('val="ED7D31"');
        // 指定的 accent1 已覆盖，不应再出现默认值
        expect(theme).not.toContain('val="4472C4"');
    });

    it('整串 XML 主题覆盖仍生效（兼容旧用法）', async () => {
        const custom = '<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Raw"><a:themeElements/></a:theme>';
        const { text } = await load({ theme: custom, slides: [{ elements: [] }] });
        expect(await text('ppt/theme/theme1.xml')).toContain('name="Raw"');
    });
});
