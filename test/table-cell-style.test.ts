import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx, pptxToHtml } from '../src/index.ts';

/** 内置表格样式 GUID */
const TABLE_GRID_ID = '{5940675A-B579-460E-94D1-54222C63F5DA}';
const ACCENT_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

/** 由 jsdom Uint8Array 安全转为可传给渲染/回读接口的 ArrayBuffer */
function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

/** 构造一个单页单表格的演示文稿 JSON */
function tableDoc(table: any) {
    return { slides: [{ elements: [table] }] };
}

/** 两行两列表格（可覆盖任意字段） */
function table(extra: any = {}) {
    return {
        type: 'table', x: 0, y: 0, width: 200, height: 100,
        colWidths: [100, 100], rowHeights: [50, 50],
        rows: [
            { cells: [{ text: 'A' }, { text: 'B' }] },
            { cells: [{ text: 'C' }, { text: 'D' }] }
        ],
        ...extra
    };
}

async function slideXml(data: Uint8Array): Promise<string> {
    return (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');
}

async function tableStylesXml(data: Uint8Array): Promise<string> {
    return (await JSZip.loadAsync(data)).file('ppt/tableStyles.xml').async('string');
}

async function renderFirstSlide(data: Uint8Array): Promise<string> {
    const res: any = await pptxToHtml(toArrayBuffer(data), { themeProcess: false });
    return res.slides[0].html;
}

/** 单格表格，便于断言单个 tcPr */
function oneCell(cell: any, tableExtra: any = {}) {
    return tableDoc(table({
        colWidths: [200], rowHeights: [100],
        rows: [{ cells: [cell] }],
        ...tableExtra
    }));
}

describe('T10 单元格内边距（生成端）', () => {
    it('inset → a:tcPr 上的 marL/marR/marT/marB（1px = 9525 EMU）', async () => {
        const data = await jsonToPptx(oneCell({ text: 'A', inset: { l: 20, r: 20, t: 10, b: 10 } }));

        expect(await slideXml(data)).toContain('<a:tcPr anchor="t" marL="190500" marR="190500" marT="95250" marB="95250"/>');
    });

    it('不再输出自造的 a:tableCellInsets 元素（非法 OOXML，PowerPoint/WPS 会整体忽略）', async () => {
        const data = await jsonToPptx(oneCell({ text: 'A', inset: { l: 5 } }));

        expect(await slideXml(data)).not.toContain('tableCellInsets');
    });

    it('未指定 inset 时不写 mar* 属性（沿用应用默认值）', async () => {
        const xml = await slideXml(await jsonToPptx(oneCell({ text: 'A' })));

        expect(xml).toContain('<a:tcPr anchor="t"');
        expect(xml).not.toContain('marL=');
    });
});

describe('T10 对角线边框（生成端）', () => {
    it('diagonal both → a:lnTlToBr 与 a:lnBlToTr 各自直接携带线属性与颜色', async () => {
        const data = await jsonToPptx(oneCell({ text: 'A', borders: { diagonal: 'both' } }));
        const xml = await slideXml(data);

        expect(xml).toContain('<a:lnTlToBr w="12700" cap="flat" cmpd="sng" algn="ctr">' +
            '<a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnTlToBr>');
        expect(xml).toContain('<a:lnBlToTr w="12700" cap="flat" cmpd="sng" algn="ctr">');
    });

    it('不再嵌套 a:ln（a:lnTlToBr 自身就是 CT_LineProperties）', async () => {
        const xml = await slideXml(await jsonToPptx(oneCell({ text: 'A', borders: { diagonal: 'tlBr' } })));

        expect(xml).not.toContain('<a:lnTlToBr><a:ln');
    });

    it('tlBr / blTr 单独指定时只输出对应一条', async () => {
        const tl = await slideXml(await jsonToPptx(oneCell({ text: 'A', borders: { diagonal: 'tlBr' } })));
        expect(tl).toContain('<a:lnTlToBr');
        expect(tl).not.toContain('<a:lnBlToTr');

        const bl = await slideXml(await jsonToPptx(oneCell({ text: 'A', borders: { diagonal: 'blTr' } })));
        expect(bl).toContain('<a:lnBlToTr');
        expect(bl).not.toContain('<a:lnTlToBr');
    });
});

describe('T10 表格样式定义（生成端 tableStyles.xml）', () => {
    it('引用的内置强调样式会写入等价定义', async () => {
        const styles = await tableStylesXml(await jsonToPptx(oneCell({ text: 'A' }, { tableStyleId: ACCENT_STYLE_ID })));

        expect(styles).toContain(`styleId="${ACCENT_STYLE_ID}"`);
        expect(styles).toContain('styleName="Medium Style 2 - Accent 1"');
        expect(styles).toContain('<a:schemeClr val="accent1"/>');   // 强调色底纹
        expect(styles).toContain('<a:schemeClr val="lt1"/>');       // 白色网格线
        expect(styles).toContain('<a:band1H>');                     // 交替行区域名（非 band1Horz）
        expect(styles).toContain(`def="${TABLE_GRID_ID}"`);
    });

    it('未指定 tableStyleId 时仍只有默认的 Table Grid 定义', async () => {
        const styles = await tableStylesXml(await jsonToPptx(oneCell({ text: 'A' })));

        expect(styles).toContain(`styleId="${TABLE_GRID_ID}"`);
        expect(styles).not.toContain(ACCENT_STYLE_ID);
    });

    it('未知 GUID 也补一份等价网格定义（否则 WPS/PowerPoint 退化成无样式无网格）', async () => {
        const unknown = '{2D1D2E6E-9063-44E1-9D16-7D619778F921}';
        const styles = await tableStylesXml(await jsonToPptx(oneCell({ text: 'A' }, { tableStyleId: unknown })));

        expect(styles).toContain(`styleId="${unknown}"`);
        expect(styles).toContain('<a:insideV>');    // Table Grid 等价定义带六向网格线
        expect(styles).toContain('<a:insideH>');
    });
});

describe('T10 单元格内边距（渲染端）', () => {
    it('未写 mar* 时用 ECMA-376 默认值：左右 0.1 英寸、上下 0.05 英寸', async () => {
        const html = await renderFirstSlide(await jsonToPptx(oneCell({ text: 'A' })));

        expect(html).toContain('padding:4.8px 9.6px 4.8px 9.6px;');
    });

    it('自定义 mar* 覆盖默认值', async () => {
        const html = await renderFirstSlide(await jsonToPptx(oneCell({ text: 'A', inset: { l: 20, r: 20, t: 10, b: 10 } })));

        expect(html).toContain('padding:10px 20px 10px 20px;');
    });
});

describe('T10 对角线边框（渲染端）', () => {
    it('对角线渲染为内联 SVG 背景，且 URL 内的单引号已编码（避免截断 style 属性）', async () => {
        const html = await renderFirstSlide(await jsonToPptx(oneCell({ text: 'A', borders: { diagonal: 'both' } })));

        expect(html).toContain('background-image:url("data:image/svg+xml,');
        expect(html).toContain('background-size:100% 100%;');
        const urls = [...html.matchAll(/url\("(data:image\/svg\+xml,[^"]*)"\)/g)].map((m) => m[1]);
        expect(urls.length).toBe(2);                     // both → 两条对角线
        for (const u of urls) {
            expect(u).not.toContain("'");                // 单引号必须编码为 %27
        }
    });

    it('无对角线时不产生 SVG 背景（无回归）', async () => {
        const html = await renderFirstSlide(await jsonToPptx(oneCell({ text: 'A' })));

        expect(html).not.toContain('data:image/svg+xml');
    });

    it('兼容旧产物把线属性嵌在 a:ln 里的写法', async () => {
        const data = await jsonToPptx(oneCell({ text: 'A' }));
        const zip = await JSZip.loadAsync(data);
        const xml = await zip.file('ppt/slides/slide1.xml').async('string');
        const patched = xml.replace(/<a:tcPr anchor="t"\/>/,
            '<a:tcPr anchor="t"><a:lnTlToBr><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:lnTlToBr></a:tcPr>');
        expect(patched).not.toBe(xml);
        zip.file('ppt/slides/slide1.xml', patched);
        const html = await renderFirstSlide(await zip.generateAsync({ type: 'uint8array' }));

        // 线宽 2pt、颜色 FF0000（URL 内引号与 # 均已编码）
        expect(html).toContain('%23FF0000');
        expect(html).toContain('stroke-width%3D%272%27');
    });
});
