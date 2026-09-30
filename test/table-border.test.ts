import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { jsonToPptx, pptxToStandard } from '../src/index.ts';

/** 1pt = 12700 EMU */
const PT = 12700;

/** 由 jsdom Uint8Array 安全转为可传给 pptxToStandard 的 ArrayBuffer */
function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

/** 构造一个单页单表格的演示文稿 JSON */
function tableDoc(table: any) {
    return { slides: [{ elements: [table] }] };
}

describe('T1 表格单元格边框线（生成端）', () => {
    it('单元格 border 生成四边 a:lnL/a:lnR/a:lnT/a:lnB 且属性正确', async () => {
        const data = await jsonToPptx(tableDoc({
            type: 'table', x: 0, y: 0, width: 100, height: 50,
            colWidths: [100], rowHeights: [50],
            rows: [{ cells: [{ text: 'A', border: { color: '#FF0000', width: 2 } }] }]
        }));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).toContain('<a:lnL');
        expect(xml).toContain('<a:lnR');
        expect(xml).toContain('<a:lnT');
        expect(xml).toContain('<a:lnB');
        // 2pt = 25400 EMU，颜色大写 HEX
        expect(xml).toContain(`w="${2 * PT}"`);
        expect(xml).toContain('val="FF0000"');
    });

    it('表格级 border 应用到未显式声明的单元格', async () => {
        const data = await jsonToPptx(tableDoc({
            type: 'table', x: 0, y: 0, width: 200, height: 100,
            colWidths: [100, 100], rowHeights: [50, 50],
            border: { color: '#000000', width: 1 },
            rows: [
                { cells: [{ text: 'A' }, { text: 'B' }] },
                { cells: [{ text: 'C' }, { text: 'D' }] }
            ]
        }));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        // 表格级默认 → 4 个单元格 × 4 边 = 16 个 ln* 节点
        const count = (xml.match(/<a:ln[LRBT]/g) || []).length;
        expect(count).toBe(16);
        expect(xml).toContain(`w="${PT}"`);
        expect(xml).toContain('val="000000"');
    });

    it('borders 分边覆盖：top:"none" 不生成 a:lnT，其他边仍生成', async () => {
        const data = await jsonToPptx(tableDoc({
            type: 'table', x: 0, y: 0, width: 100, height: 50,
            colWidths: [100], rowHeights: [50],
            border: { color: '#000000', width: 1 },
            rows: [{ cells: [{ text: 'A', borders: { top: 'none' } }] }]
        }));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        // 单单元格：应含 L/R/B 三条边，无 T
        expect((xml.match(/<a:lnL/g) || []).length).toBe(1);
        expect((xml.match(/<a:lnR/g) || []).length).toBe(1);
        expect((xml.match(/<a:lnB/g) || []).length).toBe(1);
        expect((xml.match(/<a:lnT/g) || []).length).toBe(0);
    });

    it('未声明任何边框时（兼容现状）不生成 ln* 节点', async () => {
        const data = await jsonToPptx(tableDoc({
            type: 'table', x: 0, y: 0, width: 200, height: 100,
            colWidths: [100, 100], rowHeights: [50, 50],
            rows: [
                { cells: [{ text: 'A' }, { text: 'B' }] },
                { cells: [{ text: 'C' }, { text: 'D' }] }
            ]
        }));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        expect(xml).not.toContain('<a:lnL');
        expect(xml).not.toContain('<a:lnR');
        expect(xml).not.toContain('<a:lnT');
        expect(xml).not.toContain('<a:lnB');
    });

    it('边框节点顺序位于填充之前（tcPr 内）', async () => {
        const data = await jsonToPptx(tableDoc({
            type: 'table', x: 0, y: 0, width: 100, height: 50,
            colWidths: [100], rowHeights: [50],
            rows: [{ cells: [{ text: 'A', fill: '#CCCCCC', border: { color: '#000000', width: 1 } }] }]
        }));
        const xml = await (await JSZip.loadAsync(data)).file('ppt/slides/slide1.xml').async('string');

        const lnIdx = xml.indexOf('<a:lnL');
        const fillIdx = xml.indexOf('<a:solidFill', xml.indexOf('<a:tcPr'));
        expect(lnIdx).toBeGreaterThan(0);
        expect(fillIdx).toBeGreaterThan(lnIdx); // ln 在 fill 之前
    });
});

describe('T1 表格单元格边框线（round-trip 回读）', () => {
    it('pptxToStandard 能回读单元格边框', async () => {
        const data = await jsonToPptx(tableDoc({
            type: 'table', x: 0, y: 0, width: 200, height: 100,
            colWidths: [100, 100], rowHeights: [50, 50],
            border: { color: '#000000', width: 1 },
            rows: [
                { cells: [{ text: 'A' }, { text: 'B' }] },
                { cells: [{ text: 'C', borders: { top: 'none', left: { color: '#FF0000', width: 2 } } }, { text: 'D' }] }
            ]
        }));
        const doc = await pptxToStandard(toArrayBuffer(data));

        const table: any = doc.slides[0].elements.find((e: any) => e.type === 'table');
        expect(table).toBeTruthy();

        // 表格级默认应用到 A 单元格
        const aBorders = table.rows[0].cells[0].borders;
        expect(aBorders).toBeTruthy();
        expect(aBorders.left).toEqual({ color: '000000', width: 1 });
        expect(aBorders.top).toEqual({ color: '000000', width: 1 });

        // C 单元格分边覆盖
        const cBorders = table.rows[1].cells[0].borders;
        expect(cBorders.top).toBeUndefined();           // 'none' 不回读
        expect(cBorders.left).toEqual({ color: 'FF0000', width: 2 });
        expect(cBorders.right).toEqual({ color: '000000', width: 1 });
    });

    it('再序列化回读结果不产生回归（边框保持一致）', async () => {
        const data = await jsonToPptx(tableDoc({
            type: 'table', x: 0, y: 0, width: 100, height: 50,
            colWidths: [100], rowHeights: [50],
            border: { color: '#000000', width: 1 },
            rows: [{ cells: [{ text: 'A' }] }]
        }));
        const doc = await pptxToStandard(toArrayBuffer(data));
        const reData = await jsonToPptx(doc);
        const xml = await (await JSZip.loadAsync(reData)).file('ppt/slides/slide1.xml').async('string');
        expect((xml.match(/<a:ln[LRBT]/g) || []).length).toBe(4);
    });
});
