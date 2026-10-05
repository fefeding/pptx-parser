import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToJson } from '../src/index.ts';
import { ACCENT_TABLE_STYLE_ID } from '../src/serializer/templates.ts';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

describe('表格样式解析到单元格', () => {
    it('Medium Style 2 - Accent 1 解析后首行/交替行有填充色', async () => {
        const table = {
            type: 'table' as const, x: 0, y: 0, width: 600, height: 200,
            colWidths: [200, 200, 200], rowHeights: [50, 50],
            tableStyleId: ACCENT_TABLE_STYLE_ID,
            rows: [
                { cells: [{ text: 'A' }, { text: 'B' }, { text: 'C' }] },
                { cells: [{ text: 'D' }, { text: 'E' }, { text: 'F' }] }
            ]
        };
        const data = await jsonToPptx({ slides: [{ elements: [table] }] });
        const result = await pptxToJson(toArrayBuffer(data), { mode: 'semantic' });
        const parsed = result.document.slides[0].elements[0] as any;
        expect(parsed.rows[0].cells[0].fill).toBeTruthy();
        expect(parsed.rows[1].cells[0].fill).toBeTruthy();
        expect(parsed.rows[0].cells[0].color).toBeTruthy();
        // 交替行（band1H：lumMod 40% + lumOff 60%）必须比首行更浅
        expect(parsed.rows[1].cells[0].fill).not.toBe(parsed.rows[0].cells[0].fill);
        const lum = (hex: string) => {
            const c = hex.replace('#', '');
            const [r, g, b] = [0, 2, 4].map((i) => parseInt(c.slice(i, i + 2), 16));
            return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
        };
        expect(lum(parsed.rows[1].cells[0].fill)).toBeGreaterThan(lum(parsed.rows[0].cells[0].fill));
        // 首行加粗（firstRow tcTxStyle b="on"）
        expect(parsed.rows[0].cells[0].bold).toBe(true);
    });
});
