import { describe, it, expect } from 'vitest';
import { jsonToPptx, pptxToJson } from '../src/index.ts';

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
    return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

describe('中文编号 round-trip', () => {
    it('chineseCounting 格式被保留，渲染输出包含「、」分隔符', async () => {
        const doc = {
            slides: [{
                elements: [{
                    type: 'text' as const, x: 0, y: 0, width: 400, height: 200,
                    fontSize: 20, color: '#2D3748',
                    paragraphs: [
                        { text: '编号第一项', bullet: { type: 'number' as const, fmt: 'chineseCounting' as const, start: 1 } },
                        { text: '编号第二项', bullet: { type: 'number' as const, fmt: 'chineseCounting' as const, start: 1 } }
                    ]
                }]
            }]
        };
        const data = await jsonToPptx(doc);
        const result = await pptxToJson(toArrayBuffer(data), { mode: 'semantic' });
        const parsed = result.document.slides[0].elements[0] as any;
        const fmt = parsed.paragraphs[0].bullet.fmt;
        expect(fmt).toBe('chineseCounting');
    });
});
