// @vitest-environment jsdom
// 端到端校验：真实样例经 编辑器导入(docFromPptx) → 编辑器导出(docToPptx) → jsonToPptx
// 后，主题/母版/版式/SmartArt 配色应与源文件一致（导出保真回归）。
import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { jsonToPptx, pptxToJson, pptxToStandard } from '../src/index.ts';
import { docFromPptx, docToPptx } from '../examples/editor/src/model.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SAMPLE = path.resolve(__dirname, '../examples/Sample_12.pptx');

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
  return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

describe('编辑器往返导出保真（Sample_12）', () => {
  it('保留多主题/母版/版式，SmartArt 逐块配色不因退化为单一主题而错位', async () => {    const buf = fs.readFileSync(SAMPLE);

    // 源文件：SmartArt 绑定 theme2（青色系）
    const src = await pptxToStandard(toArrayBuffer(buf));
    const srcDiagram = src.slides[10].elements.find((e: any) => e.type === 'diagram');
    expect(src.masters?.length).toBeGreaterThan(1);
    expect(srcDiagram.shapes[0].fill).toMatch(/^#1CADE4$/i);

    // 编辑器导入 → 导出
    const sem = await pptxToJson(toArrayBuffer(buf), { mode: 'semantic' });
    const edoc = docFromPptx(sem.document);
    expect((edoc as any).masters?.length).toBe(src.masters!.length);
    expect((edoc as any).themeXmls?.length).toBe(src.themeXmls!.length);

    const outDoc = docToPptx(edoc as any);
    expect(outDoc.masters!.length).toBe(src.masters!.length);
    expect(outDoc.slides[10].layout).toBe(src.slides[10].layout);

    // 再解析导出结果
    const back = await pptxToStandard(toArrayBuffer(await jsonToPptx(outDoc as any)));
    const backDiagram = back.slides[10].elements.find((e: any) => e.type === 'diagram');
    expect(backDiagram.shapes[0].fill).toBe(srcDiagram.shapes[0].fill);
  }, 60000);
});
