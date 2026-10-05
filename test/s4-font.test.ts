import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { pptxToHtml } from '../src/index';
const __dirname = path.dirname(fileURLToPath(import.meta.url));
describe('s4 preview font', () => {
  it('check font-family strings', async () => {
    (globalThis as any).URL.createObjectURL = (globalThis as any).URL.createObjectURL || (() => 'blob:mock');
    const html: any = await pptxToHtml(fs.readFileSync(path.join(__dirname, '..', 'examples/Sample_12.pptx')));
    const s = typeof html.slides[4] === 'string' ? html.slides[4] : JSON.stringify(html.slides[4]);
    const i = s.indexOf('Need more info');
    console.log('CTX:', s.slice(Math.max(0, i - 500), i + 40).match(/font-family:[^;']*/g)?.slice(-3));
    expect(true).toBe(true);
  });
});
