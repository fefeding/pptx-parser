import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { pptxToStandard } from '../src/index';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const file = path.join(__dirname, '..', 'examples', 'Sample_12.pptx');

describe('Sample_12 slide 1 rendering data', () => {
  it('preserves text box fills, text strokes, shape text and curved connectors', async () => {
    const data = await pptxToStandard(fs.readFileSync(file));
    const slide = data.slides[0];
    expect(slide.elements.length).toBeGreaterThan(0);

    // "Introduction to" 文本框应携带蓝色背景填充
    const hasText = (e: any, t: string) => {
      const all = [e.text, ...(e.paragraphs || []).map((p: any) => (p.runs || []).map((r: any) => r.text).join(''))].filter(Boolean).join('');
      return all.includes(t);
    };
    const intro = slide.elements.find((e: any) => e.type === 'text' && hasText(e, 'Introduction to'));
    expect(intro).toBeDefined();
    expect(intro!.fill).toBeTruthy();
    const introFill = (intro as any).fill;
    expect(typeof introFill === 'string' ? introFill : introFill.color).toMatch(/^(5B9BD5|2683C6|1CADE4)$/);

    // "Amir" 文字应带描边
    const amir = slide.elements.find((e: any) => e.type === 'text' && hasText(e, 'Amir'));
    expect(amir).toBeDefined();
    const amirRun = (amir as any).paragraphs?.[0]?.runs?.[0];
    expect(amirRun?.outline).toBeDefined();
    expect(amirRun?.outline?.color).toMatch(/^[0-9A-F]{6}$/);
    expect(amirRun?.outline?.width).toBeGreaterThan(0);

    // 饼图形状上应保留 "Pie" 文字
    const pie = slide.elements.find((e: any) => e.type === 'text' && hasText(e, 'Pie'));
    expect(pie).toBeDefined();
    expect((pie as any).shapeType).toBe('pie');

    // 右侧曲线连接符应保留预设类型
    const cxn = slide.elements.find((e: any) => e.type === 'shape' && e.shapeType === 'curvedConnector3');
    expect(cxn).toBeDefined();
  });
});
