// 回归：预览端（pptxToHtml）对「未显式指定字色」的 run 必须输出解析后的颜色。
//
// 背景：run 常省略 a:rPr/a:solidFill，颜色按 OOXML 语义继承
// （run → 本形状 a:lstStyle → p:style/a:fontRef → 版式/母版 lstStyle → 主题 objectDefaults/tx1）。
// 解析链全部落空时，曾经**不输出任何 color**，于是文字继承宿主页面的环境色：
// 在 VSCode 深色 UI（body { color: var(--text) } = #e6e6e6）里，白底幻灯片上的黑字变成浅灰。
// 现改为回退到主题默认文字色（tx1，经母版 clrMap 映射）。
import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { jsonToPptx, pptxToHtml } from '../src/index';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function slideHtml(slide: any): string {
  return typeof slide === 'string' ? slide : String((slide && slide.html) || '');
}

/** 预览 HTML 里的文本 span（属性用单引号） */
function textSpans(html: string): string[] {
  return [...html.matchAll(/<span[^>]*class='text-block[^']*'[^>]*>/g)].map((m) => m[0]);
}

describe('预览端字色继承（不依赖宿主环境色）', () => {
  it('无显式字色的 run 回退到主题 tx1，而不是不输出颜色', async () => {
    (globalThis as any).URL.createObjectURL = (globalThis as any).URL.createObjectURL || (() => 'blob:mock');
    // 自包含装置：一个文本框 run，完全没有 color
    const doc: any = {
      title: 't',
      slideSize: { width: 960, height: 540 },
      theme: {
        id: 'imported', name: 'Office 主题', bg: '#FFFFFF', panel: '#E7E6E6',
        text: '#000000', title: '#44546A', accent: '#5B9BD5',
        accents: ['#5B9BD5', '#ED7D31', '#A5A5A5', '#FFC000', '#4472C4', '#70AD47']
      },
      slides: [{
        elements: [{
          id: 'sp1', type: 'text', x: 100, y: 100, width: 600, height: 200,
          paragraphs: [{ runs: [{ text: '无显式字色', fontSize: 18 }] }]
        }]
      }]
    };

    const out = await jsonToPptx(doc, { outputType: 'uint8array' });
    const html: any = await pptxToHtml(Buffer.from(out as Uint8Array));
    const spans = textSpans(slideHtml(html.slides[0]));

    expect(spans.length).toBeGreaterThan(0);
    for (const s of spans) {
      // 关键断言：必须带显式颜色（缺失即会继承宿主环境色）
      expect(s).toMatch(/color\s*:/);
      expect(s).toMatch(/color:\s*#000000/i);
    }
  }, 60000);

  it('样例文件：所有预览文本 span 都带显式字色', async () => {
    (globalThis as any).URL.createObjectURL = (globalThis as any).URL.createObjectURL || (() => 'blob:mock');
    const samples = ['Sample_12.pptx', '企业微信应用介绍.pptx']
      .map((n) => path.resolve(__dirname, '../examples', n))
      .filter((p) => fs.existsSync(p));
    expect(samples.length).toBeGreaterThan(0);

    for (const p of samples) {
      const html: any = await pptxToHtml(fs.readFileSync(p));
      let total = 0;
      const missing: string[] = [];
      html.slides.forEach((s: any, i: number) => {
        for (const span of textSpans(slideHtml(s))) {
          total++;
          if (!/color\s*:/.test(span)) missing.push(`slide${i + 1}: ${span.slice(0, 100)}`);
        }
      });
      expect(total).toBeGreaterThan(0);
      expect(missing).toEqual([]);
    }
  }, 180000);
});
