/**
 * 背景图片解析测试
 */

import { describe, it, expect } from 'vitest';
import { parsePptx } from 'pptx-parser';
import { readFileSync } from 'fs';
import { join } from 'path';

describe('Background Image Parsing', () => {
  it('should parse background image from slide', async () => {
    // 注意：需要提供一个包含背景图片的 PPTX 文件
    // 这里只是测试结构，实际测试需要真实文件

    const pptxBuffer = readFileSync(join(__dirname, '../fixtures/test-with-bg.pptx'));
    const result = await parsePptx(pptxBuffer, {
      parseImages: true
    });

    // 检查是否有幻灯片
    expect(result.slides.length).toBeGreaterThan(0);

    // 检查背景是否被正确解析
    const slideWithBgImage = result.slides.find(slide => {
      const bg = slide.background;
      if (typeof bg === 'string') return false;
      return bg?.type === 'image' && bg?.value?.startsWith('data:image/');
    });

    console.log('Background image slide:', slideWithBgImage);
    expect(slideWithBgImage).toBeDefined();

    // 验证背景图片包含 base64 数据
    if (slideWithBgImage && typeof slideWithBgImage.background !== 'string') {
      const bg = slideWithBgImage.background as any;
      expect(bg.relId).toBeDefined();
      expect(bg.value).toMatch(/^data:image\/[a-z]+;base64,/);
    }
  });

  it('should parse background color fallback', async () => {
    const pptxBuffer = readFileSync(join(__dirname, '../fixtures/test-simple.pptx'));
    const result = await parsePptx(pptxBuffer);

    // 检查所有幻灯片都有背景信息
    result.slides.forEach(slide => {
      expect(slide.background).toBeDefined();
    });
  });
});
