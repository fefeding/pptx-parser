import { describe, it, expect } from 'vitest';
import { parsePptx } from '../src/core/parser';

describe('Real PPTX File Tests', () => {
  it('should parse 金腾研发架构&系统介绍.pptx', async () => {
    // 读取实际的PPTX文件
    const pptxPath = 'c:/Users/fefeding/Desktop/金腾研发架构&系统介绍/金腾研发架构&系统介绍.pptx';
    
    try {
      const fs = await import('fs');
      const fileBuffer = fs.readFileSync(pptxPath);
      
      const result = await parsePptx(fileBuffer, {
        parseImages: true,
        returnFormat: 'enhanced'
      });

      // 验证基本结构
      expect(result).toBeDefined();
      expect(result.slides).toBeDefined();
      expect(result.slides.length).toBeGreaterThan(0);
      expect(result.title).toBeDefined();
      
      // 验证幻灯片尺寸
      expect(result.props).toBeDefined();
      expect(result.props.width).toBeGreaterThan(0);
      expect(result.props.height).toBeGreaterThan(0);

      // 验证至少有一张幻灯片有元素
      const slidesWithElements = result.slides.filter(slide => 
        slide.elements && slide.elements.length > 0
      );
      expect(slidesWithElements.length).toBeGreaterThan(0);

      console.log(`Parsed ${result.slides.length} slides`);
      console.log(`Title: ${result.title}`);
      console.log(`Dimensions: ${result.props.width}x${result.props.height}`);
      console.log(`Slides with elements: ${slidesWithElements.length}`);

      // 输出第一张幻灯片的信息
      if (result.slides.length > 0) {
        const firstSlide = result.slides[0];
        console.log(`First slide title: ${firstSlide.title}`);
        console.log(`First slide background:`, firstSlide.background);
        console.log(`First slide elements count: ${firstSlide.elements.length}`);
        
        // 检查背景是否正确解析
        expect(firstSlide.background).toBeDefined();
      }
    } catch (error) {
      // 如果文件不存在，跳过测试
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
        console.log('Skipping test: PPTX file not found');
        return;
      }
      throw error;
    }
  });
});
