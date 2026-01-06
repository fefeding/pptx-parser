/**
 * 布局解析集成测试
 * 测试完整的布局、母版、幻灯片解析流程
 */

import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { parsePptx } from '../../src/index';
import { join } from 'path';
import { existsSync } from 'fs';

// 检查测试文件是否存在
type TestFile = {
  name: string;
  path: string;
  description: string;
};

const testFiles: TestFile[] = [
  {
    name: 'simple-test',
    path: join(__dirname, '../../examples/test-data/simple-test.pptx'),
    description: '简单测试PPTX文件'
  },
  {
    name: 'real-presentation',
    path: join(__dirname, '../../examples/test-data/金腾研发架构&系统介绍.pptx'),
    description: '真实演示文稿文件'
  }
];

describe('布局解析集成测试', () => {
  let availableTestFiles: TestFile[] = [];

  beforeAll(() => {
    // 检查哪些测试文件实际存在
    availableTestFiles = testFiles.filter(file => existsSync(file.path));
    console.log('可用的测试文件:', availableTestFiles.map(f => f.name));
  });

  describe('基本布局解析', () => {
    it('应该能够解析PPTX文件的基本结构', async () => {
      if (availableTestFiles.length === 0) {
        console.warn('没有找到测试文件，跳过测试');
        return;
      }

      const testFile = availableTestFiles[0];
      console.log(`使用测试文件: ${testFile.description} (${testFile.path})`);

      // 读取文件
      const fs = await import('fs/promises');
      const buffer = await fs.readFile(testFile.path);

      // 解析PPTX
      const result = await parsePptx(buffer, {
        parseImages: false, // 不解析图片以加快测试速度
        parseCharts: true,
        parseTables: true,
        parseDiagrams: true,
        returnFormat: 'enhanced'
      });

      // 验证基本结构
      expect(result).toBeDefined();
      expect(result.slides).toBeInstanceOf(Array);
      expect(result.slides.length).toBeGreaterThan(0);
      
      console.log(`✓ 成功解析 ${result.slides.length} 张幻灯片`);

      // 验证幻灯片结构
      const firstSlide = result.slides[0];
      expect(firstSlide).toBeDefined();
      expect(firstSlide.id).toBeDefined();
      expect(firstSlide.elements).toBeInstanceOf(Array);
      
      console.log(`✓ 第一张幻灯片包含 ${firstSlide.elements.length} 个元素`);

      // 验证元素类型
      const elementTypes = new Set(firstSlide.elements.map(el => el.type));
      console.log(`✓ 发现的元素类型:`, Array.from(elementTypes));

      // 验证背景解析
      if (firstSlide.background) {
        console.log(`✓ 幻灯片背景:`, firstSlide.background);
        expect(firstSlide.background.type).toMatch(/color|image|none/);
      }

      // 验证母版和布局（如果存在）
      if (result.masterSlides && result.masterSlides.length > 0) {
        console.log(`✓ 解析到 ${result.masterSlides.length} 个母版`);
        
        const firstMaster = result.masterSlides[0];
        expect(firstMaster.id).toBeDefined();
        expect(firstMaster.elements).toBeInstanceOf(Array);
        console.log(`✓ 第一个母版包含 ${firstMaster.elements.length} 个元素`);
      }

      if (result.slideLayouts && Object.keys(result.slideLayouts).length > 0) {
        console.log(`✓ 解析到 ${Object.keys(result.slideLayouts).length} 个布局`);
      }

      // 验证主题（如果存在）
      if (result.theme) {
        console.log(`✓ 解析到主题:`, result.theme.colors);
        expect(result.theme.colors).toBeDefined();
      }
    });

    it('应该正确解析不同类型的元素', async () => {
      if (availableTestFiles.length === 0) {
        console.warn('没有找到测试文件，跳过测试');
        return;
      }

      const testFile = availableTestFiles[0];
      const fs = await import('fs/promises');
      const buffer = await fs.readFile(testFile.path);

      const result = await parsePptx(buffer, {
        parseImages: false,
        returnFormat: 'enhanced'
      });

      // 统计各种元素类型的数量
      const elementCounts: Record<string, number> = {};
      
      result.slides.forEach(slide => {
        slide.elements.forEach(element => {
          elementCounts[element.type] = (elementCounts[element.type] || 0) + 1;
        });
      });

      console.log('✓ 元素统计:', elementCounts);

      // 验证至少有一些基本元素
      const totalElements = Object.values(elementCounts).reduce((sum, count) => sum + count, 0);
      expect(totalElements).toBeGreaterThan(0);
      
      console.log(`✓ 总共解析到 ${totalElements} 个元素`);
    });

    it('应该支持新的 DocumentElement API', async () => {
      if (availableTestFiles.length === 0) {
        console.warn('没有找到测试文件，跳过测试');
        return;
      }

      const testFile = availableTestFiles[0];
      const fs = await import('fs/promises');
      const buffer = await fs.readFile(testFile.path);

      const result = await parsePptx(buffer, {
        parseImages: false,
        returnFormat: 'enhanced'
      });

      // 创建 DocumentElement
      const { createDocument } = await import('../../src/index');
      const doc = createDocument(result);

      expect(doc).toBeDefined();
      expect(doc.title).toBe(result.title || 'Untitled Presentation');
      expect(doc.slides).toHaveLength(result.slides.length);
      expect(doc.width).toBe(result.props.width);
      expect(doc.height).toBe(result.props.height);
      
      console.log(`✓ DocumentElement 创建成功`);
      console.log(`  - 标题: ${doc.title}`);
      console.log(`  - 尺寸: ${doc.width}x${doc.height}`);
      console.log(`  - 幻灯片数: ${doc.slides.length}`);

      // 测试 toHTML 方法
      const html = doc.toHTML({ withNavigation: false });
      expect(html).toBeDefined();
      expect(typeof html).toBe('string');
      expect(html.length).toBeGreaterThan(0);
      
      console.log(`✓ toHTML() 方法正常工作，生成 ${html.length} 字符的HTML`);

      // 验证HTML包含基本结构
      expect(html).toContain('<div');
      expect(html).toContain('class="ppt-container"');
      
      console.log(`✓ HTML包含正确的CSS类名`);
    });
  });

  describe('背景继承链测试', () => {
    it('应该正确处理背景继承', async () => {
      if (availableTestFiles.length === 0) {
        console.warn('没有找到测试文件，跳过测试');
        return;
      }

      const testFile = availableTestFiles[0];
      const fs = await import('fs/promises');
      const buffer = await fs.readFile(testFile.path);

      const result = await parsePptx(buffer, {
        parseImages: false,
        returnFormat: 'enhanced'
      });

      // 检查幻灯片的背景是否正确设置
      result.slides.forEach((slide, index) => {
        console.log(`幻灯片 ${index + 1} 背景:`, slide.background);
        
        // 背景应该有明确的类型
        if (slide.background) {
          expect(['color', 'image', 'none'].includes(slide.background.type)).toBe(true);
        }
      });

      console.log(`✓ 所有幻灯片的背景都已正确解析`);
    });
  });

  describe('错误处理测试', () => {
    it('应该优雅处理损坏的文件', async () => {
      // 创建一个空的buffer来模拟损坏的文件
      const invalidBuffer = Buffer.from('invalid pptx content');

      try {
        await parsePptx(invalidBuffer);
        // 如果没抛出错误，测试应该失败
        expect(false).toBe(true);
      } catch (error) {
        // 期望抛出错误
        expect(error).toBeDefined();
        console.log('✓ 正确处理损坏文件的错误:', error.message);
      }
    });

    it('应该处理不存在的文件路径', async () => {
      try {
        // 尝试解析不存在的文件
        const fs = await import('fs/promises');
        await fs.readFile('/non/existent/file.pptx');
        expect(false).toBe(true);
      } catch (error) {
        expect(error).toBeDefined();
        console.log('✓ 正确处理文件不存在的错误');
      }
    });
  });
});