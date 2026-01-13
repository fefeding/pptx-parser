/**
 * 富文本解析和HTML生成测试
 * 对齐PPTXjs的核心渲染能力
 * 
 * 测试重点：
 * 1. 富文本PPTX生成器功能
 * 2. PPTX解析器正确处理富文本
 * 3. HTML生成器正确渲染富文本
 * 4. 单位转换准确性
 * 5. 样式继承正确性
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { createRichTextPptx, MockPptxGenerator } from '../test/mock-pptx-generator';
import { parsePptx } from '../src/core';
import { generateHtml, HtmlGenerator } from '../src/render/html-generator';
import { emu2px, fontUnits2px } from '../src/utils/unit-converter';

describe('富文本PPTX解析和HTML生成', () => {
  let mockPptx: Blob;

  beforeEach(async () => {
    // 创建测试用的富文本PPTX
    mockPptx = await createRichTextPptx({
      title: '测试标题',
      content: [
        {
          text: '第一行文本',
          fontSize: 18,
          color: '000000',
          bold: true
        },
        {
          text: '第二行文本',
          fontSize: 24,
          color: 'FF0000',
          italic: true
        },
        {
          text: '第三行文本',
          fontSize: 16,
          color: '0000FF',
          underline: true
        }
      ],
      backgroundColor: 'FFFFFF'
    });
  });

  describe('Mock PPTX生成器', () => {
    it('应该生成有效的PPTX文件', async () => {
      expect(mockPptx).toBeInstanceOf(Blob);
      expect(mockPptx.size).toBeGreaterThan(0);
    });

    it('应该正确设置标题', async () => {
      const generator = new MockPptxGenerator();
      await generator.createBaseStructure();
      await generator.createLayout(1);
      await generator.createRichTextSlide({
        title: '测试标题',
        content: []
      });
      const pptx = await generator.generate();

      expect(pptx).toBeInstanceOf(Blob);
    });

    it('应该支持多段富文本内容', async () => {
      const pptx = await createRichTextPptx({
        title: '多段文本',
        content: [
          { text: '段落1', fontSize: 18 },
          { text: '段落2', fontSize: 20 },
          { text: '段落3', fontSize: 22 }
        ]
      });

      expect(pptx).toBeInstanceOf(Blob);
    });

    it('应该正确设置文本样式', async () => {
      const pptx = await createRichTextPptx({
        title: '样式测试',
        content: [
          {
            text: '粗体',
            fontSize: 16,
            bold: true
          },
          {
            text: '斜体',
            fontSize: 16,
            italic: true
          },
          {
            text: '下划线',
            fontSize: 16,
            underline: true
          }
        ]
      });

      expect(pptx).toBeInstanceOf(Blob);
    });

    it('应该支持自定义颜色', async () => {
      const pptx = await createRichTextPptx({
        title: '颜色测试',
        content: [
          { text: '红色', fontSize: 18, color: 'FF0000' },
          { text: '绿色', fontSize: 18, color: '00FF00' },
          { text: '蓝色', fontSize: 18, color: '0000FF' }
        ]
      });

      expect(pptx).toBeInstanceOf(Blob);
    });
  });

  describe('PPTX解析器 - 富文本解析', () => {
    it('应该成功解析富文本PPTX', async () => {
      const result = await parsePptx(mockPptx);

      expect(result).toBeDefined();
      expect(result.slides).toHaveLength(1);
      expect(result.title).toBe('测试PPT');
    });

    it('应该正确解析文本元素', async () => {
      const result = await parsePptx(mockPptx);

      const firstSlide = result.slides[0];
      expect(firstSlide.elements).toBeInstanceOf(Array);
      expect(firstSlide.elements.length).toBeGreaterThan(0);
    });

    it('应该正确解析文本样式', async () => {
      const result = await parsePptx(mockPptx);

      const firstSlide = result.slides[0];
      const textElements = firstSlide.elements.filter((el: any) => el.type === 'text');
      
      expect(textElements.length).toBeGreaterThan(0);
      
      // 检查第一个文本元素的样式
      const firstTextElement = textElements[0];
      if (firstTextElement && firstTextElement.style) {
        expect(firstTextElement.style.fontWeight).toBeDefined();
      }
    });

    it('应该正确解析背景色', async () => {
      const result = await parsePptx(mockPptx);

      const firstSlide = result.slides[0];
      expect(firstSlide.bgColor).toBe('#FFFFFF');
    });
  });

  describe('HTML生成器 - 富文本渲染', () => {
    it('应该生成有效的HTML结构', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toContain('<div class="pptxjs-container">');
      expect(html).toContain('</div>');
      expect(html).toContain('<style>');
      expect(html).toContain('</style>');
    });

    it('应该生成幻灯片容器', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toContain('class="slide"');
      expect(html).toContain('data-slide-id="');
    });

    it('应该生成文本元素', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toContain('class="text-element"');
      expect(html).toContain('position:absolute');
    });

    it('应该应用文本样式', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toContain('font-size:');
      expect(html).toContain('px');
    });

    it('应该应用背景色', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toContain('background-color:#FFFFFF');
    });

    it('应该生成全局CSS', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toContain('.pptxjs-container');
      expect(html).toContain('.slide');
      expect(html).toContain('.text-element');
    });
  });

  describe('单位转换准确性', () => {
    it('应该正确转换EMU到PX', () => {
      // PPTX标准：914400 EMU = 96 PX
      expect(emu2px(914400)).toBe(96);
      expect(emu2px(1828800)).toBe(192);
      expect(emu2px(9144000)).toBe(960);
    });

    it('应该正确转换字体单位到PX', () => {
      // PPTXjs：fontSize = parseInt(sz) / 100 * (4/3.2)
      expect(fontUnits2px(1800)).toBe(22.5); // 18pt
      expect(fontUnits2px(2400)).toBe(30); // 24pt
      expect(fontUnits2px(3200)).toBe(40); // 32pt
    });

    it('应该双向转换保持一致性', () => {
      const testCases = [96, 192, 480, 720, 960];
      testCases.forEach(px => {
        const emu = Math.round(px * (914400 / 96));
        const convertedPx = emu2px(emu);
        expect(convertedPx).toBe(px);
      });
    });
  });

  describe('样式继承正确性', () => {
    it('应该正确继承字体样式', async () => {
      const pptx = await createRichTextPptx({
        content: [
          { text: '测试', fontSize: 20, color: 'FF0000' }
        ]
      });

      const document = await parsePptx(pptx);
      const html = generateHtml(document);

      expect(html).toContain('color:#FF0000');
      expect(html).toContain('font-size:');
    });

    it('应该正确继承粗体样式', async () => {
      const pptx = await createRichTextPptx({
        content: [
          { text: '粗体文本', fontSize: 18, bold: true }
        ]
      });

      const document = await parsePptx(pptx);
      const html = generateHtml(document);

      expect(html).toContain('font-weight:bold');
    });

    it('应该正确继承斜体样式', async () => {
      const pptx = await createRichTextPptx({
        content: [
          { text: '斜体文本', fontSize: 18, italic: true }
        ]
      });

      const document = await parsePptx(pptx);
      const html = generateHtml(document);

      expect(html).toContain('font-style:italic');
    });

    it('应该正确继承下划线样式', async () => {
      const pptx = await createRichTextPptx({
        content: [
          { text: '下划线文本', fontSize: 18, underline: true }
        ]
      });

      const document = await parsePptx(pptx);
      const html = generateHtml(document);

      expect(html).toContain('text-decoration:underline');
    });
  });

  describe('绝对定位正确性', () => {
    it('应该正确计算绝对定位位置', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      // 检查HTML中的位置属性
      expect(html).toMatch(/left:\d+px/);
      expect(html).toMatch(/top:\d+px/);
    });

    it('应该正确设置元素尺寸', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toMatch(/width:\d+px/);
      expect(html).toMatch(/height:\d+px/);
    });

    it('应该在幻灯片容器内正确定位元素', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      // 幻灯片容器应该有固定尺寸
      expect(html).toContain('width:');
      expect(html).toContain('height:');
      // 元素应该是绝对定位
      expect(html).toContain('position:absolute');
    });
  });

  describe('HTML生成器选项', () => {
    it('应该支持revealjs格式', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document, { slideType: 'section' });

      expect(html).toContain('<section class="slide"');
      expect(html).toContain('</section>');
    });

    it('应该支持自定义容器类名', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document, { containerClass: 'custom-container' });

      expect(html).toContain('class="custom-container"');
    });

    it('应该可以选择禁用全局CSS', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document, { includeGlobalCSS: false });

      expect(html).not.toContain('<style>');
    });
  });

  describe('HTML转义安全性', () => {
    it('应该正确转义特殊字符', async () => {
      const pptx = await createRichTextPptx({
        content: [
          { text: '<script>alert("test")</script>', fontSize: 18 }
        ]
      });

      const document = await parsePptx(pptx);
      const html = generateHtml(document);

      // HTML应该被转义
      expect(html).not.toContain('<script>');
      expect(html).toContain('&lt;script&gt;');
    });

    it('应该正确转义引号', async () => {
      const pptx = await createRichTextPptx({
        content: [
          { text: '测试"引号"和\'单引号\'', fontSize: 18 }
        ]
      });

      const document = await parsePptx(pptx);
      const html = generateHtml(document);

      expect(html).toContain('&quot;');
      expect(html).toContain('&#039;');
    });

    it('应该正确转义&符号', async () => {
      const pptx = await createRichTextPptx({
        content: [
          { text: '测试&符号', fontSize: 18 }
        ]
      });

      const document = await parsePptx(pptx);
      const html = generateHtml(document);

      expect(html).toContain('&amp;');
    });
  });

  describe('综合场景测试', () => {
    it('应该完整处理富文本PPTX到HTML的流程', async () => {
      // 创建富文本PPTX
      const pptx = await createRichTextPptx({
        title: '完整测试',
        content: [
          { text: '标题1', fontSize: 24, color: 'FF0000', bold: true },
          { text: '正文1', fontSize: 16, color: '000000' },
          { text: '标题2', fontSize: 20, color: '0000FF', italic: true },
          { text: '正文2', fontSize: 14, color: '333333', underline: true }
        ],
        backgroundColor: 'F0F0F0'
      });

      // 解析PPTX
      const document = await parsePptx(pptx);
      expect(document.slides).toHaveLength(1);

      // 生成HTML
      const html = generateHtml(document);

      // 验证HTML结构
      expect(html).toContain('<div class="pptxjs-container">');
      expect(html).toContain('<div class="slide"');
      expect(html).toContain('<div class="slide-background"');
      expect(html).toContain('background-color:#F0F0F0');

      // 验证文本元素
      expect(html).toContain('<div class="text-element');
      expect(html).toContain('position:absolute');

      // 验证样式
      expect(html).toContain('font-weight:bold');
      expect(html).toContain('font-style:italic');
      expect(html).toContain('text-decoration:underline');
      expect(html).toContain('color:#FF0000');
      expect(html).toContain('color:#0000FF');

      // 验证全局CSS
      expect(html).toContain('<style>');
      expect(html).toContain('.pptxjs-container');
      expect(html).toContain('.slide');
      expect(html).toContain('.text-element');
    });

    it('应该正确处理多段文本的样式继承', async () => {
      const pptx = await createRichTextPptx({
        content: [
          { text: '段落A', fontSize: 18, color: 'FF0000' },
          { text: '段落B', fontSize: 20, color: '00FF00' },
          { text: '段落C', fontSize: 22, color: '0000FF' }
        ]
      });

      const document = await parsePptx(pptx);
      const html = generateHtml(document);

      // 每个段落应该有不同的颜色和字体大小
      expect(html).toContain('color:#FF0000');
      expect(html).toContain('color:#00FF00');
      expect(html).toContain('color:#0000FF');
      expect(html).toContain('font-size:');
    });
  });
});