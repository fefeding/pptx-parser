/**
 * 图片和背景解析测试
 * 对齐PPTXjs的图片处理和背景渲染能力
 * 
 * 测试重点：
 * 1. 图片PPTX生成器功能
 * 2. 图片解析和base64嵌入
 * 3. 纯色背景渲染
 * 4. 渐变背景渲染
 * 5. 图片背景渲染
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { createImagePptx } from '../test/mock-pptx-generator';
import { parsePptx } from '../src/core';
import { generateHtml } from '../src/render/html-generator';
import { emu2px } from '../src/utils/unit-converter';

describe('图片解析和背景渲染', () => {
  // 创建一个简单的红色PNG图片（1x1像素）
  const redPixelBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIW2NkYGD4D8DwABAwEAQYAMAAAFGQw4UAAAAASUVORK5CYII=';

  let mockPptx: Blob;

  beforeEach(async () => {
    // 创建测试用的图片PPTX
    mockPptx = await createImagePptx({
      images: [
        {
          fileName: 'test-image.png',
          mimeType: 'image/png',
          data: redPixelBase64,
          x: 100,
          y: 100,
          width: 200,
          height: 150
        }
      ],
      backgroundColor: 'FFFFFF'
    });
  });

  describe('Mock PPTX图片生成器', () => {
    it('应该生成包含图片的PPTX文件', async () => {
      expect(mockPptx).toBeInstanceOf(Blob);
      expect(mockPptx.size).toBeGreaterThan(0);
    });

    it('应该正确添加媒体资源', async () => {
      const pptx = await createImagePptx({
        images: [
          {
            fileName: 'image1.png',
            mimeType: 'image/png',
            data: redPixelBase64,
            x: 50,
            y: 50,
            width: 100,
            height: 100
          }
        ]
      });

      expect(pptx).toBeInstanceOf(Blob);
    });

    it('应该支持多张图片', async () => {
      const pptx = await createImagePptx({
        images: [
          {
            fileName: 'image1.png',
            mimeType: 'image/png',
            data: redPixelBase64,
            x: 50,
            y: 50,
            width: 100,
            height: 100
          },
          {
            fileName: 'image2.png',
            mimeType: 'image/png',
            data: redPixelBase64,
            x: 200,
            y: 150,
            width: 150,
            height: 100
          }
        ]
      });

      expect(pptx).toBeInstanceOf(Blob);
    });

    it('应该正确设置图片位置', async () => {
      const pptx = await createImagePptx({
        images: [
          {
            fileName: 'position-test.png',
            mimeType: 'image/png',
            data: redPixelBase64,
            x: 300,
            y: 200,
            width: 100,
            height: 80
          }
        ]
      });

      expect(pptx).toBeInstanceOf(Blob);
    });
  });

  describe('PPTX解析器 - 图片解析', () => {
    it('应该成功解析包含图片的PPTX', async () => {
      const result = await parsePptx(mockPptx);

      expect(result).toBeDefined();
      expect(result.slides).toHaveLength(1);
    });

    it('应该正确解析图片元素', async () => {
      const result = await parsePptx(mockPptx);

      const firstSlide = result.slides[0];
      const imageElements = firstSlide.elements?.filter((el: any) => el.type === 'image');

      expect(imageElements).toBeDefined();
      expect(imageElements.length).toBeGreaterThan(0);
    });

    it('应该正确解析图片尺寸', async () => {
      const result = await parsePptx(mockPptx);

      const firstSlide = result.slides[0];
      const imageElements = firstSlide.elements?.filter((el: any) => el.type === 'image');

      if (imageElements && imageElements.length > 0) {
        const firstImage = imageElements[0];
        expect(firstImage.position).toBeDefined();
        expect(firstImage.size).toBeDefined();
      }
    });

    it('应该正确解析图片位置', async () => {
      const result = await parsePptx(mockPptx);

      const firstSlide = result.slides[0];
      const imageElements = firstSlide.elements?.filter((el: any) => el.type === 'image');

      if (imageElements && imageElements.length > 0) {
        const firstImage = imageElements[0];
        const position = firstImage.position;
        expect(position.x).toBeDefined();
        expect(position.y).toBeDefined();
      }
    });
  });

  describe('HTML生成器 - 图片渲染', () => {
    it('应该生成图片元素HTML', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toContain('<img class="slide-image"');
    });

    it('应该正确设置图片位置', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toMatch(/left:\d+px/);
      expect(html).toMatch(/top:\d+px/);
    });

    it('应该正确设置图片尺寸', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toMatch(/width:\d+px/);
      expect(html).toMatch(/height:\d+px/);
    });

    it('应该应用绝对定位', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toContain('position:absolute');
    });
  });

  describe('背景渲染 - 纯色背景', () => {
    it('应该生成纯色背景HTML', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toContain('background-color:#FFFFFF');
    });

    it('应该生成背景容器元素', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toContain('class="slide-background"');
      expect(html).toContain('width:100%');
      expect(html).toContain('height:100%');
    });

    it('应该正确应用背景定位', async () => {
      const document = await parsePptx(mockPptx);
      const html = generateHtml(document);

      expect(html).toContain('position:absolute');
      expect(html).toContain('top:0');
      expect(html).toContain('left:0');
    });
  });

  describe('背景渲染 - 渐变背景', () => {
    it('应该生成线性渐变背景HTML', async () => {
      const document = {
        slides: [
          {
            id: 1,
            width: 960,
            height: 720,
            elements: [],
            bgFill: {
              type: 'gradient',
              direction: 'to bottom',
              colors: ['#FF0000 0%', '#00FF00 50%', '#0000FF 100%']
            }
          }
        ],
        props: { width: 960, height: 720 },
        title: '渐变背景测试'
      };

      const html = generateHtml(document);

      expect(html).toContain('linear-gradient');
      expect(html).toContain('#FF0000');
      expect(html).toContain('#00FF00');
      expect(html).toContain('#0000FF');
    });

    it('应该支持自定义渐变方向', async () => {
      const document = {
        slides: [
          {
            id: 1,
            width: 960,
            height: 720,
            elements: [],
            bgFill: {
              type: 'gradient',
              direction: 'to right',
              colors: ['#000000 0%', '#FFFFFF 100%']
            }
          }
        ],
        props: { width: 960, height: 720 },
        title: '渐变方向测试'
      };

      const html = generateHtml(document);

      expect(html).toContain('linear-gradient(to right');
    });
  });

  describe('背景渲染 - 图片背景', () => {
    it('应该生成图片背景HTML', async () => {
      const document = {
        slides: [
          {
            id: 1,
            width: 960,
            height: 720,
            elements: [],
            bgFill: {
              type: 'image',
              src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIW2NkYGD4D8DwABAwEAQYAMAAAFGQw4UAAAAASUVORK5CYII='
            }
          }
        ],
        props: { width: 960, height: 720 },
        title: '图片背景测试'
      };

      const html = generateHtml(document);

      expect(html).toContain('background-image');
      expect(html).toContain('url(');
    });

    it('应该设置背景图片属性', async () => {
      const document = {
        slides: [
          {
            id: 1,
            width: 960,
            height: 720,
            elements: [],
            bgFill: {
              type: 'image',
              src: 'https://example.com/background.jpg'
            }
          }
        ],
        props: { width: 960, height: 720 },
        title: '外部图片背景测试'
      };

      const html = generateHtml(document);

      expect(html).toContain('background-size:cover');
      expect(html).toContain('background-position:center');
      expect(html).toContain('background-repeat:no-repeat');
    });

    it('应该支持base64图片背景', async () => {
      const document = {
        slides: [
          {
            id: 1,
            width: 960,
            height: 720,
            elements: [],
            bgFill: {
              type: 'image',
              src: `data:image/png;base64,${redPixelBase64}`
            }
          }
        ],
        props: { width: 960, height: 720 },
        title: 'Base64背景测试'
      };

      const html = generateHtml(document);

      expect(html).toContain('data:image/png;base64');
    });
  });

  describe('综合场景测试', () => {
    it('应该完整处理图片PPTX到HTML的流程', async () => {
      // 创建图片PPTX
      const pptx = await createImagePptx({
        images: [
          {
            fileName: 'test.png',
            mimeType: 'image/png',
            data: redPixelBase64,
            x: 100,
            y: 100,
            width: 200,
            height: 150
          }
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

      // 验证图片元素
      expect(html).toContain('<img class="slide-image"');
      expect(html).toContain('position:absolute');
      expect(html).toMatch(/left:\d+px/);
      expect(html).toMatch(/top:\d+px/);
      expect(html).toMatch(/width:\d+px/);
      expect(html).toMatch(/height:\d+px/);
    });

    it('应该正确处理多张图片', async () => {
      const pptx = await createImagePptx({
        images: [
          {
            fileName: 'image1.png',
            mimeType: 'image/png',
            data: redPixelBase64,
            x: 50,
            y: 50,
            width: 100,
            height: 100
          },
          {
            fileName: 'image2.png',
            mimeType: 'image/png',
            data: redPixelBase64,
            x: 200,
            y: 150,
            width: 150,
            height: 100
          }
        ]
      });

      const document = await parsePptx(pptx);
      const html = generateHtml(document);

      // 应该包含多张图片
      const imageCount = (html.match(/<img class="slide-image"/g) || []).length;
      expect(imageCount).toBeGreaterThanOrEqual(2);
    });

    it('应该正确处理混合背景类型', async () => {
      // 测试纯色、渐变、图片三种背景类型
      const backgrounds = [
        {
          type: 'solid',
          color: '#FF0000'
        },
        {
          type: 'gradient',
          direction: 'to bottom',
          colors: ['#000000 0%', '#FFFFFF 100%']
        },
        {
          type: 'image',
          src: `data:image/png;base64,${redPixelBase64}`
        }
      ];

      backgrounds.forEach((bgFill) => {
        const document = {
          slides: [
            {
              id: 1,
              width: 960,
              height: 720,
              elements: [],
              bgFill
            }
          ],
          props: { width: 960, height: 720 },
          title: '混合背景测试'
        };

        const html = generateHtml(document);
        expect(html).toContain('class="slide-background"');
      });
    });
  });

  describe('单位转换准确性（图片相关）', () => {
    it('应该正确转换图片位置EMU到PX', () => {
      const testCases = [
        { emu: 914400, expectedPx: 96 },    // 1英寸
        { emu: 1828800, expectedPx: 192 },  // 2英寸
        { emu: 457200, expectedPx: 48 }     // 0.5英寸
      ];

      testCases.forEach(({ emu, expectedPx }) => {
        const px = emu2px(emu);
        expect(px).toBe(expectedPx);
      });
    });

    it('应该正确转换图片尺寸EMU到PX', () => {
      // 测试常见图片尺寸
      const testCases = [
        { widthEmu: 914400, heightEmu: 685800, widthPx: 96, heightPx: 72 },   // 4:3比例
        { widthEmu: 1828800, heightEmu: 1023000, widthPx: 192, heightPx: 108 }, // 16:9比例
        { widthEmu: 2743200, heightEmu: 2057400, widthPx: 288, heightPx: 216 } // 4:3比例
      ];

      testCases.forEach(({ widthEmu, heightEmu, widthPx, heightPx }) => {
        const width = emu2px(widthEmu);
        const height = emu2px(heightEmu);
        expect(width).toBe(widthPx);
        expect(height).toBe(heightPx);
      });
    });
  });
});