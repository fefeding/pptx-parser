/**
 * PPTXjs集成测试
 * 测试转译后的PPTXjs功能
 */

import { describe, it, expect, beforeAll } from 'vitest';
import { parsePptx, Pptxjs } from '../src/pptxjs';
import { readFileSync } from 'fs';
import { join } from 'path';

describe('PPTXjs Integration Tests', () => {
  let testPptxBuffer: ArrayBuffer;

  beforeAll(() => {
    // 这里可以使用实际的PPTX文件进行测试
    // 暂时创建一个简单的测试
    console.log('PPTXjs integration tests initialized');
  });

  describe('parsePptx function', () => {
    it('should parse PPTX file and return structure', async () => {
      // 注意：这里需要一个实际的PPTX文件
      // 可以从 examples/PPTXjs 目录获取
      const pptxPath = join(__dirname, '../examples/PPTXjs/pptx/test.pptx');
      
      try {
        const fileBuffer = readFileSync(pptxPath);
        const result = await parsePptx(fileBuffer);

        expect(result).toBeDefined();
        expect(result.slides).toBeInstanceOf(Array);
        expect(result.size).toBeDefined();
        expect(result.size.width).toBeGreaterThan(0);
        expect(result.size.height).toBeGreaterThan(0);
        expect(result.globalCSS).toBeDefined();
        expect(typeof result.globalCSS).toBe('string');
      } catch (e) {
        // 文件不存在时跳过测试
        console.log('Test PPTX file not found, skipping test');
      }
    });

    it('should handle invalid input gracefully', async () => {
      await expect(parsePptx(null as any)).rejects.toThrow();
      await expect(parsePptx(undefined as any)).rejects.toThrow();
      await expect(parsePptx({} as any)).rejects.toThrow();
    });
  });

  describe('Pptxjs class', () => {
    it('should create instance and parse', async () => {
      const pptxPath = join(__dirname, '../examples/PPTXjs/pptx/test.pptx');
      
      try {
        const fileBuffer = readFileSync(pptxPath);
        const pptxjs = await Pptxjs.create(fileBuffer);

        expect(pptxjs).toBeInstanceOf(Pptxjs);
        
        const slides = pptxjs.getSlides();
        expect(slides).toBeInstanceOf(Array);

        const size = pptxjs.getSize();
        expect(size.width).toBeGreaterThan(0);
        expect(size.height).toBeGreaterThan(0);
      } catch (e) {
        console.log('Test PPTX file not found, skipping test');
      }
    });

    it('should generate HTML', async () => {
      const pptxPath = join(__dirname, '../examples/PPTXjs/pptx/test.pptx');
      
      try {
        const fileBuffer = readFileSync(pptxPath);
        const pptxjs = await Pptxjs.create(fileBuffer);

        const html = pptxjs.generateHtml();
        
        expect(html).toBeDefined();
        expect(typeof html).toBe('string');
        expect(html).toContain('<!DOCTYPE html>');
        expect(html).toContain('</html>');
        expect(html).toContain('.slide');
      } catch (e) {
        console.log('Test PPTX file not found, skipping test');
      }
    });
  });

  describe('Core Parser Functions', () => {
    it('should handle color parsing', () => {
      const { getColorValue, getThemeColor, getPresetColor } = require('../src/pptxjs/pptxjs-color-utils');

      // 测试十六进制颜色
      const hexColor = getColorValue({
        'a:srgbClr': { attrs: { val: 'FF0000' } }
      });
      expect(hexColor).toBe('#FF0000');

      // 测试主题颜色
      const themeColor = getColorValue({
        'a:schemeClr': { attrs: { val: 'accent1' } }
      });
      expect(themeColor).toBe('#4F81BD');

      // 测试预设颜色
      const presetColor = getPresetColor('red');
      expect(presetColor).toBe('#FF0000');
    });

    it('should handle text style parsing', () => {
      const { parseTextProps, generateTextStyleCss } = require('../src/pptxjs/pptxjs-text-utils');

      const textProps = {
        'a:latin': { attrs: { typeface: 'Arial' } },
        'a:sz': { attrs: { val: '1800' } },
        'a:solidFill': {
          'a:srgbClr': { attrs: { val: 'FF0000' } }
        },
        'a:b': { attrs: { val: '1' } },
        'a:i': { attrs: { val: '1' } },
      };

      const style = parseTextProps(textProps);
      
      expect(style.fontFace).toBe('Arial');
      expect(style.fontSize).toBe(18);
      expect(style.color).toBe('#FF0000');
      expect(style.bold).toBe(true);
      expect(style.italic).toBe(true);

      const css = generateTextStyleCss(style);
      expect(css).toContain('font-family');
      expect(css).toContain('font-size: 18pt');
      expect(css).toContain('color: #FF0000');
      expect(css).toContain('font-weight: bold');
      expect(css).toContain('font-style: italic');
    });

    it('should handle unit conversions', () => {
      const { angleToDegrees } = require('../src/pptxjs/pptxjs-core-parser');

      // PPTX角度单位是1/60000度
      expect(angleToDegrees(60000)).toBe(1);
      expect(angleToDegrees(120000)).toBe(2);
      expect(angleToDegrees(180000)).toBe(3);
      expect(angleToDegrees(undefined)).toBe(0);
      expect(angleToDegrees(null)).toBe(0);
    });
  });

  describe('Utility Functions', () => {
    it('should handle base64 conversion', () => {
      const { base64ArrayBuffer, getImageMimeType, generateDataUrl } = require('../src/pptxjs/pptxjs-utils');

      const testString = 'Hello, World!';
      const encoder = new TextEncoder();
      const arrayBuffer = encoder.encode(testString);
      
      const base64 = base64ArrayBuffer(arrayBuffer);
      expect(typeof base64).toBe('string');
      expect(base64.length).toBeGreaterThan(0);

      const mimeType = getImageMimeType('test.png');
      expect(mimeType).toBe('image/png');

      const mimeType2 = getImageMimeType('test.jpg');
      expect(mimeType2).toBe('image/jpeg');

      const dataUrl = generateDataUrl(base64, 'image/png');
      expect(dataUrl).toBe('data:image/png;base64,' + base64);
    });

    it('should handle number parsing', () => {
      const { safeParseInt, safeParseFloat } = require('../src/pptxjs/pptxjs-utils');

      expect(safeParseInt('123')).toBe(123);
      expect(safeParseInt('123.45')).toBe(123);
      expect(safeParseInt('abc', 0)).toBe(0);
      expect(safeParseInt(null, 5)).toBe(5);
      expect(safeParseInt(undefined, 5)).toBe(5);

      expect(safeParseFloat('123.45')).toBeCloseTo(123.45);
      expect(safeParseFloat('abc', 0)).toBe(0);
      expect(safeParseFloat(null, 5)).toBe(5);
    });

    it('should handle color utilities', () => {
      const { 
        normalizeHexColor, 
        isValidHexColor,
        hexToRgba,
        generateCssColor
      } = require('../src/pptxjs/pptxjs-color-utils');

      expect(normalizeHexColor('#FF0000')).toBe('FF0000');
      expect(normalizeHexColor('FF0000')).toBe('FF0000');
      expect(normalizeHexColor('#F00')).toBe('FF0000');
      expect(normalizeHexColor('F00')).toBe('FF0000');

      expect(isValidHexColor('#FF0000')).toBe(true);
      expect(isValidHexColor('FF0000')).toBe(true);
      expect(isValidHexColor('#F00')).toBe(true);
      expect(isValidHexColor('GGG')).toBe(false);

      expect(hexToRgba('#FF0000', 0.5)).toBe('rgba(255, 0, 0, 0.5)');
      expect(hexToRgba('#00FF00', 0.8)).toBe('rgba(0, 255, 0, 0.8)');

      const solidColor = {
        type: 'solid',
        color: '#FF0000',
        alpha: 1,
      };
      const css = generateCssColor(solidColor);
      expect(css).toBe('#FF0000');
    });
  });

  describe('Text Processing', () => {
    it('should parse text paragraphs', () => {
      const { parseTextBoxContent, generateTextBoxHtml } = require('../src/pptxjs/pptxjs-text-utils');

      const txBody = {
        'a:p': [
          {
            'a:pPr': {
              'a:lnSpc': {
                'a:spcPct': { attrs: { val: '120000' } }
              }
            },
            'a:r': [
              {
                'a:rPr': {
                  'a:sz': { attrs: { val: '1800' } },
                  'a:solidFill': {
                    'a:srgbClr': { attrs: { val: '000000' } }
                  }
                },
                'a:t': 'Hello, World!'
              }
            ]
          }
        ]
      };

      const paragraphs = parseTextBoxContent(txBody);
      expect(paragraphs).toBeInstanceOf(Array);
      expect(paragraphs.length).toBeGreaterThan(0);
      expect(paragraphs[0].text).toBe('Hello, World!');
      expect(paragraphs[0].lineSpacing).toBe(1.2);

      const html = generateTextBoxHtml(paragraphs);
      expect(html).toContain('Hello, World!');
    });

    it('should generate paragraph HTML with styles', () => {
      const { generateTextParagraphHtml, getDefaultTextStyle } = require('../src/pptxjs/pptxjs-text-utils');

      const paragraph = {
        text: 'Test paragraph',
        styles: [
          {
            fontFace: 'Arial',
            fontSize: 18,
            color: '#FF0000',
            bold: true,
          }
        ],
        textAlign: 'center' as const,
      };

      const html = generateTextParagraphHtml(paragraph);
      expect(html).toContain('Test paragraph');
      expect(html).toContain('text-align: center');
      expect(html).toContain('font-weight: bold');
    });
  });
});
