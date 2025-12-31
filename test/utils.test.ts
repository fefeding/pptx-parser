import { describe, it, expect } from 'vitest';
import { PptParseUtils } from '../src/core';

describe('PptParseUtils - 工具函数测试', () => {
  describe('generateId', () => {
    it('应该生成唯一的ID', () => {
      const id1 = PptParseUtils.generateId('test');
      const id2 = PptParseUtils.generateId('test');
      expect(id1).not.toBe(id2);
      expect(id1).toContain('test-');
      expect(id2).toContain('test-');
    });

    it('应该支持自定义前缀', () => {
      const id = PptParseUtils.generateId('slide');
      expect(id).toMatch(/^slide-\d+-\d+$/);
    });

    it('应该使用默认前缀', () => {
      const id = PptParseUtils.generateId();
      expect(id).toMatch(/^ppt-node-\d+-\d+$/);
    });
  });

  describe('parseXmlText', () => {
    it('应该处理HTML转义字符', () => {
      const text = PptParseUtils.parseXmlText('&lt;div&gt;Hello&lt;/div&gt;');
      expect(text).toBe('<div>Hello</div>');
    });

    it('应该去除首尾空格', () => {
      const text = PptParseUtils.parseXmlText('  hello world  ');
      expect(text).toBe('hello world');
    });

    it('应该处理空字符串', () => {
      const text = PptParseUtils.parseXmlText('');
      expect(text).toBe('');
    });

    it('应该处理undefined', () => {
      const text = PptParseUtils.parseXmlText(undefined as any);
      expect(text).toBe('');
    });
  });

  describe('parseXmlAttrs', () => {
    it('应该正确解析XML属性', () => {
      const xml = `<div id="test" class="container" data-value="123"></div>`;
      const parser = new DOMParser();
      const doc = parser.parseFromString(xml, 'application/xml');
      const attrs = PptParseUtils.parseXmlAttrs(doc.documentElement.attributes);

      expect(attrs.id).toBe('test');
      expect(attrs.class).toBe('container');
      expect(attrs['data-value']).toBe('123');
    });

    it('应该处理空属性', () => {
      const xml = `<div></div>`;
      const parser = new DOMParser();
      const doc = parser.parseFromString(xml, 'application/xml');
      const attrs = PptParseUtils.parseXmlAttrs(doc.documentElement.attributes);

      expect(attrs).toEqual({});
    });
  });

  describe('parseXmlToTree', () => {
    it('应该将XML字符串转换为树形结构', () => {
      const xml = `<root><child1>text1</child1><child2><sub>text2</sub></child2></root>`;
      const tree = PptParseUtils.parseXmlToTree(xml);

      expect(tree.tag).toBe('root');
      expect(tree.children).toHaveLength(2);
      expect(tree.children[0].tag).toBe('child1');
      expect(tree.children[0].text).toBe('text1');
      expect(tree.children[1].tag).toBe('child2');
      expect(tree.children[1].children[0].text).toBe('text2');
    });

    it('应该解析属性', () => {
      const xml = `<root id="main"><child name="test"/></root>`;
      const tree = PptParseUtils.parseXmlToTree(xml);

      expect(tree.attrs.id).toBe('main');
      expect(tree.children[0].attrs.name).toBe('test');
    });
  });

  describe('parseXmlRect', () => {
    it('应该正确转换EMU单位到PX', () => {
      // 914400 EMU = 96 PX (标准转换)
      const attrs = { x: '914400', y: '1828800', cx: '2743200', cy: '3657600' };
      const rect = PptParseUtils.parseXmlRect(attrs);

      expect(rect.x).toBe(96);
      expect(rect.y).toBe(192);
      expect(rect.width).toBe(288);
      expect(rect.height).toBe(384);
    });

    it('应该处理空值', () => {
      const attrs = { x: '', y: '', cx: '', cy: '' };
      const rect = PptParseUtils.parseXmlRect(attrs);

      expect(rect.x).toBe(0);
      expect(rect.y).toBe(0);
      expect(rect.width).toBe(0);
      expect(rect.height).toBe(0);
    });
  });

  describe('parseXmlStyle', () => {
    it('应该正确解析样式属性', () => {
      const attrs = {
        fontSize: '1600',
        fill: '#ff0000',
        bold: '1',
        align: 'center',
        bgFill: '#ffffff',
        border: '#000000',
        borderWidth: '2'
      };
      const style = PptParseUtils.parseXmlStyle(attrs);

      expect(style.fontSize).toBe(16);
      expect(style.color).toBe('#ff0000');
      expect(style.fontWeight).toBe('bold');
      expect(style.textAlign).toBe('center');
      expect(style.backgroundColor).toBe('#ffffff');
      expect(style.borderColor).toBe('#000000');
      expect(style.borderWidth).toBe(2);
    });

    it('应该使用默认值', () => {
      const attrs = {};
      const style = PptParseUtils.parseXmlStyle(attrs);

      expect(style.fontSize).toBe(14);
      expect(style.color).toBe('#333333');
      expect(style.fontWeight).toBe('normal');
      expect(style.textAlign).toBe('left');
      expect(style.backgroundColor).toBe('transparent');
      expect(style.borderColor).toBe('#000000');
      expect(style.borderWidth).toBe(1);
    });

    it('应该处理无效的对齐方式', () => {
      const attrs = { align: 'invalid' };
      const style = PptParseUtils.parseXmlStyle(attrs);

      expect(style.textAlign).toBe('left');
    });
  });

  describe('px2emu', () => {
    it('应该正确转换PX到EMU', () => {
      const emu = PptParseUtils.px2emu(96);
      expect(emu).toBe(914400);
    });

    it('应该正确转换不同尺寸', () => {
      expect(PptParseUtils.px2emu(100)).toBe(952500);
      expect(PptParseUtils.px2emu(200)).toBe(1905000);
      expect(PptParseUtils.px2emu(0)).toBe(0);
    });
  });

  describe('emu2px', () => {
    it('应该正确转换EMU到PX', () => {
      const px = PptParseUtils.emu2px(914400);
      expect(px).toBe(96);
    });

    it('应该正确转换不同尺寸', () => {
      expect(PptParseUtils.emu2px(952500)).toBe(100);
      expect(PptParseUtils.emu2px(1905000)).toBe(200);
      expect(PptParseUtils.emu2px(0)).toBe(0);
    });

    it('应该双向转换保持一致', () => {
      const originalPx = 100;
      const emu = PptParseUtils.px2emu(originalPx);
      const convertedPx = PptParseUtils.emu2px(emu);
      expect(convertedPx).toBe(originalPx);
    });
  });
});
