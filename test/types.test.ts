import { describe, it, expect } from 'vitest';
import type { PptDocument, PptSlide, PptElement, PptNodeType } from '../src/types';

describe('Types - 类型定义测试', () => {
  it('应该正确创建PptDocument对象', () => {
    const doc: PptDocument = {
      id: 'doc-1',
      title: '测试文档',
      slides: [],
      props: {
        width: 1280,
        height: 720,
        ratio: 1.78,
      },
    };

    expect(doc.id).toBe('doc-1');
    expect(doc.title).toBe('测试文档');
    expect(doc.props.width).toBe(1280);
    expect(doc.props.height).toBe(720);
    expect(doc.props.ratio).toBe(1.78);
  });

  it('应该正确创建PptSlide对象', () => {
    const slide: PptSlide = {
      id: 'slide-1',
      title: '第一页',
      bgColor: '#ffffff',
      elements: [],
      props: {
        width: 1280,
        height: 720,
        slideLayout: 'normal',
      },
    };

    expect(slide.id).toBe('slide-1');
    expect(slide.title).toBe('第一页');
    expect(slide.bgColor).toBe('#ffffff');
    expect(slide.props.slideLayout).toBe('normal');
  });

  it('应该正确创建文本类型的PptElement', () => {
    const element: PptElement = {
      id: 'element-1',
      type: 'text' as PptNodeType,
      rect: { x: 100, y: 100, width: 200, height: 50 },
      style: {
        fontSize: 16,
        color: '#333333',
        fontWeight: 'normal',
        textAlign: 'left',
      },
      content: '测试文本',
      props: { lineHeight: 1.5 },
    };

    expect(element.type).toBe('text');
    expect(element.content).toBe('测试文本');
    expect(element.rect.width).toBe(200);
    expect(element.style.fontSize).toBe(16);
  });

  it('应该正确创建图片类型的PptElement', () => {
    const element: PptElement = {
      id: 'element-2',
      type: 'image' as PptNodeType,
      rect: { x: 0, y: 0, width: 300, height: 200 },
      style: {},
      content: 'ppt-image-rId1',
      props: { imgId: 'rId1', alt: '图片描述' },
    };

    expect(element.type).toBe('image');
    expect(element.content).toBe('ppt-image-rId1');
    expect(element.props.imgId).toBe('rId1');
  });

  it('应该正确创建形状类型的PptElement', () => {
    const element: PptElement = {
      id: 'element-3',
      type: 'shape' as PptNodeType,
      rect: { x: 50, y: 50, width: 100, height: 100 },
      style: {
        backgroundColor: '#ff0000',
        borderColor: '#000000',
        borderWidth: 2,
      },
      content: 'rect',
      props: { borderRadius: 5 },
    };

    expect(element.type).toBe('shape');
    expect(element.content).toBe('rect');
    expect(element.style.backgroundColor).toBe('#ff0000');
  });

  it('应该正确创建表格类型的PptElement', () => {
    const element: PptElement = {
      id: 'element-4',
      type: 'table' as PptNodeType,
      rect: { x: 0, y: 0, width: 400, height: 200 },
      style: {},
      content: [
        ['表头1', '表头2', '表头3'],
        ['数据1', '数据2', '数据3'],
        ['数据4', '数据5', '数据6'],
      ],
      props: { rowCount: 3, colCount: 3 },
    };

    expect(element.type).toBe('table');
    expect(Array.isArray(element.content)).toBe(true);
    expect(element.content).toHaveLength(3);
    expect(element.props.rowCount).toBe(3);
  });

  it('应该正确创建图表类型的PptElement', () => {
    const element: PptElement = {
      id: 'element-5',
      type: 'chart' as PptNodeType,
      rect: { x: 100, y: 100, width: 500, height: 300 },
      style: {},
      content: {
        type: 'bar',
        data: [10, 20, 30, 40],
      },
      props: { legend: true, grid: true },
    };

    expect(element.type).toBe('chart');
    expect(typeof element.content).toBe('object');
    expect((element.content as any).type).toBe('bar');
  });

  it('应该正确创建容器类型的PptElement', () => {
    const element: PptElement = {
      id: 'element-6',
      type: 'container' as PptNodeType,
      rect: { x: 0, y: 0, width: 800, height: 600 },
      style: {},
      content: '',
      props: {},
      children: [
        {
          id: 'child-1',
          type: 'text' as PptNodeType,
          rect: { x: 10, y: 10, width: 100, height: 50 },
          style: {},
          content: '子元素',
          props: {},
        },
      ],
    };

    expect(element.type).toBe('container');
    expect(element.children).toBeDefined();
    expect(element.children).toHaveLength(1);
    expect(element.children![0].type).toBe('text');
  });

  it('应该正确创建媒体类型的PptElement', () => {
    const element: PptElement = {
      id: 'element-7',
      type: 'media' as PptNodeType,
      rect: { x: 0, y: 0, width: 320, height: 240 },
      style: {},
      content: 'video-url',
      props: { mediaType: 'video', duration: 30 },
    };

    expect(element.type).toBe('media');
    expect(element.props.mediaType).toBe('video');
  });

  it('应该正确处理所有支持的元素类型', () => {
    const types: PptNodeType[] = ['text', 'image', 'shape', 'table', 'chart', 'container', 'media'];

    types.forEach(type => {
      const element: PptElement = {
        id: `element-${type}`,
        type,
        rect: { x: 0, y: 0, width: 100, height: 100 },
        style: {},
        content: '',
        props: {},
      };

      expect(element.type).toBe(type);
    });
  });

  it('应该正确处理完整的文档结构', () => {
    const doc: PptDocument = {
      id: 'doc-complete',
      title: '完整文档',
      slides: [
        {
          id: 'slide-1',
          title: '第一页',
          bgColor: '#ffffff',
          elements: [
            {
              id: 'el-1',
              type: 'text',
              rect: { x: 100, y: 100, width: 200, height: 50 },
              style: { fontSize: 18, color: '#333333' },
              content: '标题',
              props: {},
            },
            {
              id: 'el-2',
              type: 'shape',
              rect: { x: 100, y: 200, width: 100, height: 100 },
              style: { backgroundColor: '#ff0000' },
              content: 'circle',
              props: {},
            },
          ],
          props: {
            width: 1280,
            height: 720,
            slideLayout: 'title',
          },
        },
        {
          id: 'slide-2',
          title: '第二页',
          bgColor: '#f0f0f0',
          elements: [
            {
              id: 'el-3',
              type: 'table',
              rect: { x: 50, y: 50, width: 400, height: 200 },
              style: {},
              content: [['A', 'B'], ['C', 'D']],
              props: { rowCount: 2, colCount: 2 },
            },
          ],
          props: {
            width: 1280,
            height: 720,
            slideLayout: 'content',
          },
        },
      ],
      props: {
        width: 1280,
        height: 720,
        ratio: 1.78,
      },
    };

    expect(doc.slides).toHaveLength(2);
    expect(doc.slides[0].elements).toHaveLength(2);
    expect(doc.slides[1].elements).toHaveLength(1);
    expect(doc.slides[1].elements[0].type).toBe('table');
  });

  it('应该支持不同的样式属性', () => {
    const element: PptElement = {
      id: 'style-test',
      type: 'text',
      rect: { x: 0, y: 0, width: 100, height: 50 },
      style: {
        fontSize: 24,
        color: '#ff0000',
        fontWeight: 'bold',
        textAlign: 'center',
        backgroundColor: '#ffffff',
        borderColor: '#000000',
        borderWidth: 3,
      },
      content: '样式测试',
      props: {},
    };

    expect(element.style.fontSize).toBe(24);
    expect(element.style.fontWeight).toBe('bold');
    expect(element.style.textAlign).toBe('center');
    expect(element.style.borderWidth).toBe(3);
  });

  it('应该支持不同的rect属性', () => {
    const rects = [
      { x: 0, y: 0, width: 100, height: 100 },
      { x: 50, y: 50, width: 200, height: 150 },
      { x: 100, y: 200, width: 300, height: 400 },
    ];

    rects.forEach(rect => {
      expect(rect.x).toBeGreaterThanOrEqual(0);
      expect(rect.y).toBeGreaterThanOrEqual(0);
      expect(rect.width).toBeGreaterThan(0);
      expect(rect.height).toBeGreaterThan(0);
    });
  });
});
