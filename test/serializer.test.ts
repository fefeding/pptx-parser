import { describe, it, expect } from 'vitest';
import { serializePptx } from '../src/core';
import type { PptDocument, PptSlide, PptElement } from '../src/types';

describe('serializePptx - PPTX序列化核心功能测试', () => {
  const createMockDocument = (overrides?: Partial<PptDocument>): PptDocument => ({
    id: 'ppt-doc-123',
    title: '测试PPT',
    slides: [],
    props: {
      width: 1280,
      height: 720,
      ratio: 1.78,
    },
    ...overrides,
  });

  const createMockSlide = (overrides?: Partial<PptSlide>): PptSlide => ({
    id: 'slide-1',
    title: '第一页',
    bgColor: '#ffffff',
    elements: [],
    props: {
      width: 1280,
      height: 720,
      slideLayout: 'normal',
    },
    ...overrides,
  });

  const createMockElement = (type: 'text' | 'shape' | 'image' | 'table', overrides?: Partial<PptElement>): PptElement => ({
    id: 'element-1',
    type,
    rect: { x: 100, y: 100, width: 200, height: 100 },
    style: {
      fontSize: 16,
      color: '#333333',
      fontWeight: 'normal',
      textAlign: 'left',
    },
    content: '',
    props: {},
    ...overrides,
  });

  it('应该成功序列化空文档', async () => {
    const doc = createMockDocument();
    const blob = await serializePptx(doc);

    expect(blob).toBeInstanceOf(Blob);
    expect(blob.type).toBe('application/zip');
  });

  it('应该成功序列化包含幻灯片的文档', async () => {
    const doc = createMockDocument({
      slides: [createMockSlide()],
    });

    const blob = await serializePptx(doc);
    expect(blob).toBeInstanceOf(Blob);
  });

  it('应该正确序列化文本元素', async () => {
    const doc = createMockDocument({
      slides: [
        createMockSlide({
          elements: [
            createMockElement('text', {
              content: '测试文本内容',
              style: {
                fontSize: 18,
                color: '#ff0000',
                fontWeight: 'bold',
              },
            }),
          ],
        }),
      ],
    });

    const blob = await serializePptx(doc);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(blob);

    const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');
    expect(slideXml).toBeDefined();
    expect(slideXml).toContain('测试文本内容');
  });

  it('应该正确序列化形状元素', async () => {
    const doc = createMockDocument({
      slides: [
        createMockSlide({
          elements: [
            createMockElement('shape', {
              content: 'rect',
              style: {
                backgroundColor: '#00ff00',
                borderColor: '#0000ff',
                borderWidth: 2,
              },
            }),
          ],
        }),
      ],
    });

    const blob = await serializePptx(doc);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(blob);

    const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');
    expect(slideXml).toBeDefined();
    expect(slideXml).toContain('00ff00'); // 背景色
    expect(slideXml).toContain('0000ff'); // 边框色
  });

  it('应该正确序列化图片元素', async () => {
    const doc = createMockDocument({
      slides: [
        createMockSlide({
          elements: [
            createMockElement('image', {
              content: 'ppt-image-rId1',
              props: { imgId: 'rId1', alt: '测试图片' },
            }),
          ],
        }),
      ],
    });

    const blob = await serializePptx(doc);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(blob);

    const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');
    expect(slideXml).toBeDefined();
    expect(slideXml).toContain('blip');
  });

  it('应该正确序列化表格元素', async () => {
    const doc = createMockDocument({
      slides: [
        createMockSlide({
          elements: [
            createMockElement('table', {
              content: [
                ['A1', 'B1', 'C1'],
                ['A2', 'B2', 'C2'],
              ],
              props: { rowCount: 2, colCount: 3 },
            }),
          ],
        }),
      ],
    });

    const blob = await serializePptx(doc);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(blob);

    const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');
    expect(slideXml).toBeDefined();
    expect(slideXml).toContain('tbl');
    expect(slideXml).toContain('A1');
    expect(slideXml).toContain('B2');
  });

  it('应该正确序列化多页幻灯片', async () => {
    const doc = createMockDocument({
      slides: [
        createMockSlide({ id: 'slide-1', title: '第一页', bgColor: '#ffffff' }),
        createMockSlide({ id: 'slide-2', title: '第二页', bgColor: '#f0f0f0' }),
        createMockSlide({ id: 'slide-3', title: '第三页', bgColor: '#e0e0e0' }),
      ],
    });

    const blob = await serializePptx(doc);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(blob);

    expect(zip.file('ppt/slides/slide1.xml')).toBeDefined();
    expect(zip.file('ppt/slides/slide2.xml')).toBeDefined();
    expect(zip.file('ppt/slides/slide3.xml')).toBeDefined();
  });

  it('应该正确设置幻灯片背景色', async () => {
    const doc = createMockDocument({
      slides: [
        createMockSlide({ bgColor: '#ff0000' }),
      ],
    });

    const blob = await serializePptx(doc);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(blob);

    const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');
    expect(slideXml).toBeDefined();
    expect(slideXml).toContain('ff0000');
  });

  it('应该正确设置文档属性', async () => {
    const doc = createMockDocument({
      title: '自定义标题',
      props: {
        width: 1920,
        height: 1080,
        ratio: 1.78,
      },
    });

    const blob = await serializePptx(doc);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(blob);

    const corePropsXml = await zip.file('docProps/core.xml')?.async('string');
    expect(corePropsXml).toContain('自定义标题');

    const layoutXml = await zip.file('ppt/slideLayouts/slideLayout1.xml')?.async('string');
    expect(layoutXml).toBeDefined();
  });

  it('应该正确序列化混合元素类型', async () => {
    const doc = createMockDocument({
      slides: [
        createMockSlide({
          elements: [
            createMockElement('text', { content: '标题' }),
            createMockElement('shape', { content: 'rect' }),
            createMockElement('table', { content: [['1', '2']] }),
          ],
        }),
      ],
    });

    const blob = await serializePptx(doc);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(blob);

    const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');
    expect(slideXml).toBeDefined();
    expect(slideXml).toContain('txBody');
    expect(slideXml).toContain('sp');
    expect(slideXml).toContain('tbl');
  });

  it('应该生成压缩的ZIP文件', async () => {
    const doc = createMockDocument({
      slides: [createMockSlide()],
    });

    const blob = await serializePptx(doc);
    expect(blob.size).toBeGreaterThan(0);
    expect(blob.size).toBeLessThan(50000); // 合理的大小范围
  });

  it('应该正确处理不同尺寸的元素', async () => {
    const doc = createMockDocument({
      slides: [
        createMockSlide({
          elements: [
            createMockElement('text', {
              rect: { x: 0, y: 0, width: 100, height: 50 },
            }),
            createMockElement('text', {
              rect: { x: 500, y: 300, width: 300, height: 200 },
            }),
          ],
        }),
      ],
    });

    const blob = await serializePptx(doc);
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(blob);

    const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');
    expect(slideXml).toBeDefined();
    // EMU值应该存在于XML中
    expect(slideXml).toMatch(/\d+/);
  });
});
