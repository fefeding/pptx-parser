/**
 * 扩展功能使用示例
 * 展示如何使用渐变填充、项目符号、超链接、阴影等高级特性
 */

import { parsePptx, serializePptx } from '../src/core';
import { PptParseUtilsExtended } from '../src/core-extended';
import type { PptDocument, PptElement } from '../src/types';

/**
 * 示例 1: 解析包含渐变填充的 PPTX
 */
async function parseGradientExample(file: File) {
  const pptDoc = await parsePptx(file);

  // 查找包含渐变填充的形状
  pptDoc.slides.forEach(slide => {
    slide.elements.forEach(element => {
      if (element.type === 'shape') {
        const fill = element.style.fill;
        if (fill?.type === 'gradient') {
          console.log('渐变填充:', {
            stops: fill.gradientStops,
            direction: fill.gradientDirection,
          });
        }
      }
    });
  });
}

/**
 * 示例 2: 解析包含项目符号的文本
 */
async function parseBulletTextExample(file: File) {
  const pptDoc = await parsePptx(file);

  pptDoc.slides.forEach(slide => {
    slide.elements.forEach(element => {
      if (element.type === 'text') {
        const content = element.content as any[];
        content.forEach((paragraph: any) => {
          if (paragraph.bullet) {
            console.log('项目符号:', {
              type: paragraph.bullet.type,
              char: paragraph.bullet.char,
              level: paragraph.bullet.level,
              text: paragraph.text,
            });
          }
        });
      }
    });
  });
}

/**
 * 示例 3: 创建包含渐变填充的形状
 */
function createGradientShapeExample(): PptDocument {
  return {
    id: 'doc-gradient',
    title: '渐变填充示例',
    slides: [
      {
        id: 'slide-1',
        title: '渐变填充',
        bgColor: '#ffffff',
        elements: [
          {
            id: 'shape-1',
            type: 'shape',
            rect: { x: 100, y: 100, width: 400, height: 300 },
            style: {
              fontSize: 16,
              color: '#ffffff',
              fill: {
                type: 'gradient',
                gradientStops: [
                  { position: 0, color: '#ff6b6b' },
                  { position: 0.5, color: '#4ecdc4' },
                  { position: 1, color: '#45b7d1' },
                ],
                gradientDirection: 45,
              },
              shadow: {
                color: '#000000',
                blur: 10,
                offsetX: 5,
                offsetY: 5,
                opacity: 0.3,
              },
            },
            content: { shapeType: 'rectangle', text: '渐变形状' },
            props: {},
          },
        ],
        props: {
          width: 1280,
          height: 720,
          slideLayout: 'blank',
        },
      },
    ],
    props: {
      width: 1280,
      height: 720,
      ratio: 1.78,
    },
  };
}

/**
 * 示例 4: 创建包含项目符号的文本
 */
function createBulletTextExample(): PptDocument {
  return {
    id: 'doc-bullet',
    title: '项目符号示例',
    slides: [
      {
        id: 'slide-1',
        title: '项目符号',
        bgColor: '#ffffff',
        elements: [
          {
            id: 'text-1',
            type: 'text',
            rect: { x: 100, y: 100, width: 1080, height: 520 },
            style: {
              fontSize: 18,
              lineHeight: 1.8,
              color: '#333333',
            },
            content: [
              { text: '• 一级项目符号', bullet: { type: 'bullet', level: 0 } },
              { text: '  • 二级项目符号', bullet: { type: 'bullet', level: 1 } },
              { text: '  • 二级项目符号', bullet: { type: 'bullet', level: 1 } },
              { text: '• 一级项目符号', bullet: { type: 'bullet', level: 0 } },
              { text: '1. 编号列表项 1', bullet: { type: 'numbered', level: 0 } },
              { text: '2. 编号列表项 2', bullet: { type: 'numbered', level: 0 } },
              { text: '  2.1 二级编号项', bullet: { type: 'numbered', level: 1 } },
            ],
            props: {},
          },
        ],
        props: {
          width: 1280,
          height: 720,
          slideLayout: 'blank',
        },
      },
    ],
    props: {
      width: 1280,
      height: 720,
      ratio: 1.78,
    },
  };
}

/**
 * 示例 5: 创建包含超链接的文本
 */
function createHyperlinkExample(): PptDocument {
  return {
    id: 'doc-hyperlink',
    title: '超链接示例',
    slides: [
      {
        id: 'slide-1',
        title: '超链接',
        bgColor: '#ffffff',
        elements: [
          {
            id: 'text-1',
            type: 'text',
            rect: { x: 100, y: 200, width: 1080, height: 300 },
            style: {
              fontSize: 24,
              color: '#333333',
              textAlign: 'center',
            },
            content: [
              {
                text: '访问 GitHub',
                hyperlink: {
                  url: 'https://github.com',
                  tooltip: '点击访问 GitHub',
                },
              },
            ],
            props: {},
          },
        ],
        props: {
          width: 1280,
          height: 720,
          slideLayout: 'blank',
        },
      },
    ],
    props: {
      width: 1280,
      height: 720,
      ratio: 1.78,
    },
  };
}

/**
 * 示例 6: 创建包含阴影效果的形状
 */
function createShadowExample(): PptDocument {
  return {
    id: 'doc-shadow',
    title: '阴影效果示例',
    slides: [
      {
        id: 'slide-1',
        title: '阴影效果',
        bgColor: '#f5f5f5',
        elements: [
          {
            id: 'shape-1',
            type: 'shape',
            rect: { x: 200, y: 150, width: 300, height: 200 },
            style: {
              backgroundColor: '#ffffff',
              fill: {
                type: 'solid',
                color: '#ffffff',
              },
              shadow: {
                color: '#000000',
                blur: 15,
                offsetX: 8,
                offsetY: 8,
                opacity: 0.4,
              },
            },
            content: { shapeType: 'rectangle' as const, text: '带阴影的形状' },
            props: {},
          },
          {
            id: 'shape-2',
            type: 'shape',
            rect: { x: 600, y: 150, width: 300, height: 200 },
            style: {
              backgroundColor: '#4ecdc4',
              fill: {
                type: 'solid',
                color: '#4ecdc4',
              },
              shadow: {
                color: '#ff6b6b',
                blur: 20,
                offsetX: 10,
                offsetY: 10,
                opacity: 0.5,
              },
            },
            content: { shapeType: 'rectangle' as const, text: '彩色阴影' },
            props: {},
          },
        ],
        props: {
          width: 1280,
          height: 720,
          slideLayout: 'blank',
        },
      },
    ],
    props: {
      width: 1280,
      height: 720,
      ratio: 1.78,
    },
  };
}

/**
 * 示例 7: 创建带旋转和翻转的形状
 */
function createTransformExample(): PptDocument {
  return {
    id: 'doc-transform',
    title: '变换效果示例',
    slides: [
      {
        id: 'slide-1',
        title: '变换效果',
        bgColor: '#ffffff',
        elements: [
          {
            id: 'shape-1',
            type: 'shape',
            rect: { x: 200, y: 150, width: 200, height: 200 },
            transform: {
              rotate: 45, // 旋转45度
            },
            style: {
              backgroundColor: '#ff6b6b',
              fill: {
                type: 'solid',
                color: '#ff6b6b',
              },
            },
            content: { shapeType: 'rectangle' as const },
            props: {},
          },
          {
            id: 'shape-2',
            type: 'shape',
            rect: { x: 500, y: 150, width: 200, height: 200 },
            transform: {
              rotate: 90,
            },
            style: {
              backgroundColor: '#4ecdc4',
              fill: {
                type: 'solid',
                color: '#4ecdc4',
              },
            },
            content: { shapeType: 'rectangle' as const },
            props: {},
          },
          {
            id: 'shape-3',
            type: 'shape',
            rect: { x: 800, y: 150, width: 200, height: 200 },
            transform: {
              flipH: true, // 水平翻转
            },
            style: {
              backgroundColor: '#45b7d1',
              fill: {
                type: 'solid',
                color: '#45b7d1',
              },
            },
            content: { shapeType: 'rectangle' as const },
            props: {},
          },
        ],
        props: {
          width: 1280,
          height: 720,
          slideLayout: 'blank',
        },
      },
    ],
    props: {
      width: 1280,
      height: 720,
      ratio: 1.78,
    },
  };
}

/**
 * 示例 8: 创建包含多种边框样式的形状
 */
function createBorderStyleExample(): PptDocument {
  return {
    id: 'doc-border',
    title: '边框样式示例',
    slides: [
      {
        id: 'slide-1',
        title: '边框样式',
        bgColor: '#ffffff',
        elements: [
          {
            id: 'shape-1',
            type: 'shape',
            rect: { x: 100, y: 100, width: 300, height: 150 },
            style: {
              backgroundColor: '#ffffff',
              borderColor: '#000000',
              borderWidth: 3,
              borderStyle: 'solid',
              fill: { type: 'solid', color: '#ffffff' },
            },
            content: { shapeType: 'rectangle' as const },
            props: {},
          },
          {
            id: 'shape-2',
            type: 'shape',
            rect: { x: 500, y: 100, width: 300, height: 150 },
            style: {
              backgroundColor: '#ffffff',
              borderColor: '#ff0000',
              borderWidth: 3,
              borderStyle: 'dashed',
              fill: { type: 'solid', color: '#ffffff' },
            },
            content: { shapeType: 'rectangle' as const },
            props: {},
          },
          {
            id: 'shape-3',
            type: 'shape',
            rect: { x: 900, y: 100, width: 300, height: 150 },
            style: {
              backgroundColor: '#ffffff',
              borderColor: '#0000ff',
              borderWidth: 3,
              borderStyle: 'dotted',
              fill: { type: 'solid', color: '#ffffff' },
            },
            content: { shapeType: 'rectangle' as const },
            props: {},
          },
          {
            id: 'shape-4',
            type: 'shape',
            rect: { x: 300, y: 350, width: 400, height: 150 },
            style: {
              backgroundColor: '#ffffff',
              borderColor: '#00ff00',
              borderWidth: 5,
              borderStyle: 'double',
              fill: { type: 'solid', color: '#ffffff' },
            },
            content: { shapeType: 'rectangle' as const },
            props: {},
          },
        ],
        props: {
          width: 1280,
          height: 720,
          ratio: 1.78,
        },
      },
    ],
    props: {
      width: 1280,
      height: 720,
      ratio: 1.78,
    },
  };
}

// 导出示例
export async function exportExamples() {
  // 导出渐变示例
  const gradientDoc = createGradientShapeExample();
  const gradientBlob = await serializePptx(gradientDoc);
  downloadBlob(gradientBlob, 'gradient-example.pptx');

  // 导出项目符号示例
  const bulletDoc = createBulletTextExample();
  const bulletBlob = await serializePptx(bulletDoc);
  downloadBlob(bulletBlob, 'bullet-example.pptx');

  // 导出超链接示例
  const hyperlinkDoc = createHyperlinkExample();
  const hyperlinkBlob = await serializePptx(hyperlinkDoc);
  downloadBlob(hyperlinkBlob, 'hyperlink-example.pptx');

  // 导出阴影示例
  const shadowDoc = createShadowExample();
  const shadowBlob = await serializePptx(shadowDoc);
  downloadBlob(shadowBlob, 'shadow-example.pptx');

  // 导出变换示例
  const transformDoc = createTransformExample();
  const transformBlob = await serializePptx(transformDoc);
  downloadBlob(transformBlob, 'transform-example.pptx');

  // 导出边框示例
  const borderDoc = createBorderStyleExample();
  const borderBlob = await serializePptx(borderDoc);
  downloadBlob(borderBlob, 'border-style-example.pptx');
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
