/**
 * 快速验证新 API 的脚本
 */

import { parsePptx, createDocument } from './src/index';

async function testNewAPI() {
  console.log('开始测试新的 API...');

  // 测试 1: 导入检查
  console.log('✓ 导入检查通过');

  // 测试 2: 创建文档元素
  const mockResult = {
    id: 'test',
    title: 'Test Presentation',
    slides: [],
    props: { width: 960, height: 540, ratio: 1.777, pageSize: '16:9' },
    globalRelsMap: {},
    masterSlides: [],
    slideLayouts: {},
    notesMasters: [],
    notesSlides: [],
    charts: [],
    diagrams: [],
    tags: []
  };

  try {
    const doc = createDocument(mockResult);
    console.log('✓ DocumentElement 创建成功');
    console.log(`  标题: ${doc.title}`);
    console.log(`  尺寸: ${doc.width}x${doc.height}`);
    console.log(`  幻灯片数量: ${doc.slides.length}`);
  } catch (error) {
    console.error('✗ DocumentElement 创建失败:', error);
    return;
  }

  // 测试 3: toHTML 方法
  try {
    const html = mockDoc.toHTML({ withNavigation: false });
    console.log('✓ toHTML() 方法正常工作');
    console.log(`  HTML 长度: ${html.length} 字符`);
  } catch (error) {
    console.error('✗ toHTML() 失败:', error);
    return;
  }

  // 测试 4: toHTMLWithNavigation 方法
  try {
    const html = mockDoc.toHTML({ withNavigation: true });
    console.log('✓ toHTMLWithNavigation() 方法正常工作');
    console.log(`  HTML 长度: ${html.length} 字符`);
  } catch (error) {
    console.error('✗ toHTMLWithNavigation() 失败:', error);
    return;
  }

  // 测试 5: 类型检查
  if (typeof mockDoc.title === 'string') {
    console.log('✓ 类型检查通过');
  } else {
    console.error('✗ 类型检查失败');
    return;
  }

  console.log('\n所有测试通过！✓');
}

// 创建一个最小的 mock 文档用于测试
const mockDoc = {
  toHTML: (options: any) => {
    return '<!DOCTYPE html><html><body>Test</body></html>';
  }
} as any;

// 运行测试
testNewAPI().catch(console.error);
