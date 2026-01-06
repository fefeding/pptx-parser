/**
 * 测试真实PPTX文件的Layout解析
 */

import { readFileSync } from 'fs';
import { parsePptx } from 'pptx-parser';

async function testLayoutParsing() {
  try {
    // 读取PPTX文件
    const pptxPath = 'c:\\Users\\fefeding\\Desktop\\金腾研发架构&系统介绍.pptx';
    const buffer = readFileSync(pptxPath);

    console.log('开始解析PPTX文件...');
    const result = await parsePptx(buffer, {
      parseImages: true,
      keepRawXml: false,
      verbose: true
    });

    console.log('\n=== 解析结果 ===');
    console.log(`标题: ${result.title}`);
    console.log(`幻灯片数量: ${result.slides.length}`);
    console.log(`布局数量: ${result.slideLayouts ? Object.keys(result.slideLayouts).length : 0}`);
    console.log(`母版数量: ${result.masterSlides ? result.masterSlides.length : 0}`);

    // 显示布局信息
    if (result.slideLayouts) {
      console.log('\n=== 布局信息 ===');
      Object.entries(result.slideLayouts).forEach(([id, layout]) => {
        console.log(`\n布局 ${id}:`);
        console.log(`  名称: ${layout.name || '无'}`);
        console.log(`  背景:`, layout.background);
      });
    }

    // 显示每个slide的背景信息
    console.log('\n=== 幻灯片背景信息 ===');
    result.slides.forEach((slide, index) => {
      console.log(`\n幻灯片 ${index + 1}:`);
      console.log(`  背景:`, slide.background);
    });

    console.log('\n解析完成！');
  } catch (error) {
    console.error('解析失败:', error);
  }
}

testLayoutParsing();
