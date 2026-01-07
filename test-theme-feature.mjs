/**
 * 主题解析功能测试
 * 验证主题解析和 CSS 生成是否正常工作
 */

import { parsePptx } from './dist/index';

async function testTheme() {
  console.log('=== 开始测试主题解析 ===\n');

  // 创建一个测试用的 PPTX 缓冲区
  // 这里应该使用实际的 PPTX 文件进行测试
  console.log('提示: 请提供一个实际的 PPTX 文件路径');
  console.log('示例代码:');

  console.log(`
  import * as fs from 'fs';
  import { parsePptx, DocumentElement } from './dist/index';

  async function runTest() {
    // 读取 PPTX 文件
    const pptxBuffer = fs.readFileSync('your-presentation.pptx');

    // 解析 PPTX
    const result = await parsePptx(pptxBuffer);

    // 创建文档元素
    const doc = DocumentElement.fromParseResult(result);

    // 检查主题
    if (doc.theme) {
      console.log('主题名称:', doc.theme.name);
      console.log('主题类前缀:', doc.theme.getThemeClassPrefix());
      console.log('主题颜色:', doc.theme.colors);

      // 生成主题 CSS
      const themeCSS = doc.theme.generateThemeCSS();
      console.log('\n生成的主题 CSS (前500字符):');
      console.log(themeCSS.substring(0, 500) + '...');

      // 生成完整 HTML
      const html = doc.toHTML({ includeStyles: true });

      // 保存文件
      fs.writeFileSync('theme-test-output.html', html);
      console.log('\nHTML 已保存到 theme-test-output.html');
      fs.writeFileSync('theme-test.css', themeCSS);
      console.log('CSS 已保存到 theme-test.css');
    } else {
      console.log('警告: PPTX 文件不包含主题');
    }
  }

  runTest().catch(console.error);
  `);
}

testTheme();
