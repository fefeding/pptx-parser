/**
 * 测试主题解析和 CSS 生成功能
 */

import { parsePptx } from './dist/index';
import { DocumentElement } from './dist/elements/DocumentElement';
import * as fs from 'fs';

async function testThemeParsing() {
  try {
    // 读取测试PPT文件
    const pptxPath = 'test.pptx';
    const pptxBuffer = fs.readFileSync(pptxPath);

    // 解析PPT
    const result = await parsePptx(pptxBuffer);

    console.log('=== 主题解析测试 ===');
    console.log('主题名称:', result.theme?.name);
    console.log('主题颜色:', result.theme?.colors);

    // 创建文档元素
    const doc = DocumentElement.fromParseResult(result);

    console.log('文档主题:', doc.theme?.name);
    console.log('主题类前缀:', doc.theme?.getThemeClassPrefix());
    console.log('主题CSS类示例:', doc.theme?.getThemeClass('accent1'));

    // 生成HTML并保存
    const html = doc.toHTML({ includeStyles: true });

    // 保存到文件
    fs.writeFileSync('test-theme-output.html', html);
    console.log('HTML已保存到 test-theme-output.html');

    // 提取主题CSS
    if (doc.theme) {
      const themeCSS = doc.theme.generateThemeCSS();
      fs.writeFileSync('test-theme-css.css', themeCSS);
      console.log('主题CSS已保存到 test-theme-css.css');
    }

  } catch (error) {
    console.error('测试失败:', error);
  }
}

testThemeParsing();
