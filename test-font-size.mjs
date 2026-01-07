/**
 * 测试字体大小转换和段落样式解析
 */

import { ShapeElement } from './src/elements/ShapeElement';

console.log('=== 字体大小转换测试 ===');
console.log('30pt 应该转换为: ', 30 * (4/3), 'px');
console.log('40pt 应该转换为: ', 40 * (4/3), 'px');
console.log('18pt 应该转换为: ', 18 * (4/3), 'px');

console.log('\n=== 段落样式测试 ===');
const paragraphStyle = {
  spaceBefore: 9,
  spaceAfter: 0,
  paddingTop: 0,
  paddingBottom: 44,
  marginLeft: 0,
  marginRight: 0
};

console.log('段落间距参数:', paragraphStyle);
console.log('转换为CSS样式:');
console.log(`  padding-top: ${paragraphStyle.spaceBefore}px`);
console.log(`  padding-bottom: ${paragraphStyle.spaceAfter}px`);
console.log(`  margin-top: ${paragraphStyle.paddingTop}px`);
console.log(`  margin-bottom: ${paragraphStyle.paddingBottom}px`);
