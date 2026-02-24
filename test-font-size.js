import { PPTXStyleUtils } from './src/js/utils/style.js';
import { PPTXXmlUtils } from './src/js/utils/xml.js';

// 模拟 warpObj 对象
const warpObj = {
    slideLayoutTables: {},
    slideMasterTables: {},
    slideMasterTextStyles: {
        'p:titleStyle': {
            'a:lvl0pPr': {
                'a:defRPr': {
                    attrs: {
                        sz: '4800', // 48pt
                        kern: '1200'
                    }
                }
            }
        },
        'p:bodyStyle': {
            'a:lvl0pPr': {
                'a:defRPr': {
                    attrs: {
                        sz: '1600', // 16pt
                        kern: '1200'
                    }
                }
            }
        }
    },
    defaultTextStyle: {
        'a:lvl0pPr': {
            'a:defRPr': {
                attrs: {
                    sz: '2400', // 24pt
                    kern: '1200'
                }
            }
        }
    }
};

// 模拟 textBodyNode 对象
const textBodyNode = {
    'a:lstStyle': {}
};

// 测试标题文本（应该使用 48pt）
const titleNode = {
    'a:rPr': {
        attrs: {}
    }
};

// 测试正文文本（应该使用 16pt）
const bodyNode = {
    'a:rPr': {
        attrs: {}
    }
};

// 测试字体大小转换
console.log('Testing font size conversion...');
console.log('==================================');

// 测试标题文本字体大小
const titleFontSize = PPTXStyleUtils.getFontSize(titleNode, textBodyNode, null, 0, 'title', warpObj);
console.log('Title font size:', titleFontSize);

// 测试正文文本字体大小
const bodyFontSize = PPTXStyleUtils.getFontSize(bodyNode, textBodyNode, null, 0, 'body', warpObj);
console.log('Body font size:', bodyFontSize);

// 验证结果
console.log('==================================');
console.log('Expected title font size: ~64px (48pt * 1.333)');
console.log('Expected body font size: ~21.33px (16pt * 1.333)');
console.log('==================================');
console.log('Test completed!');
