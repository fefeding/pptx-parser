/**
 * PPTX Parser 增强版使用示例
 *
 * 演示如何使用增强版的 parsePptx 函数解析PPTX文件
 */

import { parsePptx, type PptxParseResult } from '../src/core/parser';
import type { ParseOptions } from '../src/types-enhanced';

// ============ 基础使用示例 ============

/**
 * 示例1：最简单的使用方式
 * 上传PPTX文件并获取基本信息
 */
async function example1_BasicUsage(file: File) {
  // 最简单的调用方式
  const result = await parsePptx(file);

  console.log('PPT标题:', result.title);
  console.log('作者:', result.author);
  console.log('幻灯片数量:', result.slides.length);
  console.log('页面尺寸:', `${result.props.width}x${result.props.height}`);
}

/**
 * 示例2：带选项的解析
 * 控制解析行为，如是否解析图片、是否保留原始XML等
 */
async function example2_WithOptions(file: File) {
  const options: ParseOptions = {
    parseImages: true,    // 解析图片为Base64（默认true）
    keepRawXml: false,    // 保留原始XML字符串（默认false）
    verbose: true          // 详细日志输出（默认false）
  };

  const result = await parsePptx(file, options);
  console.log('解析完成，包含', result.slides.length, '页幻灯片');
}

// ============ 访问幻灯片元素 ============

/**
 * 示例3：遍历所有幻灯片和元素
 */
async function example3_TraverseSlides(file: File) {
  const result = await parsePptx(file);

  result.slides.forEach((slide, slideIndex) => {
    console.log(`\n=== 幻灯片 ${slideIndex + 1}: ${slide.title} ===`);
    console.log(`背景色: ${slide.background}`);
    console.log(`元素数量: ${slide.elements.length}`);

    slide.elements.forEach((element, elementIndex) => {
      console.log(`  元素${elementIndex + 1}: ${element.type}`);
      console.log(`    ID: ${element.id}`);
      console.log(`    位置: (${element.rect.x}, ${element.rect.y})`);
      console.log(`    尺寸: ${element.rect.width}x${element.rect.height}`);

      // 根据元素类型访问特定属性
      if (element.type === 'text' || element.type === 'shape') {
        console.log(`    文本: ${element.text || '(无文本)'}`);
      } else if (element.type === 'image') {
        console.log(`    图片: ${element.src || '(未解析)'}`);
        console.log(`    关联ID: ${element.relId}`);
      } else if (element.type === 'ole') {
        console.log(`    OLE对象: ${element.progId || '(未知类型)'}`);
      }
    });
  });
}

// ============ 搜索和过滤元素 ============

/**
 * 示例4：搜索包含特定文本的元素
 */
async function example4_SearchText(file: File, searchText: string) {
  const result = await parsePptx(file);
  const matches: Array<{ slideIndex: number; element: any }> = [];

  result.slides.forEach((slide, slideIndex) => {
    slide.elements.forEach(element => {
      if (element.text && element.text.includes(searchText)) {
        matches.push({ slideIndex, element });
      }
    });
  });

  console.log(`找到 ${matches.length} 个包含 "${searchText}" 的元素:`);
  matches.forEach(match => {
    console.log(`  幻灯片 ${match.slideIndex + 1}: ${match.element.text}`);
  });
}

/**
 * 示例5：过滤特定类型的元素
 */
async function example5_FilterByType(file: File) {
  const result = await parsePptx(file);

  // 获取所有图片元素
  const images = result.slides.flatMap(slide =>
    slide.elements.filter(e => e.type === 'image')
  );

  console.log(`找到 ${images.length} 个图片元素:`);
  images.forEach(img => {
    console.log(`  ${img.name || '(未命名)'}: ${img.relId}`);
  });

  // 获取所有文本元素
  const textElements = result.slides.flatMap(slide =>
    slide.elements.filter(e => e.type === 'text' || e.type === 'shape')
  );

  console.log(`找到 ${textElements.length} 个文本元素:`);
  textElements.forEach(el => {
    console.log(`  ${el.text || '(无文本)'}`);
  });
}

// ============ 处理图片资源 ============

/**
 * 示例6：提取所有图片
 */
async function example6_ExtractImages(file: File) {
  const result = await parsePptx(file, {
    parseImages: true // 启用图片解析
  });

  const images: Array<{
    slideIndex: number;
    elementId: string;
    src: string;
    relId: string;
  }> = [];

  result.slides.forEach((slide, slideIndex) => {
    slide.elements.forEach(element => {
      if (element.type === 'image') {
        const img = element as any;
        if (img.src && img.src.startsWith('data:')) {
          images.push({
            slideIndex,
            elementId: img.id,
            src: img.src,
            relId: img.relId
          });
        }
      }
    });
  });

  console.log(`提取到 ${images.length} 个图片:`);
  images.forEach((img, index) => {
    console.log(`  图片${index + 1}:`);
    console.log(`    幻灯片: ${img.slideIndex + 1}`);
    console.log(`    关联ID: ${img.relId}`);
    console.log(`    数据长度: ${img.src.length} 字符`);

    // 可以直接保存图片
    // saveImage(img.src, `image_${index}.png`);
  });
}

/**
 * 保存Base64图片到文件
 */
function saveImage(base64Data: string, filename: string) {
  const link = document.createElement('a');
  link.href = base64Data;
  link.download = filename;
  link.click();
}

// ============ 获取元数据 ============

/**
 * 示例7：获取PPT元数据
 */
async function example7_GetMetadata(file: File) {
  const result = await parsePptx(file);

  console.log('=== PPT元数据 ===');
  console.log('标题:', result.title);
  console.log('作者:', result.author);
  console.log('主题:', result.subject);
  console.log('关键词:', result.keywords);
  console.log('描述:', result.description);
  console.log('创建时间:', result.created);
  console.log('修改时间:', result.modified);
  console.log('页面尺寸:', `${result.props.width}x${result.props.height}`);
  console.log('页面比例:', result.props.pageSize);
}

// ============ 处理分组元素 ============

/**
 * 示例8：递归处理分组元素
 */
async function example8_HandleGroups(file: File) {
  const result = await parsePptx(file);

  function processElement(element: any, depth: number = 0): void {
    const indent = '  '.repeat(depth);
    console.log(`${indent}└─ ${element.type}: ${element.name || '(未命名)'}`);

    if (element.type === 'group' && element.children) {
      element.children.forEach(child => processElement(child, depth + 1));
    }
  }

  result.slides.forEach((slide, slideIndex) => {
    console.log(`\n=== 幻灯片 ${slideIndex + 1}: ${slide.title} ===`);
    slide.elements.forEach(element => processElement(element));
  });
}

// ============ 错误处理 ============

/**
 * 示例9：完善的错误处理
 */
async function example9_ErrorHandling(file: File) {
  try {
    const result = await parsePptx(file);

    if (!result || result.slides.length === 0) {
      console.warn('警告: PPT文件没有幻灯片');
      return;
    }

    console.log('解析成功');
  } catch (error) {
    console.error('解析失败:', error);

    if (error instanceof Error) {
      if (error.message.includes('ZIP')) {
        console.error('错误: 文件不是有效的PPTX格式');
      } else if (error.message.includes('XML')) {
        console.error('错误: PPT文件包含损坏的XML');
      }
    }
  }
}

// ============ 完整的Vue组件示例 ============

/**
 * 示例10：在Vue组件中使用
 */
// import { parsePptx } from 'pptx-parser';
//
// export default {
//   data() {
//     return {
//       pptxResult: null,
//       loading: false,
//       error: null
//     };
//   },
//   methods: {
//     async handleFileUpload(event: Event) {
//       const file = (event.target as HTMLInputElement).files?.[0];
//       if (!file) return;
//
//       this.loading = true;
//       try {
//         this.pptxResult = await parsePptx(file, {
//           parseImages: true,
//           verbose: true
//         });
//         console.log('解析成功:', this.pptxResult);
//       } catch (err) {
//         this.error = err.message;
//       } finally {
//         this.loading = false;
//       }
//     }
//   }
// };

// ============ 导出所有示例函数 ============
export {
  example1_BasicUsage,
  example2_WithOptions,
  example3_TraverseSlides,
  example4_SearchText,
  example5_FilterByType,
  example6_ExtractImages,
  example7_GetMetadata,
  example8_HandleGroups,
  example9_ErrorHandling
};
