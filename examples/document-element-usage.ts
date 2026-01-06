/**
 * DocumentElement 使用示例
 * 展示如何使用新的文档元素架构进行 HTML 渲染
 */

import { parsePptx, createDocument, type HtmlRenderOptions } from '../src';

async function main() {
  // 1. 解析 PPTX 文件
  const result = await parsePptx('./test.pptx');

  // 2. 创建文档元素（包含所有文档信息）
  const doc = createDocument(result);

  console.log(`文档标题: ${doc.title}`);
  console.log(`作者: ${doc.author}`);
  console.log(`幻灯片数量: ${doc.slides.length}`);
  console.log(`文档尺寸: ${doc.width}x${doc.height}`);
  console.log(`布局数量: ${Object.keys(doc.layouts).length}`);
  console.log(`母版数量: ${doc.masters.length}`);

  // 3. 转换为 HTML（默认带导航）
  const html = doc.toHTML();
  // console.log(html);

  // 4. 转换为 HTML（不带导航，静态展示）
  const staticHtml = doc.toHTML({ withNavigation: false });
  // console.log(staticHtml);

  // 5. 自定义样式
  const options: HtmlRenderOptions = {
    includeStyles: true,
    includeScripts: true,
    includeLayoutElements: true,
    withNavigation: true,
    customCss: `
      /* 自定义幻灯片样式 */
      .ppt-slide {
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
      }
      /* 自定义文本样式 */
      .ppt-text {
        line-height: 1.6;
      }
    `
  };
  const customHtml = doc.toHTML(options);

  // 6. 访问文档内容
  // 获取第一个幻灯片
  const firstSlide = doc.getSlide(0);
  if (firstSlide) {
    console.log(`第一张幻灯片标题: ${firstSlide.title}`);
    console.log(`第一张幻灯片元素数量: ${firstSlide.elements.length}`);
  }

  // 获取指定布局
  const layout = doc.getLayout('layout1');
  if (layout) {
    console.log(`布局名称: ${layout.name}`);
    console.log(`占位符数量: ${layout.placeholders.length}`);
  }

  // 获取指定母版
  const master = doc.getMaster('master1');
  if (master) {
    console.log(`母版ID: ${master.id}`);
    console.log(`母版元素数量: ${master.elements.length}`);
  }

  // 7. 访问标签和备注
  if (doc.tags.length > 0) {
    console.log(`标签数量: ${doc.tags.length}`);
    const firstTag = doc.tags[0];
    console.log(`标签值: ${firstTag.getTag('name')}`);
  }

  if (doc.notesSlides.length > 0) {
    console.log(`备注数量: ${doc.notesSlides.length}`);
    const firstNote = doc.notesSlides[0];
    console.log(`备注内容: ${firstNote.text}`);
  }

  return doc;
}

// 导出示例函数
export async function convertPptxToHtml(filePath: string, outputPath: string, options?: HtmlRenderOptions) {
  const result = await parsePptx(filePath);
  const doc = createDocument(result);
  const html = doc.toHTML(options);

  // 写入文件（需要 fs 模块）
  // fs.writeFileSync(outputPath, html, 'utf-8');
  console.log(`HTML 已生成: ${outputPath}`);

  return html;
}

// 直接运行示例
if (require.main === module) {
  main().catch(console.error);
}
