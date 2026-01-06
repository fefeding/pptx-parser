import { readFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import { parsePptx } from './dist/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

async function main() {
  const pptxBuffer = readFileSync(join(__dirname, 'examples/3guo.pptx'));
  
  console.log('Parsing PPTX with images...');
  const result = await parsePptx(pptxBuffer, { parseImages: true });
  
  console.log(`\nTotal slides: ${result.slides.length}`);
  console.log(`Total layouts: ${Object.keys(result.slideLayouts || {}).length}`);
  console.log(`Total masters: ${result.masterSlides?.length || 0}`);
  console.log(`Media map size: ${result.mediaMap?.size || 0}`);
  
  // 打印所有关系映射
  console.log('\n=== Layouts ===');
  for (const [layoutId, layout] of Object.entries(result.slideLayouts || {})) {
    console.log(`\nLayout ${layoutId}:`);
    console.log(`  Name: ${layout.name}`);
    console.log(`  Background:`, layout.background);
    console.log(`  Elements: ${layout.elements.length}`);
    console.log(`  RelsMap entries: ${Object.keys(layout.relsMap).length}`);
    for (const [relId, rel] of Object.entries(layout.relsMap)) {
      console.log(`    ${relId}: type=${rel.type}, target=${rel.target}`);
    }
    // 检查图片元素
    for (const element of layout.elements) {
      if (element.type === 'image') {
        console.log(`  Image element: relId=${element.relId}, src length=${element.src?.length || 0}`);
        if (element.src && element.src.startsWith('data:')) {
          console.log(`    Base64 prefix: ${element.src.substring(0, 60)}...`);
        }
      }
    }
  }
  
  // 打印媒体映射内容
  console.log('\n=== Media Map ===');
  if (result.mediaMap) {
    for (const [relId, url] of result.mediaMap.entries()) {
      console.log(`  ${relId}: ${url.substring(0, 80)}...`);
    }
  }
  
  // 检查是否有错误图片
  console.log('\n=== Validation ===');
  let missingImages = 0;
  for (const [layoutId, layout] of Object.entries(result.slideLayouts || {})) {
    for (const element of layout.elements) {
      if (element.type === 'image' && (!element.src || !element.src.startsWith('data:'))) {
        console.log(`  WARNING: Layout ${layoutId} image element missing base64 src`);
        missingImages++;
      }
    }
  }
  console.log(`Total missing images: ${missingImages}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});