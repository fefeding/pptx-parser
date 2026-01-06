/**
 * Parse a real PPTX file and output the result
 * Usage: npx ts-node examples/parse-real-pptx.ts
 */

import { parsePptx } from '../src/core/parser';
import * as fs from 'fs';
import * as path from 'path';

async function main() {
  const pptxPath = 'c:/Users/fefeding/Desktop/金腾研发架构&系统介绍/金腾研发架构&系统介绍.pptx';
  
  if (!fs.existsSync(pptxPath)) {
    console.error('PPTX file not found:', pptxPath);
    console.log('Please place the PPTX file at the correct path.');
    process.exit(1);
  }

  console.log('Parsing PPTX file:', pptxPath);
  console.log('='.repeat(60));

  try {
    const fileBuffer = fs.readFileSync(pptxPath);
    
    const result = await parsePptx(fileBuffer, {
      parseImages: true,
      returnFormat: 'enhanced'
    });

    console.log('\n📊 PPTX Information:');
    console.log('  Title:', result.title);
    console.log('  Author:', result.author || 'Unknown');
    console.log('  Created:', result.created?.toISOString() || 'Unknown');
    console.log('  Modified:', result.modified?.toISOString() || 'Unknown');
    
    console.log('\n📐 Dimensions:');
    console.log(`  Width: ${result.props.width}px (${Math.round(result.props.width * 914400 / 96)} EMU)`);
    console.log(`  Height: ${result.props.height}px (${Math.round(result.props.height * 914400 / 96)} EMU)`);
    console.log(`  Ratio: ${result.props.ratio.toFixed(2)}`);
    console.log(`  Page Size: ${result.props.pageSize}`);

    console.log('\n🎨 Theme:');
    if (result.theme) {
      console.log('  Background 1 (bg1):', result.theme.colors.bg1);
      console.log('  Text 1 (tx1):', result.theme.colors.tx1);
      console.log('  Accent 1:', result.theme.colors.accent1);
      console.log('  Accent 2:', result.theme.colors.accent2);
      console.log('  Accent 3:', result.theme.colors.accent3);
    } else {
      console.log('  No theme found');
    }

    console.log('\n📑 Masters:');
    if (result.masterSlides && result.masterSlides.length > 0) {
      result.masterSlides.forEach((master, index) => {
        console.log(`  Master ${index + 1}:`);
        console.log(`    ID: ${master.id}`);
        console.log(`    Background:`, master.background);
        console.log(`    Elements: ${master.elements.length}`);
        console.log(`    Color map entries: ${Object.keys(master.colorMap).length}`);
      });
    } else {
      console.log('  No master slides found');
    }

    console.log('\n📄 Slides:');
    console.log(`  Total: ${result.slides.length}`);
    
    let slidesWithBackground = 0;
    let slidesWithElements = 0;
    let totalElements = 0;
    
    result.slides.forEach((slide, index) => {
      const elementCount = slide.elements?.length || 0;
      const hasBackground = slide.background && slide.background !== '#ffffff';
      
      if (hasBackground) slidesWithBackground++;
      if (elementCount > 0) slidesWithElements++;
      totalElements += elementCount;
      
      if (index < 5 || hasBackground) {  // Show first 5 or any with special background
        console.log(`  \n  Slide ${index + 1}: ${slide.title}`);
        console.log(`    Background:`, slide.background);
        console.log(`    Elements: ${elementCount}`);
        
        if (elementCount > 0) {
          const elementTypes = new Set();
          slide.elements.forEach(el => {
            if (el.type) elementTypes.add(el.type);
          });
          console.log(`    Types:`, Array.from(elementTypes).join(', '));
        }
      }
    });

    console.log(`\n  Slides with custom background: ${slidesWithBackground}/${result.slides.length}`);
    console.log(`  Slides with elements: ${slidesWithElements}/${result.slides.length}`);
    console.log(`  Total elements: ${totalElements}`);

    // Save parsed result to JSON
    const outputPath = path.join(__dirname, 'parsed-result.json');
    fs.writeFileSync(outputPath, JSON.stringify(result, null, 2), 'utf-8');
    console.log(`\n✅ Parsed result saved to: ${outputPath}`);

    console.log('\n' + '='.repeat(60));
    console.log('✅ Parsing completed successfully!');

  } catch (error) {
    console.error('\n❌ Parsing failed:', error);
    process.exit(1);
  }
}

main();
