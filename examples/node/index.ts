#!/usr/bin/env node

/**
 * 简化版 Node.js TypeScript 示例脚本 - 解析 3guo.pptx 文件
 * 直接调用源码接口并打印返回数据
 */

import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

// 获取当前文件目录
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 动态导入模块
async function main() {
  try {
    // 设置PPTX文件路径
    const pptxFilePath = path.resolve(__dirname, '../3guo.pptx');
    
    // 检查文件是否存在
    if (!fs.existsSync(pptxFilePath)) {
      console.error(`错误: 未找到文件 ${pptxFilePath}`);
      process.exit(1);
    }
    
    console.log('开始解析 PPTX 文件...');
    
    // 动态导入模块
    const { pptxToHtml } = await import('../../src/index.ts');
    
    // 读取PPTX文件为ArrayBuffer
    const fileBuffer = fs.readFileSync(pptxFilePath);
    const arrayBuffer = fileBuffer.buffer.slice(
      fileBuffer.byteOffset, 
      fileBuffer.byteOffset + fileBuffer.byteLength
    );
    
    // 调用接口
    const result = await pptxToHtml(arrayBuffer, {
      mediaProcess: true,
      onProgress: (percent: number) => {
        console.log(`解析进度: ${percent}%`);
      }
    });
    
    // 打印返回数据
    console.log('\\n解析结果:');
    console.log(JSON.stringify({
      slideCount: result.slideCount,
      htmlLength: result.html?.length || 0,
      cssLength: result.css?.length || 0,
      hasSlides: !!result.slides?.length
    }, null, 2));
    
    console.log('\\n完整结果对象结构:');
    console.log(Object.keys(result));
    
  } catch (error) {
    console.error('解析失败:', error);
  }
}

// 执行
main();