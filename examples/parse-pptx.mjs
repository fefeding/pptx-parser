#!/usr/bin/env node
/**
 * Node.js 下使用 @fefeding/ppt-parser 解析 PPTX 的示例脚本
 *
 * 运行前请先构建产物（解析依赖 dist 目录）：
 *   npm run build          # 一次性构建
 *   # 或 npm run dev       # 监听并产出 dist（需另开终端）
 *
 * 用法：
 *   node examples/parse-pptx.mjs                 # 解析默认的 examples/test-sample.pptx
 *   node examples/parse-pptx.mjs 路径/xxx.pptx    # 解析指定文件
 *
 * 示例会输出：幻灯片数量、画布尺寸、文档元信息，以及每一页提取出的文本。
 */

import { readFile } from 'node:fs/promises';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));

// 优先从已安装的包名导入（适用于 `npm i @fefeding/ppt-parser` 的消费者）；
// 仓库内运行时回退到本地构建产物 dist。
let lib;
try {
  lib = await import('@fefeding/ppt-parser');
} catch {
  lib = await import('../dist/ppt-parser.esm.js');
}
const { pptxToJson } = lib;

async function main() {
  const argPath = process.argv[2];
  const pptxPath = resolve(__dirname, argPath || 'test-sample.pptx');

  console.log(`\n解析文件: ${pptxPath}\n`);

  // 读取为 Buffer（Uint8Array 子类），可直接传给解析函数
  const buffer = await readFile(pptxPath);

  // 解析为结构化 JSON（mediaProcess/themeProcess 控制媒体与主题处理）
  const json = await pptxToJson(buffer, {
    mediaProcess: true,
    themeProcess: true
  });

  // ---- 基础信息 ----
  console.log('=== 基础信息 ===');
  console.log('幻灯片数量:', json.slides.length);
  if (json.slideSize) {
    console.log('画布尺寸:', json.slideSize.width, 'x', json.slideSize.height);
  }
  if (json.metadata && Object.keys(json.metadata).length) {
    console.log('文档元信息:', JSON.stringify(json.metadata, null, 2));
  }

  // ---- 逐页文本 ----
  console.log('\n=== 每页文本 ===');
  for (const slide of json.slides) {
    const texts = collectText(slide.data);
    console.log(`\n[第 ${slide.slideNum} 页] ${slide.fileName}`);
    console.log(texts.length ? texts.join('\n') : '（无文本）');
  }

  // ---- 结构化数据示例（仅打印第一页，避免刷屏）----
  console.log('\n=== 第一页结构化数据（前 1200 字符）===');
  const first = JSON.stringify(json.slides[0]?.data, null, 2) || '{}';
  console.log(first.length > 1200 ? first.slice(0, 1200) + '\n...' : first);
}

/**
 * 在简化的 XmlNode 树中递归收集文本。
 * 文本段落（drawingML 的 <a:t>）在简化后表现为键 'a:t' 对应的字符串值。
 */
function collectText(node, out = []) {
  if (node == null || typeof node !== 'object') return out;
  if (Array.isArray(node)) {
    node.forEach((n) => collectText(n, out));
    return out;
  }
  for (const key of Object.keys(node)) {
    const val = node[key];
    if (key === 'a:t' && typeof val === 'string') {
      out.push(val);
    } else if (typeof val === 'object') {
      collectText(val, out);
    }
  }
  return out;
}

main().catch((err) => {
  console.error('解析失败:', err);
  console.error('\n提示：若报错找不到模块，请先执行 `npm run build` 生成 dist 目录。');
  process.exit(1);
});
