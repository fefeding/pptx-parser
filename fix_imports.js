#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

const srcDir = path.join(__dirname, 'src/js');

// 需要处理的模块映射
const moduleMapping = {
    'window.PPTXUtils.': { import: "import { PPTXUtils } from './utils/utils.js';", path: './utils/utils.js' },
    'window.PPTXColorUtils.': { import: "import { PPTXColorUtils } from './core/pptx-color-utils.js';", path: './core/pptx-color-utils.js' },
    'window.TextUtils.': { import: "import { TextUtils } from './text/text-utils.js';", path: './text/text-utils.js' },
    'window.PPTXTextElementUtils.': { import: "import { PPTXTextElementUtils } from './text/pptx-text-element-utils.js';", path: './text/pptx-text-element-utils.js' },
    'window.PPTXNodeUtils.': { import: "import { PPTXNodeUtils } from './node/pptx-node-utils.js';", path: './node/pptx-node-utils.js' }
};

function processFile(filePath) {
    let content = fs.readFileSync(filePath, 'utf8');
    let hasChanges = false;
    let importsAdded = new Set();
    
    // 检查每个模块引用
    for (const [windowRef, moduleInfo] of Object.entries(moduleMapping)) {
        if (content.includes(windowRef)) {
            // 替换 window. 引用
            const regex = new RegExp(windowRef.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
            content = content.replace(regex, windowRef.replace('window.', ''));
            hasChanges = true;
            
            // 确定正确的相对路径
            let relativePath = path.relative(path.dirname(filePath), path.join(srcDir, moduleInfo.path.replace('./', '')));
            if (!relativePath.startsWith('.')) {
                relativePath = './' + relativePath;
            }
            
            // 创建针对该文件的 import 语句
            const importStatement = `import { ${windowRef.split('.')[1].replace('.', '')} } from '${relativePath}';`;
            importsAdded.add(importStatement);
        }
    }
    
    // 如果有需要添加的 import
    if (importsAdded.size > 0) {
        const importsArray = Array.from(importsAdded);
        
        // 检查是否已有 import
        if (content.includes('import ')) {
            // 找到最后一个 import 语句
            const lines = content.split('\n');
            let lastImportIndex = -1;
            for (let i = 0; i < lines.length; i++) {
                if (lines[i].trim().startsWith('import ')) {
                    lastImportIndex = i;
                }
            }
            
            if (lastImportIndex >= 0) {
                // 在最后一个 import 后添加新 import
                lines.splice(lastImportIndex + 1, 0, ...importsArray, '');
                content = lines.join('\n');
            } else {
                // 在文件开头添加
                content = importsArray.join('\n') + '\n\n' + content;
            }
        } else {
            // 文件中没有 import，在开头添加
            content = importsArray.join('\n') + '\n\n' + content;
        }
        
        hasChanges = true;
    }
    
    // 移除 window.XXX = XXX 的赋值
    content = content.replace(/if\s*\(\s*typeof\s+window\s*!==?\s*['"]undefined['"]\s*\)\s*\{\s*window\.\w+\s*=\s*\w+;?\s*\}/g, '');
    content = content.replace(/\/\/\s*Also export to global scope for backward compatibility[\s\S]*?(window\.\w+\s*=\s*\w+;?)?/g, '');
    content = content.replace(/\/\/\s*Also export to global scope for backward compatibility/g, '');
    content = content.replace(/window\.PPTXUtils\s*=\s*PPTXUtils;?/g, '');
    content = content.replace(/window\.PPTXColorUtils\s*=\s*PPTXColorUtils;?/g, '');
    content = content.replace(/window\.TextUtils\s*=\s*TextUtils;?/g, '');
    content = content.replace(/window\.PPTXTextElementUtils\s*=\s*PPTXTextElementUtils;?/g, '');
    
    if (hasChanges) {
        fs.writeFileSync(filePath, content, 'utf8');
        console.log(`✓ 已处理: ${path.relative(__dirname, filePath)}`);
    }
}

function walkDir(dir) {
    const files = fs.readdirSync(dir);
    for (const file of files) {
        const filePath = path.join(dir, file);
        const stat = fs.statSync(filePath);
        if (stat.isDirectory()) {
            walkDir(filePath);
        } else if (file.endsWith('.js')) {
            try {
                processFile(filePath);
            } catch (error) {
                console.error(`✗ 处理失败: ${filePath}`, error.message);
            }
        }
    }
}

console.log('开始处理所有 JavaScript 文件...\n');
walkDir(srcDir);
console.log('\n✅ 所有 window. 引用已替换完成！');
