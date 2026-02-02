const fs = require('fs');
const path = require('path');

const modulesDir = '/Users/jiamao/project/github/pptx-parser/src/js/modules';

// File mappings: relative path -> global variable name
const filesToConvert = [
    { path: 'utils/color-utils.js', name: 'PPTXColorUtils', hasImport: true },
    { path: 'utils/text-utils.js', name: 'PPTXTextUtils', hasImport: true },
    { path: 'utils/image-utils.js', name: 'PPTXImageUtils', hasImport: true },
    { path: 'utils/chart-utils.js', name: 'PPTXChartUtils', hasImport: true },
    { path: 'core/pptx-processor.js', name: 'PPTXProcessor', hasImport: true },
    { path: 'core/slide-processor.js', name: 'SlideProcessor', hasImport: true },
    { path: 'core/node-processors.js', name: 'NodeProcessors', hasImport: true },
    { path: 'shapes/shape-generator.js', name: 'ShapeGenerator', hasImport: true }
];

function convertToIIFE(filePath, globalName, hasImport) {
    let content = fs.readFileSync(filePath, 'utf8');
    
    // Already converted?
    if (content.includes(`var ${globalName} = (function()`)) {
        console.log(`✓ Already converted: ${path.basename(filePath)}`);
        return;
    }
    
    // Remove imports
    if (hasImport) {
        content = content.replace(/import\s+.*?from\s+['"][^'"]+['"];\s*\n?/g, '');
    }
    
    // Remove ES6 exports
    const lines = content.split('\n');
    let newLines = [];
    let inLeadingComment = true;
    let foundFirstExport = false;
    let leadingComments = [];
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const trimmed = line.trim();
        
        // Check if we're still in leading comments
        if (inLeadingComment && trimmed !== '' && !trimmed.startsWith('/**') && !trimmed.startsWith('*')) {
            inLeadingComment = false;
            // Insert IIFE wrapper
            newLines.push('');
            newLines.push(`var ${globalName} = (function() {`);
        }
        
        if (!foundFirstExport && trimmed.match(/^export\s+/)) {
            foundFirstExport = true;
            inLeadingComment = false;
            // Insert IIFE if not already done
            if (!newLines.some(l => l.includes(`${globalName} = (function()`))) {
                newLines.push('');
                newLines.push(`var ${globalName} = (function() {`);
            }
            // Convert export to normal function
            newLines.push(line.replace(/^export\s+/, '    '));
        } else if (trimmed.match(/^export\s+/)) {
            // Convert subsequent exports
            newLines.push(line.replace(/^export\s+/, '    '));
        } else if (!trimmed.match(/^export\s+/)) {
            // Keep non-export lines as is
            newLines.push(line);
        }
    }
    
    // Add return statement at the end
    newLines.push('');
    newLines.push('    return {');
    
    // Find all exported functions
    const exportMatches = content.match(/^export\s+(?:function\s+(\w+)|(?:const|let|var)\s+(\w+)|class\s+(\w+))/gm);
    if (exportMatches) {
        exportMatches.forEach((match, idx) => {
            const funcMatch = match.match(/(?:function|const|let|var|class)\s+(\w+)/);
            if (funcMatch) {
                newLines.push(`        ${funcMatch[1]}: ${funcMatch[1]}${idx < exportMatches.length - 1 ? ',' : ''}`);
            }
        });
    }
    
    newLines.push('    };');
    newLines.push('})();');
    
    fs.writeFileSync(filePath, newLines.join('\n'), 'utf8');
    console.log(`✓ Converted: ${path.basename(filePath)} -> ${globalName}`);
}

// Convert all files
console.log('Converting modules to IIFE format...\n');

filesToConvert.forEach(({ path: relPath, name, hasImport }) => {
    const fullPath = path.join(modulesDir, relPath);
    if (fs.existsSync(fullPath)) {
        convertToIIFE(fullPath, name, hasImport);
    } else {
        console.log(`✗ File not found: ${relPath}`);
    }
});

console.log('\n✓ Module conversion complete!');
