#!/bin/bash

# 修复所有 window.PPTXUtils 引用
find src/js -name "*.js" -type f | while read file; do
    echo "处理 $file..."
    
    # 检查文件是否已经包含 PPTXUtils 的 import
    if ! grep -q "import.*PPTXUtils.*from.*utils" "$file"; then
        # 只在文件顶部没有 import 且包含 window.PPTXUtils 时才添加 import
        if grep -q "window\.PPTXUtils" "$file"; then
            # 检查是否已经有其他 import
            if grep -q "^import" "$file"; then
                # 如果已有 import，在后面添加
                sed -i '1,/^import.*from/s/^import.*$/&\nimport { PPTXUtils } from '\''..\/utils\/utils.js'\'';/' "$file" 2>/dev/null || true
            else
                # 如果没有 import，在第一行前添加
                sed -i '1s/^/import { PPTXUtils } from '\''.\/utils\/utils.js'\'';\n/' "$file" 2>/dev/null || true
            fi
        fi
    fi
    
    # 替换 window.PPTXUtils 为 PPTXUtils
    sed -i 's/window\.PPTXUtils\./PPTXUtils./g' "$file"
done

# 处理 PPTXColorUtils
find src/js -name "*.js" -type f | while read file; do
    if ! grep -q "import.*PPTXColorUtils.*from.*pptx-color-utils" "$file"; then
        if grep -q "window\.PPTXColorUtils" "$file"; then
            if grep -q "^import" "$file"; then
                sed -i '1,/^import.*from/s/^import.*$/&\nimport { PPTXColorUtils } from '\''..\/core\/pptx-color-utils.js'\'';/' "$file" 2>/dev/null || true
            else
                sed -i '1s/^/import { PPTXColorUtils } from '\''.\/core\/pptx-color-utils.js'\'';\n/' "$file" 2>/dev/null || true
            fi
        fi
    fi
    sed -i 's/window\.PPTXColorUtils\./PPTXColorUtils./g' "$file"
done

# 处理 TextUtils
find src/js -name "*.js" -type f | while read file; do
    if ! grep -q "import.*TextUtils.*from.*text-utils" "$file"; then
        if grep -q "window\.TextUtils" "$file"; then
            if grep -q "^import" "$file"; then
                sed -i '1,/^import.*from/s/^import.*$/&\nimport { TextUtils } from '\''..\/text\/text-utils.js'\'';/' "$file" 2>/dev/null || true
            else
                sed -i '1s/^/import { TextUtils } from '\''.\/text\/text-utils.js'\'';\n/' "$file" 2>/dev/null || true
            fi
        fi
    fi
    sed -i 's/window\.TextUtils\./TextUtils./g' "$file"
done

# 处理 PPTXTextElementUtils
find src/js -name "*.js" -type f | while read file; do
    if ! grep -q "import.*PPTXTextElementUtils.*from.*pptx-text-element-utils" "$file"; then
        if grep -q "window\.PPTXTextElementUtils" "$file"; then
            if grep -q "^import" "$file"; then
                sed -i '1,/^import.*from/s/^import.*$/&\nimport { PPTXTextElementUtils } from '\''..\/text\/pptx-text-element-utils.js'\'';/' "$file" 2>/dev/null || true
            else
                sed -i '1s/^/import { PPTXTextElementUtils } from '\''.\/text\/pptx-text-element-utils.js'\'';\n/' "$file" 2>/dev/null || true
            fi
        fi
    fi
    sed -i 's/window\.PPTXTextElementUtils\./PPTXTextElementUtils./g' "$file"
done

echo "所有 window. 引用已替换完成！"
