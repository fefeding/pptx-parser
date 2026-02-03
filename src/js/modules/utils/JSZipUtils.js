/**
 * JSZipUtils - Utility functions for JSZip
 * Extracted from pptx-main.js
 * IIFE format for compatibility
 */

var JSZipUtils = (function() {
    // JSZipUtils 实现
    // 这里包含从原文件中提取的 JSZipUtils 相关代码
    
    function getBinaryContent(url, callback) {
        // 原始的 getBinaryContent 实现
        try {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', url, true);
            xhr.responseType = 'arraybuffer';
            
            xhr.onload = function() {
                if (xhr.status === 200) {
                    callback(null, xhr.response);
                } else {
                    callback(new Error('Failed to load file: ' + xhr.status), null);
                }
            };
            
            xhr.onerror = function() {
                callback(new Error('Network error'), null);
            };
            
            xhr.send();
        } catch (e) {
            callback(e, null);
        }
    }
    
    // 其他 JSZipUtils 方法
    function createBinaryFile(data) {
        // 创建二进制文件的辅助方法
        return new Blob([data], { type: 'application/octet-stream' });
    }
    
    // 返回公共接口
    return {
        getBinaryContent: getBinaryContent,
        createBinaryFile: createBinaryFile
    };
})();

// 向后兼容 - 同时暴露为全局变量
if (typeof window !== 'undefined') {
    window.JSZipUtils = JSZipUtils;
}