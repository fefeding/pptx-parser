/**
 * File Utils
 * 文件处理工具函数
 */

var PPTXFileUtils = (function() {
    /**
     * 读取XML文件
     * @param {Object} zip - JSZip实例
     * @param {string} filename - 文件名
     * @param {boolean} isSlideContent - 是否为幻灯片内容
     * @param {number} appVersion - Office版本
     * @returns {Object|null} XML数据对象
     */
    function readXmlFile(zip, filename, isSlideContent, appVersion) {
    try {
        var fileContent = zip.file(filename).asText();
        if (isSlideContent && appVersion <= 12) {
            // < office2007
            // remove "<![CDATA[ ... ]]>" tag
            fileContent = fileContent.replace(/<!\[CDATA\[(.*?)\]\]>/g, '$1');
        }
        var xmlData = tXml(fileContent, { simplify: 1 });
        if (xmlData["?xml"] !== undefined) {
            return xmlData["?xml"];
        } else {
            return xmlData;
        }
    } catch (e) {
        // console.log("error readXmlFile: the file '", filename, "' not exit")
        return null;
    }
}

/**
 * 获取内容类型
 * @param {Object} zip - JSZip实例
 * @param {number} appVersion - Office版本
 * @returns {Object} 包含slides和slideLayouts的对象
 */
function getContentTypes(zip, appVersion) {
    var ContentTypesJson = readXmlFile(zip, "[Content_Types].xml", false, appVersion);
    
    var subObj = ContentTypesJson["Types"]["Override"];
    var slidesLocArray = [];
    var slideLayoutsLocArray = [];
    for (var i = 0; i < subObj.length; i++) {
        switch (subObj[i]["attrs"]["ContentType"]) {
            case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
                slidesLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
                break;
            case "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml":
                slideLayoutsLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
                break;
            default:
        }
    }
    return {
        "slides": slidesLocArray,
        "slideLayouts": slideLayoutsLocArray
    };
}

/**
 * 获取幻灯片尺寸并设置默认文本样式
 * @param {Object} zip - JSZip实例
 * @param {Object} settings - 设置对象
 * @param {number} slideFactor - 尺寸转换因子
 * @returns {Object} 包含width和height的对象
 */
function getSlideSizeAndSetDefaultTextStyle(zip, settings, slideFactor) {
    // get app version
    var app = readXmlFile(zip, "docProps/app.xml", false, settings.appVersion);
    var app_verssion_str = app["Properties"]["AppVersion"];
    settings.appVersion = parseInt(app_verssion_str);
    console.log("create by Office PowerPoint app verssion: ", app_verssion_str);

    // get slide dimensions
    var rtenObj = {};
    var content = readXmlFile(zip, "ppt/presentation.xml", false, settings.appVersion);
    var sldSzAttrs = content["p:presentation"]["p:sldSz"]["attrs"];
    var sldSzWidth = parseInt(sldSzAttrs["cx"]);
    var sldSzHeight = parseInt(sldSzAttrs["cy"]);
    var sldSzType = sldSzAttrs["type"];
    console.log("Presentation size type: ", sldSzType);

    // 1 inches  = 96px = 2.54cm
    // 1 EMU = 1 / 914400 inch
    // Pixel = EMUs * Resolution / 914400;  (Resolution = 96)

    settings.defaultTextStyle = content["p:presentation"]["p:defaultTextStyle"];

    var slideWidth = sldSzWidth * slideFactor + settings.incSlide.width | 0;
    var slideHeight = sldSzHeight * slideFactor + settings.incSlide.height | 0;
    
    rtenObj = {
        "width": slideWidth,
        "height": slideHeight
    };
    return rtenObj;
}

// Export to global namespace
/**
 * 解析媒体文件路径
 * 根据PPTX标准，处理不同上下文下的媒体文件路径
 * @param {string} mediaPath - 媒体文件路径（来自resObj[target]）
 * @param {string} context - 上下文类型：'slide', 'master', 'layout'
 * @param {string} basePath - 基础路径（通常是当前XML文件所在目录）
 * @returns {string} 解析后的完整路径
 */
function resolveMediaPath(mediaPath, context, basePath) {
    // 如果已经是绝对路径（以ppt/开头），直接返回
    if (mediaPath.startsWith('ppt/')) {
        return mediaPath;
    }
        
    // 处理相对路径
    let resolvedPath = mediaPath;
        
    // 根据上下文确定基础目录
    let baseDir = '';
    switch (context) {
        case 'slide':
            // 幻灯片中的媒体文件通常相对于ppt/slides/
            baseDir = 'ppt/slides/';
            break;
        case 'master':
            // 幻灯片母版中的媒体文件通常相对于ppt/slideMasters/
            baseDir = 'ppt/slideMasters/';
            break;
        case 'layout':
            // 版式中的媒体文件通常相对于ppt/slideLayouts/
            baseDir = 'ppt/slideLayouts/';
            break;
        default:
            // 默认情况，使用传入的基础路径
            baseDir = basePath || '';
    }
        
    // 处理路径中的../
    if (mediaPath.startsWith('../')) {
        // 移除../并构建相对于ppt/的路径
        resolvedPath = 'ppt/' + mediaPath.substring(3);
    } else if (!mediaPath.includes('/')) {
        // 如果没有路径分隔符，可能是直接在media目录下的文件
        resolvedPath = 'ppt/media/' + mediaPath;
    } else {
        // 其他相对路径，拼接基础目录
        resolvedPath = baseDir + mediaPath;
    }
        
    // 清理路径中的重复斜杠
    resolvedPath = resolvedPath.replace(/\/+/g, '/');
        
    // 移除开头的./
    if (resolvedPath.startsWith('./')) {
        resolvedPath = resolvedPath.substring(2);
    }
        
    return resolvedPath;
}
    
/**
 * 查找媒体文件（尝试多种可能的路径）
 * @param {Object} zip - JSZip实例
 * @param {string} originalPath - 原始路径
 * @param {string} context - 上下文类型
 * @param {string} basePath - 基础路径
 * @returns {Object|null} 找到的文件对象或null
 */
function findMediaFile(zip, originalPath, context, basePath) {
    // 首先尝试原始路径
    let file = zip.file(originalPath);
    if (file) {
        return file;
    }
        
    // 尝试解析后的标准路径
    const resolvedPath = resolveMediaPath(originalPath, context, basePath);
    file = zip.file(resolvedPath);
    if (file) {
        return file;
    }
        
    // 尝试常见的替代路径
    const alternativePaths = [];
        
    // 如果是media目录下的文件，尝试不同的前缀
    if (originalPath.includes('media/') || !originalPath.includes('/')) {
        const fileName = originalPath.split('/').pop();
        alternativePaths.push(
            'ppt/media/' + fileName,
            'media/' + fileName,
            fileName
        );
    }
        
    // 如果包含embeddings，也尝试相关路径
    if (originalPath.includes('embeddings/')) {
        const fileName = originalPath.split('/').pop();
        alternativePaths.push(
            'ppt/embeddings/' + fileName,
            'embeddings/' + fileName
        );
    }
        
    // 尝试所有备选路径
    for (const altPath of alternativePaths) {
        file = zip.file(altPath);
        if (file) {
            console.log(`Media file found at alternative path: ${altPath} (originally: ${originalPath})`);
            return file;
        }
    }
        
    // 如果都没找到，返回null
    return null;
}
    
return {
    readXmlFile: readXmlFile,
    getContentTypes: getContentTypes,
    getSlideSizeAndSetDefaultTextStyle: getSlideSizeAndSetDefaultTextStyle,
    resolveMediaPath: resolveMediaPath,
    findMediaFile: findMediaFile
};

})();
