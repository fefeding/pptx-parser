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
return {
    readXmlFile: readXmlFile,
    getContentTypes: getContentTypes,
    getSlideSizeAndSetDefaultTextStyle: getSlideSizeAndSetDefaultTextStyle
};

})();
