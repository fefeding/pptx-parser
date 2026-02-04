/**
 * PPTX Processor
 * PPTX文件核心处理逻辑
 */

/**
 * 处理PPTX文件
 * @param {Object} zip - JSZip实例
 * @param {Object} settings - 设置对象
 * @param {number} slideFactor - 幻灯片尺寸转换因子
 * @returns {Array} 处理结果数组
 */

var PPTXProcessor = (function() {
    function processPPTX(zip, settings, slideFactor) {
    var post_ary = [];
    var dateBefore = new Date();

    // 处理缩略图
    if (zip.file("docProps/thumbnail.jpeg") !== null) {
        var pptxThumbImg = base64ArrayBuffer(zip.file("docProps/thumbnail.jpeg").asArrayBuffer());
        post_ary.push({
            "type": "pptx-thumb",
            "data": pptxThumbImg,
            "slide_num": -1
        });
    }

    // 获取文件信息和幻灯片尺寸
    var filesInfo = PPTXFileUtils.getContentTypes(zip, settings.appVersion);
    var slideSize = getSlideSizeAndSetDefaultTextStyle(zip, settings, slideFactor);
    
    // 读取表格样式
    var tableStyles = PPTXFileUtils.readXmlFile(zip, "ppt/tableStyles.xml", false, settings.appVersion);
    
    post_ary.push({
        "type": "slideSize",
        "data": slideSize,
        "slide_num": 0
    });

    // 处理所有幻灯片
    var numOfSlides = filesInfo["slides"].length;
    for (var i = 0; i < numOfSlides; i++) {
        var filename = filesInfo["slides"][i];
        var filename_no_path = "";
        var filename_no_path_ary = [];
        if (filename.indexOf("/") != -1) {
            filename_no_path_ary = filename.split("/");
            filename_no_path = filename_no_path_ary.pop();
        } else {
            filename_no_path = filename;
        }
        
        var filename_no_path_no_ext = "";
        if (filename_no_path.indexOf(".") != -1) {
            var filename_no_path_no_ext_ary = filename_no_path.split(".");
            var slide_ext = filename_no_path_no_ext_ary.pop();
            filename_no_path_no_ext = filename_no_path_no_ext_ary.join(".");
        }
        
        var slide_number = 1;
        if (filename_no_path_no_ext != "" && filename_no_path.indexOf("slide") != -1) {
            slide_number = Number(filename_no_path_no_ext.substr(5));
        }

        // 处理单个幻灯片
        var slideHtml = processSingleSlide(zip, filename, i, slideSize, settings, slideFactor);
        
        post_ary.push({
            "type": "slide",
            "data": slideHtml,
            "slide_num": slide_number,
            "file_name": filename_no_path_no_ext
        });
        
        post_ary.push({
            "type": "progress-update",
            "slide_num": (numOfSlides + i + 1),
            "data": (i + 1) * 100 / numOfSlides
        });
    }

    // 排序
    post_ary.sort(function (a, b) {
        return a.slide_num - b.slide_num;
    });

    // 添加全局CSS
    post_ary.push({
        "type": "globalCSS",
        "data": genGlobalCSS()
    });

    // 添加执行时间
    var dateAfter = new Date();
    post_ary.push({
        "type": "ExecutionTime",
        "data": dateAfter - dateBefore
    });
    
    return post_ary;
}

/**
 * 将ArrayBuffer转换为base64
 * @param {ArrayBuffer} buffer - ArrayBuffer
 * @returns {string} base64字符串
 */
function base64ArrayBuffer(arrayBuffer) {
    var base64 = '';
    var encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
    var bytes = new Uint8Array(arrayBuffer);
    var byteLength = bytes.byteLength;
    var byteRemainder = byteLength % 3;
    var mainLength = byteLength - byteRemainder;

    var a, b, c, d;
    var chunk;

    for (var i = 0; i < mainLength; i = i + 3) {
        chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | (bytes[i + 2]);
        a = (chunk & 16515072) >> 18;
        b = (chunk & 258048) >> 12;
        c = (chunk & 4032) >> 6;
        d = chunk & 63;
        base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d];
    }

    if (byteRemainder == 1) {
        chunk = bytes[mainLength];
        a = (chunk & 252) >> 2;
        b = (chunk & 3) << 4;
        base64 += encodings[a] + encodings[b] + '==';
    } else if (byteRemainder == 2) {
        chunk = (bytes[mainLength] << 8) | (bytes[mainLength + 1]);
        a = (chunk & 64512) >> 10;
        b = (chunk & 1008) >> 4;
        c = (chunk & 15) << 2;
        base64 += encodings[a] + encodings[b] + encodings[c] + '=';
    }

    return base64;
}

/**
 * 生成全局CSS
 * @returns {string} CSS字符串
 */
function genGlobalCSS() {
    var cssText = "";
    for (var key in styleTable) {
        var tagname = "";
        cssText += tagname + " ." + styleTable[key]["name"] +
            ((styleTable[key]["suffix"]) ? styleTable[key]["suffix"] : "") +
            "{" + styleTable[key]["text"] + "}\n";
    }

    if (settings.slideMode && settings.slideType == "divs2slidesjs") {
        cssText += "#all_slides_warpper{margin-right: auto;margin-left: auto;padding-top:10px;width: " + slideWidth + "px;}\n";
    }
    return cssText;
}


    return {
        processPPTX: processPPTX
    };
})();

window.PPTXProcessor = PPTXProcessor;