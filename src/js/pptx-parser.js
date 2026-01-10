/**
 * PPTXParser - PPTX 解析逻辑模块
 * 提取自 pptxjs.js
 */

(function () {
    var $ = window.jQuery;

    // 全局变量，将在初始化时设置
    var app_verssion;
    var defaultTextStyle = null;
    var tableStyles;
    var styleTable = {};
    var slideFactor = 96 / 914400;
    var fontSizeFactor = 4 / 3.2;
    var slideWidth = 0;
    var slideHeight = 0;
    var isSlideMode = false;
    var processFullTheme = true;
    var settings;

    // 工具函数引用
    var PPTXUtils = window.PPTXUtils;

    // 解析器配置
    function configure(config) {
        settings = config;
        processFullTheme = settings.themeProcess;
    }

    // 主解析函数
    function processPPTX(zip) {
        var post_ary = [];
        var dateBefore = new Date();

        if (zip.file("docProps/thumbnail.jpeg") !== null) {
            var pptxThumbImg = PPTXUtils.base64ArrayBuffer(zip.file("docProps/thumbnail.jpeg").asArrayBuffer());
            post_ary.push({
                "type": "pptx-thumb",
                "data": pptxThumbImg,
                "slide_num": -1
            });
        }

        var filesInfo = getContentTypes(zip);
        var slideSize = getSlideSizeAndSetDefaultTextStyle(zip);
        tableStyles = readXmlFile(zip, "ppt/tableStyles.xml");
        //console.log("slideSize: ", slideSize)
        post_ary.push({
            "type": "slideSize",
            "data": slideSize,
            "slide_num": 0
        });

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
            var slideHtml = processSingleSlide(zip, filename, i, slideSize);
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

        post_ary.sort(function (a, b) {
            return a.slide_num - b.slide_num;
        });

        // 注意：genGlobalCSS 将在 pptx-html.js 中定义
        post_ary.push({
            "type": "globalCSS",
            "data": window.PPTXHtml ? window.PPTXHtml.genGlobalCSS() : ''
        });

        var dateAfter = new Date();
        post_ary.push({
            "type": "ExecutionTime",
            "data": dateAfter - dateBefore
        });
        return post_ary;
    }

    // 读取 XML 文件
    function readXmlFile(zip, filename, isSlideContent) {
        try {
            // 尝试解析文件路径，处理相对路径问题
            var fileEntry = zip.file(filename);
            if (!fileEntry && !filename.startsWith("ppt/") && !filename.startsWith("[Content_Types].xml") && !filename.startsWith("docProps/")) {
                // 尝试添加 ppt/ 前缀
                fileEntry = zip.file("ppt/" + filename);
            }
            if (!fileEntry) {
                // 如果仍然找不到，返回 null
                console.warn("XML file not found:", filename);
                return null;
            }
            var fileContent = fileEntry.asText();
            if (isSlideContent && app_verssion <= 12) {
                //< office2007
                //remove "<![CDATA[ ... ]]>" tag
                fileContent = fileContent.replace(/<!\[CDATA\[(.*?)\]\]>/g, '$1');
            }
            var xmlData = tXml(fileContent, { simplify: 1 });
            if (xmlData["?xml"] !== undefined) {
                return xmlData["?xml"];
            } else {
                return xmlData;
            }
        } catch (e) {
            //console.log("error readXmlFile: the file '", filename, "' not exit")
            return null;
        }
    }

    // 获取内容类型
    function getContentTypes(zip) {
        var ContentTypesJson = readXmlFile(zip, "[Content_Types].xml");

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

    // 获取幻灯片尺寸并设置默认文本样式
    function getSlideSizeAndSetDefaultTextStyle(zip) {
        //get app version
        var app = readXmlFile(zip, "docProps/app.xml");
        var app_verssion_str = app["Properties"]["AppVersion"]
        app_verssion = parseInt(app_verssion_str);
        console.log("create by Office PowerPoint app verssion: ", app_verssion_str)

        //get slide dimensions
        var rtenObj = {};
        var content = readXmlFile(zip, "ppt/presentation.xml");
        var sldSzAttrs = content["p:presentation"]["p:sldSz"]["attrs"];
        var sldSzWidth = parseInt(sldSzAttrs["cx"]);
        var sldSzHeight = parseInt(sldSzAttrs["cy"]);
        var sldSzType = sldSzAttrs["type"];
        console.log("Presentation size type: ", sldSzType)

        //1 inches  = 96px = 2.54cm
        // 1 EMU = 1 / 914400 inch
        // Pixel = EMUs * Resolution / 914400;  (Resolution = 96)
        //var standardHeight = 6858000;
        //console.log("slideFactor: ", slideFactor, "standardHeight:", standardHeight, (standardHeight - sldSzHeight) / standardHeight)
        
        //slideFactor = (96 * (1 + ((standardHeight - sldSzHeight) / standardHeight))) / 914400 ;

        //slideFactor = slideFactor + sldSzHeight*((standardHeight - sldSzHeight) / standardHeight) ;

        //var ration = sldSzWidth / sldSzHeight;
        
        //Scale
        // var viewProps = readXmlFile(zip, "ppt/viewProps.xml");
        // var scaleLoc = getTextByPathList(viewProps, ["p:viewPr", "p:slideViewPr", "p:cSldViewPr", "p:cViewPr","p:scale"]);
        // var scaleXnodes, scaleX = 1, scaleYnode, scaleY = 1;
        // if (scaleLoc !== undefined){
        //     scaleXnodes = scaleLoc["a:sx"]["attrs"];
        //     var scaleXnodesN = scaleXnodes["n"];
        //     var scaleXnodesD = scaleXnodes["d"];
        //     if (scaleXnodesN !== undefined && scaleXnodesD !== undefined && scaleXnodesN != 0){
        //         scaleX = parseInt(scaleXnodesD)/parseInt(scaleXnodesN);
        //     }
        //     scaleYnode = scaleLoc["a:sy"]["attrs"];
        //     var scaleYnodeN = scaleYnode["n"];
        //     var scaleYnodeD = scaleYnode["d"];
        //     if (scaleYnodeN !== undefined && scaleYnodeD !== undefined && scaleYnodeN != 0) {
        //         scaleY = parseInt(scaleYnodeD) / parseInt(scaleYnodeN) ;
        //     }

        // }
        //console.log("scaleX: ", scaleX, "scaleY:", scaleY)
        //slideFactor = slideFactor * scaleX;

        defaultTextStyle = content["p:presentation"]["p:defaultTextStyle"];

        slideWidth = sldSzWidth * slideFactor + settings.incSlide.width|0;// * scaleX;//parseInt(sldSzAttrs["cx"]) * 96 / 914400;
        slideHeight = sldSzHeight * slideFactor + settings.incSlide.height|0;// * scaleY;//parseInt(sldSzAttrs["cy"]) * 96 / 914400;
        rtenObj = {
            "width": slideWidth,
            "height": slideHeight
        };
        return rtenObj;
    }

    // 公开 API
    window.PPTXParser = {
        configure: configure,
        processPPTX: processPPTX,
        readXmlFile: readXmlFile,
        getContentTypes: getContentTypes,
        getSlideSizeAndSetDefaultTextStyle: getSlideSizeAndSetDefaultTextStyle,
        slideFactor: slideFactor,
        fontSizeFactor: fontSizeFactor,
        slideWidth: slideWidth,
        slideHeight: slideHeight,
        isSlideMode: isSlideMode,
        processFullTheme: processFullTheme,
        styleTable: styleTable,
        tableStyles: tableStyles,
        defaultTextStyle: defaultTextStyle,
        app_verssion: app_verssion
    };

})();