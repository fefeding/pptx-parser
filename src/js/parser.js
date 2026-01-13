

    import { PPTXUtils } from './utils/utils.js';
    import { PPTXColorUtils } from './core/color-utils.js';

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

    // 解析器配置
    function configure(config) {
        settings = config;
        processFullTheme = settings.themeProcess;
        if (config.processSingleSlide) {
            window._processSingleSlideCallback = config.processSingleSlide;
        }
        if (config.processNodesInSlide) {
            window._processNodesInSlideCallback = config.processNodesInSlide;
        }
        if (config.getBackground) {
            window._getBackgroundCallback = config.getBackground;
        }
        if (config.getSlideBackgroundFill) {
            window._getSlideBackgroundFillCallback = config.getSlideBackgroundFill;
        }
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
            // Use internal processSingleSlide function if no callback provided
            var slideHtml;
            if (typeof window._processSingleSlideCallback === 'function') {
                slideHtml = window._processSingleSlideCallback(zip, filename, i, slideSize);
            } else {
                slideHtml = processSingleSlide(zip, filename, i, slideSize);
            }
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

        // globalCSS 将在主文件中处理，此时 styleTable 已经被填充
        // post_ary.push({
        //     "type": "globalCSS",
        //     "data": window.PPTXHtml ? window.PPTXHtml.genGlobalCSS() : ''
        // });

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
        // 根据 PPTX 规范（ECMA-376），p:defaultTextStyle 是可选元素
        // 当不存在时，提供一个空对象以避免 null/undefined 错误
        if (defaultTextStyle === undefined || defaultTextStyle === null) {
            defaultTextStyle = {};
        }

        slideWidth = sldSzWidth * slideFactor + settings.incSlide.width|0;// * scaleX;//parseInt(sldSzAttrs["cx"]) * 96 / 914400;
        slideHeight = sldSzHeight * slideFactor + settings.incSlide.height|0;// * scaleY;//parseInt(sldSzAttrs["cy"]) * 96 / 914400;
        rtenObj = {
            "width": slideWidth,
            "height": slideHeight
        };
        return rtenObj;
    }

    // 索引节点
    function indexNodes(content) {
        var keys = Object.keys(content);
        var spTreeNode = content[keys[0]]["p:cSld"]["p:spTree"];

        var idTable = {};
        var idxTable = {};
        var typeTable = {};

        for (var key in spTreeNode) {
            if (key == "p:nvGrpSpPr" || key == "p:grpSpPr") {
                continue;
            }

            var targetNode = spTreeNode[key];

            if (targetNode.constructor === Array) {
                for (var i = 0; i < targetNode.length; i++) {
                    var nvSpPrNode = targetNode[i]["p:nvSpPr"];
                    var id = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                    var idx = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                    var type = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

                    if (id !== undefined) {
                        idTable[id] = targetNode[i];
                    }
                    if (idx !== undefined) {
                        idxTable[idx] = targetNode[i];
                    }
                    if (type !== undefined) {
                        typeTable[type] = targetNode[i];
                    }
                }
            } else {
                var nvSpPrNode = targetNode["p:nvSpPr"];
                var id = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                var idx = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                var type = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

                if (id !== undefined) {
                    idTable[id] = targetNode;
                }
                if (idx !== undefined) {
                    idxTable[idx] = targetNode;
                }
                if (type !== undefined) {
                    typeTable[type] = targetNode;
                }
            }
        }

        return { "idTable": idTable, "idxTable": idxTable, "typeTable": typeTable };
    }

    // 获取背景
    function getBackground(warpObj, slideSize, index) {
        var bgResult = "";
        if (processFullTheme === true) {
            // 读取 slide 节点中的背景
            var bgNode = PPTXUtils.getTextByPathList(warpObj.slideContent, ["p:sld", "p:cSld", "p:bg"]);
            if (bgNode) {
                var bgPr = bgNode["p:bgPr"];
                if (bgPr) {
                    // 纯色填充
                    var solidFill = PPTXUtils.getTextByPathList(bgPr, ["a:solidFill"]);
                    if (solidFill) {
                        var color = window.PPTXColorUtils ? PPTXColorUtils.getFillColor(solidFill, warpObj.themeContent, warpObj.themeResObj, warpObj.slideLayoutClrOvride) : "";
                        if (color) {
                            bgResult = "<div class='slide-background-" + index + "' style='position:absolute;width:" + slideSize.width + "px;height:" + slideSize.height + "px;background-color:" + color + ";'></div>";
                        }
                    }
                    // 图片填充等可在此扩展
                }
            }
        }
        return bgResult;
    }

    // 获取幻灯片背景填充
    function getSlideBackgroundFill(warpObj, index) {
        var bgColor = "";
        if (processFullTheme == "colorsAndImageOnly") {
            var bgNode = PPTXUtils.getTextByPathList(warpObj.slideContent, ["p:sld", "p:cSld", "p:bg"]);
            if (bgNode) {
                var bgPr = bgNode["p:bgPr"];
                if (bgPr) {
                    var solidFill = PPTXUtils.getTextByPathList(bgPr, ["a:solidFill"]);
                    if (solidFill) {
                        var color = window.PPTXColorUtils ? PPTXColorUtils.getFillColor(solidFill, warpObj.themeContent, warpObj.themeResObj, warpObj.slideLayoutClrOvride) : "";
                        if (color) {
                            bgColor = "background-color:" + color + ";";
                        }
                    }
                }
            }
        }
        return bgColor;
    }

    // 处理单个幻灯片
    function processSingleSlide(zip, sldFileName, index, slideSize) {
        var resName = sldFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
        var resContent = PPTXParser.readXmlFile(zip, resName);
        var RelationshipArray = resContent["Relationships"]["Relationship"];
        var layoutFilename = "";
        var diagramFilename = "";
        var slideResObj = {};
        if (RelationshipArray.constructor === Array) {
            for (var i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i]["attrs"]["Type"]) {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                        layoutFilename = PPTXUtils.resolveRelationshipTarget(resName, RelationshipArray[i]["attrs"]["Target"]);
                        break;
                    case "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing":
                        diagramFilename = PPTXUtils.resolveRelationshipTarget(resName, RelationshipArray[i]["attrs"]["Target"]);
                        slideResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                            "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": PPTXUtils.resolveRelationshipTarget(resName, RelationshipArray[i]["attrs"]["Target"])
                        };
                        break;
                    default:
                        slideResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                            "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": PPTXUtils.resolveRelationshipTarget(resName, RelationshipArray[i]["attrs"]["Target"])
                        };
                }
            }
        } else {
            layoutFilename = PPTXUtils.resolveRelationshipTarget(resName, RelationshipArray["attrs"]["Target"]);
        }

        var slideLayoutContent = PPTXParser.readXmlFile(zip, layoutFilename);
        var slideLayoutTables = indexNodes(slideLayoutContent);
        var sldLayoutClrOvr = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);

        if (sldLayoutClrOvr !== undefined) {
            slideLayoutClrOvride = sldLayoutClrOvr["attrs"];
        }

        var slideLayoutResFilename = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
        var slideLayoutResContent = PPTXParser.readXmlFile(zip, slideLayoutResFilename);
        RelationshipArray = slideLayoutResContent["Relationships"]["Relationship"];
        var masterFilename = "";
        var layoutResObj = {};
        if (RelationshipArray.constructor === Array) {
            for (var i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i]["attrs"]["Type"]) {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
                        masterFilename = PPTXUtils.resolveRelationshipTarget(slideLayoutResFilename, RelationshipArray[i]["attrs"]["Target"]);
                        break;
                    default:
                        layoutResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                            "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": PPTXUtils.resolveRelationshipTarget(slideLayoutResFilename, RelationshipArray[i]["attrs"]["Target"])
                        };
                }
            }
        } else {
            masterFilename = PPTXUtils.resolveRelationshipTarget(slideLayoutResFilename, RelationshipArray["attrs"]["Target"]);
        }

        var slideMasterContent = PPTXParser.readXmlFile(zip, masterFilename);
        var slideMasterTextStyles = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
        var slideMasterTables = indexNodes(slideMasterContent);

        var slideMasterResFilename = masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster") + ".rels";
        var slideMasterResContent = PPTXParser.readXmlFile(zip, slideMasterResFilename);
        RelationshipArray = slideMasterResContent["Relationships"]["Relationship"];
        var themeFilename = "";
        var masterResObj = {};
        if (RelationshipArray.constructor === Array) {
            for (var i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i]["attrs"]["Type"]) {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                        themeFilename = PPTXUtils.resolveRelationshipTarget(slideMasterResFilename, RelationshipArray[i]["attrs"]["Target"]);
                        break;
                    default:
                        masterResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                            "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": PPTXUtils.resolveRelationshipTarget(slideMasterResFilename, RelationshipArray[i]["attrs"]["Target"])
                        };
                }
            }
        } else {
            themeFilename = PPTXUtils.resolveRelationshipTarget(slideMasterResFilename, RelationshipArray["attrs"]["Target"]);
        }

        var themeResObj = {};
        var themeContent = {};
        if (themeFilename !== undefined) {
            var themeName = themeFilename.split("/").pop();
            var themeResFileName = themeFilename.replace(themeName, "_rels/" + themeName) + ".rels";
            themeContent = PPTXParser.readXmlFile(zip, themeFilename);
            var themeResContent = PPTXParser.readXmlFile(zip, themeResFileName);
            if (themeResContent !== null) {
                var relationshipArray = themeResContent["Relationships"]["Relationship"];
                if (relationshipArray !== undefined){
                    if (relationshipArray.constructor === Array) {
                        for (var i = 0; i < relationshipArray.length; i++) {
                            themeResObj[relationshipArray[i]["attrs"]["Id"]] = {
                                "type": relationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                "target": PPTXUtils.resolveRelationshipTarget(themeResFileName, relationshipArray[i]["attrs"]["Target"])
                            };
                        }
                    } else {
                        themeResObj[relationshipArray["attrs"]["Id"]] = {
                            "type": relationshipArray["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": PPTXUtils.resolveRelationshipTarget(themeResFileName, relationshipArray["attrs"]["Target"])
                        };
                    }
                }
            }
        }

        var diagramResObj = {};
        var digramFileContent = {};
        if (diagramFilename !== undefined) {
            var diagName = diagramFilename.split("/").pop();
            var diagramResFileName = diagramFilename.replace(diagName, "_rels/" + diagName) + ".rels";
            digramFileContent = PPTXParser.readXmlFile(zip, diagramFilename);
            if (digramFileContent !== null && digramFileContent !== undefined && digramFileContent != "") {
                var digramFileContentObjToStr = JSON.stringify(digramFileContent);
                digramFileContentObjToStr = digramFileContentObjToStr.replace(/dsp:/g, "p:");
                digramFileContent = JSON.parse(digramFileContentObjToStr);
            }

            var digramResContent = PPTXParser.readXmlFile(zip, diagramResFileName);
            if (digramResContent !== null) {
                var relationshipArray = digramResContent["Relationships"]["Relationship"];
                if (relationshipArray.constructor === Array) {
                    for (var i = 0; i < relationshipArray.length; i++) {
                        diagramResObj[relationshipArray[i]["attrs"]["Id"]] = {
                            "type": relationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": PPTXUtils.resolveRelationshipTarget(diagramResFileName, relationshipArray[i]["attrs"]["Target"])
                        };
                    }
                } else {
                    diagramResObj[relationshipArray["attrs"]["Id"]] = {
                        "type": relationshipArray["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        "target": PPTXUtils.resolveRelationshipTarget(diagramResFileName, relationshipArray["attrs"]["Target"])
                    };
                }
            }
        }

        var slideContent = PPTXParser.readXmlFile(zip, sldFileName, true);
        var nodes = slideContent["p:sld"]["p:cSld"]["p:spTree"];
        var warpObj = {
            "zip": zip,
            "sldFileName": sldFileName,
            "slideLayoutContent": slideLayoutContent,
            "slideLayoutTables": slideLayoutTables,
            "slideMasterContent": slideMasterContent,
            "slideMasterTables": slideMasterTables,
            "slideContent": slideContent,
            "slideResObj": slideResObj,
            "slideMasterTextStyles": slideMasterTextStyles,
            "layoutResObj": layoutResObj,
            "masterResObj": masterResObj,
            "themeContent": themeContent,
            "themeResObj": themeResObj,
            "digramFileContent": digramFileContent,
            "diagramResObj": diagramResObj,
            "tableStyles": tableStyles,
            "defaultTextStyle": PPTXParser.defaultTextStyle
        };
        var bgResult = "";
        if (processFullTheme === true) {
            // Use callback if provided, otherwise use internal function
            if (typeof window._getBackgroundCallback === 'function') {
                bgResult = window._getBackgroundCallback(warpObj, slideSize, index);
            } else {
                bgResult = getBackground(warpObj, slideSize, index);
            }
        }

        var bgColor = "";
        if (processFullTheme == "colorsAndImageOnly") {
            // Use callback if provided, otherwise use internal function
            if (typeof window._getSlideBackgroundFillCallback === 'function') {
                bgColor = window._getSlideBackgroundFillCallback(warpObj, index);
            } else {
                bgColor = getSlideBackgroundFill(warpObj, index);
            }
        }

        if (settings.slideMode && settings.slideType == "revealjs") {
            var result = "<section class='slide' style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;" + bgColor + "'>"
        } else {
            var result = "<div class='slide' style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;" + bgColor + "'>"
        }
        result += bgResult;
        // Use callback for processNodesInSlide if provided
        var processNodesFunc = window._processNodesInSlideCallback || processNodesInSlide;
        for (var nodeKey in nodes) {
            if (nodes[nodeKey].constructor === Array) {
                for (var i = 0; i < nodes[nodeKey].length; i++) {
                    result += processNodesFunc(nodeKey, nodes[nodeKey][i], nodes, warpObj, "slide");
                }
            } else {
                result += processNodesFunc(nodeKey, nodes[nodeKey], nodes, warpObj, "slide");
            }
        }
        if (settings.slideMode && settings.slideType == "revealjs") {
            return result + "</div></section>";
        } else {
            return result + "</div></div>";
        }
    }

    // 公开 API
    const PPTXParser = {
        configure: configure,
        processPPTX: processPPTX,
        readXmlFile: readXmlFile,
        getContentTypes: getContentTypes,
        getSlideSizeAndSetDefaultTextStyle: getSlideSizeAndSetDefaultTextStyle,
        indexNodes: indexNodes,
        processSingleSlide: processSingleSlide,
        getBackground: getBackground,
        getSlideBackgroundFill: getSlideBackgroundFill,
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


export { PPTXParser };

// Also export to global scope for backward compatibility
window.PPTXParser = PPTXParser;
