







    import { PPTXUtils, getAttrs, getAttr } from './utils/utils.js';
    import { PPTXColorUtils } from './core/color-utils.js';
    import { XMLParser } from 'fast-xml-parser';

    // 创建 XML 解析器实例
    const xmlParser = new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: '',
        textNodeName: '#text',
        allowBooleanAttributes: true,
        parseTagValue: true,
        parseAttributeValue: true,
        trimValues: true,
        cdataPropName: '__cdata',
        attributesGroupName: 'attrs',  // 将属性放在 attrs 对象中，与 tXml 兼容
        ignoreDeclaration: true,
        ignorePiTags: true
    });

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
    async function processPPTX(zip) {
        var post_ary = [];
        var dateBefore = new Date();

        // 调试：验证 ZIP 对象
        console.log("ZIP object in processPPTX:", zip);
        console.log("ZIP files count:", Object.keys(zip.files).length);
        console.log("ZIP has file method:", typeof zip.file);

        if (zip.file("docProps/thumbnail.jpeg") !== null) {
            var pptxThumbImg = PPTXUtils.base64ArrayBuffer(await zip.file("docProps/thumbnail.jpeg").async("arraybuffer"));
            post_ary.push({
                "type": "pptx-thumb",
                "data": pptxThumbImg,
                "slide_num": -1
            });
        }

        var filesInfo = await getContentTypes(zip);
        var slideSize = await getSlideSizeAndSetDefaultTextStyle(zip);
        tableStyles = await readXmlFile(zip, "ppt/tableStyles.xml");
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
                slideHtml = await processSingleSlide(zip, filename, i, slideSize);
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
    async function readXmlFile(zip, filename, isSlideContent) {
        try {
            // 调试：第一次调用时打印 ZIP 文件列表
            if (!readXmlFile._debugPrinted) {
                console.log("readXmlFile - ZIP files:", Object.keys(zip.files));
                readXmlFile._debugPrinted = true;
            }

            // 调试：打印当前查找的文件名
            console.log("Looking for file:", filename);

            // 尝试解析文件路径，处理相对路径问题
            var fileEntry = zip.file(filename);

            // 调试：打印查找结果
            console.log("Direct lookup result:", !!fileEntry);

            // 如果直接路径找不到，尝试各种可能的路径变体
            if (!fileEntry) {
                var possiblePaths = [];

                if (filename.startsWith("[Content_Types].xml")) {
                    possiblePaths = [
                        filename
                    ];
                } else if (filename.startsWith("docProps/")) {
                    possiblePaths = [
                        filename,
                        "docProps/" + filename.split("/").pop()
                    ];
                } else if (filename.startsWith("ppt/")) {
                    possiblePaths = [
                        filename,
                        filename.split("/").pop()
                    ];
                } else {
                    possiblePaths = [
                        filename,
                        "ppt/" + filename
                    ];
                }

                // 尝试所有可能的路径
                for (var i = 0; i < possiblePaths.length; i++) {
                    fileEntry = zip.file(possiblePaths[i]);
                    console.log("Trying path:", possiblePaths[i], "->", !!fileEntry);
                    if (fileEntry) {
                        console.log("Found file using alternate path:", possiblePaths[i]);
                        break;
                    }
                }
            }

            if (!fileEntry) {
                // 如果仍然找不到，返回 null
                console.warn("XML file not found:", filename);
                console.warn("Available files in ZIP:", Object.keys(zip.files));
                return null;
            }
            var fileContent = await fileEntry.async("text");
            if (isSlideContent && app_verssion <= 12) {
                //< office2007
                //remove "<![CDATA[ ... ]]>" tag
                fileContent = fileContent.replace(/<!\[CDATA\[(.*?)\]\]>/g, '$1');
            }
            var xmlData = xmlParser.parse(fileContent);
            // fast-xml-parser 不返回 ?xml 节点，直接返回解析结果
            return xmlData;
        } catch (e) {
            console.error("Error reading/parsing XML file:", filename, e);
            console.error("Error details:", e.message, e.stack);
            return null;
        }
    }

    // 获取内容类型
    async function getContentTypes(zip) {
        var ContentTypesJson = await readXmlFile(zip, "[Content_Types].xml");



        if (!ContentTypesJson || !ContentTypesJson["Types"]) {
            console.error("Failed to read [Content_Types].xml");
            return {
                "slides": [],
                "slideLayouts": []
            };
        }

        var subObj = ContentTypesJson["Types"]["Override"];
        var slidesLocArray = [];
        var slideLayoutsLocArray = [];
        var overrides = Array.isArray(subObj) ? subObj : [subObj];
        for (var i = 0; i < overrides.length; i++) {
            var item = overrides[i];
            var contentType = getAttr(item, "ContentType");
            var partName = getAttr(item, "PartName");
            switch (contentType) {
                case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
                    slidesLocArray.push(partName.substr(1));
                    break;
                case "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml":
                    slideLayoutsLocArray.push(partName.substr(1));
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
    async function getSlideSizeAndSetDefaultTextStyle(zip) {
        //get app version
        var app = await readXmlFile(zip, "docProps/app.xml");
        if (!app || !app["Properties"]) {
            console.error("Failed to read docProps/app.xml");
            app_verssion = 16; // 默认版本
        } else {
            var app_verssion_str = getAttr(app["Properties"], "AppVersion");
            app_verssion = parseInt(app_verssion_str);
            console.log("create by Office PowerPoint app verssion: ", app_verssion_str);
        }

        //get slide dimensions
        var rtenObj = {};
        var content = await readXmlFile(zip, "ppt/presentation.xml");
        if (!content || !content["p:presentation"] || !content["p:presentation"]["p:sldSz"]) {
            console.error("Failed to read ppt/presentation.xml");
            return {
                "width": 960,
                "height": 540
            };
        }
        var sldSzAttrs = content["p:presentation"]["p:sldSz"];
        var sldSzWidth = parseInt(getAttr(sldSzAttrs, "cx"));
        var sldSzHeight = parseInt(getAttr(sldSzAttrs, "cy"));
        var sldSzType = getAttr(sldSzAttrs, "type");
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
        var firstKey = keys[0];
        if (!content[firstKey] || !content[firstKey]["p:cSld"]) {
            console.error("indexNodes: Invalid content structure", content);
            return { idTable: {}, idxTable: {}, typeTable: {} };
        }
        var spTreeNode = content[firstKey]["p:cSld"]["p:spTree"];
        if (!spTreeNode) {
            console.error("indexNodes: p:spTree not found", content[firstKey]["p:cSld"]);
            return { idTable: {}, idxTable: {}, typeTable: {} };
        }

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
    async function processSingleSlide(zip, sldFileName, index, slideSize) {
        var resName = sldFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
        var resContent = await PPTXParser.readXmlFile(zip, resName);
        var RelationshipArray = resContent["Relationships"]["Relationship"];
        var layoutFilename = "";
        var diagramFilename = "";
        var slideResObj = {};
        if (RelationshipArray.constructor === Array) {
            for (var i = 0; i < RelationshipArray.length; i++) {
                var rel = RelationshipArray[i];
                var type = getAttr(rel, "Type");
                var target = getAttr(rel, "Target");
                var id = getAttr(rel, "Id");
                switch (type) {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                        layoutFilename = PPTXUtils.resolveRelationshipTarget(resName, target);
                        break;
                    case "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing":
                        diagramFilename = PPTXUtils.resolveRelationshipTarget(resName, target);
                        slideResObj[id] = {
                            "type": type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": PPTXUtils.resolveRelationshipTarget(resName, target)
                        };
                        break;
                    default:
                        slideResObj[id] = {
                            "type": type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": PPTXUtils.resolveRelationshipTarget(resName, target)
                        };
                }
            }
        } else {
            layoutFilename = PPTXUtils.resolveRelationshipTarget(resName, RelationshipArray["attrs"]["Target"]);
        }

        var slideLayoutContent = await PPTXParser.readXmlFile(zip, layoutFilename);
        var slideLayoutTables = indexNodes(slideLayoutContent);
        var sldLayoutClrOvr = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);

        if (sldLayoutClrOvr !== undefined) {
            slideLayoutClrOvride = sldLayoutClrOvr["attrs"];
        }

        var slideLayoutResFilename = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
        var slideLayoutResContent = await PPTXParser.readXmlFile(zip, slideLayoutResFilename);
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

        var slideMasterContent = await PPTXParser.readXmlFile(zip, masterFilename);
        var slideMasterTextStyles = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
        var slideMasterTables = indexNodes(slideMasterContent);

        var slideMasterResFilename = masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster") + ".rels";
        var slideMasterResContent = await PPTXParser.readXmlFile(zip, slideMasterResFilename);
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
            themeContent = await PPTXParser.readXmlFile(zip, themeFilename);
            var themeResContent = await PPTXParser.readXmlFile(zip, themeResFileName);
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
            digramFileContent = await PPTXParser.readXmlFile(zip, diagramFilename);
            if (digramFileContent !== null && digramFileContent !== undefined && digramFileContent != "") {
                var digramFileContentObjToStr = JSON.stringify(digramFileContent);
                digramFileContentObjToStr = digramFileContentObjToStr.replace(/dsp:/g, "p:");
                digramFileContent = JSON.parse(digramFileContentObjToStr);
            }

            var digramResContent = await PPTXParser.readXmlFile(zip, diagramResFileName);
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

        var slideContent = await PPTXParser.readXmlFile(zip, sldFileName, true);
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
