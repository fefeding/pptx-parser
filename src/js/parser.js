


    import { PPTXConstants } from './core/constants.js';
    import { PPTXUtils } from './core/utils.js';

    import { PPTXColorUtils } from './core/color.js';
    import parseXml from './core/xml-parser.js';

    // 全局变量，将在初始化时设置
    let app_verssion;
    let defaultTextStyle = null;
    let tableStyles;
    let styleTable = {};
    let slideFactor = PPTXConstants.SLIDE_FACTOR;
    let fontSizeFactor = PPTXConstants.FONT_SIZE_FACTOR;
    let slideWidth = 0;
    let slideHeight = 0;
    let isSlideMode = false;
    let processFullTheme = true;
    let settings;
    let slideLayoutClrOvride;

    // 回调函数变量（替换 window._ 全局变量）
    let _processSingleSlideCallback = null;
    let _processNodesInSlideCallback = null;
    let _getBackgroundCallback = null;
    let _getSlideBackgroundFillCallback = null;

    /**
     * 配置解析器
     * @param {Object} config - 配置对象
     */
    function configure(config) {
        settings = config;
        processFullTheme = settings.themeProcess;
        if (config.processSingleSlide) {
            _processSingleSlideCallback = config.processSingleSlide;
        }
        if (config.processNodesInSlide) {
            _processNodesInSlideCallback = config.processNodesInSlide;
        }
        if (config.getBackground) {
            _getBackgroundCallback = config.getBackground;
        }
        if (config.getSlideBackgroundFill) {
            _getSlideBackgroundFillCallback = config.getSlideBackgroundFill;
        }
    }

    /**
     * 主解析函数 - 解析 PPTX 文件并返回处理结果数组
     * @param {JSZip} zip - JSZip 实例
     * @returns {Array} 包含幻灯片、缩略图等数据的数组
     */
    function processPPTX(zip) {
        const post_ary = [];
        const dateBefore = new Date();

        parseXml.resetOrder();

        if (zip.file("docProps/thumbnail.jpeg") !== null) {
            const pptxThumbImg = PPTXUtils.base64ArrayBuffer(zip.file("docProps/thumbnail.jpeg").asArrayBuffer());
            post_ary.push({
                type: "pptx-thumb",
                data: pptxThumbImg,
                slide_num: -1
            });
        }

        const filesInfo = getContentTypes(zip);
        const slideSize = getSlideSizeAndSetDefaultTextStyle(zip);
        tableStyles = readXmlFile(zip, "ppt/tableStyles.xml");

        post_ary.push({
            type: "slideSize",
            data: slideSize,
            slide_num: 0
        });

        const numOfSlides = filesInfo.slides.length;
        for (let i = 0; i < numOfSlides; i++) {
            const filename = filesInfo.slides[i];
            let filename_no_path = "";
            let filename_no_path_ary = [];
            if (filename.indexOf("/") !== -1) {
                filename_no_path_ary = filename.split("/");
                filename_no_path = filename_no_path_ary.pop();
            } else {
                filename_no_path = filename;
            }
            let filename_no_path_no_ext = "";
            if (filename_no_path.indexOf(".") !== -1) {
                const filename_no_path_no_ext_ary = filename_no_path.split(".");
                filename_no_path_no_ext_ary.pop(); // remove extension
                filename_no_path_no_ext = filename_no_path_no_ext_ary.join(".");
            }
            let slide_number = 1;
            if (filename_no_path_no_ext !== "" && filename_no_path.indexOf("slide") !== -1) {
                slide_number = Number(filename_no_path_no_ext.substring(5));
            }

            let slideHtml;
            if (typeof _processSingleSlideCallback === 'function') {
                slideHtml = _processSingleSlideCallback(zip, filename, i, slideSize);
            } else {
                slideHtml = processSingleSlide(zip, filename, i, slideSize);
            }
            post_ary.push({
                type: "slide",
                data: slideHtml,
                slide_num: slide_number,
                file_name: filename_no_path_no_ext
            });
            post_ary.push({
                type: "progress-update",
                slide_num: (numOfSlides + i + 1),
                data: (i + 1) * 100 / numOfSlides
            });
        }

        post_ary.sort((a, b) => a.slide_num - b.slide_num);

        const dateAfter = new Date();
        post_ary.push({
            type: "ExecutionTime",
            data: dateAfter - dateBefore
        });
        return post_ary;
    }

    /**
     * 读取 XML 文件
     * @param {JSZip} zip - JSZip 实例
     * @param {string} filename - 文件名
     * @param {boolean} isSlideContent - 是否为幻灯片内容
     * @returns {Object|null} 解析后的 XML 对象
     */
    function readXmlFile(zip, filename, isSlideContent) {
        try {
            let fileEntry = zip.file(filename);
            if (!fileEntry && !filename.startsWith("ppt/") && !filename.startsWith("[Content_Types].xml") && !filename.startsWith("docProps/")) {
                // 尝试添加 ppt/ 前缀
                fileEntry = zip.file("ppt/" + filename);
            }
            if (!fileEntry) {
                return null;
            }
            let fileContent = fileEntry.asText();
            if (isSlideContent && app_verssion <= 12) {
                // Office 2007: 移除 "<![CDATA[ ... ]]>" 标签
                fileContent = fileContent.replace(/<!\[CDATA\[(.*?)\]\]>/g, '$1');
            }

            const xmlData = parseXml(fileContent, { simplify: 1 });

            if (xmlData["?xml"] !== undefined) {
                return xmlData["?xml"];
            } else {
                return xmlData;
            }
        } catch (e) {
            return null;
        }
    }

    /**
     * 获取内容类型 - 解析 [Content_Types].xml
     * @param {JSZip} zip - JSZip 实例
     * @returns {Object} 包含 slides 和 slideLayouts 数组
     */
    function getContentTypes(zip) {
        const ContentTypesJson = readXmlFile(zip, "[Content_Types].xml");

        const subObj = ContentTypesJson.Types.Override;
        const slidesLocArray = [];
        const slideLayoutsLocArray = [];
        for (let i = 0; i < subObj.length; i++) {
            switch (subObj[i].attrs.ContentType) {
                case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
                    slidesLocArray.push(subObj[i].attrs.PartName.substr(1));
                    break;
                case "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml":
                    slideLayoutsLocArray.push(subObj[i].attrs.PartName.substr(1));
                    break;
                default:
                    break;
            }
        }
        return {
            slides: slidesLocArray,
            slideLayouts: slideLayoutsLocArray
        };
    }

    /**
     * 获取幻灯片尺寸并设置默认文本样式
     * @param {JSZip} zip - JSZip 实例
     * @returns {Object} 包含 width 和 height 的对象
     */
    function getSlideSizeAndSetDefaultTextStyle(zip) {
        const app = readXmlFile(zip, "docProps/app.xml");
        if (!app) {
            return null;
        }
        const app_verssion_str = app.Properties.AppVersion;
        app_verssion = parseInt(app_verssion_str);

        const content = readXmlFile(zip, "ppt/presentation.xml");
        const sldSzAttrs = content["p:presentation"]["p:sldSz"].attrs;
        const sldSzWidth = parseInt(sldSzAttrs.cx);
        const sldSzHeight = parseInt(sldSzAttrs.cy);
        const sldSzType = sldSzAttrs.type;

        // 1 inches = 96px = 2.54cm
        // 1 EMU = 1 / 914400 inch
        // Pixel = EMUs * Resolution / 914400 (Resolution = 96)

        defaultTextStyle = content["p:presentation"]["p:defaultTextStyle"];
        // 根据 PPTX 规范（ECMA-376），p:defaultTextStyle 是可选元素
        if (defaultTextStyle === undefined || defaultTextStyle === null) {
            defaultTextStyle = {};
        }

        slideWidth = Math.floor(sldSzWidth * slideFactor + settings.incSlide.width);
        slideHeight = Math.floor(sldSzHeight * slideFactor + settings.incSlide.height);
        const rtenObj = {
            width: slideWidth,
            height: slideHeight
        };
        return rtenObj;
    }

    /**
     * 索引节点 - 为幻灯片节点创建索引表
     * @param {Object} content - 幻灯片内容
     * @returns {Object} 包含 idTable、idxTable 和 typeTable 的对象
     */
    function indexNodes(content) {
        const keys = Object.keys(content);
        const spTreeNode = content[keys[0]]["p:cSld"]["p:spTree"];

        const idTable = {};
        const idxTable = {};
        const typeTable = {};

        for (const key in spTreeNode) {
            if (key === "p:nvGrpSpPr" || key === "p:grpSpPr") {
                continue;
            }

            const targetNode = spTreeNode[key];

            if (Array.isArray(targetNode)) {
                for (let i = 0; i < targetNode.length; i++) {
                    const nvSpPrNode = targetNode[i]["p:nvSpPr"];
                    const id = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                    const idx = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                    const type = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

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
                const nvSpPrNode = targetNode["p:nvSpPr"];
                const id = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                const idx = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                const type = PPTXUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

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

        return { idTable, idxTable, typeTable };
    }

    /**
     * 获取背景 - 生成幻灯片背景 HTML
     * @param {Object} warpObj - 包含解析信息的对象
     * @param {Object} slideSize - 幻灯片尺寸
     * @param {number} index - 幻灯片索引
     * @returns {string} 背景 HTML
     */
    function getBackground(warpObj, slideSize, index) {
        let bgResult = "";
        if (processFullTheme === true) {
            // 读取 slide 节点中的背景
            const bgNode = PPTXUtils.getTextByPathList(warpObj.slideContent, ["p:sld", "p:cSld", "p:bg"]);
            if (bgNode) {
                const bgPr = bgNode["p:bgPr"];
                if (bgPr) {
                    // 纯色填充
                    const solidFill = PPTXUtils.getTextByPathList(bgPr, ["a:solidFill"]);
                    if (solidFill) {
                        const color = PPTXColorUtils.getFillColor(solidFill, warpObj.themeContent, warpObj.themeResObj, warpObj.slideLayoutClrOvride);
                        if (color) {
                            bgResult = `<div class='slide-background-${index}' style='position:absolute;width:${slideSize.width}px;height:${slideSize.height}px;background-color:${color};'></div>`;
                        }
                    }
                }
            }
        }
        return bgResult;
    }

    /**
     * 获取幻灯片背景填充 - 返回 CSS 背景样式
     * @param {Object} warpObj - 包含解析信息的对象
     * @param {number} index - 幻灯片索引
     * @returns {string} CSS 背景样式字符串
     */
    function getSlideBackgroundFill(warpObj, index) {
        let bgColor = "";
        if (processFullTheme === "colorsAndImageOnly") {
            const bgNode = PPTXUtils.getTextByPathList(warpObj.slideContent, ["p:sld", "p:cSld", "p:bg"]);
            if (bgNode) {
                const bgPr = bgNode["p:bgPr"];
                if (bgPr) {
                    const solidFill = PPTXUtils.getTextByPathList(bgPr, ["a:solidFill"]);
                    if (solidFill) {
                        const color = PPTXColorUtils.getFillColor(solidFill, warpObj.themeContent, warpObj.themeResObj, warpObj.slideLayoutClrOvride);
                        if (color) {
                            bgColor = `background-color:${color};`;
                        }
                    }
                }
            }
        }
        return bgColor;
    }

    /**
     * 处理单个幻灯片
     * @param {JSZip} zip - JSZip 实例
     * @param {string} sldFileName - 幻灯片文件名
     * @param {number} index - 幻灯片索引
     * @param {Object} slideSize - 幻灯片尺寸
     * @returns {string} 幻灯片 HTML
     */
    function processSingleSlide(zip, sldFileName, index, slideSize) {
        const resName = sldFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
        const resContent = PPTXParser.readXmlFile(zip, resName);
        let RelationshipArray = resContent.Relationships.Relationship;
        let layoutFilename = "";
        let diagramFilename = "";
        const slideResObj = {};
        if (Array.isArray(RelationshipArray)) {
            for (let i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i].attrs.Type) {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                        layoutFilename = PPTXUtils.resolveRelationshipTarget(resName, RelationshipArray[i].attrs.Target);
                        break;
                    case "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing":
                        diagramFilename = PPTXUtils.resolveRelationshipTarget(resName, RelationshipArray[i].attrs.Target);
                        slideResObj[RelationshipArray[i].attrs.Id] = {
                            type: RelationshipArray[i].attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target: PPTXUtils.resolveRelationshipTarget(resName, RelationshipArray[i].attrs.Target)
                        };
                        break;
                    default:
                        slideResObj[RelationshipArray[i].attrs.Id] = {
                            type: RelationshipArray[i].attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target: PPTXUtils.resolveRelationshipTarget(resName, RelationshipArray[i].attrs.Target)
                        };
                }
            }
        } else {
            layoutFilename = PPTXUtils.resolveRelationshipTarget(resName, RelationshipArray.attrs.Target);
        }

        const slideLayoutContent = PPTXParser.readXmlFile(zip, layoutFilename);
        const slideLayoutTables = indexNodes(slideLayoutContent);
        const sldLayoutClrOvr = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);

        if (sldLayoutClrOvr !== undefined) {
            slideLayoutClrOvride = sldLayoutClrOvr.attrs;
        }

        const slideLayoutResFilename = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
        const slideLayoutResContent = PPTXParser.readXmlFile(zip, slideLayoutResFilename);
        RelationshipArray = slideLayoutResContent.Relationships.Relationship;
        let masterFilename = "";
        const layoutResObj = {};
        if (Array.isArray(RelationshipArray)) {
            for (let i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i].attrs.Type) {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
                        masterFilename = PPTXUtils.resolveRelationshipTarget(slideLayoutResFilename, RelationshipArray[i].attrs.Target);
                        break;
                    default:
                        layoutResObj[RelationshipArray[i].attrs.Id] = {
                            type: RelationshipArray[i].attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target: PPTXUtils.resolveRelationshipTarget(slideLayoutResFilename, RelationshipArray[i].attrs.Target)
                        };
                }
            }
        } else {
            masterFilename = PPTXUtils.resolveRelationshipTarget(slideLayoutResFilename, RelationshipArray.attrs.Target);
        }

        const slideMasterContent = PPTXParser.readXmlFile(zip, masterFilename);
        const slideMasterTextStyles = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
        const slideMasterTables = indexNodes(slideMasterContent);

        const slideMasterResFilename = masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster") + ".rels";
        const slideMasterResContent = PPTXParser.readXmlFile(zip, slideMasterResFilename);
        RelationshipArray = slideMasterResContent.Relationships.Relationship;
        let themeFilename = "";
        const masterResObj = {};
        if (Array.isArray(RelationshipArray)) {
            for (let i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i].attrs.Type) {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                        themeFilename = PPTXUtils.resolveRelationshipTarget(slideMasterResFilename, RelationshipArray[i].attrs.Target);
                        break;
                    default:
                        masterResObj[RelationshipArray[i].attrs.Id] = {
                            type: RelationshipArray[i].attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target: PPTXUtils.resolveRelationshipTarget(slideMasterResFilename, RelationshipArray[i].attrs.Target)
                        };
                }
            }
        } else {
            themeFilename = PPTXUtils.resolveRelationshipTarget(slideMasterResFilename, RelationshipArray.attrs.Target);
        }

        const themeResObj = {};
        let themeContent = {};
        if (themeFilename !== undefined) {
            const themeName = themeFilename.split("/").pop();
            const themeResFileName = themeFilename.replace(themeName, "_rels/" + themeName) + ".rels";
            themeContent = PPTXParser.readXmlFile(zip, themeFilename);
            const themeResContent = PPTXParser.readXmlFile(zip, themeResFileName);
            if (themeResContent !== null) {
                const relationshipArray = themeResContent.Relationships.Relationship;
                if (relationshipArray !== undefined) {
                    if (Array.isArray(relationshipArray)) {
                        for (let i = 0; i < relationshipArray.length; i++) {
                            themeResObj[relationshipArray[i].attrs.Id] = {
                                type: relationshipArray[i].attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                target: PPTXUtils.resolveRelationshipTarget(themeResFileName, relationshipArray[i].attrs.Target)
                            };
                        }
                    } else {
                        themeResObj[relationshipArray.attrs.Id] = {
                            type: relationshipArray.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target: PPTXUtils.resolveRelationshipTarget(themeResFileName, relationshipArray.attrs.Target)
                        };
                    }
                }
            }
        }

        const diagramResObj = {};
        let digramFileContent = {};
        if (diagramFilename) {
            const diagName = diagramFilename.split("/").pop();
            const diagramResFileName = diagramFilename.replace(diagName, "_rels/" + diagName) + ".rels";
            digramFileContent = PPTXParser.readXmlFile(zip, diagramFilename);
            if (digramFileContent !== null && digramFileContent !== undefined && digramFileContent !== "") {
                let digramFileContentObjToStr = JSON.stringify(digramFileContent);
                digramFileContentObjToStr = digramFileContentObjToStr.replace(/dsp:/g, "p:");
                digramFileContent = JSON.parse(digramFileContentObjToStr);
            }

            const digramResContent = PPTXParser.readXmlFile(zip, diagramResFileName);
            if (digramResContent !== null) {
                const relationshipArray = digramResContent.Relationships.Relationship;
                if (Array.isArray(relationshipArray)) {
                    for (let i = 0; i < relationshipArray.length; i++) {
                        diagramResObj[relationshipArray[i].attrs.Id] = {
                            type: relationshipArray[i].attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target: PPTXUtils.resolveRelationshipTarget(diagramResFileName, relationshipArray[i].attrs.Target)
                        };
                    }
                } else {
                    diagramResObj[relationshipArray.attrs.Id] = {
                        type: relationshipArray.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target: PPTXUtils.resolveRelationshipTarget(diagramResFileName, relationshipArray.attrs.Target)
                    };
                }
            }
        }

        const slideContent = PPTXParser.readXmlFile(zip, sldFileName, true);
        const nodes = slideContent["p:sld"]["p:cSld"]["p:spTree"];
        const warpObj = {
            zip,
            sldFileName,
            slideLayoutContent,
            slideLayoutTables,
            slideMasterContent,
            slideMasterTables,
            slideContent,
            slideResObj,
            slideMasterTextStyles,
            layoutResObj,
            masterResObj,
            themeContent,
            themeResObj,
            digramFileContent,
            diagramResObj,
            tableStyles,
            defaultTextStyle: PPTXParser.defaultTextStyle
        };
        let bgResult = "";
        if (processFullTheme === true) {
            if (typeof _getBackgroundCallback === 'function') {
                bgResult = _getBackgroundCallback(warpObj, slideSize, index);
            } else {
                bgResult = getBackground(warpObj, slideSize, index);
            }
        }

        let bgColor = "";
        if (processFullTheme === "colorsAndImageOnly") {
            if (typeof _getSlideBackgroundFillCallback === 'function') {
                bgColor = _getSlideBackgroundFillCallback(warpObj, index);
            } else {
                bgColor = getSlideBackgroundFill(warpObj, index);
            }
        }

        let result;
        if (settings.slideMode && settings.slideType === "revealjs") {
            result = `<section class='slide' style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;
        } else {
            result = `<div class='slide' data-slide-index='${index}' style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;
        }
        result += bgResult;

        const processNodesFunc = _processNodesInSlideCallback;
        if (!processNodesFunc) {
            return result + (settings.slideMode && settings.slideType === "revealjs" ? "</section>" : "</div>");
        }
        for (const nodeKey in nodes) {
            if (Array.isArray(nodes[nodeKey])) {
                for (let i = 0; i < nodes[nodeKey].length; i++) {
                    const nodeResult = processNodesFunc(nodeKey, nodes[nodeKey][i], nodes, warpObj, "slide");
                    result += nodeResult;
                }
            } else {
                const nodeResult = processNodesFunc(nodeKey, nodes[nodeKey], nodes, warpObj, "slide");
                result += nodeResult;
            }
        }
        return result + (settings.slideMode && settings.slideType === "revealjs" ? "</section>" : "</div>");
    }

    /**
     * PPTX 解析器模块
     * 提供解析 PPTX 文件的核心功能
     */
    const PPTXParser = {
        configure,
        processPPTX,
        readXmlFile,
        getContentTypes,
        getSlideSizeAndSetDefaultTextStyle,
        indexNodes,
        processSingleSlide,
        getBackground,
        getSlideBackgroundFill,
        slideFactor,
        fontSizeFactor,
        slideWidth,
        slideHeight,
        isSlideMode,
        processFullTheme,
        styleTable,
        tableStyles,
        defaultTextStyle,
        app_verssion
    };


export { PPTXParser };