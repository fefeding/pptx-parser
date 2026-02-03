/**
 * Slide Processor
 * 幻灯片处理模块
 */
var xmlUtils = PPTXXmlUtils;
/**
 * 处理单个幻灯片
 * @param {Object} zip - JSZip实例
 * @param {string} sldFileName - 幻灯片文件名
 * @param {number} index - 幻灯片索引
 * @param {Object} slideSize - 幻灯片尺寸
 * @param {Object} settings - 设置对象
 * @param {number} slideFactor - 尺寸转换因子
 * @returns {string} HTML字符串
 */

var SlideProcessor = (function() {
    function processSingleSlide(zip, sldFileName, index, slideSize, settings, slideFactor) {
    // =====< Step 1 >=====
    // Read relationship filename of the slide (Get slideLayoutXX.xml)
    // @sldFileName: ppt/slides/slide1.xml
    // @resName: ppt/slides/_rels/slide1.xml.rels
    var resName = sldFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
    var resContent = xmlUtils.readXmlFile(zip, resName, false, settings.appVersion);
    var RelationshipArray = resContent["Relationships"]["Relationship"];
    // console.log("RelationshipArray: ", RelationshipArray);
    
    var layoutFilename = "";
    var diagramFilename = "";
    var slideResObj = {};
    
    if (RelationshipArray.constructor === Array) {
        for (var i = 0; i < RelationshipArray.length; i++) {
            switch (RelationshipArray[i]["attrs"]["Type"]) {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                    layoutFilename = RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
                    break;
                case "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing":
                    diagramFilename = RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
                    /* falls through */
                default:
                    slideResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                        "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        "target": RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/")
                    };
            }
        }
    } else {
        layoutFilename = RelationshipArray["attrs"]["Target"].replace("../", "ppt/");
    }

    // Open slideLayoutXX.xml
    var slideLayoutContent = xmlUtils.readXmlFile(zip, layoutFilename, false, settings.appVersion);
    var slideLayoutTables = indexNodes(slideLayoutContent);
    var sldLayoutClrOvr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);
    
    var slideLayoutClrOvride = null;
    if (sldLayoutClrOvr !== undefined) {
        slideLayoutClrOvride = sldLayoutClrOvr["attrs"];
    }

    // =====< Step 2 >=====
    // Read slide master filename of the slidelayout (Get slideMasterXX.xml)
    // @resName: ppt/slideLayouts/slideLayout1.xml
    // @masterName: ppt/slideLayouts/_rels/slideLayout1.xml.rels
    var slideLayoutResFilename = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
    var slideLayoutResContent = xmlUtils.readXmlFile(zip, slideLayoutResFilename, false, settings.appVersion);
    RelationshipArray = slideLayoutResContent["Relationships"]["Relationship"];
    
    var masterFilename = "";
    var layoutResObj = {};
    
    if (RelationshipArray.constructor === Array) {
        for (var j = 0; j < RelationshipArray.length; j++) {
            switch (RelationshipArray[j]["attrs"]["Type"]) {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
                    masterFilename = RelationshipArray[j]["attrs"]["Target"].replace("../", "ppt/");
                    break;
                default:
                    layoutResObj[RelationshipArray[j]["attrs"]["Id"]] = {
                        "type": RelationshipArray[j]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        "target": RelationshipArray[j]["attrs"]["Target"].replace("../", "ppt/")
                    };
            }
        }
    } else {
        masterFilename = RelationshipArray["attrs"]["Target"].replace("../", "ppt/");
    }

    // Open slideMasterXX.xml
    var slideMasterContent = xmlUtils.readXmlFile(zip, masterFilename, false, settings.appVersion);
    var slideMasterTextStyles = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
    var slideMasterTables = indexNodes(slideMasterContent);

    // /////////////////Amir/////////////
    // Open slideMasterXX.xml.rels
    var slideMasterResFilename = masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster") + ".rels";
    var slideMasterResContent = xmlUtils.readXmlFile(zip, slideMasterResFilename, false, settings.appVersion);
    RelationshipArray = slideMasterResContent["Relationships"]["Relationship"];
    
    var themeFilename = "";
    var masterResObj = {};
    
    if (RelationshipArray.constructor === Array) {
        for (var k = 0; k < RelationshipArray.length; k++) {
            switch (RelationshipArray[k]["attrs"]["Type"]) {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                    themeFilename = RelationshipArray[k]["attrs"]["Target"].replace("../", "ppt/");
                    break;
                default:
                    masterResObj[RelationshipArray[k]["attrs"]["Id"]] = {
                        "type": RelationshipArray[k]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        "target": RelationshipArray[k]["attrs"]["Target"].replace("../", "ppt/")
                    };
            }
        }
    } else {
        themeFilename = RelationshipArray["attrs"]["Target"].replace("../", "ppt/");
    }

    // Load Theme file
    var themeResObj = {};
    var themeContent = null;
    
    if (themeFilename !== undefined) {
        var themeName = themeFilename.split("/").pop();
        var themeResFileName = themeFilename.replace(themeName, "_rels/" + themeName) + ".rels";
        themeContent = xmlUtils.readXmlFile(zip, themeFilename, false, settings.appVersion);
        var themeResContent = xmlUtils.readXmlFile(zip, themeResFileName, false, settings.appVersion);
        
        if (themeResContent !== null) {
            var relationshipArray = themeResContent["Relationships"]["Relationship"];
            if (relationshipArray !== undefined) {
                if (relationshipArray.constructor === Array) {
                    for (var l = 0; l < relationshipArray.length; l++) {
                        themeResObj[relationshipArray[l]["attrs"]["Id"]] = {
                            "type": relationshipArray[l]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": relationshipArray[l]["attrs"]["Target"].replace("../", "ppt/")
                        };
                    }
                } else {
                    themeResObj[relationshipArray["attrs"]["Id"]] = {
                        "type": relationshipArray["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        "target": relationshipArray["attrs"]["Target"].replace("../", "ppt/")
                    };
                }
            }
        }
    }

    // Load diagram file
    var diagramResObj = {};
    var digramFileContent = {};
    var diagramFilename = null;
    
    // Diagram processing logic here...
    // 省略了diagram处理的部分代码以保持简洁

    // =====< Step 3 >=====
    var slideContent = xmlUtils.readXmlFile(zip, sldFileName, true, settings.appVersion);
    var nodes = slideContent["p:sld"]["p:cSld"]["p:spTree"];
    
    var warpObj = {
        "zip": zip,
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
        "defaultTextStyle": settings.defaultTextStyle
    };

    // 处理背景
    var bgResult = "";
    if (settings.processFullTheme === true) {
        bgResult = getBackground(warpObj, slideSize, index);
    }

    var bgColor = "";
    if (settings.processFullTheme === "colorsAndImageOnly") {
        bgColor = getSlideBackgroundFill(warpObj, index);
    }

    // 生成幻灯片HTML
    var slideClass = settings.slideMode && settings.slideType === "revealjs" ? "section" : "div";
    var result = `<${slideClass} class='slide' style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;
    result += bgResult;

    // 处理所有节点
    for (var nodeKey in nodes) {
        if (nodes[nodeKey].constructor === Array) {
            for (var i = 0; i < nodes[nodeKey].length; i++) {
                result += processNodesInSlide(nodeKey, nodes[nodeKey][i], nodes, warpObj, "slide");
            }
        } else {
            result += processNodesInSlide(nodeKey, nodes[nodeKey], nodes, warpObj, "slide");
        }
    }

    result += settings.slideMode && settings.slideType === "revealjs" ? "</div></section>" : "</div></div>";

    return result;
}

/**
 * 索引节点
 * @param {Object} content - 内容对象
 * @returns {Object} 索引表
 */
function indexNodes(content) {
    var keys = Object.keys(content);
    var spTreeNode = content[keys[0]]["p:cSld"]["p:spTree"];

    var idTable = {};
    var idxTable = {};
    var typeTable = {};

    for (var key in spTreeNode) {
        if (key === "p:nvGrpSpPr" || key === "p:grpSpPr") {
            continue;
        }

        var targetNode = spTreeNode[key];

        if (targetNode.constructor === Array) {
            for (var i = 0; i < targetNode.length; i++) {
                var nvSpPrNode = targetNode[i]["p:nvSpPr"];
                var id = getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                var idx = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                var type = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

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
            var id = getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
            var idx = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
            var type = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

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

        /**
         * getTextByPathList
         * @param {Object} node
         * @param {string Array} path
         */
        function getTextByPathList(node, path) {
    
            if (path.constructor !== Array) {
                throw Error("Error of path type! path is not array.");
            }

            if (node === undefined) {
                return undefined;
            }

            var l = path.length;
            for (var i = 0; i < l; i++) {
                node = node[path[i]];
                if (node === undefined) {
                    return undefined;
                }
            }

            return node;
}

function getBackground(warpObj, slideSize, index) {
    var slideContent = warpObj["slideContent"];
    var slideLayoutContent = warpObj["slideLayoutContent"];
    var slideMasterContent = warpObj["slideMasterContent"];

    var nodesSldLayout = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:spTree"]);
    var nodesSldMaster = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:spTree"]);
    var showMasterSp = getTextByPathList(slideLayoutContent, ["p:sldLayout", "attrs", "showMasterSp"]);
    var bgColor = getSlideBackgroundFill(warpObj, index);
    var result = "<div class='slide-background-" + index + "' style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;" + bgColor + "'>"
    var node_ph_type_ary = [];
    if (nodesSldLayout !== undefined) {
        for (var nodeKey in nodesSldLayout) {
            if (nodesSldLayout[nodeKey].constructor === Array) {
                for (var i = 0; i < nodesSldLayout[nodeKey].length; i++) {
                    var ph_type = getTextByPathList(nodesSldLayout[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    if (ph_type != "pic") {
                        result += processNodesInSlide(nodeKey, nodesSldLayout[nodeKey][i], nodesSldLayout, warpObj, "slideLayoutBg");
                    }
                }
            } else {
                var ph_type = getTextByPathList(nodesSldLayout[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                if (ph_type != "pic") {
                    result += processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], nodesSldLayout, warpObj, "slideLayoutBg");
                }
            }
        }
    }
    if (nodesSldMaster !== undefined && (showMasterSp == "1" || showMasterSp === undefined)) {
        for (var nodeKey in nodesSldMaster) {
            if (nodesSldMaster[nodeKey].constructor === Array) {
                for (var i = 0; i < nodesSldMaster[nodeKey].length; i++) {
                    var ph_type = getTextByPathList(nodesSldMaster[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    result += processNodesInSlide(nodeKey, nodesSldMaster[nodeKey][i], nodesSldMaster, warpObj, "slideMasterBg");
                }
            } else {
                var ph_type = getTextByPathList(nodesSldMaster[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                result += processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], nodesSldMaster, warpObj, "slideMasterBg");
            }
        }
    }
    return result;
}

function getSlideBackgroundFill(warpObj, index) {
    var slideContent = warpObj["slideContent"];
    var slideLayoutContent = warpObj["slideLayoutContent"];
    var slideMasterContent = warpObj["slideMasterContent"];

    var bgPr = getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgPr"]);
    var bgRef = getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgRef"]);
    var bgcolor;
    if (bgPr !== undefined) {
        var bgFillTyp = getFillType(bgPr);

        if (bgFillTyp == "SOLID_FILL") {
            var sldFill = bgPr["a:solidFill"];
            var clrMapOvr;
            var sldClrMapOvr = getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                clrMapOvr = sldClrMapOvr;
            } else {
                var sldClrMapOvr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    clrMapOvr = sldClrMapOvr;
                } else {
                    clrMapOvr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                }
            }
            var sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
            bgcolor = "background: #" + sldBgClr + ";";
        } else if (bgFillTyp == "GRADIENT_FILL") {
            bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
        } else if (bgFillTyp == "PIC_FILL") {
            bgcolor = getBgPicFill(bgPr, "slideBg", warpObj, undefined, index);
        }
    } else if (bgRef !== undefined) {
        var clrMapOvr;
        var sldClrMapOvr = getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
        if (sldClrMapOvr !== undefined) {
            clrMapOvr = sldClrMapOvr;
        } else {
            var sldClrMapOvr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                clrMapOvr = sldClrMapOvr;
            } else {
                clrMapOvr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
            }
        }
        var phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);
        var idx = Number(bgRef["attrs"]["idx"]);

        if (idx == 0 || idx == 1000) {
        } else if (idx > 0 && idx < 1000) {
        } else if (idx > 1000) {
            var trueIdx = idx - 1000;
            var bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
            var sortblAry = [];
            Object.keys(bgFillLst).forEach(function (key) {
                var bgFillLstTyp = bgFillLst[key];
                if (key != "attrs") {
                    if (bgFillLstTyp.constructor === Array) {
                        for (var i = 0; i < bgFillLstTyp.length; i++) {
                            var obj = {};
                            obj[key] = bgFillLstTyp[i];
                            obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                            obj["attrs"] = {
                                "order": bgFillLstTyp[i]["attrs"]["order"]
                            }
                            sortblAry.push(obj)
                        }
                    } else {
                        var obj = {};
                        obj[key] = bgFillLstTyp;
                        obj["idex"] = bgFillLstTyp["attrs"]["order"];
                        obj["attrs"] = {
                            "order": bgFillLstTyp["attrs"]["order"]
                        }
                        sortblAry.push(obj)
                    }
                }
            });
            var sortByOrder = sortblAry.slice(0);
            sortByOrder.sort(function (a, b) {
                return a.idex - b.idex;
            });
            var bgFillLstIdx = sortByOrder[trueIdx - 1];
            var bgFillTyp = getFillType(bgFillLstIdx);
            if (bgFillTyp == "SOLID_FILL") {
                var sldFill = bgFillLstIdx["a:solidFill"];
                var sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                bgcolor = "background: #" + sldBgClr + ";";
            } else if (bgFillTyp == "GRADIENT_FILL") {
                bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
            } else {
                console.log(bgFillTyp)
            }
        }
    } else {
        bgPr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr"]);
        bgRef = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgRef"]);
        var clrMapOvr;
        var sldClrMapOvr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
        if (sldClrMapOvr !== undefined) {
            clrMapOvr = sldClrMapOvr;
        } else {
            clrMapOvr = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
        }
        var phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);
        var idx = Number(bgRef["attrs"]["idx"]);

        if (idx == 0 || idx == 1000) {
        } else if (idx > 0 && idx < 1000) {
        } else if (idx > 1000) {
            var trueIdx = idx - 1000;
            var bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
            var sortblAry = [];
            Object.keys(bgFillLst).forEach(function (key) {
                var bgFillLstTyp = bgFillLst[key];
                if (key != "attrs") {
                    if (bgFillLstTyp.constructor === Array) {
                        for (var i = 0; i < bgFillLstTyp.length; i++) {
                            var obj = {};
                            obj[key] = bgFillLstTyp[i];
                            obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                            obj["attrs"] = {
                                "order": bgFillLstTyp[i]["attrs"]["order"]
                            }
                            sortblAry.push(obj)
                        }
                    } else {
                        var obj = {};
                        obj[key] = bgFillLstTyp;
                        obj["idex"] = bgFillLstTyp["attrs"]["order"];
                        obj["attrs"] = {
                            "order": bgFillLstTyp["attrs"]["order"]
                        }
                        sortblAry.push(obj)
                    }
                }
            });
            var sortByOrder = sortblAry.slice(0);
            sortByOrder.sort(function (a, b) {
                return a.idex - b.idex;
            });
            var bgFillLstIdx = sortByOrder[trueIdx - 1];
            var bgFillTyp = getFillType(bgFillLstIdx);
            if (bgFillTyp == "SOLID_FILL") {
                var sldFill = bgFillLstIdx["a:solidFill"];
                var sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                bgcolor = "background: #" + sldBgClr + ";";
            } else if (bgFillTyp == "GRADIENT_FILL") {
                bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
            } else {
                console.log(bgFillTyp)
            }
        }
    }
    return bgcolor;
}


    return {
        processSingleSlide: processSingleSlide
    };
})();