/**
 * Node Processors
 * 节点处理器模块 - 处理各种幻灯片节点
 */

/**
 * 处理Slide中的节点
 * @param {string} nodeKey - 节点键
 * @param {Object} nodeValue - 节点值
 * @param {Object} nodes - 节点集合
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 来源
 * @param {string} sType - 类型
 * @returns {string} HTML结果
 */

var NodeProcessors = (function() {
    function processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType) {
    var result = "";

    switch (nodeKey) {
        case "p:sp":    // Shape, Text
            result = processSpNode(nodeValue, nodes, warpObj, source, sType);
            break;
        case "p:cxnSp":    // Shape, Text (with connection)
            result = processCxnSpNode(nodeValue, nodes, warpObj, source, sType);
            break;
        case "p:pic":    // Picture
            result = processPicNode(nodeValue, warpObj, source, sType);
            break;
        case "p:graphicFrame":    // Chart, Diagram, Table
            result = processGraphicFrameNode(nodeValue, warpObj, source, sType);
            break;
        case "p:grpSp":
            result = processGroupSpNode(nodeValue, warpObj, source);
            break;
        case "mc:AlternateContent": // Equations and formulas as Image
            var mcFallbackNode = getTextByPathList(nodeValue, ["mc:Fallback"]);
            result = processGroupSpNode(mcFallbackNode, warpObj, source);
            break;
        default:
            // console.log("nodeKey: ", nodeKey)
    }

    return result;
}

/**
 * 处理形状节点
 * @param {Object} node - 节点
 * @param {Object} pNode - 父节点
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 来源
 * @param {string} sType - 类型
 * @returns {string} HTML结果
 */
    function processSpNode(node, pNode, warpObj, source, sType) {
    var id = getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "id"]);
    var name = getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "name"]);
    var idx = (getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]) === undefined) ? undefined : getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]);
    var type = (getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]) === undefined) ? undefined : getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
    var order = getTextByPathList(node, ["attrs", "order"]);
    var isUserDrawnBg;
    if (source == "slideLayoutBg" || source == "slideMasterBg") {
        var userDrawn = getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "attrs", "userDrawn"]);
        if (userDrawn == "1") {
            isUserDrawnBg = true;
        } else {
            isUserDrawnBg = false;
        }
    }
    var slideLayoutSpNode = undefined;
    var slideMasterSpNode = undefined;

    if (idx !== undefined) {
        slideLayoutSpNode = warpObj["slideLayoutTables"]["idxTable"][idx];
        if (type !== undefined) {
            slideMasterSpNode = warpObj["slideMasterTables"]["typeTable"][type];
        } else {
            slideMasterSpNode = warpObj["slideMasterTables"]["idxTable"][idx];
        }
    } else {
        if (type !== undefined) {
            slideLayoutSpNode = warpObj["slideLayoutTables"]["typeTable"][type];
            slideMasterSpNode = warpObj["slideMasterTables"]["typeTable"][type];
        }
    }

    if (type === undefined) {
        txBoxVal = getTextByPathList(node, ["p:nvSpPr", "p:cNvSpPr", "attrs", "txBox"]);
        if (txBoxVal == "1") {
            type = "textBox";
        }
    }
    if (type === undefined) {
        type = getTextByPathList(slideLayoutSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
        if (type === undefined) {
            if (source == "diagramBg") {
                type = "diagram";
            } else {
                type = "obj";
            }
        }
    }
    return shapeGenerator.genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, isUserDrawnBg, sType, source);
}

/**
 * 处理连接形状节点
 * @param {Object} node - 节点
 * @param {Object} pNode - 父节点
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 来源
 * @param {string} sType - 类型
 * @returns {string} HTML结果
 */
    function processCxnSpNode(node, pNode, warpObj, source, sType) {
    var id = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
    var name = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
    var idx = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
    var type = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
    var order = node["attrs"]["order"];

    return shapeGenerator.genShape(node, pNode, undefined, undefined, id, name, idx, type, order, warpObj, undefined, sType, source);
}

/**
 * 处理图片节点
 * @param {Object} node - 节点
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 来源
 * @param {string} sType - 类型
 * @returns {string} HTML结果
 */
    function processPicNode(node, warpObj, source, sType) {
    var rtrnData = "";
    var mediaPicFlag = false;
    var order = node["attrs"]["order"];

    var rid = node["p:blipFill"]["a:blip"]["attrs"]["r:embed"];
    var resObj;
    if (source == "slideMasterBg") {
        resObj = warpObj["masterResObj"];
    } else if (source == "slideLayoutBg") {
        resObj = warpObj["layoutResObj"];
    } else {
        resObj = warpObj["slideResObj"];
    }
    var imgName = (resObj[rid] !== undefined) ? resObj[rid]["target"] : undefined;

    if (imgName === undefined) {
        console.warn("Image reference not found in resObj for rid:", rid);
        return "";
    }

    var imgFileExt = PPTXImageUtils.extractFileExtension(imgName).toLowerCase();
    var zip = warpObj["zip"];
    
    var context = 'slide';
    if (source == "slideMasterBg") {
        context = 'master';
    } else if (source == "slideLayoutBg") {
        context = 'layout';
    }
    
    var imgFile = PPTXFileUtils.findMediaFile(zip, imgName, context, '');
    if (imgFile === null) {
        console.warn("Image file not found in processPicNode:", imgName);
        return "";
    }
    var imgArrayBuffer = imgFile.asArrayBuffer();
    var mimeType = "";
    var xfrmNode = node["p:spPr"]["a:xfrm"];
    if (xfrmNode === undefined) {
        var idx = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "idx"]);
        var type = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "type"]);
        if (idx !== undefined) {
            xfrmNode = getTextByPathList(warpObj["slideLayoutTables"], ["idxTable", idx, "p:spPr", "a:xfrm"]);
        }
    }
    var rotate = 0;
    var rotateNode = getTextByPathList(node, ["p:spPr", "a:xfrm", "attrs", "rot"]);
    if (rotateNode !== undefined) {
        rotate = angleToDegrees(rotateNode);
    }
    var vdoNode = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:videoFile"]);
    var vdoRid, vdoFile, vdoFileExt, vdoMimeType, uInt8Array, blob, vdoBlob, mediaSupportFlag = false, isVdeoLink = false;
    var mediaProcess = settings.mediaProcess;
    if (vdoNode !== undefined & mediaProcess) {
        vdoRid = vdoNode["attrs"]["r:link"];
        vdoFile = resObj[vdoRid]["target"];
        var checkIfLink = PPTXImageUtils.IsVideoLink(vdoFile);
        if (checkIfLink) {
            vdoFile = escapeHtml(vdoFile);
            isVdeoLink = true;
            mediaSupportFlag = true;
            mediaPicFlag = true;
        } else {
            vdoFileExt = PPTXImageUtils.extractFileExtension(vdoFile).toLowerCase();
            if (vdoFileExt == "mp4" || vdoFileExt == "webm" || vdoFileExt == "ogg") {
                var vdoFileObj = PPTXFileUtils.findMediaFile(zip, vdoFile, context, '');
                if (vdoFileObj === null) {
                    console.warn("Video file not found:", vdoFile);
                } else {
                    uInt8Array = vdoFileObj.asArrayBuffer();
                    vdoMimeType = PPTXImageUtils.getMimeType(vdoFileExt);
                    blob = new Blob([uInt8Array], {
                        type: vdoMimeType
                    });
                    vdoBlob = URL.createObjectURL(blob);
                    mediaSupportFlag = true;
                    mediaPicFlag = true;
                }
            }
        }
    }
    var audioNode = getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:audioFile"]);
    var audioRid, audioFile, audioFileExt, audioMimeType, uInt8ArrayAudio, blobAudio, audioBlob;
    var audioPlayerFlag = false;
    var audioObjc;
    if (audioNode !== undefined & mediaProcess) {
        audioRid = audioNode["attrs"]["r:link"];
        audioFile = resObj[audioRid]["target"];
        audioFileExt = PPTXImageUtils.extractFileExtension(audioFile).toLowerCase();
        if (audioFileExt == "mp3" || audioFileExt == "wav" || audioFileExt == "ogg") {
            var audioFileObj = PPTXFileUtils.findMediaFile(zip, audioFile, context, '');
            if (audioFileObj === null) {
                console.warn("Audio file not found:", audioFile);
            } else {
                uInt8ArrayAudio = audioFileObj.asArrayBuffer();
                blobAudio = new Blob([uInt8ArrayAudio]);
                audioBlob = URL.createObjectURL(blobAudio);
                var cx = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * 20;
                var cy = xfrmNode["a:ext"]["attrs"]["cy"];
                var x = parseInt(xfrmNode["a:off"]["attrs"]["x"]) / 2.5;
                var y = xfrmNode["a:off"]["attrs"]["y"];
                audioObjc = {
                    "a:ext": {
                        "attrs": {
                            "cx": cx,
                            "cy": cy
                        }
                    },
                    "a:off": {
                        "attrs": {
                            "x": x,
                            "y": y
                        }
                    }
                }
                audioPlayerFlag = true;
                mediaSupportFlag = true;
                mediaPicFlag = true;
            }
        }
    }
    mimeType = PPTXImageUtils.getMimeType(imgFileExt);
    rtrnData = "<div class='block content' style='" +
        ((mediaProcess && audioPlayerFlag) ? getPosition(audioObjc, node, undefined, undefined) : getPosition(xfrmNode, node, undefined, undefined)) +
        ((mediaProcess && audioPlayerFlag) ? getSize(audioObjc, undefined, undefined) : getSize(xfrmNode, undefined, undefined)) +
        " z-index: " + order + ";" +
        "transform: rotate(" + rotate + "deg);'>";
    if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
        rtrnData += "<img src='data:" + mimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%; height: 100%'/>";
    } else if ((vdoNode !== undefined || audioNode !== undefined) && mediaProcess && mediaSupportFlag) {
        if (vdoNode !== undefined && !isVdeoLink) {
            rtrnData += "<video  src='" + vdoBlob + "' controls style='width: 100%; height: 100%'>Your browser does not support video tag.</video>";
        } else if (vdoNode !== undefined && isVdeoLink) {
            rtrnData += "<iframe   src='" + vdoFile + "' controls style='width: 100%; height: 100%'></iframe >";
        }
        if (audioNode !== undefined) {
            rtrnData += '<audio id="audio_player" controls ><source src="' + audioBlob + '"></audio>';
        }
    }
    if (!mediaSupportFlag && mediaPicFlag) {
        rtrnData += "<span style='color:red;font-size:40px;position: absolute;'>This media file Not supported by HTML5</span>";
    }
    if ((vdoNode !== undefined || audioNode !== undefined) && !mediaProcess && mediaSupportFlag) {
        console.log("Founded supported media file but media process disabled (mediaProcess=false)");
    }
    rtrnData += "</div>";
    return rtrnData;
}

/**
 * 处理图形框架节点
 * @param {Object} node - 节点
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 来源
 * @param {string} sType - 类型
 * @returns {string} HTML结果
 */
    function processGraphicFrameNode(node, warpObj, source, sType) {
    var result = "";
    var graphicTypeUri = getTextByPathList(node, ["a:graphic", "a:graphicData", "attrs", "uri"]);

    switch (graphicTypeUri) {
        case "http://schemas.openxmlformats.org/drawingml/2006/table":
            result = genTable(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/chart":
            result = genChart(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
            result = genDiagram(node, warpObj, source, sType);
            break;
        case "http://schemas.openxmlformats.org/presentationml/2006/ole":
            var oleObjNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent", "mc:Fallback","p:oleObj"]);
            
            if (oleObjNode === undefined) {
                oleObjNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "p:oleObj"]);
            }
            if (oleObjNode !== undefined){
                result = processGroupSpNode(oleObjNode, warpObj, source);
            }
            break;
        default:
    }

    return result;
}

/**
 * 处理组合形状节点
 * @param {Object} node - 节点
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 来源
 * @returns {string} HTML结果
 */
    function processGroupSpNode(node, warpObj, source) {
    var xfrmNode = getTextByPathList(node, ["p:grpSpPr", "a:xfrm"]);
    if (xfrmNode !== undefined) {
        var x = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * slideFactor;
        var y = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * slideFactor;
        
        var chx, chy, chcx, chcy;
        
        if (xfrmNode["a:chOff"] !== undefined && xfrmNode["a:chOff"]["attrs"] !== undefined) {
            chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * slideFactor;
            chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * slideFactor;
        } else {
            chx = x;
            chy = y;
        }
        
        var cx = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * slideFactor;
        var cy = parseInt(xfrmNode["a:ext"]["attrs"]["cy"]) * slideFactor;
        
        if (xfrmNode["a:chExt"] !== undefined && xfrmNode["a:chExt"]["attrs"] !== undefined) {
            chcx = parseInt(xfrmNode["a:chExt"]["attrs"]["cx"]) * slideFactor;
            chcy = parseInt(xfrmNode["a:chExt"]["attrs"]["cy"]) * slideFactor;
        } else {
            chcx = cx;
            chcy = cy;
        }
        var rotate = parseInt(xfrmNode["attrs"]["rot"])
        var rotStr = "";
        var top = y - chy,
            left = x - chx,
            width = cx - chcx,
            height = cy - chcy;

        var sType = "group";
        if (!isNaN(rotate)) {
            rotate = angleToDegrees(rotate);
            rotStr += "transform: rotate(" + rotate + "deg) ; transform-origin: center;";
            if (rotate != 0) {
                top = y;
                left = x;
                width = cx;
                height = cy;
                sType = "group-rotate";
            }
        }
    }
    var grpStyle = "";

    if (rotStr !== undefined && rotStr != "") {
        grpStyle += rotStr;
    }

    if (top !== undefined) {
        grpStyle += "top: " + top + "px;";
    }
    if (left !== undefined) {
        grpStyle += "left: " + left + "px;";
    }
    if (width !== undefined) {
        grpStyle += "width:" + width + "px;";
    }
    if (height !== undefined) {
        grpStyle += "height: " + height + "px;";
    }
    var order = node["attrs"]["order"];

    var result = "<div class='block group' style='z-index: " + order + ";" + grpStyle + "'>";

    for (var nodeKey in node) {
        if (node[nodeKey].constructor === Array) {
            for (var i = 0; i < node[nodeKey].length; i++) {
                result += processNodesInSlide(nodeKey, node[nodeKey][i], node, warpObj, source, sType);
            }
        } else {
            result += processNodesInSlide(nodeKey, node[nodeKey], node, warpObj, source, sType);
        }
    }

    result += "</div>";

    return result;
}

// Helper function - 需要迁移或导入
function getTextByPathList(obj, pathList) {
    // TODO: 实现getTextByPathList逻辑
    return null;
}


    /**
     * 生成表格
     * @param {Object} node - 节点
     * @param {Object} warpObj - 包装对象
     * @returns {string} HTML表格
     */
    function genTable(node, warpObj) {
        var order = node["attrs"]["order"];
        var tableNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
        var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
        /////////////////////////////////////////Amir////////////////////////////////////////////////
        var getTblPr = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
        var getColsGrid = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
        var tblDir = "";
        if (getTblPr !== undefined) {
            var isRTL = getTblPr["attrs"]["rtl"];
            tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
        }
        var firstRowAttr = getTblPr["attrs"]["firstRow"]; //associated element <a:firstRow> in the table styles
        var firstColAttr = getTblPr["attrs"]["firstCol"]; //associated element <a:firstCol> in the table styles
        var lastRowAttr = getTblPr["attrs"]["lastRow"]; //associated element <a:lastRow> in the table styles
        var lastColAttr = getTblPr["attrs"]["lastCol"]; //associated element <a:lastCol> in the table styles
        var bandRowAttr = getTblPr["attrs"]["bandRow"]; //associated element <a:band1H>, <a:band2H> in the table styles
        var bandColAttr = getTblPr["attrs"]["bandCol"]; //associated element <a:band1V>, <a:band2V> in the table styles
        //console.log("getTblPr: ", getTblPr);
        var tblStylAttrObj = {
            isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
            isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
            isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
            isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
            isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
            isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
        };

        var thisTblStyle;
        var tbleStyleId = getTblPr["a:tableStyleId"];
        if (tbleStyleId !== undefined) {
            var tbleStylList = tableStyles["a:tblStyleLst"]["a:tblStyle"];
            if (tbleStylList !== undefined) {
                if (tbleStylList.constructor === Array) {
                    for (var k = 0; k < tbleStylList.length; k++) {
                        if (tbleStylList[k]["attrs"]["styleId"] == tbleStyleId) {
                            thisTblStyle = tbleStylList[k];
                        }
                    }
                } else {
                    if (tbleStylList["attrs"]["styleId"] == tbleStyleId) {
                        thisTblStyle = tbleStylList;
                    }
                }
            }
        }
        if (thisTblStyle !== undefined) {
            thisTblStyle["tblStylAttrObj"] = tblStylAttrObj;
            warpObj["thisTbiStyle"] = thisTblStyle;
        }
        var tblStyl = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
        var tblBorderStyl = getTextByPathList(tblStyl, ["a:tcBdr"]);
        var tbl_borders = "";
        if (tblBorderStyl !== undefined) {
            tbl_borders = getTableBorders(tblBorderStyl, warpObj);
        }
        var tbl_bgcolor = "";
        var tbl_opacity = 1;
        var tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
        //console.log( "thisTblStyle:", thisTblStyle, "warpObj:", warpObj)
        if (tbl_bgFillschemeClr !== undefined) {
            tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
        }
        if (tbl_bgFillschemeClr === undefined) {
            tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
            tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
        }
        if (tbl_bgcolor !== "") {
            tbl_bgcolor = "background-color: #" + tbl_bgcolor + ";";
        }
        ////////////////////////////////////////////////////////////////////////////////////////////
        var tableHtml = "<table " + tblDir + " style='border-collapse: collapse;" +
            getPosition(xfrmNode, node, undefined, undefined) +
            getSize(xfrmNode, undefined, undefined) +
            " z-index: " + order + ";" +
            tbl_borders + ";" +
            tbl_bgcolor + "'>";

        var trNodes = tableNode["a:tr"];
        if (trNodes.constructor !== Array) {
            trNodes = [trNodes];
        }
        //if (trNodes.constructor === Array) {
            //multi rows
            var totalrowSpan = 0;
            var rowSpanAry = [];
            for (var i = 0; i < trNodes.length; i++) {
                //////////////rows Style ////////////Amir
                var rowHeightParam = trNodes[i]["attrs"]["h"];
                var rowHeight = 0;
                var rowsStyl = "";
                if (rowHeightParam !== undefined) {
                    rowHeight = parseInt(rowHeightParam) * slideFactor;
                    rowsStyl += "height:" + rowHeight + "px;";
                }
                var fillColor = "";
                var row_borders = "";
                var fontClrPr = "";
                var fontWeight = "";
                var band_1H_fillColor;
                var band_2H_fillColor;

                if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
                    var bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                    if (bgFillschemeClr !== undefined) {
                        var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                        if (local_fillColor !== undefined) {
                            fillColor = local_fillColor;
                        }
                    }
                    var rowTxtStyl = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        var local_fontColor = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                        if (local_fontColor !== undefined) {
                            fontClrPr = local_fontColor;
                        }

                        var local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                        if (local_fontWeight != "") {
                            fontWeight = local_fontWeight;
                        }
                    }
                }

                if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {

                    var bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                    if (bgFillschemeClr !== undefined) {
                        var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                        if (local_fillColor !== undefined) {
                            fillColor = local_fillColor;
                        }
                    }
                    var borderStyl = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
                    if (borderStyl !== undefined) {
                        var local_row_borders = getTableBorders(borderStyl, warpObj);
                        if (local_row_borders != "") {
                            row_borders = local_row_borders;
                        }
                    }
                    var rowTxtStyl = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        var local_fontClrPr = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                        if (local_fontClrPr !== undefined) {
                            fontClrPr = local_fontClrPr;
                        }
                        var local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                        if (local_fontWeight !== "") {
                            fontWeight = local_fontWeight;
                        }
                    }

                } else if (i > 0 && tblStylAttrObj["isBandRowAttr"] == 1 && thisTblStyle !== undefined) {
                    fillColor = "";
                    row_borders = undefined;
                    if ((i % 2) == 0 && thisTblStyle["a:band2H"] !== undefined) {
                        // console.log("i: ", i, 'thisTblStyle["a:band2H"]:', thisTblStyle["a:band2H"])
                        //check if there is a row bg
                        var bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== "") {
                                fillColor = local_fillColor;
                                band_2H_fillColor = local_fillColor;
                            }
                        }


                        var borderStyl = getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            var local_row_borders = getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        var rowTxtStyl = getTextByPathList(thisTblStyle, ["a:band2H", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            var local_fontClrPr = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }
                        }

                        var local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");

                        if (local_fontWeight !== "") {
                            fontWeight = local_fontWeight;
                        }
                    }
                    if ((i % 2) != 0 && thisTblStyle["a:band1H"] !== undefined) {
                        var bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                                band_1H_fillColor = local_fillColor;
                            }
                        }
                        var borderStyl = getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            var local_row_borders = getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        var rowTxtStyl = getTextByPathList(thisTblStyle, ["a:band1H", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            var local_fontClrPr = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }
                            var local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight != "") {
                                fontWeight = local_fontWeight;
                            }
                        }
                    }

                }
                //last row
                if (i == (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1 && thisTblStyle !== undefined) {
                    var bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                    if (bgFillschemeClr !== undefined) {
                        var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                        if (local_fillColor !== undefined) {
                            fillColor = local_fillColor;
                        }
                        // var local_colorOpacity = getColorOpacity(bgFillschemeClr);
                        // if(local_colorOpacity !== undefined){
                        //     colorOpacity = local_colorOpacity;
                        // }
                    }
                    var borderStyl = getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:tcBdr"]);
                    if (borderStyl !== undefined) {
                        var local_row_borders = getTableBorders(borderStyl, warpObj);
                        if (local_row_borders != "") {
                            row_borders = local_row_borders;
                        }
                    }
                    var rowTxtStyl = getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        var local_fontClrPr = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                        if (local_fontClrPr !== undefined) {
                            fontClrPr = local_fontClrPr;
                        }

                        var local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                        if (local_fontWeight !== "") {
                            fontWeight = local_fontWeight;
                        }
                    }
                }
                rowsStyl += ((row_borders !== undefined) ? row_borders : "");
                rowsStyl += ((fontClrPr !== undefined) ? " color: #" + fontClrPr + ";" : "");
                rowsStyl += ((fontWeight != "") ? " font-weight:" + fontWeight + ";" : "");
                if (fillColor !== undefined && fillColor != "") {
                    //rowsStyl += "background-color: rgba(" + hexToRgbNew(fillColor) + "," + colorOpacity + ");";
                    rowsStyl += "background-color: #" + fillColor + ";";
                }
                tableHtml += "<tr style='" + rowsStyl + "'>";
                ////////////////////////////////////////////////

                var tcNodes = trNodes[i]["a:tc"];
                if (tcNodes !== undefined) {
                    if (tcNodes.constructor === Array) {
                        //multi columns
                        var j = 0;
                        if (rowSpanAry.length == 0) {
                            rowSpanAry = Array.apply(null, Array(tcNodes.length)).map(function () { return 0 });
                        }
                        var totalColSpan = 0;
                        while (j < tcNodes.length) {
                            if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                                var a_sorce;
                                //j=0 : first col
                                if (j == 0 && tblStylAttrObj["isFrstColAttr"] == 1) {
                                    a_sorce = "a:firstCol";
                                    if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) &&
                                        getTextByPathList(thisTblStyle, ["a:seCell"]) !== undefined) {
                                        a_sorce = "a:seCell";
                                    } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 &&
                                        getTextByPathList(thisTblStyle, ["a:neCell"]) !== undefined) {
                                        a_sorce = "a:neCell";
                                    }
                                } else if ((j > 0 && tblStylAttrObj["isBandColAttr"] == 1) &&
                                    !(tblStylAttrObj["isFrstColAttr"] == 1 && i == 0) &&
                                    !(tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1)) &&
                                    j != (tcNodes.length - 1)) {

                                    if ((j % 2) != 0) {

                                        var aBandNode = getTextByPathList(thisTblStyle, ["a:band2V"]);
                                        if (aBandNode === undefined) {
                                            aBandNode = getTextByPathList(thisTblStyle, ["a:band1V"]);
                                            if (aBandNode !== undefined) {
                                                a_sorce = "a:band2V";
                                            }
                                        } else {
                                            a_sorce = "a:band2V";
                                        }

                                    }
                                }

                                if (j == (tcNodes.length - 1) && tblStylAttrObj["isLstColAttr"] == 1) {
                                    a_sorce = "a:lastCol";
                                    if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) && getTextByPathList(thisTblStyle, ["a:swCell"]) !== undefined) {
                                        a_sorce = "a:swCell";
                                    } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 && getTextByPathList(thisTblStyle, ["a:nwCell"]) !== undefined) {
                                        a_sorce = "a:nwCell";
                                    }
                                }

                                var cellParmAry = getTableCellParams(tcNodes[j], getColsGrid, i, j, thisTblStyle, a_sorce, warpObj);
                                var text = cellParmAry[0];
                                var colStyl = cellParmAry[1];
                                var cssName = cellParmAry[2];
                                var rowSpan = cellParmAry[3];
                                var colSpan = cellParmAry[4];

                                if (rowSpan !== undefined) {
                                    totalrowSpan++;
                                    rowSpanAry[j] = parseInt(rowSpan) - 1;
                                    tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' rowspan ='" +
                                        parseInt(rowSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                                } else if (colSpan !== undefined) {
                                    tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' colspan = '" +
                                        parseInt(colSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                                    totalColSpan = parseInt(colSpan) - 1;
                                } else {
                                    tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' style = '" + colStyl + "'>" + text + "</td>";
                                }

                            } else {
                                if (rowSpanAry[j] != 0) {
                                    rowSpanAry[j] -= 1;
                                }
                                if (totalColSpan != 0) {
                                    totalColSpan--;
                                }
                            }
                            j++;
                        }
                    } else {
                        //single column 

                        var a_sorce;
                        if (tblStylAttrObj["isFrstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                            a_sorce = "a:firstCol";

                        } else if ((tblStylAttrObj["isBandColAttr"] == 1) && !(tblStylAttrObj["isLstRowAttr"] == 1)) {

                            var aBandNode = getTextByPathList(thisTblStyle, ["a:band2V"]);
                            if (aBandNode === undefined) {
                                aBandNode = getTextByPathList(thisTblStyle, ["a:band1V"]);
                                if (aBandNode !== undefined) {
                                    a_sorce = "a:band2V";
                                }
                            } else {
                                a_sorce = "a:band2V";
                            }
                        }

                        if (tblStylAttrObj["isLstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                            a_sorce = "a:lastCol";
                        }

                        var cellParmAry = getTableCellParams(tcNodes, getColsGrid, i, undefined, thisTblStyle, a_sorce, warpObj);
                        var text = cellParmAry[0];
                        var colStyl = cellParmAry[1];
                        var cssName = cellParmAry[2];
                        var rowSpan = cellParmAry[3];

                        if (rowSpan !== undefined) {
                            tableHtml += "<td  class='" + cssName + "' rowspan='" + parseInt(rowSpan) + "' style = '" + colStyl + "'>" + text + "</td>";
                        } else {
                            tableHtml += "<td class='" + cssName + "' style='" + colStyl + "'>" + text + "</td>";
                        }
                    }
                }
                tableHtml += "</tr>";
            }
            //////////////////////////////////////////////////////////////////////////////////
        

            return tableHtml;
        }
    
    /**
     * 生成图表
     * @param {Object} node - 节点
     * @param {Object} warpObj - 包装对象
     * @returns {string} HTML图表容器
     */
    function genChart(node, warpObj) {

        var order = node["attrs"]["order"];
        var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
        var result = "<div id='chart" + chartID + "' class='block content' style='" +
            getPosition(xfrmNode, node, undefined, undefined) + getSize(xfrmNode, undefined, undefined) +
            " z-index: " + order + ";'></div>";

        var rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
        var refName = warpObj["slideResObj"][rid]["target"];
        var content = PPTXXmlUtils.readXmlFile(warpObj["zip"], refName);
        var plotArea = getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);

        var chartData = null;
        for (var key in plotArea) {
            switch (key) {
                case "c:lineChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "lineChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:barChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "barChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:pieChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "pieChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:pie3DChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "pie3DChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:areaChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "areaChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:scatterChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "scatterChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:catAx":
                    break;
                case "c:valAx":
                    break;
                default:
            }
        }

        if (chartData !== null) {
            MsgQueue.push(chartData);
        }

        chartID++;
        return result;
    }

    /**
     * 生成图表
     * @param {Object} node - 节点
     * @param {Object} warpObj - 包装对象
     * @param {string} source - 来源
     * @param {string} sType - 类型
     * @returns {string} HTML图表容器
     */
    function genDiagram(node, warpObj, source, sType) {
        //console.log(warpObj)
        //PPTXXmlUtils.readXmlFile(zip, sldFileName)
        /**files define the diagram:
         * 1-colors#.xml,
         * 2-data#.xml, 
         * 3-layout#.xml,
         * 4-quickStyle#.xml.
         * 5-drawing#.xml, which Microsoft added as an extension for persisting diagram layout information.
         */
        ///get colors#.xml, data#.xml , layout#.xml , quickStyle#.xml
        var order = node["attrs"]["order"];
        var zip = warpObj["zip"];
        var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
        var dgmRelIds = getTextByPathList(node, ["a:graphic", "a:graphicData", "dgm:relIds", "attrs"]);
        //console.log(dgmRelIds)
        var dgmClrFileId = dgmRelIds["r:cs"];
        var dgmDataFileId = dgmRelIds["r:dm"];
        var dgmLayoutFileId = dgmRelIds["r:lo"];
        var dgmQuickStyleFileId = dgmRelIds["r:qs"];
        var dgmClrFileName = warpObj["slideResObj"][dgmClrFileId].target,
            dgmDataFileName = warpObj["slideResObj"][dgmDataFileId].target,
            dgmLayoutFileName = warpObj["slideResObj"][dgmLayoutFileId].target;
        dgmQuickStyleFileName = warpObj["slideResObj"][dgmQuickStyleFileId].target;
        //console.log("dgmClrFileName: " , dgmClrFileName,", dgmDataFileName: ",dgmDataFileName,", dgmLayoutFileName: ",dgmLayoutFileName,", dgmQuickStyleFileName: ",dgmQuickStyleFileName);
        var dgmClr = PPTXXmlUtils.readXmlFile(zip, dgmClrFileName);
        var dgmData = PPTXXmlUtils.readXmlFile(zip, dgmDataFileName);
        var dgmLayout = PPTXXmlUtils.readXmlFile(zip, dgmLayoutFileName);
        var dgmQuickStyle = PPTXXmlUtils.readXmlFile(zip, dgmQuickStyleFileName);
        //console.log(dgmClr,dgmData,dgmLayout,dgmQuickStyle)
        ///get drawing#.xml
        // var dgmDrwFileName = "";
        // var dataModelExt = getTextByPathList(dgmData, ["dgm:dataModel", "dgm:extLst", "a:ext", "dsp:dataModelExt", "attrs"]);
        // if (dataModelExt !== undefined) {
        //     var dgmDrwFileId = dataModelExt["relId"];
        //     dgmDrwFileName = warpObj["slideResObj"][dgmDrwFileId]["target"];
        // }
        // var dgmDrwFile = "";
        // if (dgmDrwFileName != "") {
        //     dgmDrwFile = PPTXXmlUtils.readXmlFile(zip, dgmDrwFileName);
        // }
        // var dgmDrwSpArray = getTextByPathList(dgmDrwFile, ["dsp:drawing", "dsp:spTree", "dsp:sp"]);
        //var dgmDrwSpArray = getTextByPathList(warpObj["digramFileContent"], ["dsp:drawing", "dsp:spTree", "dsp:sp"]);
        var dgmDrwSpArray = getTextByPathList(warpObj["digramFileContent"], ["p:drawing", "p:spTree", "p:sp"]);
        var rslt = "";
        if (dgmDrwSpArray !== undefined) {
            var dgmDrwSpArrayLen = dgmDrwSpArray.length;
            for (var i = 0; i < dgmDrwSpArrayLen; i++) {
                var dspSp = dgmDrwSpArray[i];
                // var dspSpObjToStr = JSON.stringify(dspSp);
                // var pSpStr = dspSpObjToStr.replace(/dsp:/g, "p:");
                // var pSpStrToObj = JSON.parse(pSpStr);
                //console.log("pSpStrToObj[" + i + "]: ", pSpStrToObj);
                //rslt += processSpNode(pSpStrToObj, node, warpObj, "diagramBg", sType)
                rslt += processSpNode(dspSp, node, warpObj, "diagramBg", sType);
            }
            // dgmDrwFile: "dsp:"-> "p:"
        }

        return "<div class='block diagram-content' style='" +
            getPosition(xfrmNode, node, undefined, undefined, sType) +
            getSize(xfrmNode, undefined, undefined) +
            "'>" + rslt + "</div>";
    }

    return {
        processNodesInSlide: processNodesInSlide,
        processSpNode: processSpNode,
        processCxnSpNode: processCxnSpNode,
        processPicNode: processPicNode,
        processGraphicFrameNode: processGraphicFrameNode,
        processGroupSpNode: processGroupSpNode,
        genTable: genTable,
        genChart: genChart,
        genDiagram: genDiagram
    };
})();