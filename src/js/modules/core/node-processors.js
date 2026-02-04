/**
 * Node Processors
 * 节点处理器模块 - 处理各种幻灯片节点
 */

var shapeGenerator = ShapeGenerator;
var colorUtils = PPTXColorUtils;
var xmlUtils = PPTXXmlUtils;
var fileUtils = PPTXFileUtils;
var imageUtils = PPTXImageUtils;

// 定义 slideFactor 变量，用于位置和大小计算
var slideFactor = 96 / 914400;

// 定义 fontSizeFactor 变量，用于字体大小计算
var fontSizeFactor = 4 / 3.2;

// 定义 chartID 变量，用于图表ID生成
var chartID = 0;

// 定义 MsgQueue 变量，用于消息队列
var MsgQueue = new Array();

// 定义 styleTable 变量，用于样式表
var styleTable = {};

// 定义 is_first_br 变量，用于处理换行符
var is_first_br = true;

// 辅助函数：获取文件的MIME类型
function getMimeType(ext) {
    var mimeTypes = {
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'png': 'image/png',
        'gif': 'image/gif',
        'bmp': 'image/bmp',
        'svg': 'image/svg+xml',
        'webp': 'image/webp'
    };
    return mimeTypes[ext.toLowerCase()] || 'application/octet-stream';
}

// 辅助函数：将ArrayBuffer转换为base64
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
        chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
        a = (chunk & 16515072) >> 18; // 262144
        b = (chunk & 258048) >> 12;   // 4096
        c = (chunk & 4032) >> 6;      // 64
        d = chunk & 63;               // 63
        base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d];
    }

    if (byteRemainder == 1) {
        chunk = bytes[mainLength];
        a = (chunk & 252) >> 2; // 15<<2
        b = (chunk & 3) << 4;   // 3<<4
        base64 += encodings[a] + encodings[b] + '==';
    } else if (byteRemainder == 2) {
        chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];
        a = (chunk & 64512) >> 10; // 15<<10
        b = (chunk & 1008) >> 4;   // 63<<4
        c = (chunk & 15) << 2;     // 15<<2
        base64 += encodings[a] + encodings[b] + encodings[c] + '=';
    }

    return base64;
}

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
    function processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType, settings) {
    var result = "";

    switch (nodeKey) {
        case "p:sp":    // Shape, Text
            result = processSpNode(nodeValue, nodes, warpObj, source, sType);
            break;
        case "p:cxnSp":    // Shape, Text (with connection)
            result = processCxnSpNode(nodeValue, nodes, warpObj, source, sType);
            break;
        case "p:pic":    // Picture
            result = processPicNode(nodeValue, warpObj, source, sType, settings);
            break;
        case "p:graphicFrame":    // Chart, Diagram, Table
            result = processGraphicFrameNode(nodeValue, warpObj, source, sType);
            break;
        case "p:grpSp":
            result = processGroupSpNode(nodeValue, warpObj, source);
            break;
        case "mc:AlternateContent": // Equations and formulas as Image
            var mcFallbackNode = xmlUtils.getTextByPathList(nodeValue, ["mc:Fallback"]);
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
    var id = xmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "id"]);
    var name = xmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "name"]);
    var idx = (xmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]) === undefined) ? undefined : xmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]);
    var type = (xmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]) === undefined) ? undefined : xmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
    var order = xmlUtils.getTextByPathList(node, ["attrs", "order"]);
    var isUserDrawnBg;
    if (source == "slideLayoutBg" || source == "slideMasterBg") {
        var userDrawn = xmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "attrs", "userDrawn"]);
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
        txBoxVal = xmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvSpPr", "attrs", "txBox"]);
        if (txBoxVal == "1") {
            type = "textBox";
        }
    }
    if (type === undefined) {
        type = xmlUtils.getTextByPathList(slideLayoutSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
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
    function processPicNode(node, warpObj, source, sType, settings) {
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

    var imgFileExt = imageUtils.extractFileExtension(imgName).toLowerCase();
    var zip = warpObj["zip"];
    
    var context = 'slide';
    if (source == "slideMasterBg") {
        context = 'master';
    } else if (source == "slideLayoutBg") {
        context = 'layout';
    }
    
    var imgFile = fileUtils.findMediaFile(zip, imgName, context, '');
    if (imgFile === null) {
        console.warn("Image file not found in processPicNode:", imgName);
        return "";
    }
    var imgArrayBuffer = imgFile.asArrayBuffer();
    var mimeType = "";
    var xfrmNode = node["p:spPr"]["a:xfrm"];
    if (xfrmNode === undefined) {
        var idx = xmlUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "idx"]);
        var type = xmlUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "type"]);
        if (idx !== undefined) {
            xfrmNode = xmlUtils.getTextByPathList(warpObj["slideLayoutTables"], ["idxTable", idx, "p:spPr", "a:xfrm"]);
        }
    }
    var rotate = 0;
    var rotateNode = xmlUtils.getTextByPathList(node, ["p:spPr", "a:xfrm", "attrs", "rot"]);
    if (rotateNode !== undefined) {
        rotate = colorUtils.angleToDegrees(rotateNode);
    }
    var vdoNode = xmlUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:videoFile"]);
    var vdoRid, vdoFile, vdoFileExt, vdoMimeType, uInt8Array, blob, vdoBlob, mediaSupportFlag = false, isVdeoLink = false;
    var mediaProcess = settings.mediaProcess;
    if (vdoNode !== undefined & mediaProcess) {
        vdoRid = vdoNode["attrs"]["r:link"];
        vdoFile = resObj[vdoRid]["target"];
        var checkIfLink = imageUtils.IsVideoLink(vdoFile);
        if (checkIfLink) {
            vdoFile = xmlUtils.escapeHtml(vdoFile);
            isVdeoLink = true;
            mediaSupportFlag = true;
            mediaPicFlag = true;
        } else {
            vdoFileExt = imageUtils.extractFileExtension(vdoFile).toLowerCase();
            if (vdoFileExt == "mp4" || vdoFileExt == "webm" || vdoFileExt == "ogg") {
                var vdoFileObj = fileUtils.findMediaFile(zip, vdoFile, context, '');
                if (vdoFileObj === null) {
                    console.warn("Video file not found:", vdoFile);
                } else {
                    uInt8Array = vdoFileObj.asArrayBuffer();
                    vdoMimeType = imageUtils.getMimeType(vdoFileExt);
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
    var audioNode = xmlUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:audioFile"]);
    var audioRid, audioFile, audioFileExt, audioMimeType, uInt8ArrayAudio, blobAudio, audioBlob;
    var audioPlayerFlag = false;
    var audioObjc;
    if (audioNode !== undefined & mediaProcess) {
        audioRid = audioNode["attrs"]["r:link"];
        audioFile = resObj[audioRid]["target"];
        audioFileExt = imageUtils.extractFileExtension(audioFile).toLowerCase();
        if (audioFileExt == "mp3" || audioFileExt == "wav" || audioFileExt == "ogg") {
            var audioFileObj = fileUtils.findMediaFile(zip, audioFile, context, '');
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
    mimeType = imageUtils.getMimeType(imgFileExt);
    rtrnData = "<div class='block content' style='" +
        ((mediaProcess && audioPlayerFlag) ? getPosition(audioObjc, node, undefined, undefined) : getPosition(xfrmNode, node, undefined, undefined)) +
        ((mediaProcess && audioPlayerFlag) ? getSize(audioObjc, undefined, undefined) : getSize(xfrmNode, undefined, undefined)) +
        " z-index: " + order + ";" +
        "transform: rotate(" + rotate + "deg);'>";
    if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
        rtrnData += "<img src='data:" + mimeType + ";base64," + imageUtils.base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%; height: 100%'/>";
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
    var graphicTypeUri = xmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "attrs", "uri"]);

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
            var oleObjNode = xmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent", "mc:Fallback","p:oleObj"]);
            
            if (oleObjNode === undefined) {
                oleObjNode = xmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "p:oleObj"]);
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
    var xfrmNode = xmlUtils.getTextByPathList(node, ["p:grpSpPr", "a:xfrm"]);
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
            rotate = colorUtils.angleToDegrees(rotate);
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
                result += processNodesInSlide(nodeKey, node[nodeKey][i], node, warpObj, source, undefined, undefined);
            }
        } else {
            result += processNodesInSlide(nodeKey, node[nodeKey], node, warpObj, source, undefined, undefined);
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
        var tableNode = xmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
        var xfrmNode = xmlUtils.getTextByPathList(node, ["p:xfrm"]);
        /////////////////////////////////////////Amir////////////////////////////////////////////////
        var getTblPr = xmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
        var getColsGrid = xmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
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
        if (tbleStyleId !== undefined && warpObj["tableStyles"] !== null && warpObj["tableStyles"] !== undefined) {
            var tbleStylList = warpObj["tableStyles"]["a:tblStyleLst"]["a:tblStyle"];
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
        var tblStyl = xmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
        var tblBorderStyl = xmlUtils.getTextByPathList(tblStyl, ["a:tcBdr"]);
        var tbl_borders = "";
        if (tblBorderStyl !== undefined) {
            tbl_borders = getTableBorders(tblBorderStyl, warpObj);
        }
        var tbl_bgcolor = "";
        var tbl_opacity = 1;
        var tbl_bgFillschemeClr = xmlUtils.getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
        //console.log( "thisTblStyle:", thisTblStyle, "warpObj:", warpObj)
        if (tbl_bgFillschemeClr !== undefined) {
            tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
        }
        if (tbl_bgFillschemeClr === undefined) {
            tbl_bgFillschemeClr = xmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
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
                    var bgFillschemeClr = xmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                    if (bgFillschemeClr !== undefined) {
                        var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                        if (local_fillColor !== undefined) {
                            fillColor = local_fillColor;
                        }
                    }
                    var rowTxtStyl = xmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        var local_fontColor = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                        if (local_fontColor !== undefined) {
                            fontClrPr = local_fontColor;
                        }

                        var local_fontWeight = ((xmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                        if (local_fontWeight != "") {
                            fontWeight = local_fontWeight;
                        }
                    }
                }

                if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {

                    var bgFillschemeClr = xmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                    if (bgFillschemeClr !== undefined) {
                        var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                        if (local_fillColor !== undefined) {
                            fillColor = local_fillColor;
                        }
                    }
                    var borderStyl = xmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
                    if (borderStyl !== undefined) {
                        var local_row_borders = getTableBorders(borderStyl, warpObj);
                        if (local_row_borders != "") {
                            row_borders = local_row_borders;
                        }
                    }
                    var rowTxtStyl = xmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        var local_fontClrPr = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                        if (local_fontClrPr !== undefined) {
                            fontClrPr = local_fontClrPr;
                        }
                        var local_fontWeight = ((xmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
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
                        var bgFillschemeClr = xmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== "") {
                                fillColor = local_fillColor;
                                band_2H_fillColor = local_fillColor;
                            }
                        }


                        var borderStyl = xmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            var local_row_borders = getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        var rowTxtStyl = xmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            var local_fontClrPr = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }
                        }

                        var local_fontWeight = ((xmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");

                        if (local_fontWeight !== "") {
                            fontWeight = local_fontWeight;
                        }
                    }
                    if ((i % 2) != 0 && thisTblStyle["a:band1H"] !== undefined) {
                        var bgFillschemeClr = xmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                                band_1H_fillColor = local_fillColor;
                            }
                        }
                        var borderStyl = xmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            var local_row_borders = getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        var rowTxtStyl = xmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            var local_fontClrPr = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }
                            var local_fontWeight = ((xmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight != "") {
                                fontWeight = local_fontWeight;
                            }
                        }
                    }

                }
                //last row
                if (i == (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1 && thisTblStyle !== undefined) {
                    var bgFillschemeClr = xmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:fill", "a:solidFill"]);
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
                    var borderStyl = xmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:tcBdr"]);
                    if (borderStyl !== undefined) {
                        var local_row_borders = getTableBorders(borderStyl, warpObj);
                        if (local_row_borders != "") {
                            row_borders = local_row_borders;
                        }
                    }
                    var rowTxtStyl = xmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        var local_fontClrPr = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                        if (local_fontClrPr !== undefined) {
                            fontClrPr = local_fontClrPr;
                        }

                        var local_fontWeight = ((xmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
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
                                        xmlUtils.getTextByPathList(thisTblStyle, ["a:seCell"]) !== undefined) {
                                        a_sorce = "a:seCell";
                                    } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 &&
                                        xmlUtils.getTextByPathList(thisTblStyle, ["a:neCell"]) !== undefined) {
                                        a_sorce = "a:neCell";
                                    }
                                } else if ((j > 0 && tblStylAttrObj["isBandColAttr"] == 1) &&
                                    !(tblStylAttrObj["isFrstColAttr"] == 1 && i == 0) &&
                                    !(tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1)) &&
                                    j != (tcNodes.length - 1)) {

                                    if ((j % 2) != 0) {

                                        var aBandNode = xmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                                        if (aBandNode === undefined) {
                                            aBandNode = xmlUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
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
                                    if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) && xmlUtils.getTextByPathList(thisTblStyle, ["a:swCell"]) !== undefined) {
                                        a_sorce = "a:swCell";
                                    } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 && xmlUtils.getTextByPathList(thisTblStyle, ["a:nwCell"]) !== undefined) {
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

                            var aBandNode = xmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                            if (aBandNode === undefined) {
                                aBandNode = xmlUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
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
        var xfrmNode = xmlUtils.getTextByPathList(node, ["p:xfrm"]);
        var result = "<div id='chart" + chartID + "' class='block content' style='" +
            getPosition(xfrmNode, node, undefined, undefined) + getSize(xfrmNode, undefined, undefined) +
            " z-index: " + order + ";'></div>";

        var rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
        var refName = warpObj["slideResObj"][rid]["target"];
        var content = xmlUtils.readXmlFile(warpObj["zip"], refName);
        var plotArea = xmlUtils.getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);

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

    function extractChartData(serNode) {
        var dataMat = new Array();

        if (serNode === undefined) {
            return dataMat;
        }

        if (serNode["c:xVal"] !== undefined) {
            var dataRow = new Array();
            eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                dataRow.push(parseFloat(innerNode["c:v"]));
                return "";
            });
            dataMat.push(dataRow);
            dataRow = new Array();
            eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                dataRow.push(parseFloat(innerNode["c:v"]));
                return "";
            });
            dataMat.push(dataRow);
        } else {
            eachElement(serNode, function (innerNode, index) {
                var dataRow = new Array();
                var colName = xmlUtils.getTextByPathList(innerNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]) || index;

                var rowNames = {};
                if (xmlUtils.getTextByPathList(innerNode, ["c:cat", "c:strRef", "c:strCache", "c:pt"]) !== undefined) {
                    eachElement(innerNode["c:cat"]["c:strRef"]["c:strCache"]["c:pt"], function (innerNode, index) {
                        rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                        return "";
                    });
                } else if (xmlUtils.getTextByPathList(innerNode, ["c:cat", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                    eachElement(innerNode["c:cat"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                        rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                        return "";
                    });
                }

                if (xmlUtils.getTextByPathList(innerNode, ["c:val", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                    eachElement(innerNode["c:val"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                        dataRow.push({ x: innerNode["attrs"]["idx"], y: parseFloat(innerNode["c:v"]) });
                        return "";
                    });
                }

                dataMat.push({ key: colName, values: dataRow, xlabels: rowNames });
                return "";
            });
        }

        return dataMat;
    }

    function eachElement(node, doFunction) {
        if (node === undefined) {
            return;
        }
        var result = "";
        if (node.constructor === Array) {
            var l = node.length;
            for (var i = 0; i < l; i++) {
                result += doFunction(node[i], i);
            }
        } else {
            result += doFunction(node, 0);
        }
        return result;
    }

    function getTableBorders(node, warpObj) {
        var borderStyle = "";
        if (node["a:bottom"] !== undefined) {
            var obj = {
                "p:spPr": {
                    "a:ln": node["a:bottom"]["a:ln"]
                }
            }
            var borders = getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-bottom");
        }
        if (node["a:top"] !== undefined) {
            var obj = {
                "p:spPr": {
                    "a:ln": node["a:top"]["a:ln"]
                }
            }
            var borders = getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-top");
        }
        if (node["a:right"] !== undefined) {
            var obj = {
                "p:spPr": {
                    "a:ln": node["a:right"]["a:ln"]
                }
            }
            var borders = getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-right");
        }
        if (node["a:left"] !== undefined) {
            var obj = {
                "p:spPr": {
                    "a:ln": node["a:left"]["a:ln"]
                }
            }
            var borders = getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-left");
        }

        return borderStyle;
    }

    function getBorder(node, pNode, isSvgMode, bType, warpObj) {
        var cssText, lineNode, subNodeTxt;

        if (bType == "shape") {
            cssText = "border: ";
            lineNode = node["p:spPr"]["a:ln"];
        } else if (bType == "text") {
            cssText = "";
            lineNode = node["a:rPr"]["a:ln"];
        }

        var is_noFill = xmlUtils.getTextByPathList(lineNode, ["a:noFill"]);
        if (is_noFill !== undefined) {
            return "hidden";
        }

        if (lineNode == undefined) {
            var lnRefNode = xmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
            if (lnRefNode !== undefined) {
                var lnIdx = xmlUtils.getTextByPathList(lnRefNode, ["attrs", "idx"]);
                lineNode = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"][Number(lnIdx) - 1];
            }
        }
        if (lineNode == undefined) {
            cssText = "";
            lineNode = node;
        }

        var borderColor;
        if (lineNode !== undefined) {
            var borderWidth = parseInt(xmlUtils.getTextByPathList(lineNode, ["attrs", "w"])) / 12700;
            if (isNaN(borderWidth) || borderWidth < 1) {
                cssText += (4/3) + "px ";
            } else {
                cssText += borderWidth + "px ";
            }
            var borderType = xmlUtils.getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
            if (borderType === undefined) {
                borderType = xmlUtils.getTextByPathList(lineNode, ["attrs", "cmpd"]);
            }
            var strokeDasharray = "0";
            switch (borderType) {
                case "solid":
                    cssText += "solid";
                    strokeDasharray = "0";
                    break;
                case "dash":
                    cssText += "dashed";
                    strokeDasharray = "5";
                    break;
                case "dashDot":
                    cssText += "dashed";
                    strokeDasharray = "5, 5, 1, 5";
                    break;
                case "dot":
                    cssText += "dotted";
                    strokeDasharray = "1, 5";
                    break;
                case "lgDash":
                    cssText += "dashed";
                    strokeDasharray = "10, 5";
                    break;
                case "dbl":
                    cssText += "double";
                    strokeDasharray = "0";
                    break;
                case "lgDashDotDot":
                    cssText += "dashed";
                    strokeDasharray = "10, 5, 1, 5, 1, 5";
                    break;
                case "sysDash":
                    cssText += "dashed";
                    strokeDasharray = "5, 2";
                    break;
                case "sysDashDot":
                    cssText += "dashed";
                    strokeDasharray = "5, 2, 1, 5";
                    break;
                case "sysDashDotDot":
                    cssText += "dashed";
                    strokeDasharray = "5, 2, 1, 5, 1, 5";
                    break;
                case "sysDot":
                    cssText += "dotted";
                    strokeDasharray = "2, 5";
                    break;
                case undefined:
                default:
                    cssText += "solid";
                    strokeDasharray = "0";
            }
            var fillTyp = colorUtils.getFillType(lineNode);
            if (fillTyp == "NO_FILL") {
                borderColor = isSvgMode ? "none" : "";
            } else if (fillTyp == "SOLID_FILL") {
                borderColor = getSolidFill(lineNode["a:solidFill"], undefined, undefined, warpObj);
            } else if (fillTyp == "GRADIENT_FILL") {
                borderColor = getGradientFill(lineNode["a:gradFill"], warpObj);
            } else if (fillTyp == "PATTERN_FILL") {
                borderColor = getPatternFill(lineNode["a:pattFill"], warpObj);
            }
        }

        if (borderColor === undefined) {
            var lnRefNode = xmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
            if (lnRefNode !== undefined) {
                borderColor = getSolidFill(lnRefNode, undefined, undefined, warpObj);
            }
        }

        if (borderColor === undefined) {
            if (isSvgMode) {
                borderColor = "none";
            } else {
                borderColor = "hidden";
            }
        } else {
            borderColor = "#" + borderColor;
        }
        cssText += " " + borderColor + " ";

        if (isSvgMode) {
            return { "color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray };
        } else {
            return cssText + ";";
        }
    }

    function getSolidFill(node, clrMap, phClr, warpObj) {
        if (node === undefined) {
            return undefined;
        }

        var color = "";
        var clrNode;
        if (node["a:srgbClr"] !== undefined) {
            clrNode = node["a:srgbClr"];
            color = xmlUtils.getTextByPathList(clrNode, ["attrs", "val"]);
        } else if (node["a:schemeClr"] !== undefined) {
            clrNode = node["a:schemeClr"];
            var schemeClr = xmlUtils.getTextByPathList(clrNode, ["attrs", "val"]);
            color = getSchemeColorFromTheme("a:" + schemeClr, clrMap, phClr, warpObj);
        } else if (node["a:scrgbClr"] !== undefined) {
            clrNode = node["a:scrgbClr"];
            var defBultColorVals = clrNode["attrs"];
            var red = (defBultColorVals["r"].indexOf("%") != -1) ? defBultColorVals["r"].split("%").shift() : defBultColorVals["r"];
            var green = (defBultColorVals["g"].indexOf("%") != -1) ? defBultColorVals["g"].split("%").shift() : defBultColorVals["g"];
            var blue = (defBultColorVals["b"].indexOf("%") != -1) ? defBultColorVals["b"].split("%").shift() : defBultColorVals["b"];
            color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100));
        } else if (node["a:prstClr"] !== undefined) {
            clrNode = node["a:prstClr"];
            var prstClr = xmlUtils.getTextByPathList(clrNode, ["attrs", "val"]);
            color = getColorName2Hex(prstClr);
        } else if (node["a:hslClr"] !== undefined) {
            clrNode = node["a:hslClr"];
            var defBultColorVals = clrNode["attrs"];
            var hue = Number(defBultColorVals["hue"]) / 100000;
            var sat = Number((defBultColorVals["sat"].indexOf("%") != -1) ? defBultColorVals["sat"].split("%").shift() : defBultColorVals["sat"]) / 100;
            var lum = Number((defBultColorVals["lum"].indexOf("%") != -1) ? defBultColorVals["lum"].split("%").shift() : defBultColorVals["lum"]) / 100;
            var hsl2rgb = hslToRgb(hue, sat, lum);
            color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b);
        } else if (node["a:sysClr"] !== undefined) {
            clrNode = node["a:sysClr"];
            var sysClr = xmlUtils.getTextByPathList(clrNode, ["attrs", "lastClr"]);
            if (sysClr !== undefined) {
                color = sysClr;
            }
        }

        var isAlpha = false;
        var alpha = parseInt(xmlUtils.getTextByPathList(clrNode, ["a:alpha", "attrs", "val"])) / 100000;
        if (!isNaN(alpha)) {
            var al_color = tinycolor(color);
            al_color.setAlpha(alpha);
            color = al_color.toHex8();
            isAlpha = true;
        }

        return color;
    }

    function getGradientFill(node, warpObj) {
        var gsLst = node["a:gsLst"]["a:gs"];
        var color_ary = [];
        for (var i = 0; i < gsLst.length; i++) {
            var lo_color = getSolidFill(gsLst[i], undefined, undefined, warpObj);
            color_ary[i] = lo_color;
        }
        var lin = node["a:lin"];
        var rot = 0;
        if (lin !== undefined) {
            rot = colorUtils.angleToDegrees(lin["attrs"]["ang"]) + 90;
        }
        return {
            "color": color_ary,
            "rot": rot
        };
    }

    function getPatternFill(node, warpObj) {
        var fgColor = "", bgColor = "", prst = "";
        var bgClr = node["a:bgClr"];
        var fgClr = node["a:fgClr"];
        prst = node["attrs"]["prst"];
        fgColor = getSolidFill(fgClr, undefined, undefined, warpObj);
        bgColor = getSolidFill(bgClr, undefined, undefined, warpObj);
        var linear_gradient = getLinerGrandient(prst, bgColor, fgColor);
        return linear_gradient;
    }

    function getLinerGrandient(prst, bgColor, fgColor) {
        return "repeating-linear-gradient(45deg, #" + bgColor + ", #" + fgColor + " 2px, #" + bgColor + " 4px);";
    }

    function getSchemeColorFromTheme(schemeClr, clrMap, phClr, warpObj) {
        return "000000";
    }

    function toHex(n) {
        var hex = n.toString(16);
        while (hex.length < 2) {
            hex = "0" + hex;
        }
        return hex;
    }

    function getColorName2Hex(colorName) {
        var colorMap = {
            "black": "000000",
            "white": "FFFFFF",
            "red": "FF0000",
            "green": "00FF00",
            "blue": "0000FF",
            "yellow": "FFFF00",
            "cyan": "00FFFF",
            "magenta": "FF00FF"
        };
        return colorMap[colorName.toLowerCase()] || "000000";
    }

    function hslToRgb(h, s, l) {
        var r, g, b;
        if (s == 0) {
            r = g = b = l;
        } else {
            var hue2rgb = function(p, q, t) {
                if (t < 0) t += 1;
                if (t > 1) t -= 1;
                if (t < 1/6) return p + (q - p) * 6 * t;
                if (t < 1/2) return q;
                if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
                return p;
            };
            var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            var p = 2 * l - q;
            r = hue2rgb(p, q, h + 1/3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1/3);
        }
        return {
            r: Math.round(r * 255),
            g: Math.round(g * 255),
            b: Math.round(b * 255)
        };
    }

    function escapeHtml(text) {
        if (!text) return text;
        return text
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function getPicFill(type, node, warpObj) {
        var img;
        var rId = node["a:blip"]["attrs"]["r:embed"];
        var imgPath;
        if (type == "slideBg" || type == "slide") {
            imgPath = xmlUtils.getTextByPathList(warpObj, ["slideResObj", rId, "target"]);
        } else if (type == "slideLayoutBg") {
            imgPath = xmlUtils.getTextByPathList(warpObj, ["layoutResObj", rId, "target"]);
        } else if (type == "slideMasterBg") {
            imgPath = xmlUtils.getTextByPathList(warpObj, ["masterResObj", rId, "target"]);
        } else if (type == "themeBg") {
            imgPath = xmlUtils.getTextByPathList(warpObj, ["themeResObj", rId, "target"]);
        } else if (type == "diagramBg") {
            imgPath = xmlUtils.getTextByPathList(warpObj, ["diagramResObj", rId, "target"]);
        }
        if (imgPath === undefined) {
            return undefined;
        }
        img = xmlUtils.getTextByPathList(warpObj, ["loaded-images", imgPath]);
        if (img === undefined) {
            imgPath = escapeHtml(imgPath);
            var imgExt = imgPath.split(".").pop();
            if (imgExt == "xml") {
                return undefined;
            }
            var imgFile = warpObj["zip"].file(imgPath);
            if (imgFile === null || imgFile === undefined) {
                console.warn("Image file not found:", imgPath);
                return undefined;
            }
            var imgArrayBuffer = imgFile.asArrayBuffer();
            var imgMimeType = getMimeType(imgExt);
            img = "data:" + imgMimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer);
            xmlUtils.setTextByPathList(warpObj, ["loaded-images", imgPath], img);
        }
        return img;
    }

    function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, tbl_col_width) {
        var text = "";
        var slideMasterTextStyles = warpObj["slideMasterTextStyles"];

        if (textBodyNode === undefined) {
            return text;
        }

        var pFontStyle = xmlUtils.getTextByPathList(spNode, ["p:style", "a:fontRef"]);

        var apNode = textBodyNode["a:p"];
        if (apNode.constructor !== Array) {
            apNode = [apNode];
        }

        for (var i = 0; i < apNode.length; i++) {
            var pNode = apNode[i];
            var rNode = pNode["a:r"];
            var fldNode = pNode["a:fld"];
            var brNode = pNode["a:br"];
            if (rNode !== undefined) {
                rNode = (rNode.constructor === Array) ? rNode : [rNode];
            }
            if (rNode !== undefined && fldNode !== undefined) {
                fldNode = (fldNode.constructor === Array) ? fldNode : [fldNode];
                rNode = rNode.concat(fldNode)
            }
            if (rNode !== undefined && brNode !== undefined) {
                is_first_br = true;
                brNode = (brNode.constructor === Array) ? brNode : [brNode];
                brNode.forEach(function (item, indx) {
                    item.type = "br";
                });
                if (brNode.length > 1) {
                    brNode.shift();
                }
                rNode = rNode.concat(brNode)
                rNode.sort(function (a, b) {
                    return a.attrs.order - b.attrs.order;
                });
            }

            var styleText = "";
            var marginsVer = getVerticalMargins(pNode, textBodyNode, type, idx, warpObj);
            if (marginsVer != "") {
                styleText = marginsVer;
            }
            if (type == "body" || type == "obj" || type == "shape") {
                styleText += "font-size: 0px;";
                styleText += "font-weight: 100;";
                styleText += "font-style: normal;";
            }
            var cssName = "";

            if (styleText in styleTable) {
                cssName = styleTable[styleText]["name"];
            } else {
                cssName = "_css_" + (Object.keys(styleTable).length + 1);
                styleTable[styleText] = {
                    "name": cssName,
                    "text": styleText
                };
            }

            var prg_width_node = xmlUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
            var prg_height_node;
            var sld_prg_width = ((prg_width_node !== undefined) ? ("width:" + (parseInt(prg_width_node) * slideFactor) + "px;") : "width:inherit;");
            var sld_prg_height = ((prg_height_node !== undefined) ? ("height:" + (parseInt(prg_height_node) * slideFactor) + "px;") : "");
            var prg_dir = getPregraphDir(pNode, textBodyNode, idx, type, warpObj);
            text += "<div style='display: flex;" + sld_prg_width + sld_prg_height + "' class='slide-prgrph " + getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir, warpObj) + " " +
                prg_dir + " " + cssName + "' >";
            var buText_ary = genBuChar(pNode, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj);
            var isBullate = (buText_ary[0] !== undefined && buText_ary[0] !== null && buText_ary[0] != "" ) ? true : false;
            var bu_width = (buText_ary[1] !== undefined && buText_ary[1] !== null && isBullate) ? buText_ary[1] + buText_ary[2] : 0;
            text += (buText_ary[0] !== undefined) ? buText_ary[0]:"";

            var margin_ary = getPregraphMargn(pNode, idx, type, isBullate, warpObj);
            var margin = margin_ary[0];
            var mrgin_val = margin_ary[1];
            if (prg_width_node === undefined && tbl_col_width !== undefined && prg_width_node != 0){
                prg_width_node = tbl_col_width;
            }

            var prgrph_text = "";
            var total_text_len = 0;
            if (rNode === undefined && pNode !== undefined) {
                var prgr_text = genSpanElement(pNode, undefined, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, 1, warpObj, isBullate);
                if (isBullate) {
                    var txt_obj = $(prgr_text)
                        .css({ 'position': 'absolute', 'float': 'left', 'white-space': 'nowrap', 'visibility': 'hidden' })
                        .appendTo($('body'));
                    total_text_len += txt_obj.outerWidth();
                    txt_obj.remove();
                }
                prgrph_text += prgr_text;
            } else if (rNode !== undefined) {
                for (var j = 0; j < rNode.length; j++) {
                    var prgr_text = genSpanElement(rNode[j], j, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNode.length, warpObj, isBullate);
                    if (isBullate) {
                        var txt_obj = $(prgr_text)
                            .css({ 'position': 'absolute', 'float': 'left', 'white-space': 'nowrap', 'visibility': 'hidden'})
                            .appendTo($('body'));
                        total_text_len += txt_obj.outerWidth();
                        txt_obj.remove();
                    }
                    prgrph_text += prgr_text;
                }
            }

            prg_width_node = parseInt(prg_width_node) * slideFactor - bu_width - mrgin_val;
            if (isBullate) {
                if (total_text_len < prg_width_node ){
                    prg_width_node = total_text_len + bu_width;
                }
            }
            var prg_width = ((prg_width_node !== undefined) ? ("width:" + (prg_width_node )) + "px;" : "width:inherit;");
            text += "<div style='height: 100%;direction: initial;overflow-wrap:break-word;word-wrap: break-word;" + prg_width + margin + "' >";
            text += prgrph_text;
            text += "</div>";
            text += "</div>";
        }

        return text;
    }

    function getVerticalMargins(pNode, textBodyNode, type, idx, warpObj) {
        var lvl = 1
        var spcBefNode = xmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPts", "attrs", "val"]);
        var spcAftNode = xmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPts", "attrs", "val"]);
        var lnSpcNode = xmlUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPct", "attrs", "val"]);
        var lnSpcNodeType = "Pct";
        if (lnSpcNode === undefined) {
            lnSpcNode = xmlUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
            if (lnSpcNode !== undefined) {
                lnSpcNodeType = "Pts";
            }
        }
        var lvlNode = xmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvl = parseInt(lvlNode) + 1;
        }
        var fontSize;
        if (xmlUtils.getTextByPathList(pNode, ["a:r"]) !== undefined) {
            var fontSizeStr = getFontSize(pNode["a:r"], textBodyNode,undefined, lvl, type, warpObj);
            if (fontSizeStr != "inherit") {
                fontSize = parseInt(fontSizeStr, "px");
            }
        }

        var isInLayoutOrMaster = true;
        if(type == "shape" || type == "textBox"){
            isInLayoutOrMaster = false;
        }
        if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
            if (idx !== undefined) {
                var laypPrNode = xmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "idxTable", idx, "p:txBody", "a:p", (lvl - 1), "a:pPr"]);

                if (spcBefNode === undefined) {
                    spcBefNode = xmlUtils.getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                }

                if (spcAftNode === undefined) {
                    spcAftNode = xmlUtils.getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                }

                if (lnSpcNode === undefined) {
                    lnSpcNode = xmlUtils.getTextByPathList(laypPrNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                    if (lnSpcNode === undefined) {
                        lnSpcNode = xmlUtils.getTextByPathList(laypPrNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                        if (lnSpcNode !== undefined) {
                            lnSpcNodeType = "Pts";
                        }
                    }
                }
            }
        }
        if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
            var slideMasterTextStyles = warpObj["slideMasterTextStyles"];
            var dirLoc = "";
            var lvlStr = "a:lvl" + lvl + "pPr";
            switch (type) {
                case "title":
                case "ctrTitle":
                    dirLoc = "p:titleStyle";
                    break;
                case "body":
                case "obj":
                case "dt":
                case "ftr":
                case "sldNum":
                case "textBox":
                    dirLoc = "p:bodyStyle";
                    break;
                case "shape":
                default:
                    dirLoc = "p:otherStyle";
            }
            var inLvlNode = xmlUtils.getTextByPathList(slideMasterTextStyles, [dirLoc, lvlStr]);
            if (inLvlNode !== undefined) {
                if (spcBefNode === undefined) {
                    spcBefNode = xmlUtils.getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                }

                if (spcAftNode === undefined) {
                    spcAftNode = xmlUtils.getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                }

                if (lnSpcNode === undefined) {
                    lnSpcNode = xmlUtils.getTextByPathList(inLvlNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                    if (lnSpcNode === undefined) {
                        lnSpcNode = xmlUtils.getTextByPathList(inLvlNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                        if (lnSpcNode !== undefined) {
                            lnSpcNodeType = "Pts";
                        }
                    }
                }
            }
        }
        var spcBefor = 0, spcAfter = 0, spcLines = 0;
        var marginTopBottomStr = "";
        if (spcBefNode !== undefined) {
            spcBefor = parseInt(spcBefNode) / 100;
        }
        if (spcAftNode !== undefined) {
            spcAfter = parseInt(spcAftNode) / 100;
        }
        
        if (lnSpcNode !== undefined && fontSize !== undefined) {
            if (lnSpcNodeType == "Pts") {
                marginTopBottomStr += "padding-top: " + ((parseInt(lnSpcNode) / 100) - fontSize) + "px;";
            } else {
                var fct = parseInt(lnSpcNode) / 100000;
                spcLines = fontSize * (fct - 1) - fontSize;
                var pTop = (fct > 1) ? spcLines : 0;
                var pBottom = (fct > 1) ? fontSize : 0;
                marginTopBottomStr += "padding-top: " + pBottom + "px;";
                marginTopBottomStr += "padding-bottom: " + spcLines + "px;";
            }
        }

        marginTopBottomStr += "margin-top: " + (spcBefor - 1) + "px;";
        if (spcAftNode !== undefined || lnSpcNode !== undefined) {
            marginTopBottomStr += "margin-bottom: " + spcAfter  + "px;";
        }

        return marginTopBottomStr;
    }

    function getHorizontalAlign(node, textBodyNode, idx, type, prg_dir, warpObj) {
        var algn = xmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
        if (algn === undefined) {
            var lvlIdx = 1;
            var lvlNode = xmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvlIdx = parseInt(lvlNode) + 1;
            }
            var lvlStr = "a:lvl" + lvlIdx + "pPr";

            var lstStyle = textBodyNode["a:lstStyle"];
            algn = xmlUtils.getTextByPathList(lstStyle, [lvlStr, "attrs", "algn"]);

            if (algn === undefined && idx !== undefined ) {
                algn = xmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                if (algn === undefined) {
                    algn = xmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
                    if (algn === undefined) {
                        algn = xmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvlIdx - 1), "a:pPr", "attrs", "algn"]);
                    }
                }
            }
            if (algn === undefined) {
                if (type !== undefined) {
                    algn = xmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);

                    if (algn === undefined) {
                        if (type == "title" || type == "ctrTitle") {
                            algn = xmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                        } else if (type == "body" || type == "obj" || type == "subTitle") {
                            algn = xmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                        } else if (type == "shape" || type == "diagram") {
                            algn = xmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr, "attrs", "algn"]);
                        } else if (type == "textBox") {
                            algn = xmlUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr, "attrs", "algn"]);
                        } else {
                            algn = xmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                        }
                    }
                } else {
                    algn = xmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                }
            }
        }

        if (algn === undefined) {
            if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                return "h-mid";
            } else if (type == "sldNum") {
                return "h-right";
            }
        }
        if (algn !== undefined) {
            switch (algn) {
                case "l":
                    if (prg_dir == "pregraph-rtl"){
                        return "h-left-rtl";
                    }else{
                        return "h-left";
                    }
                    break;
                case "r":
                    if (prg_dir == "pregraph-rtl") {
                        return "h-right-rtl";
                    }else{
                        return "h-right";
                    }
                    break;
                case "ctr":
                    return "h-mid";
                    break;
                case "just":
                case "dist":
                default:
                    return "h-" + algn;
            }
        }
    }

    function getPregraphDir(node, textBodyNode, idx, type, warpObj) {
        var rtl = xmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "rtl"]);

        if (rtl === undefined) {
            var layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
            var pPrNodeLaout = layoutMasterNode.nodeLaout;
            var pPrNodeMaster = layoutMasterNode.nodeMaster;
            rtl = xmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
            if (rtl === undefined && type != "shape") {
                rtl = xmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
            }
        }

        if (rtl == "1") {
            return "pregraph-rtl";
        } else if (rtl == "0") {
            return "pregraph-ltr";
        }
        return "pregraph-inherit";
    }

    function getLayoutAndMasterNode(node, idx, type, warpObj) {
        var pPrNodeLaout, pPrNodeMaster;
        var pPrNode = node["a:pPr"];
        var lvl = 1;
        var lvlNode = xmlUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvl = parseInt(lvlNode) + 1;
        }
        if (idx !== undefined) {
            pPrNodeLaout = xmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", "a:lvl" + lvl + "pPr"]);
            if (pPrNodeLaout === undefined) {
                pPrNodeLaout = xmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr"]);
                if (pPrNodeLaout === undefined) {
                    pPrNodeLaout = xmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvl - 1), "a:pPr"]);
                }
            }
        }
        if (type !== undefined) {
            var lvlStr = "a:lvl" + lvl + "pPr";
            if (pPrNodeLaout === undefined) {
                pPrNodeLaout = xmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
            }
            if (type == "title" || type == "ctrTitle") {
                pPrNodeMaster = xmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr]);
            } else if (type == "body" || type == "obj" || type == "subTitle") {
                pPrNodeMaster = xmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr]);
            } else if (type == "shape" || type == "diagram") {
                pPrNodeMaster = xmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr]);
            } else if (type == "textBox") {
                pPrNodeMaster = xmlUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr]);
            } else {
                pPrNodeMaster = xmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
            }
        }
        return {
            "nodeLaout": pPrNodeLaout,
            "nodeMaster": pPrNodeMaster
        };
    }

    function genBuChar(node, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj) {
        return "";
    }

    function getPregraphMargn(pNode, idx, type, isBullate, warpObj) {
        if (!isBullate){
            return ["",0];
        }
        return ["",0];
    }

    function genSpanElement(node, rIndex, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNodeLength, warpObj, isBullate) {
        var text = xmlUtils.getTextByPathList(node, ["a:t"]);
        var styleText = "";
        var lstStyle = textBodyNode["a:lstStyle"];

        if (text === undefined && node["type"] !== undefined) {
            if (is_first_br) {
                is_first_br = false;
                return "<span class='line-break-br' ></span>";
            }
            styleText += "display: block;";
        } else {
            is_first_br = true;
        }

        if (typeof text !== 'string') {
            text = xmlUtils.getTextByPathList(node, ["a:fld", "a:t"]);
            if (typeof text !== 'string') {
                text = "&nbsp;";
            }
        }

        var pPrNode = pNode["a:pPr"];
        var lvl = 1;
        var lvlNode = xmlUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvl = parseInt(lvlNode) + 1;
        }

        var layoutMasterNode = getLayoutAndMasterNode(pNode, idx, type, warpObj);
        var pPrNodeLaout = layoutMasterNode.nodeLaout;
        var pPrNodeMaster = layoutMasterNode.nodeMaster;

        var linkID = xmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
        var linkTooltip = "";
        if (linkID !== undefined) {
            linkTooltip = xmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]);
            if (linkTooltip !== undefined) {
                linkTooltip = "title='" + linkTooltip + "'";
            }
        }

        var fontClrPr = getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
        var fontClrType = fontClrPr[2];

        if (fontClrType == "solid") {
            if (linkID === undefined && fontClrPr[0] !== undefined && fontClrPr[0] != "") {
                styleText += "color: #" + fontClrPr[0] + ";";
            }
            if (fontClrPr[1] !== undefined && fontClrPr[1] != "" && fontClrPr[1] != ";") {
                styleText += "text-shadow:" + fontClrPr[1] + ";";
            }
            if (fontClrPr[3] !== undefined && fontClrPr[3] != "") {
                styleText += "background-color: #" + fontClrPr[3] + ";";
            }
        }

        var font_size = getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj);
        styleText += "font-size:" + font_size + ";";

        var font_type = getFontType(node, type, warpObj, pFontStyle);
        styleText += "font-family:" + font_type + ";";

        var font_bold = getFontBold(node, type, warpObj["slideMasterTextStyles"]);
        styleText += "font-weight:" + font_bold + ";";

        var font_italic = getFontItalic(node, type, warpObj["slideMasterTextStyles"]);
        styleText += "font-style:" + font_italic + ";";

        var font_decoration = getFontDecoration(node, type, warpObj["slideMasterTextStyles"]);
        styleText += "text-decoration:" + font_decoration + ";";

        var highlight = xmlUtils.getTextByPathList(node, ["a:rPr", "a:highlight"]);
        if (highlight !== undefined) {
            styleText += "background-color:#" + getSolidFill(highlight, undefined, undefined, warpObj) + ";";
        }

        var cssName = "";
        if (styleText in styleTable) {
            cssName = styleTable[styleText]["name"];
        } else {
            cssName = "_css_" + (Object.keys(styleTable).length + 1);
            styleTable[styleText] = {
                "name": cssName,
                "text": styleText
            };
        }

        if (linkID !== undefined && linkID != "") {
            var linkURL = warpObj["slideResObj"][linkID]["target"];
            linkURL = escapeHtml(linkURL);
            return "<span class='text-block " + cssName + "'><a href='" + linkURL + "' " + linkTooltip + " target='_blank'>" +
                    text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + "</a></span>";
        } else {
            return "<span class='text-block " + cssName + "'>" + text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + "</span>";
        }
    }

    function getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj) {
        var rPrNode = xmlUtils.getTextByPathList(node, ["a:rPr"]);
        var filTyp, color, colorType = "", highlightColor = "";

        if (rPrNode !== undefined) {
            filTyp = colorUtils.getFillType(rPrNode);
            if (filTyp == "SOLID_FILL") {
                var solidFillNode = rPrNode["a:solidFill"];
                color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
                var highlightNode = rPrNode["a:highlight"];
                if (highlightNode !== undefined) {
                    highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                }
                colorType = "solid";
            } else if (filTyp == "PATTERN_FILL") {
                var pattFill = rPrNode["a:pattFill"];
                color = getPatternFill(pattFill, warpObj);
                colorType = "pattern";
            } else if (filTyp == "PIC_FILL") {
                color = getPicFill("slide", rPrNode["a:blipFill"], warpObj);
                colorType = "pic";
            } else if (filTyp == "GRADIENT_FILL") {
                var shpFill = rPrNode["a:gradFill"];
                color = getGradientFill(shpFill, warpObj);
                colorType = "gradient";
            }
        }

        if (color === undefined && xmlUtils.getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]) !== undefined) {
            var lstStyledefRPr = xmlUtils.getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]);
            filTyp = colorUtils.getFillType(lstStyledefRPr);
            if (filTyp == "SOLID_FILL") {
                var solidFillNode = lstStyledefRPr["a:solidFill"];
                color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
                var highlightNode = lstStyledefRPr["a:highlight"];
                if (highlightNode !== undefined) {
                    highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                }
                colorType = "solid";
            } else if (filTyp == "PATTERN_FILL") {
                var pattFill = lstStyledefRPr["a:pattFill"];
                color = getPatternFill(pattFill, warpObj);
                colorType = "pattern";
            } else if (filTyp == "PIC_FILL") {
                color = getPicFill("slide", lstStyledefRPr["a:blipFill"], warpObj);
                colorType = "pic";
            } else if (filTyp == "GRADIENT_FILL") {
                var shpFill = lstStyledefRPr["a:gradFill"];
                color = getGradientFill(shpFill, warpObj);
                colorType = "gradient";
            }
        }

        if (color === undefined) {
            var sPstyle = xmlUtils.getTextByPathList(pNode, ["p:style", "a:fontRef"]);
            if (sPstyle !== undefined) {
                color = getSolidFill(sPstyle, undefined, undefined, warpObj);
                if (color !== undefined) {
                    colorType = "solid";
                }
                var highlightNode = sPstyle["a:highlight"];
                if (highlightNode !== undefined) {
                    highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                }
            }
            if (color === undefined) {
                if (pFontStyle !== undefined) {
                    color = getSolidFill(pFontStyle, undefined, undefined, warpObj);
                    if (color !== undefined) {
                        colorType = "solid";
                    }
                }
            }
        }

        if (color === undefined) {
            var layoutMasterNode = getLayoutAndMasterNode(pNode, idx, type, warpObj);
            var pPrNodeLaout = layoutMasterNode.nodeLaout;
            var pPrNodeMaster = layoutMasterNode.nodeMaster;

            if (pPrNodeLaout !== undefined) {
                var defRpRLaout = xmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:solidFill"]);
                if (defRpRLaout !== undefined) {
                    color = getSolidFill(defRpRLaout, undefined, undefined, warpObj);
                    var highlightNode = xmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:highlight"]);
                    if (highlightNode !== undefined) {
                        highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                    colorType = "solid";
                }
            }
            if (color === undefined) {
                if (pPrNodeMaster !== undefined) {
                    var defRprMaster = xmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:solidFill"]);
                    if (defRprMaster !== undefined) {
                        color = getSolidFill(defRprMaster, undefined, undefined, warpObj);
                        var highlightNode = xmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:highlight"]);
                        if (highlightNode !== undefined) {
                            highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                        }
                        colorType = "solid";
                    }
                }
            }
        }

        return [color, "", colorType, highlightColor];
    }

    function getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj) {
        var lstStyle = (textBodyNode !== undefined)? textBodyNode["a:lstStyle"] : undefined;
        var lvlpPr = "a:lvl" + lvl + "pPr";
        var fontSize = undefined;
        var sz, kern;

        if (node["a:rPr"] !== undefined) {
            fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
        }
        if (isNaN(fontSize) || fontSize === undefined && node["a:fld"] !== undefined) {
            sz = xmlUtils.getTextByPathList(node["a:fld"], ["a:rPr", "attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
        }
        if ((isNaN(fontSize) || fontSize === undefined) && node["a:t"] === undefined) {
            sz = xmlUtils.getTextByPathList(node["a:endParaRPr"], [ "attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
        }
        if ((isNaN(fontSize) || fontSize === undefined) && lstStyle !== undefined) {
            sz = xmlUtils.getTextByPathList(lstStyle, [lvlpPr, "a:defRPr", "attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
        }

        var isAutoFit = false;
        var isKerning = false;
        if (textBodyNode !== undefined){
            var spAutoFitNode = xmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:spAutoFit"]);
            if (spAutoFitNode !== undefined){
                isAutoFit = true;
                isKerning = true;
            }
        }

        if (isNaN(fontSize) || fontSize === undefined) {
            sz = xmlUtils.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
            kern = xmlUtils.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            if (isKerning && kern !== undefined && !isNaN(fontSize) && (fontSize - parseInt(kern) / 100) > 0){
                fontSize = fontSize - parseInt(kern) / 100;
            }
        }

        if (isNaN(fontSize) || fontSize === undefined) {
            sz = xmlUtils.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
            kern = xmlUtils.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            if (sz === undefined) {
                if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                    sz = xmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = xmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                } else if (type == "body" || type == "obj" || type == "dt" || type == "sldNum" || type === "textBox") {
                    sz = xmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = xmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                }
                else if (type == "shape") {
                    sz = xmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = xmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    isKerning = false;
                }

                if (sz === undefined) {
                    sz = xmlUtils.getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = (kern === undefined)? xmlUtils.getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "kern"]) : undefined;
                    isKerning = false;
                }
            } 
            fontSize = parseInt(sz) / 100;
            if (isKerning && kern !== undefined && !isNaN(fontSize) && ((fontSize - parseInt(kern) / 100) > parseInt(kern) / 100 )) {
                fontSize = fontSize - parseInt(kern) / 100;
            }
        }

        var baseline = xmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
        if (baseline !== undefined && !isNaN(fontSize)) {
            var baselineVl = parseInt(baseline) / 100000;
            fontSize -= baselineVl;
        }

        if (!isNaN(fontSize)){
            var normAutofit = xmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit", "attrs", "fontScale"]);
            if (normAutofit !== undefined && normAutofit != 0){
                fontSize = Math.round(fontSize * (normAutofit / 100000))
            }
        }

        return isNaN(fontSize) ? ((type == "br") ? "initial" : "inherit") : (fontSize * fontSizeFactor + "px");
    }

    function getFontType(node, type, warpObj, pFontStyle) {
        var typeface = xmlUtils.getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);

        if (typeface === undefined) {
            var fontIdx = "";
            var fontGrup = "";
            if (pFontStyle !== undefined) {
                fontIdx = xmlUtils.getTextByPathList(pFontStyle, ["attrs", "idx"]);
            }
            var fontSchemeNode = xmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:fontScheme"]);
            if (fontIdx == "") {
                if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                    fontIdx = "major";
                } else {
                    fontIdx = "minor";
                }
            }
            fontGrup = "a:" + fontIdx + "Font";
            typeface = xmlUtils.getTextByPathList(fontSchemeNode, [fontGrup, "a:latin", "attrs", "typeface"]);
        }

        return (typeface === undefined) ? "inherit" : typeface;
    }

    function getFontBold(node, type, slideMasterTextStyles) {
        return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["b"] === "1") ? "bold" : "inherit";
    }

    function getFontItalic(node, type, slideMasterTextStyles) {
        return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["i"] === "1") ? "italic" : "inherit";
    }

    function getFontDecoration(node, type, slideMasterTextStyles) {
        if (node["a:rPr"] !== undefined) {
            var underLine = node["a:rPr"]["attrs"]["u"] !== undefined ? node["a:rPr"]["attrs"]["u"] : "none";
            var strikethrough = node["a:rPr"]["attrs"]["strike"] !== undefined ? node["a:rPr"]["attrs"]["strike"] : 'noStrike';

            if (underLine != "none" && strikethrough == "noStrike") {
                return "underline";
            } else if (underLine == "none" && strikethrough != "noStrike") {
                return "line-through";
            } else if (underLine != "none" && strikethrough != "noStrike") {
                return "underline line-through";
            } else {
                return "inherit";
            }
        } else {
            return "inherit";
        }
    }

    function getShapeFill(node, pNode, isSvgMode, warpObj, source) {
        var fillType = colorUtils.getFillType(xmlUtils.getTextByPathList(node, ["p:spPr"]));
        var fillColor;
        if (fillType == "NO_FILL") {
            return isSvgMode ? "none" : "";
        } else if (fillType == "SOLID_FILL") {
            var shpFill = node["p:spPr"]["a:solidFill"];
            fillColor = getSolidFill(shpFill, undefined, undefined, warpObj);
        } else if (fillType == "GRADIENT_FILL") {
            var shpFill = node["p:spPr"]["a:gradFill"];
            fillColor = getGradientFill(shpFill, warpObj);
        } else if (fillType == "PATTERN_FILL") {
            var shpFill = node["p:spPr"]["a:pattFill"];
            fillColor = getPatternFill(shpFill, warpObj);
        } else if (fillType == "PIC_FILL") {
            var shpFill = node["p:spPr"]["a:blipFill"];
            fillColor = getPicFill(source, shpFill, warpObj);
        }

        if (fillColor === undefined) {
            var clrName = xmlUtils.getTextByPathList(node, ["p:style", "a:fillRef"]);
            var idx = parseInt(xmlUtils.getTextByPathList(node, ["p:style", "a:fillRef", "attrs", "idx"]));
            if (idx == 0 || idx == 1000) {
                return isSvgMode ? "none" : "";
            }
            fillColor = getSolidFill(clrName, undefined, undefined, warpObj);
        }

        if (fillColor === undefined) {
            var grpFill = xmlUtils.getTextByPathList(node, ["p:spPr", "a:grpFill"]);
            if (grpFill !== undefined) {
                var grpShpFill = pNode["p:grpSpPr"];
                var spShpNode = { "p:spPr": grpShpFill };
                return getShapeFill(spShpNode, node, isSvgMode, warpObj, source);
            } else if (fillType == "NO_FILL") {
                return isSvgMode ? "none" : "";
            }
        }

        if (fillColor !== undefined) {
            if (fillType == "GRADIENT_FILL") {
                if (isSvgMode) {
                    return fillColor;
                } else {
                    var colorAry = fillColor.color;
                    var rot = fillColor.rot;
                    var bgcolor = "background: linear-gradient(" + rot + "deg,";
                    for (var i = 0; i < colorAry.length; i++) {
                        if (i == colorAry.length - 1) {
                            bgcolor += "#" + colorAry[i] + ");";
                        } else {
                            bgcolor += "#" + colorAry[i] + ", ";
                        }
                    }
                    return bgcolor;
                }
            } else if (fillType == "PIC_FILL") {
                if (isSvgMode) {
                    return fillColor;
                } else {
                    return "background-image:url(" + fillColor + ");";
                }
            } else if (fillType == "PATTERN_FILL") {
                var bgPtrn = "", bgSize = "", bgPos = "";
                bgPtrn = fillColor[0];
                if (fillColor[1] !== null && fillColor[1] !== undefined && fillColor[1] != "") {
                    bgSize = " background-size:" + fillColor[1] + ";";
                }
                if (fillColor[2] !== null && fillColor[2] !== undefined && fillColor[2] != "") {
                    bgPos = " background-position:" + fillColor[2] + ";";
                }
                return "background: " + bgPtrn + ";" + bgSize + bgPos;
            } else {
                if (isSvgMode) {
                    var color = tinycolor(fillColor);
                    fillColor = color.toRgbString();
                    return fillColor;
                } else {
                    return "background-color: #" + fillColor + ";";
                }
            }
        } else {
            if (isSvgMode) {
                return "none";
            } else {
                return "background-color: inherit;";
            }
        }
    }

    function getTableCellParams(tcNodes, getColsGrid, row_idx, col_idx, thisTblStyle, cellSource, warpObj) {
        var rowSpan = xmlUtils.getTextByPathList(tcNodes, ["attrs", "rowSpan"]);
        var colSpan = xmlUtils.getTextByPathList(tcNodes, ["attrs", "gridSpan"]);
        var vMerge = xmlUtils.getTextByPathList(tcNodes, ["attrs", "vMerge"]);
        var hMerge = xmlUtils.getTextByPathList(tcNodes, ["attrs", "hMerge"]);
        var colStyl = "word-wrap: break-word;";
        var colWidth;
        var celFillColor = "";
        var col_borders = "";
        var colFontClrPr = "";
        var colFontWeight = "";
        var lin_bottm = "", lin_top = "", lin_left = "", lin_right = "";
        
        var colSapnInt = parseInt(colSpan);
        var total_col_width = 0;
        if (!isNaN(colSapnInt) && colSapnInt > 1) {
            for (var k = 0; k < colSapnInt; k++) {
                total_col_width += parseInt(xmlUtils.getTextByPathList(getColsGrid[col_idx + k], ["attrs", "w"]));
            }
        } else {
            total_col_width = xmlUtils.getTextByPathList((col_idx === undefined) ? getColsGrid : getColsGrid[col_idx], ["attrs", "w"]);
        }

        var text = genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj, total_col_width);

        if (total_col_width != 0) {
            colWidth = parseInt(total_col_width) * slideFactor;
            colStyl += "width:" + colWidth + "px;";
        }

        lin_bottm = xmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnB"]);
        if (lin_bottm === undefined && cellSource !== undefined) {
            if (cellSource !== undefined)
                lin_bottm = xmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
            if (lin_bottm === undefined) {
                lin_bottm = xmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
            }
        }
        lin_top = xmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnT"]);
        if (lin_top === undefined) {
            if (cellSource !== undefined)
                lin_top = xmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
            if (lin_top === undefined) {
                lin_top = xmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
            }
        }
        lin_left = xmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnL"]);
        if (lin_left === undefined) {
            if (cellSource !== undefined)
                lin_left = xmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
            if (lin_left === undefined) {
                lin_left = xmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
            }
        }
        lin_right = xmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnR"]);
        if (lin_right === undefined) {
            if (cellSource !== undefined)
                lin_right = xmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
            if (lin_right === undefined) {
                lin_right = xmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
            }
        }

        if (lin_bottm !== undefined && lin_bottm != "") {
            var bottom_line_border = getBorder(lin_bottm, undefined, false, "", warpObj);
            if (bottom_line_border != "") {
                colStyl += "border-bottom:" + bottom_line_border + ";";
            }
        }
        if (lin_top !== undefined && lin_top != "") {
            var top_line_border = getBorder(lin_top, undefined, false, "", warpObj);
            if (top_line_border != "") {
                colStyl += "border-top: " + top_line_border + ";";
            }
        }
        if (lin_left !== undefined && lin_left != "") {
            var left_line_border = getBorder(lin_left, undefined, false, "", warpObj);
            if (left_line_border != "") {
                colStyl += "border-left: " + left_line_border + ";";
            }
        }
        if (lin_right !== undefined && lin_right != "") {
            var right_line_border = getBorder(lin_right, undefined, false, "", warpObj);
            if (right_line_border != "") {
                colStyl += "border-right:" + right_line_border + ";";
            }
        }

        var getCelFill = xmlUtils.getTextByPathList(tcNodes, ["a:tcPr"]);
        if (getCelFill !== undefined && getCelFill != "") {
            var cellObj = {
                "p:spPr": getCelFill
            };
            celFillColor = getShapeFill(cellObj, undefined, false, warpObj, "slide");
        }

        if (celFillColor == "" || celFillColor == "background-color: inherit;") {
            var bgFillschemeClr;
            if (cellSource !== undefined)
                bgFillschemeClr = xmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:fill", "a:solidFill"]);
            if (bgFillschemeClr !== undefined) {
                var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                if (local_fillColor !== undefined) {
                    celFillColor = " background-color: #" + local_fillColor + ";";
                }
            }
        }
        var cssName = "";
        if (celFillColor !== undefined && celFillColor != "") {
            if (celFillColor in styleTable) {
                cssName = styleTable[celFillColor]["name"];
            } else {
                cssName = "_tbl_cell_css_" + (Object.keys(styleTable).length + 1);
                styleTable[celFillColor] = {
                    "name": cssName,
                    "text": celFillColor
                };
            }
        }

        var rowTxtStyl;
        if (cellSource !== undefined) {
            rowTxtStyl = xmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcTxStyle"]);
        }
        if (rowTxtStyl !== undefined) {
            var local_fontClrPr = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
            if (local_fontClrPr !== undefined) {
                colFontClrPr = local_fontClrPr;
            }
            var local_fontWeight = ((xmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
            if (local_fontWeight !== "") {
                colFontWeight = local_fontWeight;
            }
        }
        colStyl += ((colFontClrPr !== "") ? "color: #" + colFontClrPr + ";" : "");
        colStyl += ((colFontWeight != "") ? " font-weight:" + colFontWeight + ";" : "");

        return [text, colStyl, cssName, rowSpan, colSpan];
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
        var xfrmNode = xmlUtils.getTextByPathList(node, ["p:xfrm"]);
        var dgmRelIds = xmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "dgm:relIds", "attrs"]);
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
        var dgmClr = xmlUtils.readXmlFile(zip, dgmClrFileName);
        var dgmData = xmlUtils.readXmlFile(zip, dgmDataFileName);
        var dgmLayout = xmlUtils.readXmlFile(zip, dgmLayoutFileName);
        var dgmQuickStyle = xmlUtils.readXmlFile(zip, dgmQuickStyleFileName);
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
        var dgmDrwSpArray = xmlUtils.getTextByPathList(warpObj["digramFileContent"], ["p:drawing", "p:spTree", "p:sp"]);
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

    /**
     * 获取位置信息
     * @param {Object} slideSpNode - 幻灯片形状节点
     * @param {Object} pNode - 父节点
     * @param {Object} slideLayoutSpNode - 幻灯片布局形状节点
     * @param {Object} slideMasterSpNode - 幻灯片母版形状节点
     * @param {string} sType - 类型
     * @returns {string} 位置样式字符串
     */
    function getPosition(slideSpNode, pNode, slideLayoutSpNode, slideMasterSpNode, sType) {
        var off;
        var x = -1, y = -1;

        if (slideSpNode !== undefined) {
            off = slideSpNode["a:off"]["attrs"];
        }

        if (off === undefined && slideLayoutSpNode !== undefined) {
            off = slideLayoutSpNode["a:off"]["attrs"];
        } else if (off === undefined && slideMasterSpNode !== undefined) {
            off = slideMasterSpNode["a:off"]["attrs"];
        }
        var offX = 0, offY = 0;
        var grpX = 0, grpY = 0;
        if (sType == "group") {

            var grpXfrmNode = xmlUtils.getTextByPathList(pNode, ["p:grpSpPr", "a:xfrm"]);
            if (grpXfrmNode !== undefined) {
                grpX = parseInt(grpXfrmNode["a:off"]["attrs"]["x"]) * slideFactor;
                grpY = parseInt(grpXfrmNode["a:off"]["attrs"]["y"]) * slideFactor;
            }
        }
        if (sType == "group-rotate" && pNode["p:grpSpPr"] !== undefined) {
            var xfrmNode = pNode["p:grpSpPr"]["a:xfrm"];
            var chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * slideFactor;
            var chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * slideFactor;

            offX = chx;
            offY = chy;
        }
        if (off === undefined) {
            return "";
        } else {
            x = parseInt(off["x"]) * slideFactor;
            y = parseInt(off["y"]) * slideFactor;
            return (isNaN(x) || isNaN(y)) ? "" : "top:" + (y - offY + grpY) + "px; left:" + (x - offX + grpX) + "px;";
        }

    }

    /**
     * 获取大小信息
     * @param {Object} slideSpNode - 幻灯片形状节点
     * @param {Object} slideLayoutSpNode - 幻灯片布局形状节点
     * @param {Object} slideMasterSpNode - 幻灯片母版形状节点
     * @returns {string} 大小样式字符串
     */
    function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
        var ext = undefined;
        var w = -1, h = -1;

        if (slideSpNode !== undefined) {
            ext = slideSpNode["a:ext"]["attrs"];
        } else if (slideLayoutSpNode !== undefined) {
            ext = slideLayoutSpNode["a:ext"]["attrs"];
        } else if (slideMasterSpNode !== undefined) {
            ext = slideMasterSpNode["a:ext"]["attrs"];
        }

        if (ext === undefined) {
            return "";
        } else {
            w = parseInt(ext["cx"]) * slideFactor;
            h = parseInt(ext["cy"]) * slideFactor;
            return (isNaN(w) || isNaN(h)) ? "" : "width:" + w + "px; height:" + h + "px;";
        }

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

window.NodeProcessors = NodeProcessors;