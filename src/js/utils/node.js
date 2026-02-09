/**
 * 节点工具函数模块
 * 提供PPTX节点处理和索引功能
 */

import { PPTXXmlUtils } from './xml.js';
import { PPTXStyleUtils } from './style.js';
import { PPTXTextUtils } from './text.js';
import { PPTXShapeUtils } from '../shape/shape.js';
import { SLIDE_FACTOR } from '../core/constants.js';

export const PPTXNodeUtils = (function() {

    /**
     * genDiagram - 生成 Diagram HTML
     * @param {Object} node - 节点
     * @param {Object} warpObj - 包装对象
     * @param {string} source - 源类型
     * @param {string} sType - 形状类型
     * @param {Object} settings - 设置对象
     * @returns {string} 生成的HTML
     */
    function genDiagram(node, warpObj, source, sType, settings) {
        var order = node["attrs"]["order"];
        var zip = warpObj["zip"];
        var xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);
        var dgmRelIds = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "dgm:relIds", "attrs"]);
        var dgmClrFileId = dgmRelIds["r:cs"];
        var dgmDataFileId = dgmRelIds["r:dm"];
        var dgmLayoutFileId = dgmRelIds["r:lo"];
        var dgmQuickStyleFileId = dgmRelIds["r:qs"];
        var dgmClrFileName = warpObj["slideResObj"][dgmClrFileId].target,
            dgmDataFileName = warpObj["slideResObj"][dgmDataFileId].target,
            dgmLayoutFileName = warpObj["slideResObj"][dgmLayoutFileId].target;
        const dgmQuickStyleFileName = warpObj["slideResObj"][dgmQuickStyleFileId].target;
        
        var dgmClr = PPTXXmlUtils.readXmlFile(zip, dgmClrFileName);
        var dgmData = PPTXXmlUtils.readXmlFile(zip, dgmDataFileName);
        var dgmLayout = PPTXXmlUtils.readXmlFile(zip, dgmLayoutFileName);
        var dgmQuickStyle = PPTXXmlUtils.readXmlFile(zip, dgmQuickStyleFileName);
        
        var dgmDrwSpArray = PPTXXmlUtils.getTextByPathList(warpObj["digramFileContent"], ["p:drawing", "p:spTree", "p:sp"]);
        var rslt = "";
        if (dgmDrwSpArray !== undefined) {
            var dgmDrwSpArrayLen = dgmDrwSpArray.length;
            for (var i = 0; i < dgmDrwSpArrayLen; i++) {
                var dspSp = dgmDrwSpArray[i];
                rslt += processSpNode(dspSp, node, warpObj, "diagramBg", sType);
            }
        }
        
        return "<div class='block diagram-content' style='" +
            PPTXXmlUtils.getPosition(xfrmNode, node, undefined, undefined, sType) +
            PPTXXmlUtils.getSize(xfrmNode, undefined, undefined) +
            "'>" + rslt + "</div>";
    }

    /**
     * indexNodes - 索引幻灯片节点
     * @param {Object} content - 幻灯片内容
     * @returns {Object} 包含idTable、idxTable和typeTable的对象
     */
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
                    var id = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                    var idx = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                    var type = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

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
                var id = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                var idx = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                var type = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

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
     * processGroupSpNode - 处理组形状节点
     * @param {Object} node - 组形状节点
     * @param {Object} warpObj - 包装对象
     * @param {string} source - 源
     * @returns {string} 生成的HTML
     */
    function processGroupSpNode(node, warpObj, source, settings) {
        var xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:grpSpPr", "a:xfrm"]);
        if (xfrmNode !== undefined) {
            var x = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * SLIDE_FACTOR;
            var y = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * SLIDE_FACTOR;

            // 根据ECMA-376标准，a:chOff和a:chExt是可选元素
            // 当不存在时，应该使用父元素的对应值作为默认值
            var chx, chy, chcx, chcy;

            if (xfrmNode["a:chOff"] !== undefined && xfrmNode["a:chOff"]["attrs"] !== undefined) {
                chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * SLIDE_FACTOR;
                chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * SLIDE_FACTOR;
            } else {
                // 当a:chOff不存在时，使用a:off的值作为默认值
                chx = x;
                chy = y;
            }

            var cx = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * SLIDE_FACTOR;
            var cy = parseInt(xfrmNode["a:ext"]["attrs"]["cy"]) * SLIDE_FACTOR;

            if (xfrmNode["a:chExt"] !== undefined && xfrmNode["a:chExt"]["attrs"] !== undefined) {
                chcx = parseInt(xfrmNode["a:chExt"]["attrs"]["cx"]) * SLIDE_FACTOR;
                chcy = parseInt(xfrmNode["a:chExt"]["attrs"]["cy"]) * SLIDE_FACTOR;
            } else {
                // 当a:chExt不存在时，使用a:ext的值作为默认值
                chcx = cx;
                chcy = cy;
            }
            var rotate = parseInt(xfrmNode["attrs"]["rot"]);
            var rotStr = "";
            var top = y - chy,
                left = x - chx,
                width = cx - chcx,
                height = cy - chcy;

            var sType = "group";
            if (!isNaN(rotate)) {
                rotate = PPTXXmlUtils.angleToDegrees(rotate);
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

        // Procsee all child nodes
        for (var nodeKey in node) {
            if (node[nodeKey].constructor === Array) {
                for (var i = 0; i < node[nodeKey].length; i++) {
                    result += processNodesInSlide(nodeKey, node[nodeKey][i], node, warpObj, source, sType, settings);
                }
            } else {
                result += processNodesInSlide(nodeKey, node[nodeKey], node, warpObj, source, sType, settings);
            }
        }

        result += "</div>";

        return result;
    }

    /**
     * processNodesInSlide - 处理幻灯片中的节点
     * @param {string} nodeKey - 节点键
     * @param {Object} nodeValue - 节点值
     * @param {Object} nodes - 节点集合
     * @param {Object} warpObj - 包装对象
     * @param {string} source - 源
     * @param {string} sType - 形状类型
     * @returns {string} 生成的HTML
     */
    function processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType, settings) {
        var result = "";

        switch (nodeKey) {
            case "p:sp":    // Shape, Text
                result = processSpNode(nodeValue, nodes, warpObj, source, sType, settings);
                break;
            case "p:cxnSp":    // Shape, Text (with connection)
                result = processCxnSpNode(nodeValue, nodes, warpObj, source, sType, settings);
                break;
            case "p:pic":    // Picture
                result = processPicNode(nodeValue, warpObj, source, sType, settings);
                break;
            case "p:graphicFrame":    // Chart, Diagram, Table
                result = processGraphicFrameNode(nodeValue, warpObj, source, sType, settings);
                break;
            case "p:grpSp":
                result = processGroupSpNode(nodeValue, warpObj, source, settings);
                break;
            case "mc:AlternateContent": //Equations and formulas as Image
                var mcFallbackNode = PPTXXmlUtils.getTextByPathList(nodeValue, ["mc:Fallback"]);
                result = processGroupSpNode(mcFallbackNode, warpObj, source, settings);
                break;
            default:
                //console.log("nodeKey: ", nodeKey)
        }

        return result;
    }
   

        function processSpNode(node, pNode, warpObj, source, sType, settings) {

            /*
            *  958    <xsd:complexType name="CT_GvmlShape">
            *  959   <xsd:sequence>
            *  960     <xsd:element name="nvSpPr" type="CT_GvmlShapeNonVisual"     minOccurs="1" maxOccurs="1"/>
            *  961     <xsd:element name="spPr"   type="CT_ShapeProperties"        minOccurs="1" maxOccurs="1"/>
            *  962     <xsd:element name="txSp"   type="CT_GvmlTextShape"          minOccurs="0" maxOccurs="1"/>
            *  963     <xsd:element name="style"  type="CT_ShapeStyle"             minOccurs="0" maxOccurs="1"/>
            *  964     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
            *  965   </xsd:sequence>
            *  966 </xsd:complexType>
            */

            var id = PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "id"]);
            var name = PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "name"]);
            var idx = (PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]) === undefined) ? undefined : PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]);
            var type = (PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]) === undefined) ? undefined : PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
            var order = PPTXXmlUtils.getTextByPathList(node, ["attrs", "order"]);
            var isUserDrawnBg;
            if (source == "slideLayoutBg" || source == "slideMasterBg") {
                var userDrawn = PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "attrs", "userDrawn"]);
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
                const txBoxVal = PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvSpPr", "attrs", "txBox"]);
                if (txBoxVal == "1") {
                    type = "textBox";
                }
            }
            if (type === undefined) {
                type = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                if (type === undefined) {
                    //type = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    if (source == "diagramBg") {
                        type = "diagram";
                    } else {

                        type = "obj"; //default type
                    }
                }
            }
            //console.log("processSpNode type:", type, "idx:", idx);
            return PPTXShapeUtils.genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, isUserDrawnBg, sType, source, settings);
        }

        function processCxnSpNode(node, pNode, warpObj, source, sType, settings) {

            var id = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
            var name = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
            var idx = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
            var type = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
            //<p:cNvCxnSpPr>(<p:cNvCxnSpPr>, <a:endCxn>)
            var order = node["attrs"]["order"];

            return PPTXShapeUtils.genShape(node, pNode, undefined, undefined, id, name, idx, type, order, warpObj, undefined, sType, source, settings);
        }
    function processPicNode(node, warpObj, source, sType, settings) {
            //console.log("processPicNode node:", node, "source:", source, "sType:", sType, "warpObj;", warpObj);
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
                //imgName = warpObj["slideResObj"][rid]["target"];
                resObj = warpObj["slideResObj"];
            }
            var imgName = (resObj[rid] !== undefined) ? resObj[rid]["target"] : undefined;

            if (imgName === undefined) {
                console.warn("Image reference not found in resObj for rid:", rid);
                return "";
            }

            //console.log("processPicNode imgName:", imgName);
            var imgFileExt = PPTXXmlUtils.extractFileExtension(imgName).toLowerCase();
            var zip = warpObj["zip"];
            
            // 确定上下文类型用于路径解析
            var context = 'slide';
            if (source == "slideMasterBg") {
                context = 'master';
            } else if (source == "slideLayoutBg") {
                context = 'layout';
            }
            
            // 使用改进的媒体文件查找方法
            var imgFile = PPTXXmlUtils.findMediaFile(zip, imgName, context, '');
            if (imgFile === null) {
                console.warn("Image file not found in processPicNode:", imgName);
                return "";
            }
            var imgArrayBuffer = imgFile.asArrayBuffer();
            var mimeType = "";
            var xfrmNode = node["p:spPr"]["a:xfrm"];
            if (xfrmNode === undefined) {
                var idx = PPTXXmlUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "idx"]);
                var type = PPTXXmlUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "type"]);
                if (idx !== undefined) {
                    xfrmNode = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"], ["idxTable", idx, "p:spPr", "a:xfrm"]);
                }
            }
            ///////////////////////////////////////Amir//////////////////////////////
            var rotate = 0;
            var rotateNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:xfrm", "attrs", "rot"]);
            if (rotateNode !== undefined) {
                rotate = PPTXXmlUtils.angleToDegrees(rotateNode);
            }
            //video
            var vdoNode = PPTXXmlUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:videoFile"]);
            var vdoRid, vdoFile, vdoFileExt, vdoMimeType, uInt8Array, blob, vdoBlob, mediaSupportFlag = false, isVdeoLink = false;
            var mediaProcess = settings.mediaProcess;
            if (vdoNode !== undefined & mediaProcess) {
                vdoRid = vdoNode["attrs"]["r:link"];
                vdoFile = resObj[vdoRid]["target"];
                var checkIfLink = PPTXXmlUtils.IsVideoLink(vdoFile);
                if (checkIfLink) {
                    vdoFile = PPTXXmlUtils.escapeHtml(vdoFile);
                    //vdoBlob = vdoFile;
                    isVdeoLink = true;
                    mediaSupportFlag = true;
                    mediaPicFlag = true;
                } else {
                    vdoFileExt = PPTXXmlUtils.extractFileExtension(vdoFile).toLowerCase();
                    if (vdoFileExt == "mp4" || vdoFileExt == "webm" || vdoFileExt == "ogg") {
                        // 使用改进的媒体文件查找方法
                        var vdoFileObj = PPTXXmlUtils.findMediaFile(zip, vdoFile, context, '');
                        if (vdoFileObj === null) {
                            console.warn("Video file not found:", vdoFile);
                        } else {
                            uInt8Array = vdoFileObj.asArrayBuffer();
                            vdoMimeType = PPTXXmlUtils.getMimeType(vdoFileExt);
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
            //Audio
            var audioNode = PPTXXmlUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:audioFile"]);
            var audioRid, audioFile, audioFileExt, audioMimeType, uInt8ArrayAudio, blobAudio, audioBlob;
            var audioPlayerFlag = false;
            var audioObjc;
            if (audioNode !== undefined & mediaProcess) {
                audioRid = audioNode["attrs"]["r:link"];
                audioFile = resObj[audioRid]["target"];
                audioFileExt = PPTXXmlUtils.extractFileExtension(audioFile).toLowerCase();
                if (audioFileExt == "mp3" || audioFileExt == "wav" || audioFileExt == "ogg") {
                    // 使用改进的媒体文件查找方法
                    var audioFileObj = PPTXXmlUtils.findMediaFile(zip, audioFile, context, '');
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
            //console.log(node)
            //////////////////////////////////////////////////////////////////////////
            mimeType = PPTXXmlUtils.getMimeType(imgFileExt);
            rtrnData = "<div class='block content' style='" +
                ((mediaProcess && audioPlayerFlag) ? PPTXXmlUtils.getPosition(audioObjc, node, undefined, undefined) : PPTXXmlUtils.getPosition(xfrmNode, node, undefined, undefined)) +
                ((mediaProcess && audioPlayerFlag) ? PPTXXmlUtils.getSize(audioObjc, undefined, undefined) : PPTXXmlUtils.getSize(xfrmNode, undefined, undefined)) +
                " z-index: " + order + ";" +
                "transform: rotate(" + rotate + "deg);'>";
            if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
                rtrnData += "<img src='data:" + mimeType + ";base64," + PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%; height: 100%'/>";
            } else if ((vdoNode !== undefined || audioNode !== undefined) && mediaProcess && mediaSupportFlag) {
                if (vdoNode !== undefined && !isVdeoLink) {
                    rtrnData += "<video  src='" + vdoBlob + "' controls style='width: 100%; height: 100%'>Your browser does not support the video tag.</video>";
                } else if (vdoNode !== undefined && isVdeoLink) {
                    rtrnData += "<iframe   src='" + vdoFile + "' controls style='width: 100%; height: 100%'></iframe >";
                }
                if (audioNode !== undefined) {
                    rtrnData += '<audio id="audio_player" controls ><source src="' + audioBlob + '"></audio>';
                    //'<button onclick="audio_player.play()">Play</button>'+
                    //'<button onclick="audio_player.pause()">Pause</button>';
                }
            }
            if (!mediaSupportFlag && mediaPicFlag) {
                rtrnData += "<span style='color:red;font-size:40px;position: absolute;'>This media file Not supported by HTML5</span>";
            }
            if ((vdoNode !== undefined || audioNode !== undefined) && !mediaProcess && mediaSupportFlag) {
                console.log("Founded supported media file but media process disabled (mediaProcess=false)");
            }
            rtrnData += "</div>";
            //console.log(rtrnData)
            return rtrnData;
        }

        function processGraphicFrameNode(node, warpObj, source, sType, settings) {

            var result = "";
            var graphicTypeUri = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "attrs", "uri"]);

            switch (graphicTypeUri) {
                case "http://schemas.openxmlformats.org/drawingml/2006/table":
                    result = PPTXTextUtils.genTable(node, warpObj);
                    break;
                case "http://schemas.openxmlformats.org/drawingml/2006/chart":
                    result = PPTXTextUtils.genChart(node, warpObj);
                    break;
                case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
                    result = genDiagram(node, warpObj, source, sType, settings);
                    break;
                case "http://schemas.openxmlformats.org/presentationml/2006/ole":
                    //result = genDiagram(node, warpObj, source, sType);
                    var oleObjNode = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent", "mc:Fallback","p:oleObj"]);
                    
                    if (oleObjNode === undefined) {
                        oleObjNode = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "p:oleObj"]);
                    }
                    //console.log("node:", node, "oleObjNode:", oleObjNode)
                    if (oleObjNode !== undefined){
                        result = PPTXNodeUtils.processGroupSpNode(oleObjNode, warpObj, source, sType, settings);
                    }
                    break;
                default:
            }

            return result;
        }

        function processSpPrNode(node, warpObj) {

            /*
            * 2241 <xsd:complexType name="CT_ShapeProperties">
            * 2242   <xsd:sequence>
            * 2243     <xsd:element name="xfrm" type="CT_Transform2D"  minOccurs="0" maxOccurs="1"/>
            * 2244     <xsd:group   ref="EG_Geometry"                  minOccurs="0" maxOccurs="1"/>
            * 2245     <xsd:group   ref="EG_FillProperties"            minOccurs="0" maxOccurs="1"/>
            * 2246     <xsd:element name="ln" type="CT_LineProperties" minOccurs="0" maxOccurs="1"/>
            * 2247     <xsd:group   ref="EG_EffectProperties"          minOccurs="0" maxOccurs="1"/>
            * 2248     <xsd:element name="scene3d" type="CT_Scene3D"   minOccurs="0" maxOccurs="1"/>
            * 2249     <xsd:element name="sp3d" type="CT_Shape3D"      minOccurs="0" maxOccurs="1"/>
            * 2250     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
            * 2251   </xsd:sequence>
            * 2252   <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
            * 2253 </xsd:complexType>
            */

            // TODO:
        }

    function getBackground(warpObj, slideSize, index, settings) {
            //var rslt = "";
            var slideContent = warpObj["slideContent"];
            var slideLayoutContent = warpObj["slideLayoutContent"];
            var slideMasterContent = warpObj["slideMasterContent"];

            var nodesSldLayout = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:spTree"]);
            var nodesSldMaster = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:spTree"]);
            // console.log("slideContent : ", slideContent)
            // console.log("slideLayoutContent : ", slideLayoutContent)
            // console.log("slideMasterContent : ", slideMasterContent)
            //console.log("warpObj : ", warpObj)
            var showMasterSp = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "attrs", "showMasterSp"]);
            //console.log("slideLayoutContent : ", slideLayoutContent, ", showMasterSp: ", showMasterSp)
            var bgColor = PPTXStyleUtils.getSlideBackgroundFill(warpObj, index);
            var result = "<div class='slide-background-" + index + "' style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;" + bgColor + "'>"
            var node_ph_type_ary = [];
            if (nodesSldLayout !== undefined) {
                for (var nodeKey in nodesSldLayout) {
                    if (nodesSldLayout[nodeKey].constructor === Array) {
                        for (var i = 0; i < nodesSldLayout[nodeKey].length; i++) {
                            var ph_type = PPTXXmlUtils.getTextByPathList(nodesSldLayout[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                            // if (ph_type !== undefined && ph_type != "pic") {
                            //     node_ph_type_ary.push(ph_type);
                            // }
                            if (ph_type != "pic") {
                                result += PPTXNodeUtils.processNodesInSlide(nodeKey, nodesSldLayout[nodeKey][i], nodesSldLayout, warpObj, "slideLayoutBg", 'group', settings); //slideLayoutBg , slideMasterBg
                            }
                        }
                    } else {
                        var ph_type = PPTXXmlUtils.getTextByPathList(nodesSldLayout[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                        // if (ph_type !== undefined && ph_type != "pic") {
                        //     node_ph_type_ary.push(ph_type);
                        // }
                        if (ph_type != "pic") {
                            result += PPTXNodeUtils.processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], nodesSldLayout, warpObj, "slideLayoutBg", 'group', settings); //slideLayoutBg, slideMasterBg
                        }
                    }
                }
            }
            if (nodesSldMaster !== undefined && (showMasterSp == "1" || showMasterSp === undefined)) {
                for (var nodeKey in nodesSldMaster) {
                    if (nodesSldMaster[nodeKey].constructor === Array) {
                        for (var i = 0; i < nodesSldMaster[nodeKey].length; i++) {
                            var ph_type = PPTXXmlUtils.getTextByPathList(nodesSldMaster[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                            //if (node_ph_type_ary.indexOf(ph_type) > -1) {
                            result += PPTXNodeUtils.processNodesInSlide(nodeKey, nodesSldMaster[nodeKey][i], nodesSldMaster, warpObj, "slideMasterBg", 'group', settings); //slideLayoutBg , slideMasterBg
                            //}
                        }
                    } else {
                        var ph_type = PPTXXmlUtils.getTextByPathList(nodesSldMaster[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                        //if (node_ph_type_ary.indexOf(ph_type) > -1) {
                        result += PPTXNodeUtils.processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], nodesSldMaster, warpObj, "slideMasterBg", 'group', settings); //slideLayoutBg, slideMasterBg
                        //}
                    }
                }
            }
            return result;

        }

    return {
        indexNodes: indexNodes,
        processGroupSpNode: processGroupSpNode,
        processNodesInSlide: processNodesInSlide,
        processSpNode,
        processCxnSpNode,
        processPicNode,
        processGraphicFrameNode,
        processSpPrNode,
        getBackground,
        genDiagram
    };
})();