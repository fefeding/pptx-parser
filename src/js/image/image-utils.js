import { PPTXUtils } from '../utils/utils.js';
import { PPTXHtml } from '../html-generator.js';

var PPTXImageUtils = {};

    /**
 * 处理图片节点
 * @param {Object} node - 图片节点
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 来源
 * @param {string} sType - 类型
 * @param {Function} getPosition - 获取位置的函数
 * @param {Function} getSize - 获取尺寸的函数
 * @param {Object} settings - 配置设置
 * @returns {string} HTML字符串
 */
PPTXImageUtils.processPicNode = function(node, warpObj, source, sType, getPosition, getSize, settings) {
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
    var imgName = resObj[rid]["target"];

    var imgFileExt = PPTXUtils.extractFileExtension(imgName).toLowerCase();
    var zip = warpObj["zip"];

    // 尝试解析图片路径,处理相对路径问题
    var imgFile = zip.file(imgName);
    if (!imgFile && !imgName.startsWith("ppt/")) {
        imgFile = zip.file("ppt/" + imgName);
    }
    if (!imgFile) {
        console.error("Image file not found:", imgName);
        return "";
    }

    var imgArrayBuffer = imgFile.async("arraybuffer");
    var mimeType = "";
    var xfrmNode = node["p:spPr"]["a:xfrm"];
    if (xfrmNode === undefined) {
        var idx = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "idx"]);
        if (idx !== undefined) {
            xfrmNode = PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"], ["idxTable", idx, "p:spPr", "a:xfrm"]);
        }
    }

    // 处理旋转
    var rotate = 0;
    var rotateNode = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:xfrm", "attrs", "rot"]);
    if (rotateNode !== undefined) {
        rotate = PPTXUtils.angleToDegrees(rotateNode);
    }

    // 处理视频
    var vdoNode = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:videoFile"]);
    var vdoRid, vdoFile, vdoFileExt, vdoMimeType, uInt8Array, blob, vdoBlob, mediaSupportFlag = false, isVdeoLink = false;
    var mediaProcess = settings.mediaProcess;

    if (vdoNode !== undefined && mediaProcess) {
        vdoRid = vdoNode["attrs"]["r:link"];
        vdoFile = resObj[vdoRid]["target"];
        var checkIfLink = PPTXUtils.IsVideoLink(vdoFile);
        if (checkIfLink) {
            vdoFile = PPTXUtils.escapeHtml(vdoFile);
            isVdeoLink = true;
            mediaSupportFlag = true;
            mediaPicFlag = true;
        } else {
            vdoFileExt = PPTXUtils.extractFileExtension(vdoFile).toLowerCase();
            if (vdoFileExt == "mp4" || vdoFileExt == "webm" || vdoFileExt == "ogg") {
                var vdoFileEntry = zip.file(vdoFile);
                if (!vdoFileEntry && !vdoFile.startsWith("ppt/")) {
                    vdoFileEntry = zip.file("ppt/" + vdoFile);
                }
                if (!vdoFileEntry) {
                    console.error("Video file not found:", vdoFile);
                } else {
                    uInt8Array = vdoFileEntry.async("arraybuffer");
                    vdoMimeType = PPTXUtils.getMimeType(vdoFileExt);
                    blob = new Blob([uInt8Array], { type: vdoMimeType });
                    vdoBlob = URL.createObjectURL(blob);
                    mediaSupportFlag = true;
                    mediaPicFlag = true;
                }
            }
        }
    }

    // 处理音频
    var audioNode = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:audioFile"]);
    var audioRid, audioFile, audioFileExt, audioMimeType, uInt8ArrayAudio, blobAudio, audioBlob;
    var audioPlayerFlag = false;
    var audioObjc;

    if (audioNode !== undefined && mediaProcess) {
        audioRid = audioNode["attrs"]["r:link"];
        audioFile = resObj[audioRid]["target"];
        audioFileExt = PPTXUtils.extractFileExtension(audioFile).toLowerCase();
        if (audioFileExt == "mp3" || audioFileExt == "wav" || audioFileExt == "ogg") {
            var audioFileEntry = zip.file(audioFile);
            if (!audioFileEntry && !audioFile.startsWith("ppt/")) {
                audioFileEntry = zip.file("ppt/" + audioFile);
            }
            if (!audioFileEntry) {
                console.error("Audio file not found:", audioFile);
            } else {
                uInt8ArrayAudio = audioFileEntry.async("arraybuffer");
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
                };
                audioPlayerFlag = true;
                mediaSupportFlag = true;
                mediaPicFlag = true;
            }
        }
    }

    // 生成HTML
    mimeType = PPTXUtils.getMimeType(imgFileExt);
    rtrnData = "<div class='block content' style='" +
        ((mediaProcess && audioPlayerFlag) ? getPosition(audioObjc, node, undefined, undefined) : getPosition(xfrmNode, node, undefined, undefined)) +
        ((mediaProcess && audioPlayerFlag) ? getSize(audioObjc, undefined, undefined) : getSize(xfrmNode, undefined, undefined)) +
        " z-index: " + order + ";" +
        "transform: rotate(" + rotate + "deg);'>";

    if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
        rtrnData += "<img src='data:" + mimeType + ";base64," + PPTXUtils.base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%; height: 100%'/>";
    } else if ((vdoNode !== undefined || audioNode !== undefined) && mediaProcess && mediaSupportFlag) {
        if (vdoNode !== undefined && !isVdeoLink) {
            rtrnData += "<video src='" + vdoBlob + "' controls style='width: 100%; height: 100%'>Your browser does not support the video tag.</video>";
        } else if (vdoNode !== undefined && isVdeoLink) {
            rtrnData += "<iframe src='" + vdoFile + "' controls style='width: 100%; height: 100%'></iframe>";
        }
        if (audioNode !== undefined) {
            rtrnData += '<audio id="audio_player" controls><source src="' + audioBlob + '"></audio>';
        }
    }

    if (!mediaSupportFlag && mediaPicFlag) {
        rtrnData += "<span style='color:red;font-size:40px;position: absolute;'>This media file is not supported by HTML5</span>";
    }

    if ((vdoNode !== undefined || audioNode !== undefined) && !mediaProcess && mediaSupportFlag) {
        console.log("Found supported media file but media process disabled (mediaProcess=false)");
    }

    rtrnData += "</div>";
    return rtrnData;
};

    /**
 * 处理图形框架节点(表格、图表、图解)
 * @param {Object} node - 图形框架节点
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 来源
 * @param {string} sType - 类型
 * @param {Function} genTableInternal - 生成表格的函数
 * @param {Function} genDiagram - 生成图解的函数
 * @param {Function} processGroupSpNode - 处理组节点的函数
 * @returns {string} HTML字符串
 */
PPTXImageUtils.processGraphicFrameNode = function(node, warpObj, source, sType, genTableInternal, genDiagram, processGroupSpNode) {
    var result = "";
    var graphicTypeUri = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "attrs", "uri"]);

    switch (graphicTypeUri) {
        case "http://schemas.openxmlformats.org/drawingml/2006/table":
            result = genTableInternal(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/chart":
            result = PPTXHtml.genChart(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
            result = genDiagram(node, warpObj, source, sType);
            break;
        case "http://schemas.openxmlformats.org/presentationml/2006/ole":
            var oleObjNode = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent", "mc:Fallback", "p:oleObj"]);
            if (oleObjNode === undefined) {
                oleObjNode = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "p:oleObj"]);
            }
            if (oleObjNode !== undefined) {
                result = processGroupSpNode(oleObjNode, warpObj, source);
            }
            break;
        default:
    }

    return result;
};


export { PPTXImageUtils };

// Also export to global scope for backward compatibility
window.PPTXImageUtils = PPTXImageUtils;