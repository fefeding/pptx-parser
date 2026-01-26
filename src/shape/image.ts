import { PPTXUtils } from '../core/utils.js';
import { PPTXHtml } from '../html.js';

interface PPTXImageUtilsType {
    processPicNode: (node: any, warpObj: any, source: string, sType: string, getPosition: Function, getSize: Function, settings: any) => Promise<string>;
    processGraphicFrameNode: (node: any, warpObj: any, source: string, sType: string, genTableInternal: Function, genDiagram: Function, processGroupSpNode: Function) => string | Promise<string>;
}

const PPTXImageUtils = {} as PPTXImageUtilsType;

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
PPTXImageUtils.processPicNode = async function(node: any, warpObj: any, source: string, sType: string, getPosition: Function, getSize: Function, settings: any): Promise<string> {
    let rtrnData: string = "";
    let mediaPicFlag: boolean = false;
    const order: string = node["attrs"]["order"];

    const rid: string = node["p:blipFill"]["a:blip"]["attrs"]["r:embed"];
    let resObj: any;
    if (source == "slideMasterBg") {
        resObj = warpObj["masterResObj"];
    } else if (source == "slideLayoutBg") {
        resObj = warpObj["layoutResObj"];
    } else {
        resObj = warpObj["slideResObj"];
    }
    const imgName: string = resObj[rid]["target"];

    const imgFileExt: string = PPTXUtils.extractFileExtension(imgName).toLowerCase();
    const zip: any = warpObj["zip"];

    // 尝试解析图片路径,处理相对路径问题
    let imgFile: any = zip.file(imgName);
    if (!imgFile && !imgName.startsWith("ppt/")) {
        imgFile = zip.file("ppt/" + imgName);
    }
    if (!imgFile) {
        return "";
    }

    const imgArrayBuffer: ArrayBuffer = await imgFile.async('arraybuffer');
    let mimeType: string = "";
    let xfrmNode: any = node["p:spPr"]["a:xfrm"];
    if (xfrmNode === undefined) {
        const idx: any = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "idx"]);
        if (idx !== undefined) {
            xfrmNode = PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"], ["idxTable", idx, "p:spPr", "a:xfrm"]);
        }
    }

    // 处理旋转
    let rotate: number = 0;
    const rotateNode: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:xfrm", "attrs", "rot"]);
    if (rotateNode !== undefined) {
        rotate = PPTXUtils.angleToDegrees(rotateNode);
    }

    // 处理视频
    const vdoNode: any = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:videoFile"]);
    let vdoRid: string, vdoFile: string, vdoFileExt: string, vdoMimeType: string, uInt8Array: any, blob: Blob, vdoBlob: string, mediaSupportFlag: boolean = false, isVdeoLink: boolean = false;
    const mediaProcess: boolean = settings.mediaProcess;

    if (vdoNode !== undefined && mediaProcess) {
        vdoRid = vdoNode["attrs"]["r:link"];
        vdoFile = resObj[vdoRid]["target"];
        const checkIfLink: boolean = PPTXUtils.IsVideoLink(vdoFile);
        if (checkIfLink) {
            vdoFile = PPTXUtils.escapeHtml(vdoFile);
            isVdeoLink = true;
            mediaSupportFlag = true;
            mediaPicFlag = true;
        } else {
            vdoFileExt = PPTXUtils.extractFileExtension(vdoFile).toLowerCase();
            if (vdoFileExt == "mp4" || vdoFileExt == "webm" || vdoFileExt == "ogg") {
                let vdoFileEntry: any = zip.file(vdoFile);
                if (!vdoFileEntry && !vdoFile.startsWith("ppt/")) {
                    vdoFileEntry = zip.file("ppt/" + vdoFile);
                }
                if (!vdoFileEntry) {
                    //
                } else {
                    uInt8Array = new Uint8Array(await vdoFileEntry.async('arraybuffer'));
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
    const audioNode: any = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:audioFile"]);
    let audioRid: string, audioFile: string, audioFileExt: string, audioMimeType: string, uInt8ArrayAudio: any, blobAudio: Blob, audioBlob: string;
    let audioPlayerFlag: boolean = false;
    let audioObjc: any;

    if (audioNode !== undefined && mediaProcess) {
        audioRid = audioNode["attrs"]["r:link"];
        audioFile = resObj[audioRid]["target"];
        audioFileExt = PPTXUtils.extractFileExtension(audioFile).toLowerCase();
        if (audioFileExt == "mp3" || audioFileExt == "wav" || audioFileExt == "ogg") {
            let audioFileEntry: any = zip.file(audioFile);
            if (!audioFileEntry && !audioFile.startsWith("ppt/")) {
                audioFileEntry = zip.file("ppt/" + audioFile);
            }
            if (!audioFileEntry) {
                //
            } else {
                uInt8ArrayAudio = new Uint8Array(await audioFileEntry.async('arraybuffer'));
                blobAudio = new Blob([uInt8ArrayAudio]);
                audioBlob = URL.createObjectURL(blobAudio);
                const cx: number = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * 20;
                const cy: any = xfrmNode["a:ext"]["attrs"]["cy"];
                let x: number = parseInt(xfrmNode["a:off"]["attrs"]["x"]) / 2.5;
                let y: any = xfrmNode["a:off"]["attrs"]["y"];
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
        " z-index: " + order + `;transform: rotate(` + rotate + "deg);'>";

    if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
        rtrnData += "<img src='" + PPTXUtils.arrayBufferToBlobUrl(imgArrayBuffer, mimeType) + "' style='width: 100%; height: 100%'/>";
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
        //
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
PPTXImageUtils.processGraphicFrameNode = async function(node: any, warpObj: any, source: string, sType: string, genTableInternal: Function, genDiagram: Function, processGroupSpNode: Function): Promise<string> {
    let result: string = "";
    const graphicTypeUri: string = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "attrs", "uri"]);

    switch (graphicTypeUri) {
        case "http://schemas.openxmlformats.org/drawingml/2006/table":
            result = genTableInternal(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/chart":
            result = await PPTXHtml.genChart(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
            result = genDiagram(node, warpObj, source, sType);
            break;
        case "http://schemas.openxmlformats.org/presentationml/2006/ole":
            let oleObjNode: any = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent", "mc:Fallback", "p:oleObj"]);
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