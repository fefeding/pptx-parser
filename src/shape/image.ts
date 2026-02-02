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
            result = await genTableInternal(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/chart":
            result = await PPTXHtml.genChart(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
            result = await genDiagram(node, warpObj, source, sType);
            break;
        case "http://schemas.openxmlformats.org/presentationml/2006/ole":
            // 尝试获取OLE对象
            let oleObjNode: any = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent", "mc:Fallback", "p:oleObj"]);
            if (oleObjNode === undefined) {
                oleObjNode = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "p:oleObj"]);
            }
            
            // 检查是否是数学公式对象
            let isMathFormula = false;
            if (oleObjNode) {
                // 检查对象类型是否为数学公式
                const progId = PPTXUtils.getTextByPathList(oleObjNode, ["attrs", "progId"]);
                if (progId && (progId.includes("Equation") || progId.includes("Math"))) {
                    isMathFormula = true;
                }
                
                // 对于数学公式OLE对象，即使没有显式标记，也可能是数学公式
                if (!isMathFormula) {
                    // 检查OLE对象的其他特征
                    const type = PPTXUtils.getTextByPathList(oleObjNode, ["attrs", "type"]);
                    if (type && (type.includes("equation") || type.includes("math"))) {
                        isMathFormula = true;
                    }
                }
            }
            
            // 尝试从OLE对象相关的图形数据中获取图像数据（如您示例中的数学公式图像）
            let imageData = null;
            let imageResourceId = null;
            
            if (isMathFormula) {
                // 检查备用内容中是否有图像数据
                const alternateContent = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent"]);
                if (alternateContent) {
                    // 尝试从选择和备用内容中获取图像数据
                    imageData = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "pic:pict", "v:shape", "v:imagedata"]);
                    if (!imageData) {
                        imageData = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "pic:pict", "v:shape", "v:imagedata"]);
                    }
                    
                    // 如果找到图像数据，获取资源ID
                    if (imageData && imageData["attrs"]) {
                        imageResourceId = imageData["attrs"]["r:id"] || imageData["attrs"]["r:pic"]; 
                    }
                }
                
                // 如果备用内容中没有，尝试直接从图形数据中获取
                if (!imageData) {
                    imageData = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "pic:pict", "v:shape", "v:imagedata"]);
                    if (imageData && imageData["attrs"]) {
                        imageResourceId = imageData["attrs"]["r:id"] || imageData["attrs"]["r:pic"]; 
                    }
                }
                
                // 尝试其他可能的图像数据路径（针对不同的PPTX生成方式）
                if (!imageData) {
                    // 尝试从 blipFill 获取图像数据
                    const blipData = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:pic", "a:blipFill", "a:blip"]);
                    if (blipData && blipData["attrs"] && blipData["attrs"]["r:embed"]) {
                        imageResourceId = blipData["attrs"]["r:embed"];
                        // 创建模拟的 imageData 结构
                        imageData = { "attrs": { "r:id": imageResourceId } };
                    }
                }
                
                if (!imageData) {
                    // 尝试从其他OLE对象可能的图像路径获取
                    const directBlip = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:blip"]);
                    if (directBlip && directBlip["attrs"] && directBlip["attrs"]["r:embed"]) {
                        imageResourceId = directBlip["attrs"]["r:embed"];
                        imageData = { "attrs": { "r:id": imageResourceId } };
                    }
                }
                
                // 有时图像资源ID可能存储在不同的属性中
                if (!imageResourceId && oleObjNode && oleObjNode["attrs"]) {
                    // 尝试从OLE对象属性中获取资源ID
                    imageResourceId = oleObjNode["attrs"]["r:id"];
                }
            }
            
            // 如果是数学公式，尝试从备用内容中获取OMath
            if (isMathFormula) {
                const alternateContent = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent"]);
                if (alternateContent) {
                    // 尝试从OMath中获取数学内容
                    const oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "m:oMathPara"]);
                    if (oMathContent) {
                        // 如果找到OMath内容，生成替代文本或SVG表示
                        result = generateMathPlaceholder(oMathContent);
                    } else {
                        // 尝试其他OMath路径
                        const oMathContent2 = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "m:oMathPara"]);
                        if (oMathContent2) {
                            result = generateMathPlaceholder(oMathContent2);
                        } else {
                            // 再尝试其他可能的OMath路径
                            const oMathContent3 = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "m:oMath"]);
                            if (oMathContent3) {
                                result = generateMathPlaceholder(oMathContent3);
                            } else {
                                const oMathContent4 = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "m:oMath"]);
                                if (oMathContent4) {
                                    result = generateMathPlaceholder(oMathContent4);
                                } else {
                                    // 如果找不到OMath，但有图像数据，尝试使用图像数据
                                    if (imageData) {
                                        // 获取图像资源ID
                                        let imgId = PPTXUtils.getTextByPathList(imageData, ["attrs", "r:id"]);
                                        
                                        // 如果r:id没有找到，尝试其他可能的属性
                                        if (!imgId && imageData["attrs"]) {
                                            const attrs = imageData["attrs"];
                                            imgId = attrs["r:embed"] || attrs["r:link"];
                                        }
                                        
                                        if (imgId) {
                                            // 尝试从不同的资源对象中获取图像
                                            let resObj = warpObj["slideResObj"];
                                            let imgName = resObj[imgId]?.["target"];
                                            
                                            // 如果在slideResObj中没找到，尝试在其他资源对象中查找
                                            if (!imgName) {
                                                resObj = warpObj["masterResObj"];
                                                imgName = resObj[imgId]?.["target"];
                                            }
                                            if (!imgName) {
                                                resObj = warpObj["layoutResObj"];
                                                imgName = resObj[imgId]?.["target"];
                                            }
                                            
                                            // 对于OLE对象，有时资源ID直接就是文件路径
                                            if (!imgName) {
                                                // 尝试将imgId作为文件路径
                                                if (imgId.startsWith('ppt/media/') || imgId.startsWith('ppt/')) {
                                                    imgName = imgId;
                                                    resObj = warpObj["slideResObj"]; // 使用任意resObj变量
                                                } else {
                                                    // 尝试构建可能的媒体路径
                                                    const possiblePaths = [
                                                        `ppt/media/${imgId}.png`,
                                                        `ppt/media/${imgId}.jpg`,
                                                        `ppt/media/${imgId}.jpeg`,
                                                        `ppt/media/${imgId}.gif`,
                                                        `ppt/media/${imgId}.bmp`,
                                                        `ppt/embeddings/${imgId}.png`,
                                                        `ppt/embeddings/${imgId}.jpg`,
                                                        `ppt/embeddings/${imgId}.jpeg`,
                                                        `ppt/embeddings/${imgId}.gif`,
                                                        `ppt/embeddings/${imgId}.bmp`,
                                                        `${imgId}.png`,
                                                        `${imgId}.jpg`,
                                                        `${imgId}.jpeg`,
                                                        `${imgId}.gif`,
                                                        `${imgId}.bmp`
                                                    ];
                                                    
                                                    const zip = warpObj["zip"];
                                                    for (const path of possiblePaths) {
                                                        if (zip.file(path)) {
                                                            imgName = path;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                            
                                            const zip = warpObj["zip"];
                                            
                                            if (imgName && zip) {
                                                // 尝试获取图像数据
                                                let imgFile = zip.file(imgName);
                                                if (!imgFile && !imgName.startsWith("ppt/")) {
                                                    imgFile = zip.file("ppt/" + imgName);
                                                }
                                                if (!imgFile && !imgName.startsWith("/")) {
                                                    imgFile = zip.file("/" + imgName);
                                                }
                                                
                                                if (imgFile) {
                                                    const imgArrayBuffer = await imgFile.async('arraybuffer');
                                                    const mimeType = PPTXUtils.getMimeType(PPTXUtils.extractFileExtension(imgName));
                                                    const imgBase64 = PPTXUtils.base64ArrayBuffer(imgArrayBuffer);
                                                    
                                                    // 生成图像HTML
                                                    const xfrmNode = PPTXUtils.getTextByPathList(node, ["p:xfrm"]);
                                                    const order = node["attrs"]["order"];
                                                    const positionStyle = PPTXUtils.getPosition(xfrmNode, node, undefined, undefined, undefined);
                                                    const sizeStyle = PPTXUtils.getSize(xfrmNode, undefined, undefined);
                                                    
                                                    result = `<div class='block content' style='${positionStyle}${sizeStyle} z-index: ${order};'>
                                                                <img src='data:${mimeType};base64,${imgBase64}' style='width: 100%; height: 100%' alt='数学公式'/>
                                                              </div>`;
                                                } else {
                                                    console.warn(`Could not find image file for id: ${imgId}, name: ${imgName}`);
                                                    // 如果无法从图像数据生成，尝试从OMath内容生成数学公式占位符
                                                    const alternateContent = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent"]);
                                                    if (alternateContent) {
                                                        const oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                                        if (oMathContent) {
                                                            result = generateMathPlaceholder(oMathContent);
                                                        } else {
                                                            const fallbackOMath = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                                            if (fallbackOMath) {
                                                                result = generateMathPlaceholder(fallbackOMath);
                                                            }
                                                        }
                                                    }
                                                }
                                            } else {
                                                console.warn(`Could not find image resource for id: ${imgId}`);
                                                // 如果无法从图像数据生成，尝试从OMath内容生成数学公式占位符
                                                const alternateContent = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent"]);
                                                if (alternateContent) {
                                                    const oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                                    if (oMathContent) {
                                                        result = generateMathPlaceholder(oMathContent);
                                                    } else {
                                                        const fallbackOMath = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                                        if (fallbackOMath) {
                                                            result = generateMathPlaceholder(fallbackOMath);
                                                        }
                                                    }
                                                }
                                            }
                                        } else {
                                            console.warn('Could not find image ID in imagedata attributes');
                                            // 如果无法从图像数据生成，尝试从OMath内容生成数学公式占位符
                                            const alternateContent = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent"]);
                                            if (alternateContent) {
                                                const oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                                if (oMathContent) {
                                                    result = generateMathPlaceholder(oMathContent);
                                                } else {
                                                    const fallbackOMath = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                                    if (fallbackOMath) {
                                                        result = generateMathPlaceholder(fallbackOMath);
                                                    }
                                                }
                                            }
                                        }
                                    } else {
                                        // 如果找不到OMath，回退到常规处理
                                        if (oleObjNode !== undefined) {
                                            result = await processGroupSpNode(oleObjNode, warpObj, source);
                                        }
                                    }
                                }
                            }
                        }
                    }
                } else {
                    // 如果没有备用内容，但有图像数据，尝试使用图像数据
                    if (imageData) {
                        // 获取图像资源ID
                        let imgId = PPTXUtils.getTextByPathList(imageData, ["attrs", "r:id"]);
                        
                        // 如果r:id没有找到，尝试其他可能的属性
                        if (!imgId && imageData["attrs"]) {
                            const attrs = imageData["attrs"];
                            imgId = attrs["r:embed"] || attrs["r:link"];
                        }
                        
                        if (imgId) {
                            // 尝试从不同的资源对象中获取图像
                            let resObj = warpObj["slideResObj"];
                            let imgName = resObj[imgId]?.["target"];
                            
                            // 如果在slideResObj中没找到，尝试在其他资源对象中查找
                            if (!imgName) {
                                resObj = warpObj["masterResObj"];
                                imgName = resObj[imgId]?.["target"];
                            }
                            if (!imgName) {
                                resObj = warpObj["layoutResObj"];
                                imgName = resObj[imgId]?.["target"];
                            }
                            
                            // 对于OLE对象，有时资源ID直接就是文件路径
                            if (!imgName) {
                                // 尝试将imgId作为文件路径
                                if (imgId.startsWith('ppt/media/') || imgId.startsWith('ppt/')) {
                                    imgName = imgId;
                                    resObj = warpObj["slideResObj"]; // 使用任意resObj变量
                                } else {
                                    // 尝试构建可能的媒体路径
                                    const possiblePaths = [
                                        `ppt/media/${imgId}.png`,
                                        `ppt/media/${imgId}.jpg`,
                                        `ppt/media/${imgId}.jpeg`,
                                        `ppt/media/${imgId}.gif`,
                                        `ppt/media/${imgId}.bmp`,
                                        `ppt/embeddings/${imgId}.png`,
                                        `ppt/embeddings/${imgId}.jpg`,
                                        `ppt/embeddings/${imgId}.jpeg`,
                                        `ppt/embeddings/${imgId}.gif`,
                                        `ppt/embeddings/${imgId}.bmp`,
                                        `${imgId}.png`,
                                        `${imgId}.jpg`,
                                        `${imgId}.jpeg`,
                                        `${imgId}.gif`,
                                        `${imgId}.bmp`
                                    ];
                                    
                                    const zip = warpObj["zip"];
                                    for (const path of possiblePaths) {
                                        if (zip.file(path)) {
                                            imgName = path;
                                            break;
                                        }
                                    }
                                }
                            }
                            
                            const zip = warpObj["zip"];
                            
                            if (imgName && zip) {
                                // 尝试获取图像数据
                                let imgFile = zip.file(imgName);
                                if (!imgFile && !imgName.startsWith("ppt/")) {
                                    imgFile = zip.file("ppt/" + imgName);
                                }
                                if (!imgFile && !imgName.startsWith("/")) {
                                    imgFile = zip.file("/" + imgName);
                                }
                                
                                if (imgFile) {
                                    const imgArrayBuffer = await imgFile.async('arraybuffer');
                                    const mimeType = PPTXUtils.getMimeType(PPTXUtils.extractFileExtension(imgName));
                                    const imgBase64 = PPTXUtils.base64ArrayBuffer(imgArrayBuffer);
                                    
                                    // 生成图像HTML
                                    const xfrmNode = PPTXUtils.getTextByPathList(node, ["p:xfrm"]);
                                    const order = node["attrs"]["order"];
                                    const positionStyle = PPTXUtils.getPosition(xfrmNode, node, undefined, undefined, undefined);
                                    const sizeStyle = PPTXUtils.getSize(xfrmNode, undefined, undefined);
                                    
                                    result = `<div class='block content' style='${positionStyle}${sizeStyle} z-index: ${order};'>
                                                <img src='data:${mimeType};base64,${imgBase64}' style='width: 100%; height: 100%' alt='数学公式'/>
                                              </div>`;
                                } else {
                                    console.warn(`Could not find image file for id: ${imgId}, name: ${imgName}`);
                                    // 如果无法从图像数据生成，尝试从OMath内容生成数学公式占位符
                                    const alternateContent = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent"]);
                                    if (alternateContent) {
                                        const oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                        if (oMathContent) {
                                            result = generateMathPlaceholder(oMathContent);
                                        } else {
                                            const fallbackOMath = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                            if (fallbackOMath) {
                                                result = generateMathPlaceholder(fallbackOMath);
                                            }
                                        }
                                    }
                                }
                            } else {
                                console.warn(`Could not find image resource for id: ${imgId}, name: ${imgName}`);
                                // 如果无法从图像数据生成，尝试从OMath内容生成数学公式占位符
                                const alternateContent = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent"]);
                                if (alternateContent) {
                                    const oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                    if (oMathContent) {
                                        result = generateMathPlaceholder(oMathContent);
                                    } else {
                                        const fallbackOMath = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                        if (fallbackOMath) {
                                            result = generateMathPlaceholder(fallbackOMath);
                                        }
                                    }
                                }
                            }
                        } else {
                            console.warn('Could not find image ID in imagedata attributes');
                            // 如果无法从图像数据生成，尝试从OMath内容生成数学公式占位符
                            const alternateContent = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent"]);
                            if (alternateContent) {
                                const oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                if (oMathContent) {
                                    result = generateMathPlaceholder(oMathContent);
                                } else {
                                    const fallbackOMath = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "m:oMathPara"]);
                                    if (fallbackOMath) {
                                        result = generateMathPlaceholder(fallbackOMath);
                                    }
                                }
                            }
                        }
                    } else {
                        // 如果没有备用内容，尝试常规处理
                        if (oleObjNode !== undefined) {
                            result = await processGroupSpNode(oleObjNode, warpObj, source);
                        }
                    }
                }
            } else {
                // 非数学公式OLE对象，使用常规处理
                if (oleObjNode !== undefined) {
                    result = await processGroupSpNode(oleObjNode, warpObj, source);
                }
            }
            
            // 如果结果仍然为空，尝试直接从图形数据中查找OMath内容
            if (!result || result.trim() === '') {
                const alternateContent = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent"]);
                if (alternateContent) {
                    const oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "m:oMathPara"]);
                    if (oMathContent) {
                        result = generateMathPlaceholder(oMathContent);
                    } else {
                        const fallbackOMath = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "m:oMathPara"]);
                        if (fallbackOMath) {
                            result = generateMathPlaceholder(fallbackOMath);
                        } else {
                            // 尝试直接在图形数据中查找OMath
                            const directOMath = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "m:oMathPara"]);
                            if (directOMath) {
                                result = generateMathPlaceholder(directOMath);
                            } else {
                                // 尝试在其他可能的位置查找OMath内容
                                const directOMath2 = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "m:oMath"]);
                                if (directOMath2) {
                                    result = generateMathPlaceholder(directOMath2);
                                }
                            }
                        }
                    }
                }
            }
            
            // 如果仍然没有结果，检查是否OLE对象中包含可访问的图像资源
            if (!result || result.trim() === '') {
                // 尝试从OLE对象属性中直接获取图像资源
                if (oleObjNode && oleObjNode["attrs"]) {
                    const oleAttrs = oleObjNode["attrs"];
                    const possibleResourceIds = [oleAttrs["r:id"], oleAttrs["r:embed"], oleAttrs["r:link"]];
                                
                    for (const resourceId of possibleResourceIds) {
                        if (resourceId) {
                            // 尝试从不同资源对象中查找
                            const resourceObjects = [warpObj["slideResObj"], warpObj["masterResObj"], warpObj["layoutResObj"]];
            
                            for (const resObj of resourceObjects) {
                                if (resObj && resObj[resourceId] && resObj[resourceId]["target"]) {
                                    const imgName = resObj[resourceId]["target"];
                                    const zip = warpObj["zip"];
                                                
                                    if (zip) {
                                        let imgFile = zip.file(imgName);
                                        if (!imgFile && !imgName.startsWith("ppt/")) {
                                            imgFile = zip.file("ppt/" + imgName);
                                        }
                                        if (!imgFile && !imgName.startsWith("/")) {
                                            imgFile = zip.file("/" + imgName);
                                        }
                                                    
                                        if (imgFile) {
                                            const imgArrayBuffer = await imgFile.async('arraybuffer');
                                            const mimeType = PPTXUtils.getMimeType(PPTXUtils.extractFileExtension(imgName));
                                            const imgBase64 = PPTXUtils.base64ArrayBuffer(imgArrayBuffer);
                                                        
                                            // 生成图像HTML
                                            const xfrmNode = PPTXUtils.getTextByPathList(node, ["p:xfrm"]);
                                            const order = node["attrs"]["order"];
                                            const positionStyle = PPTXUtils.getPosition(xfrmNode, node, undefined, undefined, undefined);
                                            const sizeStyle = PPTXUtils.getSize(xfrmNode, undefined, undefined);
                                                        
                                            result = `<div class='block content' style='${positionStyle}${sizeStyle} z-index: ${order};'>
                                                        <img src='data:${mimeType};base64,${imgBase64}' style='width: 100%; height: 100%' alt='OLE Object'/>
                                                      </div>`;
                                            break; // 找到图像后退出循环
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
                        
            // 如果仍然没有找到合适的图像数据，生成一个包含数学公式的占位符
            if (!result || result.trim() === '') {
                // 尝试从OMath内容生成数学公式表示
                const alternateContent = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "mc:AlternateContent"]);
                let oMathContent = null;
                            
                if (alternateContent) {
                    oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "m:oMathPara"]); 
                    if (!oMathContent) {
                        oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "m:oMathPara"]);
                    }
                    if (!oMathContent) {
                        oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Choice", "a:graphic", "a:graphicData", "m:oMath"]);
                    }
                    if (!oMathContent) {
                        oMathContent = PPTXUtils.getTextByPathList(alternateContent, ["mc:Fallback", "a:graphic", "a:graphicData", "m:oMath"]);
                    }
                }
                            
                if (oMathContent) {
                    // 使用更高级的数学公式处理
                    result = generateMathPlaceholder(oMathContent);
                } else {
                    // 如果所有方法都失败，至少显示一个数学公式占位符
                    const xfrmNode = PPTXUtils.getTextByPathList(node, ["p:xfrm"]);
                    const order = node["attrs"]["order"];
                    const positionStyle = PPTXUtils.getPosition(xfrmNode, node, undefined, undefined, undefined);
                    const sizeStyle = PPTXUtils.getSize(xfrmNode, undefined, undefined);
                                
                    result = `<div class='block content' style='${positionStyle}${sizeStyle} z-index: ${order};'>
                                <div style='display: flex; align-items: center; justify-content: center; width: 100%; height: 100%; background-color: #f0f0f0; border: 1px dashed #ccc;'>
                                    <span style='color: #666; font-style: italic;'>数学公式</span>
                                </div>
                              </div>`;
                }
            }
                        
            break;
        default:
    }

    return result;
};

// 辅助函数：生成数学公式占位符
function generateMathPlaceholder(oMathContent: any): string {
    // 如果无法解析OMath，生成一个占位符
    // 这里可以扩展为实际解析OMath内容并转换为可视化的数学公式
    if (oMathContent && typeof oMathContent === 'object') {
        // 尝试从OMath内容中提取简单的文本表示
        const mathText = extractSimpleMathText(oMathContent);
        if (mathText) {
            return "<span class='math-placeholder' style='font-style:italic; color:#0066cc;' title='数学公式'>" + mathText + "</span>";
        }
    }
    return "<span class='math-placeholder' style='font-style:italic; color:#0066cc;' title='数学公式'>[数学公式]</span>";
}

// 辅助函数：从OMath内容中提取简单文本
function extractSimpleMathText(oMathContent: any): string | null {
    try {
        // 递归搜索文本节点
        const findTextNodes = (node: any): string[] => {
            let texts: string[] = [];
            
            if (typeof node === 'object' && node !== null) {
                for (const key in node) {
                    if (key === 'm:t' && typeof node[key] === 'string') {
                        // 直接文本节点
                        texts.push(node[key]);
                    } else if (key === 'm:r' && typeof node[key] === 'object') {
                        // 寻找文本运行中的文本
                        const run = node[key];
                        if (Array.isArray(run)) {
                            for (const r of run) {
                                if (r && r['m:t']) {
                                    texts.push(r['m:t']);
                                }
                            }
                        } else if (run && run['m:t']) {
                            texts.push(run['m:t']);
                        }
                    } else if (key === 'm:sSub' || key === 'm:sSup' || key === 'm:f' || key === 'm:e' || key === 'm:num' || key === 'm:den' || key === 'm:d' || key === 'm:r') {
                        // 数学结构节点：下标、上标、分数、分子、分母、分隔符等
                        texts = texts.concat(findTextNodes(node[key]));
                    } else if (typeof node[key] === 'object') {
                        texts = texts.concat(findTextNodes(node[key]));
                    }
                }
            }
            
            return texts;
        };

        const textNodes = findTextNodes(oMathContent);
        return textNodes.join('');
    } catch (e) {
        return null;
    }
}

export { PPTXImageUtils };