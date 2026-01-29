/**
 * pptx-parser - PPTX 文件解析为 HTML 的主入口
 * 支持 File、Blob 或 ArrayBuffer 输入
 */

// Import dependencies
import { PPTXConstants } from './core/constants.js';
import { PPTXUtils, PPTXFileReader } from './core/utils.js';
import { PPTXParser } from './parser.js';
import { PPTXHtml } from './html.js';
import { PPTXStyleManager } from './core/style-manager.js';
import { PPTXShapeUtils } from './shape/shape.js';
import { PPTXShapePropertyExtractor } from './shape/property-extractor.js';
import { PPTXShapeFillsUtils } from './shape/fills.js';
import { PPTXBasicShapes } from './shape/basic.js';
import { PPTXNodeUtils } from './node/node.js';
import { PPTXBackgroundUtils } from './core/background.js';
import { PPTXImageUtils } from './shape/image.js';
import { PPTXDiagramUtils } from './shape/diagram.js';
import { TextUtils } from './text/text.js';
import { PPTXCSSUtils } from './core/css.js';
import { PPTXColorUtils } from './core/color.js';
import { PPTXTableUtils } from './shape/table.js';
import { PPTXTextStyleUtils } from './text/style.js';
import { PPTXTextElementUtils } from './text/element.js';
import { PPTXShapeContainer } from './shape/container.js';
import { PPTXStarShapes } from './shape/star.js';
import { PPTXCalloutShapes } from './shape/callout.js';
import { PPTXArrowShapes } from './shape/arrow.js';
import { PPTXMathShapes } from './shape/math.js';
import { processSpNode as processSpNodeModule, processCxnSpNode as processCxnSpNodeModule } from './shape/node-processor.js';
import { genShape as genShapeModule } from './shape/generator.js';


/**
 * 主函数 - 将 PPTX 文件转换为 HTML
 * @param {File|Blob|ArrayBuffer} file - 输入文件
 * @param {Object} options - 配置选项
 * @returns {Promise<Object>} 包含 html、css、slides 等的对象
 */
function pptxToHtml(file, options) {
    let isDone = false;
    const MsgQueue = [];
    PPTXHtml.MsgQueue = MsgQueue;

    const slideFactor = PPTXConstants.SLIDE_FACTOR;
    const fontSizeFactor = PPTXConstants.FONT_SIZE_FACTOR;

    let slideWidth = 0;
    let slideHeight = 0;
    let isSlideMode = false;

    // API object for external control
    const api = {
        get isSlideMode() { return isSlideMode; },
        set isSlideMode(value) { isSlideMode = value; },
        initSlideMode() { initSlideMode(options.container, options); },
        exitSlideMode() { exitSlideMode(options.container, options); },
        updateProgress(percent) { updateProgressBar(percent); }
    };

    // 计算元素位置和尺寸 - 使用 PPTXUtils 中的函数
    const getPosition = PPTXUtils ? PPTXUtils.getPosition : () => "";
    const getSize = PPTXUtils ? PPTXUtils.getSize : () => "";

    let processFullTheme = true;
    const styleTable = {};

    /**
     * 深度合并对象
     */
    function deepExtend(destination) {
        for (let i = 1; i < arguments.length; i++) {
            const source = arguments[i];
            for (const property in source) {
                if (source.hasOwnProperty(property)) {
                    if (source[property] && typeof source[property] === 'object' && source[property].constructor === Object) {
                        destination[property] = destination[property] || {};
                        deepExtend(destination[property], source[property]);
                    } else {
                        destination[property] = source[property];
                    }
                }
            }
        }
        return destination;
    }

    const settings = deepExtend({
        slidesScale: "", // 缩放百分比
        mediaProcess: true, // 是否处理视频和音频文件
        themeProcess: true, // 主题处理：true, false, "colorsAndImageOnly"
        incSlide: {
            width: 0,
            height: 0
        }
    }, options);

    processFullTheme = settings.themeProcess;

    /**
     * 更新进度条
     */
    function updateProgressBar(percent) {
        if (settings.onProgress) {
            settings.onProgress(percent);
        }
    }

    /**
     * 转换为 HTML - 将 File/Blob/ArrayBuffer 转换为 ArrayBuffer
     */
    function convertToHtml(file) {
        let fileArrayBuffer = file;
        if (file instanceof File || file instanceof Blob) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (event) => {
                    fileArrayBuffer = event.target.result;
                    processZip(fileArrayBuffer, resolve, reject);
                };
                reader.onerror = (event) => {
                    reject(new Error("Failed to read file: " + event.target.error));
                };
                reader.readAsArrayBuffer(file);
            });
        } else if (file instanceof ArrayBuffer) {
            return new Promise((resolve, reject) => {
                processZip(file, resolve, reject);
            });
        } else {
            return Promise.reject(new Error("Invalid file type: must be File, Blob, or ArrayBuffer"));
        }
    }

    /**
     * 处理 ZIP 文件 - 解析 PPTX 并生成 HTML
     */
    function processZip(fileArrayBuffer, resolve, reject) {
        if (fileArrayBuffer.byteLength < 10) {
            reject(new Error("Invalid file: too small"));
            return;
        }

        const zip = new JSZip();
        zip.load(fileArrayBuffer);

        // 配置 PPTXParser 模块 - 传递必要的回调函数
        PPTXParser.configure({
            ...settings,
            processNodesInSlide,
            getBackground,
            getSlideBackgroundFill
        });

        const rslt_ary = PPTXParser.processPPTX(zip);

        // 收集生成的 HTML、CSS 和数据
        const result = {
            html: "",
            css: "",
            slides: [],
            slideSize: null,
            chartQueue: []
        };

        for (let i = 0; i < rslt_ary.length; i++) {
            switch (rslt_ary[i].type) {
                case "slide":
                    result.html += rslt_ary[i].data;
                    result.slides.push(rslt_ary[i].data);
                    break;
                case "pptx-thumb":
                    // 缩略图可以在这里处理
                    break;
                case "slideSize":
                    slideWidth = rslt_ary[i].data.width;
                    slideHeight = rslt_ary[i].data.height;
                    result.slideSize = rslt_ary[i].data;
                    break;
                case "globalCSS":
                    result.css += rslt_ary[i].data;
                    break;
                case "ExecutionTime":
                    // 生成并添加全局 CSS
                    if (typeof PPTXCSSUtils.genGlobalCSS === 'function') {
                        result.css += PPTXCSSUtils.genGlobalCSS(styleTable, settings, slideWidth);
                    }
                    result.chartQueue = MsgQueue.slice(); // 复制图表队列
                    isDone = true;
                    break;
                case "progress-update":
                    updateProgressBar(rslt_ary[i].data);
                    break;
                default:
                    break;
            }
        }

        resolve(result);
    }

    function initSlideMode(divId, settings) {
    }

    function exitSlideMode(divId) {
    }



    function processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType) {
        const handlers = {
            processSpNode: (node, pNode, warpObj, source, sType) => {
                return processSpNodeModule(node, pNode, warpObj, source, sType, (n, pn, sLSN, sMSN, id, nm, idx, typ, ord, wo, uDBg, sTy, src) => {
                    return genShapeModule(n, pn, sLSN, sMSN, id, nm, idx, typ, ord, wo, uDBg, sTy, src, styleTable);
                });
            },
            processCxnSpNode: (node, pNode, warpObj, source, sType) => {
                return processCxnSpNodeModule(node, pNode, warpObj, source, sType, (n, pn, sLSN, sMSN, id, nm, idx, typ, ord, wo, uDBg, sTy, src) => {
                    return genShapeModule(n, pn, undefined, undefined, id, nm, idx, typ, ord, wo, undefined, sTy, src, styleTable);
                });
            },
            processPicNode,
            processGraphicFrameNode,
            processGroupSpNode
        };
        return PPTXNodeUtils.processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType, handlers);
    }

    function processGroupSpNode(node, warpObj, source) {
        return PPTXNodeUtils.processGroupSpNode(node, warpObj, source, slideFactor, processNodesInSlide);
    }

    // processSpNode 和 processCxnSpNode 已移至 pptx-shape-node-processor.js 模块
    // 这些函数现在通过 processNodesInSlide 中的 handlers 调用








    function processPicNode(node, warpObj, source, sType) {
        let rtrnData = "";
        let mediaPicFlag = false;
        const order = node.attrs.order;

        const rid = node["p:blipFill"]["a:blip"].attrs["r:embed"];
        let resObj;
        if (source === "slideMasterBg") {
            resObj = warpObj.masterResObj;
        } else if (source === "slideLayoutBg") {
            resObj = warpObj.layoutResObj;
        } else {
            resObj = warpObj.slideResObj;
        }
        const imgName = resObj[rid].target;

        const imgFileExt = PPTXUtils.extractFileExtension(imgName).toLowerCase();
        const zip = warpObj.zip;
        let imgFile = zip.file(imgName);
        if (!imgFile && !imgName.startsWith("ppt/")) {
            imgFile = zip.file("ppt/" + imgName);
        }
        if (!imgFile) {
            return "";
        }
        const imgArrayBuffer = imgFile.asArrayBuffer();
        let mimeType = "";
        let xfrmNode = node["p:spPr"]["a:xfrm"];
        if (xfrmNode === undefined) {
            const idx = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "idx"]);
            const type = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "type"]);
            if (idx !== undefined) {
                xfrmNode = PPTXUtils.getTextByPathList(warpObj.slideLayoutTables, ["idxTable", idx, "p:spPr", "a:xfrm"]);
            }
        }

        let rotate = 0;
        const rotateNode = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:xfrm", "attrs", "rot"]);
        if (rotateNode !== undefined) {
            rotate = PPTXUtils.angleToDegrees(rotateNode);
        }

        // 视频处理
        const vdoNode = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:videoFile"]);
        let vdoRid, vdoFile, vdoFileExt, vdoMimeType, uInt8Array, blob, vdoBlob, mediaSupportFlag = false, isVdeoLink = false;
        const mediaProcess = settings.mediaProcess;
        if (vdoNode !== undefined && mediaProcess) {
            vdoRid = vdoNode.attrs["r:link"];
            vdoFile = resObj[vdoRid].target;
            const checkIfLink = PPTXUtils.IsVideoLink(vdoFile);
            if (checkIfLink) {
                vdoFile = PPTXUtils.escapeHtml(vdoFile);
                isVdeoLink = true;
                mediaSupportFlag = true;
                mediaPicFlag = true;
            } else {
                vdoFileExt = PPTXUtils.extractFileExtension(vdoFile).toLowerCase();
                if (vdoFileExt === "mp4" || vdoFileExt === "webm" || vdoFileExt === "ogg") {
                    let vdoFileEntry = zip.file(vdoFile);
                    if (!vdoFileEntry && !vdoFile.startsWith("ppt/")) {
                        vdoFileEntry = zip.file("ppt/" + vdoFile);
                    }
                    if (!vdoFileEntry) {
                    } else {
                        uInt8Array = vdoFileEntry.asArrayBuffer();
                        vdoMimeType = PPTXUtils.getMimeType(vdoFileExt);
                        blob = new Blob([uInt8Array], { type: vdoMimeType });
                        vdoBlob = URL.createObjectURL(blob);
                            mediaSupportFlag = true;
                            mediaPicFlag = true;
                    }
                }
            }
        }
        // 音频处理
        const audioNode = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:audioFile"]);
        let audioRid, audioFile, audioFileExt, audioMimeType, uInt8ArrayAudio, blobAudio, audioBlob;
        let audioPlayerFlag = false;
        let audioObjc;
        if (audioNode !== undefined && mediaProcess) {
            audioRid = audioNode.attrs["r:link"];
            audioFile = resObj[audioRid].target;
            audioFileExt = PPTXUtils.extractFileExtension(audioFile).toLowerCase();
            if (audioFileExt === "mp3" || audioFileExt === "wav" || audioFileExt === "ogg") {
                let audioFileEntry = zip.file(audioFile);
                if (!audioFileEntry && !audioFile.startsWith("ppt/")) {
                    audioFileEntry = zip.file("ppt/" + audioFile);
                }
                if (!audioFileEntry) {
                } else {
                    uInt8ArrayAudio = audioFileEntry.asArrayBuffer();
                    blobAudio = new Blob([uInt8ArrayAudio]);
                    audioBlob = URL.createObjectURL(blobAudio);
                    const cx = parseInt(xfrmNode["a:ext"].attrs.cx) * 20;
                    const cy = xfrmNode["a:ext"].attrs.cy;
                    const x = parseInt(xfrmNode["a:off"].attrs.x) / 2.5;
                    const y = xfrmNode["a:off"].attrs.y;
                    audioObjc = {
                        "a:ext": { attrs: { cx, cy } },
                        "a:off": { attrs: { x, y } }
                    };
                    audioPlayerFlag = true;
                    mediaSupportFlag = true;
                    mediaPicFlag = true;
                }
            }
        }

        mimeType = PPTXUtils.getMimeType(imgFileExt);
        rtrnData = `<div class='block content' style='${(mediaProcess && audioPlayerFlag) ? getPosition(audioObjc, node, undefined, undefined) : getPosition(xfrmNode, node, undefined, undefined)}${(mediaProcess && audioPlayerFlag) ? getSize(audioObjc, undefined, undefined) : getSize(xfrmNode, undefined, undefined)} z-index: ${order};transform: rotate(${rotate}deg);'>`;

        if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
                rtrnData += "<img src='" + PPTXUtils.arrayBufferToBlobUrl(imgArrayBuffer, mimeType) + "' style='width: 100%; height: 100%'/>";
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
            }
            rtrnData += "</div>";
            //console.log(rtrnData)
            return rtrnData;
    }

    function processGraphicFrameNode(node, warpObj, source, sType) {
        // 使用 PPTXImageUtils 模块处理图形框架节点
        return PPTXImageUtils.processGraphicFrameNode(node, warpObj, source, sType, genTableInternal, genDiagram, processGroupSpNode);
    }





    // genGlobalCSS 已移至 PPTXCSSUtils 模块

    function genTableInternal(node, warpObj) {
        return PPTXTableUtils.genTableInternal(node, warpObj, styleTable);
    }


    function genDiagram(node, warpObj, source, sType) {
        const readXmlFileFunc = PPTXParser && PPTXParser.readXmlFile ? PPTXParser.readXmlFile : () => null;
        return PPTXDiagramUtils.genDiagram(node, warpObj, source, sType, readXmlFileFunc, getPosition, getSize, null);
    }

    function getBackground(warpObj, slideSize, index) {
        return PPTXBackgroundUtils.getBackground(warpObj, slideSize, index, processNodesInSlide);
    }

    function getSlideBackgroundFill(warpObj, index) {
        return PPTXBackgroundUtils.getSlideBackgroundFill(warpObj, index);
    }

    // 调用主处理函数并返回 Promise
    return convertToHtml(file);
}

// Enhanced API functions
async function parsePptx(file, options = {}) {
    const result = await pptxToHtml(file, options);
    
    // Return structured data as per new API specification
    return {
        html: result.html || result,
        slides: result.slides || [],
        size: result.slideSize || { width: 0, height: 0 },
        thumb: result.thumb || null,
        globalCSS: result.css || '',
        title: options.title || 'Presentation',
        author: options.author || 'Unknown',
        elements: []
    };
}

// Export for use in ES6 modules
export { pptxToHtml, parsePptx };