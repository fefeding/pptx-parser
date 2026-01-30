/**
 * pptx-parser - PPTX 文件解析为 HTML 的主入口
 * 支持 File、Blob 或 ArrayBuffer 输入
 */
// Import dependencies
import { PPTXConstants } from './core/constants';
import { PPTXUtils, PPTXFileReader } from './core/utils';
import { PPTXParser } from './parser';
import { PPTXHtml } from './html';
import { PPTXStyleManager } from './core/style-manager';
import { PPTXShapeUtils } from './shape/shape.js';
import { PPTXShapePropertyExtractor } from './shape/property-extractor.js';
import { PPTXShapeFillsUtils } from './shape/fills';
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
import { genShape as genShapeModule } from './shape/generator';
import JSZip from 'jszip';

interface PptxToHtmlOptions {
    container?: HTMLElement;
    onProgress?: (percent: number) => void;
    slidesScale?: string;
    mediaProcess?: boolean;
    themeProcess?: boolean | string;
    incSlide?: {
        width?: number;
        height?: number;
    };
}

interface ApiObject {
    isSlideMode: boolean;
    initSlideMode(): void;
    exitSlideMode(): void;
    updateProgress(percent: number): void;
}

/**
 * 主函数 - 将 PPTX 文件转换为 HTML
 * @param file - 输入文件
 * @param options - 配置选项
 * @returns 包含 html、css、slides 等的对象
 */
function pptxToHtml(file: File | Blob | ArrayBuffer, options: PptxToHtmlOptions = {}): Promise<any> {
    let isDone = false;
    const MsgQueue: any[] = [];
    (PPTXHtml as any).MsgQueue = MsgQueue;
    const slideFactor = PPTXConstants.SLIDE_FACTOR;
    const fontSizeFactor = PPTXConstants.FONT_SIZE_FACTOR;
    let slideWidth = 0;
    let slideHeight = 0;
    let isSlideMode = false;
    
    // API object for external control
    const api: ApiObject = {
        get isSlideMode() { return isSlideMode; },
        set isSlideMode(value: boolean) { isSlideMode = value; },
        initSlideMode() { initSlideMode(options.container!, options); },
        exitSlideMode() { exitSlideMode(options.container!, options); },
        updateProgress(percent: number) { updateProgressBar(percent); }
    };
    
    // 计算元素位置和尺寸 - 使用 PPTXUtils 中的函数
    const getPosition = PPTXUtils ? PPTXUtils.getPosition : () => "";
    const getSize = PPTXUtils ? PPTXUtils.getSize : () => "";
    let processFullTheme: boolean | string = true;
    const styleTable: any = {};
    
    /**
     * 深度合并对象
     */
    function deepExtend(destination: any, ...sources: any[]): any {
        for (let i = 0; i < sources.length; i++) {
            const source = sources[i];
            if (!source) continue;
            for (const property in source) {
                if (source.hasOwnProperty(property)) {
                    if (source[property] && typeof source[property] === 'object' && source[property].constructor === Object) {
                        destination[property] = destination[property] || {};
                        deepExtend(destination[property], source[property]);
                    }
                    else {
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
    function updateProgressBar(percent: number): void {
        if (settings.onProgress) {
            settings.onProgress(percent);
        }
    }
    
    /**
     * 转换为 HTML - 将 File/Blob/ArrayBuffer 转换为 ArrayBuffer
     */
    function convertToHtml(file: File | Blob | ArrayBuffer): Promise<any> {
        let fileArrayBuffer: ArrayBuffer;
        if (file instanceof File || (file as any) instanceof Blob) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = async (event) => {
                    fileArrayBuffer = event.target.result as ArrayBuffer;
                    await processZip(fileArrayBuffer, resolve, reject);
                };
                reader.onerror = (event) => {
                    reject(new Error("Failed to read file: " + (event.target as any).error));
                };
                reader.readAsArrayBuffer(file as File | Blob);
            });
        }
        else if (file instanceof ArrayBuffer) {
            return new Promise(async (resolve, reject) => {
                await processZip(file, resolve, reject);
            });
        }
        else {
            return Promise.reject(new Error("Invalid file type: must be File, Blob, or ArrayBuffer"));
        }
    }
    
    /**
     * 处理 ZIP 文件 - 解析 PPTX 并生成 HTML
     */
    async function processZip(fileArrayBuffer: ArrayBuffer, resolve: (value: any) => void, reject: (reason?: any) => void) {
        if (fileArrayBuffer.byteLength < 10) {
            reject(new Error("Invalid file: too small"));
            return;
        }
        try {
            // 使用异步方式加载ZIP，但预先缓存所有文件内容以模拟同步访问
            const zip = await JSZip.loadAsync(fileArrayBuffer);
            
            // 配置 PPTXParser 模块 - 传递必要的回调函数
            (PPTXParser as any).configure({
                ...settings,
                processNodesInSlide,
                getBackground,
                getSlideBackgroundFill
            });
            
        const rslt_ary = await (PPTXParser as any).processPPTX(zip);
            
            // 收集生成的 HTML、CSS 和数据
            const result: any = {
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
                        if (typeof (PPTXCSSUtils as any).genGlobalCSS === 'function') {
                            result.css += (PPTXCSSUtils as any).genGlobalCSS(styleTable, settings, slideWidth);
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
        } catch (error) {
            reject(error);
        }
    }
    
    function initSlideMode(divId: HTMLElement | undefined, settings: any): void {
        // Implementation
    }
    
    function exitSlideMode(divId: HTMLElement | undefined, settings: any): void {
        // Implementation
    }
    
    interface ProcessNodeHandlers {
        processSpNode: (node: any, pNode: any, warpObj: any, source: any, sType: any) => string | Promise<string>;
        processCxnSpNode: (node: any, pNode: any, warpObj: any, source: any, sType: any) => string | Promise<string>;
        processPicNode: (node: any, warpObj: any, source: any, sType: any) => string | Promise<string>;
        processGraphicFrameNode: (node: any, warpObj: any, source: any, sType: any) => string | Promise<string>;
        processGroupSpNode: (node: any, warpObj: any, source: any) => string | Promise<string>;
    }
    
    async function processNodesInSlide(nodeKey: string, nodeValue: any, nodes: any, warpObj: any, source: any, sType: any): Promise<string> {
        const handlers: ProcessNodeHandlers = {
            processSpNode: async (node: any, pNode: any, warpObj: any, source: any, sType: any) => {
                return await processSpNodeModule(node, pNode, warpObj, source, sType, async (n: any, pn: any, sLSN: any, sMSN: any, id: any, nm: any, idx: any, typ: any, ord: any, wo: any, uDBg: any, sTy: any, src: any) => {
                    return await genShapeModule(n, pn, sLSN, sMSN, id, nm, idx, typ, ord, wo, uDBg, sTy, src, styleTable);
                });
            },
            processCxnSpNode: async (node: any, pNode: any, warpObj: any, source: any, sType: any) => {
                return await processCxnSpNodeModule(node, pNode, warpObj, source, sType, async (n: any, pn: any, sLSN: any, sMSN: any, id: any, nm: any, idx: any, typ: any, ord: any, wo: any, uDBg: any, sTy: any, src: any) => {
                    return await genShapeModule(n, pn, undefined, undefined, id, nm, idx, typ, ord, wo, undefined, sTy, src, styleTable);
                });
            },
            processPicNode,
            processGraphicFrameNode,
            processGroupSpNode
        };
        return await (PPTXNodeUtils as any).processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType, handlers);
    }
    
    async function processGroupSpNode(node: any, warpObj: any, source: any): Promise<string> {
        return await (PPTXNodeUtils as any).processGroupSpNode(node, warpObj, source, slideFactor, processNodesInSlide);
    }
    
    // processSpNode 和 processCxnSpNode 已移至 pptx-shape-node-processor.js 模块
    // 这些函数现在通过 processNodesInSlide 中的 handlers 调用
    async function processPicNode(node: any, warpObj: any, source: any, sType: any): Promise<string> {
        let rtrnData = "";
        let mediaPicFlag = false;
        const order = node.attrs.order;
        const rid = node["p:blipFill"]["a:blip"].attrs["r:embed"];
        let resObj: any;
        if (source === "slideMasterBg") {
            resObj = warpObj.masterResObj;
        }
        else if (source === "slideLayoutBg") {
            resObj = warpObj.layoutResObj;
        }
        else {
            resObj = warpObj.slideResObj;
        }
        const imgName = resObj[rid].target;
        const imgFileExt = (PPTXUtils as any).extractFileExtension(imgName).toLowerCase();
        const zip = warpObj.zip;
        let imgFile = zip.file(imgName);
        if (!imgFile && !imgName.startsWith("ppt/")) {
            imgFile = zip.file("ppt/" + imgName);
        }
        if (!imgFile) {
            return "";
        }
        const imgArrayBuffer = await imgFile.async('arraybuffer');
        let xfrmNode = node["p:spPr"]["a:xfrm"];
        if (xfrmNode === undefined) {
            const idx = (PPTXUtils as any).getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "idx"]);
            const type = (PPTXUtils as any).getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "type"]);
            if (idx !== undefined) {
                xfrmNode = (PPTXUtils as any).getTextByPathList(warpObj.slideLayoutTables, ["idxTable", idx, "p:spPr", "a:xfrm"]);
            }
        }
        let rotate = 0;
        const rotateNode = (PPTXUtils as any).getTextByPathList(node, ["p:spPr", "a:xfrm", "attrs", "rot"]);
        if (rotateNode !== undefined) {
            rotate = (PPTXUtils as any).angleToDegrees(rotateNode);
        }
        // 视频处理
        const vdoNode = (PPTXUtils as any).getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:videoFile"]);
        let vdoRid: string, vdoFile: string, vdoFileExt: string, vdoMimeType: string, uInt8Array: ArrayBuffer, blob: Blob, vdoBlob: string, mediaSupportFlag = false, isVdeoLink = false;
        const mediaProcess = settings.mediaProcess;
        if (vdoNode !== undefined && mediaProcess) {
            vdoRid = vdoNode.attrs["r:link"];
            vdoFile = resObj[vdoRid].target;
            const checkIfLink = (PPTXUtils as any).IsVideoLink(vdoFile);
            if (checkIfLink) {
                vdoFile = (PPTXUtils as any).escapeHtml(vdoFile);
                isVdeoLink = true;
                mediaSupportFlag = true;
                mediaPicFlag = true;
            }
            else {
                    vdoFileExt = (PPTXUtils as any).extractFileExtension(vdoFile).toLowerCase();
                    if (vdoFileExt === "mp4" || vdoFileExt === "webm" || vdoFileExt === "ogg") {
                        let vdoFileEntry = zip.file(vdoFile);
                        if (!vdoFileEntry && !vdoFile.startsWith("ppt/")) {
                            vdoFileEntry = zip.file("ppt/" + vdoFile);
                        }
                        if (!vdoFileEntry) {
                            // File not found
                        }
                        else {
                            uInt8Array = await vdoFileEntry.async('arraybuffer');
                        vdoMimeType = (PPTXUtils as any).getMimeType(vdoFileExt);
                        blob = new Blob([uInt8Array], { type: vdoMimeType });
                        vdoBlob = URL.createObjectURL(blob);
                        mediaSupportFlag = true;
                        mediaPicFlag = true;
                    }
                }
            }
        }
        // 音频处理
        const audioNode = (PPTXUtils as any).getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:audioFile"]);
        let audioRid: string, audioFile: string, audioFileExt: string, audioMimeType: string, uInt8ArrayAudio: ArrayBuffer, blobAudio: Blob, audioBlob: string;
        let audioPlayerFlag = false;
        let audioObjc: any;
        if (audioNode !== undefined && mediaProcess) {
            audioRid = audioNode.attrs["r:link"];
            audioFile = resObj[audioRid].target;
            audioFileExt = (PPTXUtils as any).extractFileExtension(audioFile).toLowerCase();
            if (audioFileExt === "mp3" || audioFileExt === "wav" || audioFileExt === "ogg") {
                let audioFileEntry = zip.file(audioFile);
                if (!audioFileEntry && !audioFile.startsWith("ppt/")) {
                    audioFileEntry = zip.file("ppt/" + audioFile);
                }
                if (!audioFileEntry) {
                    // File not found
                }
                else {
                    uInt8ArrayAudio = await audioFileEntry.async('arraybuffer');
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
        const mimeType = (PPTXUtils as any).getMimeType(imgFileExt);
        rtrnData = `<div class='block content' style='${(mediaProcess && audioPlayerFlag) ? getPosition(audioObjc, node, undefined, undefined, undefined) : getPosition(xfrmNode, node, undefined, undefined, undefined)}${(mediaProcess && audioPlayerFlag) ? getSize(audioObjc, undefined, undefined) : getSize(xfrmNode, undefined, undefined)} z-index: ${order};transform: rotate(${rotate}deg);'>`;
        if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
            rtrnData += "<img src='" + (PPTXUtils as any).arrayBufferToBlobUrl(imgArrayBuffer, mimeType) + "' style='width: 100%; height: 100%'/>";
        }
        else if ((vdoNode !== undefined || audioNode !== undefined) && mediaProcess && mediaSupportFlag) {
            if (vdoNode !== undefined && !isVdeoLink) {
                rtrnData += "<video  src='" + vdoBlob + "' controls style='width: 100%; height: 100%'>Your browser does not support the video tag.</video>";
            }
            else if (vdoNode !== undefined && isVdeoLink) {
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
            // Media process disabled
        }
        rtrnData += "</div>";
        return rtrnData;
    }
    
    async function processGraphicFrameNode(node: any, warpObj: any, source: any, sType: any): Promise<string> {
        // 使用 PPTXImageUtils 模块处理图形框架节点
        // 需要将此函数设为异步，以便处理异步的图表生成
        return await (PPTXImageUtils as any).processGraphicFrameNode(node, warpObj, source, sType, genTableInternal, genDiagram, processGroupSpNode);
    }
    
    // genGlobalCSS 已移至 PPTXCSSUtils 模块
    async function genTableInternal(node: any, warpObj: any): Promise<string> {
        return await (PPTXTableUtils as any).genTableInternal(node, warpObj, styleTable);
    }
    
    async function genDiagram(node: any, warpObj: any, source: any, sType: any): Promise<string> {
        const readXmlFileFunc = PPTXParser && (PPTXParser as any).readXmlFile ? (PPTXParser as any).readXmlFile : () => null;
        return await (PPTXDiagramUtils as any).genDiagram(node, warpObj, source, sType, readXmlFileFunc, getPosition, getSize, null);
    }
    
    async function getBackground(warpObj: any, slideSize: any, index: number): Promise<string> {
        return await (PPTXBackgroundUtils as any).getBackground(warpObj, slideSize, index, processNodesInSlide);
    }
    
    async function getSlideBackgroundFill(warpObj: any, index: number): Promise<string> {
        return await (PPTXBackgroundUtils as any).getSlideBackgroundFill(warpObj, index);
    }
    
    // 调用主处理函数并返回 Promise
    return convertToHtml(file);
}

interface ParsePptxResult {
    html: string;
    css: string;
    slides: any[];
    size: {
        width: number;
        height: number;
    };
    globalCSS: string;
    chartQueue: any[];
}

/**
 * 解析PPTX文件并返回结构化数据
 * @param file - 输入文件 (File | Blob | ArrayBuffer)
 * @returns 包含HTML、CSS、幻灯片数据和尺寸的对象
 */
async function parsePptx(file: File | Blob | ArrayBuffer): Promise<ParsePptxResult> {
    const result = await pptxToHtml(file);
    
    return {
        html: result.html,
        css: result.css,
        slides: result.slides,
        size: result.slideSize,
        globalCSS: result.css,
        chartQueue: result.chartQueue
    };
}

// Export for use in ES6 modules
export { pptxToHtml, parsePptx, PPTXHtml };
