/**
 * pptxjs.js
 * Ver. : 1.21.1
 * last update: 16/11/2021
 * Author: meshesha , https://github.com/meshesha
 * LICENSE: MIT
 * url:https://pptx.js.org/
 * fix issues:
 * [#16](https://github.com/meshesha/PPTXjs/issues/16)
 *
 * New:
 *  - supports both jQuery and vanilla JavaScript
 */



// Import dependencies
import { PPTXConstants } from './core/pptx-constants.js';
import { PPTXUtils, PPTXFileReader } from './utils/utils.js';
import { PPTXParser } from './pptx-parser.js';
import { PPTXHtml } from './pptx-html.js';
import { PPTXStyleManager } from './core/pptx-style-manager.js';
import { PPTXShapeUtils } from './shape/pptx-shape-utils.js';
import { PPTXShapePropertyExtractor } from './shape/pptx-shape-property-extractor.js';
import { PPTXShapeFillsUtils } from './shape/pptx-shape-fills-utils.js'
import { PPTXBasicShapes } from './shape/pptx-basic-shapes.js';
import { PPTXNodeUtils } from './node/pptx-node-utils.js';
import { PPTXBackgroundUtils } from './core/pptx-background-utils.js';
import { PPTXImageUtils } from './image/pptx-image-utils.js';
import { PPTXDiagramUtils } from './diagram/pptx-diagram-utils.js';
import { TextUtils } from './text/text-utils.js';
import { PPTXUIUtils } from './ui/pptx-ui-utils.js';
import { PPTXCSSUtils } from './core/pptx-css-utils.js';
import { PPTXColorUtils } from './core/pptx-color-utils.js';
import { PPTXTableUtils } from './table/pptx-table-utils.js';
import { PPTXTextStyleUtils } from './text/pptx-text-style-utils.js';
import { PPTXTextElementUtils } from './text/pptx-text-element-utils.js';
import { PPTXShapeContainer } from './shape/pptx-shape-container.js';
import { PPTXStarShapes } from './shape/pptx-star-shapes.js';
import { PPTXCalloutShapes } from './shape/pptx-callout-shapes.js';
import { PPTXArrowShapes } from './shape/pptx-arrow-shapes.js';
import { PPTXMathShapes } from './shape/pptx-math-shapes.js';
import { initSlideMode as initSlideModeModule, exitSlideMode as exitSlideModeModule } from './ui/pptx-slide-mode.js';
import { processSpNode as processSpNodeModule, processCxnSpNode as processCxnSpNodeModule } from './node/pptx-shape-node-processor.js';
import { genShape as genShapeModule } from './shape/pptx-shape-generator.js';


        //var slideLayoutClrOvride = "";
        var defaultTextStyle = null;

        var chartID = 0;

        var _order = 1;

        var app_verssion ;


    // Main pptxToHtml function
    function pptxToHtml(element, options) {
        //var worker;
        var $result = typeof element === 'string' ? document.querySelector(element) : (element && element.jquery ? element[0] : element);
        var divId = element.id || element.getAttribute("id");

        var isDone = false;

        var MsgQueue = new Array();
        PPTXHtml.MsgQueue = MsgQueue;

        //var slideLayoutClrOvride = "";

        var defaultTextStyle = null;

        var chartID = 0;

        var _order = 1;

        var app_verssion ;

        var slideFactor = PPTXConstants.SLIDE_FACTOR;
        var fontSizeFactor = PPTXConstants.FONT_SIZE_FACTOR;
        ////////////////////// 
        var slideWidth = 0;
    var slideHeight = 0;
    var isSlideMode = false;
    
    // API object for external control
    var api = {
        get isSlideMode() { return isSlideMode; },
        set isSlideMode(value) { isSlideMode = value; },
        initSlideMode: function() { initSlideMode(divId, settings); },
        exitSlideMode: function() { exitSlideMode(divId); },
        updateProgress: function(percent) { updateProgressBar(percent); },
        removeLoading: function() { PPTXUIUtils.removeLoadingMessage(); }
    };

    // 计算元素位置和尺寸 - 使用 PPTXUtils 中的函数
    var getPosition = PPTXUtils ? PPTXUtils.getPosition : function() { return ""; };
    var getSize = PPTXUtils ? PPTXUtils.getSize : function() { return ""; };

    var processFullTheme = true;
        var styleTable = {};
        
        // Deep extend function
        function deepExtend(destination) {
            for (var i = 1; i < arguments.length; i++) {
                var source = arguments[i];
                for (var property in source) {
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
        
        var settings = deepExtend({
            // These are the defaults.
            pptxFileUrl: "",
            fileInputId: "",
            slidesScale: "", //Change Slides scale by percent
            slideMode: false, /** true,false - enable slideshow mode */
            slideType: "divs2slidesjs",  /*'divs2slidesjs' (default), 'revealjs' - slideshow engine */
            revealjsPath: "./revealjs/reveal.js", /* path to reveal.js library */
            mediaProcess: true, /** true,false: if true then process video and audio files */
            jsZipV2: false,
            themeProcess: true, /*true (default) , false, "colorsAndImageOnly"*/
            incSlide:{
                width: 0,
                height: 0
            },
            slideModeConfig: {
                first: 1,
                nav: true, /** true,false : show or not nav buttons*/
                navTxtColor: "black", /** color */
                keyBoardShortCut: true, /** true,false ,condition: */
                showSlideNum: true, /** true,false */
                showTotalSlideNum: true, /** true,false */
                autoSlide: true, /** false or seconds , F8 to active ,keyBoardShortCut: true */
                randomAutoSlide: false, /** true,false ,autoSlide:true */
                loop: false,  /** true,false */
                background: false, /** false or color*/
                transition: "default", /** transition type: "slid","fade","default","random" , to show transition efects :transitionTime > 0.5 */
                transitionTime: 1 /** transition time between slides in seconds */
            },
            revealjsConfig: {}
        }, options);

        processFullTheme = settings.themeProcess;

        var container = document.getElementById(divId);
        if (Array.isArray(container)) {
            container = container[0];
        }
            if (container) {
            var loadingMsg = document.createElement("div");
            loadingMsg.className = "slides-loadnig-msg";
            loadingMsg.style.display = "block";
            loadingMsg.style.width = "100%";
            loadingMsg.style.color = "white";
            loadingMsg.style.backgroundColor = "#ddd";

            var progressBar = document.createElement("div");
            progressBar.className = "slides-loading-progress-bar";
            progressBar.style.width = "1%";
            progressBar.style.backgroundColor = "#4775d1";
            progressBar.innerHTML = "<span style='text-align: center;'>Loading... (1%)</span>";

            loadingMsg.appendChild(progressBar);
            container.prepend(loadingMsg);
        } else {
            console.warn("Container not found for loading message");
        }

        // 动态加载脚本已改为通过配置提供，调用方应自行加载依赖
        // 这里保留向后兼容性，但发出警告
        if (settings.slideMode && typeof window.pptxjslideObj === 'undefined') {
            console.warn('Slide mode is enabled but divs2slides.js is not loaded. Please load it manually or use the appropriate configuration.');
        }
        if (settings.jsZipV2 !== false) {
            console.warn('jsZipV2 option is deprecated. Please configure JSZip properly in your build system or load it explicitly.');
        }

        if (settings.keyBoardShortCut && $result) {
            var keyHandler = function(event) {
                if (event.key === 'F5') {
                    event.preventDefault();
                    if (!isSlideMode) {
                        api.initSlideMode();
                    } else {
                        api.exitSlideMode();
                    }
                }
            };
            document.addEventListener("keydown", keyHandler);
        }
        
        if (settings.pptxFileUrl != "") {
            try{
                JSZipUtils.getBinaryContent(settings.pptxFileUrl, function (err, content) {
                    var blob = new Blob([content]);
                    var file_name = settings.pptxFileUrl;
                    var fArry = file_name.split(".");
                    fArry.pop();
                    blob.name = fArry[0];
                    PPTXFileReader.setupBlob(blob, {
                        on: {
                            load: function (arrayBuffer) {
                                convertToHtml(arrayBuffer);
                            }
                        }
                    });
                });
            }catch(e){
                console.error("file url error (" + settings.pptxFileUrl+ "0)")
                PPTXUIUtils.removeLoadingMessage();
            }
        } else {
            PPTXUIUtils.removeLoadingMessage();
        }
        if (settings.fileInputId != "") {
            document.getElementById(settings.fileInputId).addEventListener("change", function (evt) {
                $result.innerHTML = "";
                var file = evt.target.files[0];
                // var fileName = file[0].name;
                //var fileSize = file[0].size;
                var fileType = file.type;
                if (fileType == "application/vnd.openxmlformats-officedocument.presentationml.presentation") {
                    PPTXFileReader.setupBlob(file, {
                        on: {
                            load: function (arrayBuffer) {
                                convertToHtml(arrayBuffer);
                            }
                        }
                    });
                } else {
                    alert("This is not pptx file");
                }
            });
        } else {
            console.warn("fileInputId not provided, file upload listener not attached");
        }

        function updateProgressBar(percent) {
            if (options.onProgress) {
                options.onProgress(percent);
            }
            // 同时更新 PPTXUIUtils 中的进度条（向后兼容）
            PPTXUIUtils.updateProgressBar(percent);
        }

        function convertToHtml(file) {
            //'use strict';
            //console.log("file", file, "size:", file.byteLength);
            if (file.byteLength < 10){
                console.error("file url error (" + settings.pptxFileUrl + "0)")
                PPTXUIUtils.removeLoadingMessage();
                return;
            }
            var zip = new JSZip(), s;
            //if (typeof file === 'string') { // Load
            zip = zip.load(file);  //zip.load(file, { base64: true });

            // 配置 PPTXParser 模块 - 传递必要的回调函数
            PPTXParser.configure({
                ...settings,
                processNodesInSlide: processNodesInSlide,
                getBackground: getBackground,
                getSlideBackgroundFill: getSlideBackgroundFill
            });

            var rslt_ary = PPTXParser.processPPTX(zip);

            // 收集生成的 HTML、CSS 和数据
            var result = {
                html: "",
                css: "",
                slides: [],
                slideSize: null,
                chartQueue: []
            };

            //s = readXmlFile(zip, 'ppt/tableStyles.xml');
            //var slidesHeight = $("#" + divId + " .slide").height();
            for (var i = 0; i < rslt_ary.length; i++) {
                switch (rslt_ary[i]["type"]) {
                    case "slide":
                        result.html += rslt_ary[i]["data"];
                        result.slides.push(rslt_ary[i]["data"]);
                        break;
                    case "pptx-thumb":
                        //$("#pptx-thumb").attr("src", "data:image/jpeg;base64," +rslt_ary[i]["data"]);
                        break;
                    case "slideSize":
                        slideWidth = rslt_ary[i]["data"].width;
                        slideHeight = rslt_ary[i]["data"].height;
                        result.slideSize = rslt_ary[i]["data"];
                        /*
                        $("#"+divId).css({
                            'width': slideWidth + 80,
                            'height': slideHeight + 60
                        });
                        */
                        break;
                    case "globalCSS":
                        //console.log(rslt_ary[i]["data"])
                        result.css += rslt_ary[i]["data"];
                        break;
                    case "ExecutionTime":
                        // 生成并添加全局 CSS
                        if (typeof PPTXCSSUtils.genGlobalCSS === 'function') {
                            result.css += PPTXCSSUtils.genGlobalCSS(styleTable, settings, slideWidth);
                        }
                        result.chartQueue = MsgQueue.slice(); // 复制图表队列

                        // 如果调用方提供了 DOM 容器，则插入（向后兼容）
                        if ($result && typeof $result.insertAdjacentHTML === 'function') {
                            $result.insertAdjacentHTML('beforeend', result.html);
                            $result.insertAdjacentHTML('beforeend', "<style>" + result.css + "</style>");
                            PPTXHtml.processMsgQueue(MsgQueue);
                            PPTXHtml.setNumericBullets($result.querySelectorAll ? $result.querySelectorAll(".block") : document.querySelectorAll(".block"));
                            PPTXHtml.setNumericBullets($result.querySelectorAll ? $result.querySelectorAll("table td") : document.querySelectorAll("table td"));

                            if (!settings.slideMode || (settings.slideMode && settings.slideType == "revealjs")) {
                                PPTXUIUtils.getSlidesWrapper(divId);

                                if (settings.slideMode && settings.slideType == "revealjs") {
                                    PPTXUIUtils.addRevealClass(divId);
                                }
                            }

                            PPTXUIUtils.updateWrapperHeight(divId, settings.slidesScale, false, settings.slideType, null);
                        }

                        isDone = true;

                        if (settings.slideMode && !isSlideMode) {
                            isSlideMode = true;
                            initSlideMode(divId, settings);
                        } else if (!settings.slideMode) {
                            PPTXUIUtils.removeLoadingMessage();
                        }
                        break;
                    case "progress-update":
                        //console.log(rslt_ary[i]["data"]); //update progress bar
                        updateProgressBar(rslt_ary[i]["data"])
                        break;
                    default:
                }
            }

            // 返回结果对象供调用方使用
            return result;
        }

        function initSlideMode(divId, settings) {
            // 使用外部模块的 initSlideMode 函数
            return initSlideModeModule(divId, settings, PPTXUIUtils.updateWrapperHeight.bind(PPTXUIUtils));
        }

        function exitSlideMode(divId) {
            // 使用外部模块的 exitSlideMode 函数
            return exitSlideModeModule(divId, settings, PPTXUIUtils.updateWrapperHeight.bind(PPTXUIUtils));
        }


        function processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType) {
            // 使用 PPTXNodeUtils 模块处理节点
            // processSpNode, processCxnSpNode 现在从外部模块导入
            var handlers = {
                processSpNode: function(node, pNode, warpObj, source, sType) {
                    return processSpNodeModule(node, pNode, warpObj, source, sType, function(n, pn, sLSN, sMSN, id, nm, idx, typ, ord, wo, uDBg, sTy, src) {
                        return genShapeModule(n, pn, sLSN, sMSN, id, nm, idx, typ, ord, wo, uDBg, sTy, src, styleTable);
                    });
                },
                processCxnSpNode: function(node, pNode, warpObj, source, sType) {
                    return processCxnSpNodeModule(node, pNode, warpObj, source, sType, function(n, pn, sLSN, sMSN, id, nm, idx, typ, ord, wo, uDBg, sTy, src) {
                        return genShapeModule(n, pn, undefined, undefined, id, nm, idx, typ, ord, wo, undefined, sTy, src, styleTable);
                    });
                },
                processPicNode: processPicNode,
                processGraphicFrameNode: processGraphicFrameNode,
                processGroupSpNode: processGroupSpNode
            };
            return PPTXNodeUtils.processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType, handlers);
        }

        function processGroupSpNode(node, warpObj, source) {
            // 使用 PPTXNodeUtils 模块处理组节点
            return PPTXNodeUtils.processGroupSpNode(node, warpObj, source, slideFactor, processNodesInSlide);
        }

        // processSpNode 和 processCxnSpNode 已移至 pptx-shape-node-processor.js 模块
        // 这些函数现在通过 processNodesInSlide 中的 handlers 调用

        






        function processPicNode(node, warpObj, source, sType) {
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
            var imgName = resObj[rid]["target"];

            //console.log("processPicNode imgName:", imgName);
            var imgFileExt =PPTXUtils.extractFileExtension(imgName).toLowerCase();
            var zip = warpObj["zip"];
            // 尝试解析图片路径，处理相对路径问题
            var imgFile = zip.file(imgName);
            if (!imgFile && !imgName.startsWith("ppt/")) {
                // 尝试添加 ppt/ 前缀
                imgFile = zip.file("ppt/" + imgName);
            }
            if (!imgFile) {
                // 如果仍然找不到，记录错误并返回空字符串
                console.error("Image file not found:", imgName);
                return "";
            }
            var imgArrayBuffer = imgFile.asArrayBuffer();
            var mimeType = "";
            var xfrmNode = node["p:spPr"]["a:xfrm"];
            if (xfrmNode === undefined) {
                var idx = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "idx"]);
                var type = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "type"]);
                if (idx !== undefined) {
                    xfrmNode = PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"], ["idxTable", idx, "p:spPr", "a:xfrm"]);
                }
            }
            ///////////////////////////////////////Amir//////////////////////////////
            var rotate = 0;
            var rotateNode = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:xfrm", "attrs", "rot"]);
            if (rotateNode !== undefined) {
                rotate = PPTXUtils.angleToDegrees(rotateNode);
            }
            //video
            var vdoNode = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:videoFile"]);
            var vdoRid, vdoFile, vdoFileExt, vdoMimeType, uInt8Array, blob, vdoBlob, mediaSupportFlag = false, isVdeoLink = false;
            var mediaProcess = settings.mediaProcess;
            if (vdoNode !== undefined & mediaProcess) {
                vdoRid = vdoNode["attrs"]["r:link"];
                vdoFile = resObj[vdoRid]["target"];
                var checkIfLink = PPTXUtils.IsVideoLink(vdoFile);
                if (checkIfLink) {
                    vdoFile = PPTXUtils.escapeHtml(vdoFile);
                    //vdoBlob = vdoFile;
                    isVdeoLink = true;
                    mediaSupportFlag = true;
                    mediaPicFlag = true;
                } else {
                    vdoFileExt =PPTXUtils.extractFileExtension(vdoFile).toLowerCase();
                    if (vdoFileExt == "mp4" || vdoFileExt == "webm" || vdoFileExt == "ogg") {
                        // 尝试解析视频路径，处理相对路径问题
                        var vdoFileEntry = zip.file(vdoFile);
                        if (!vdoFileEntry && !vdoFile.startsWith("ppt/")) {
                            // 尝试添加 ppt/ 前缀
                            vdoFileEntry = zip.file("ppt/" + vdoFile);
                        }
                        if (!vdoFileEntry) {
                            // 如果仍然找不到，记录错误并跳过
                            console.error("Video file not found:", vdoFile);
                        } else {
                            uInt8Array = vdoFileEntry.asArrayBuffer();
                            vdoMimeType = PPTXUtils.getMimeType(vdoFileExt);
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
            var audioNode = PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:audioFile"]);
            var audioRid, audioFile, audioFileExt, audioMimeType, uInt8ArrayAudio, blobAudio, audioBlob;
            var audioPlayerFlag = false;
            var audioObjc;
            if (audioNode !== undefined & mediaProcess) {
                audioRid = audioNode["attrs"]["r:link"];
                audioFile = resObj[audioRid]["target"];
                audioFileExt =PPTXUtils.extractFileExtension(audioFile).toLowerCase();
                if (audioFileExt == "mp3" || audioFileExt == "wav" || audioFileExt == "ogg") {
                    // 尝试解析音频路径，处理相对路径问题
                    var audioFileEntry = zip.file(audioFile);
                    if (!audioFileEntry && !audioFile.startsWith("ppt/")) {
                        // 尝试添加 ppt/ 前缀
                        audioFileEntry = zip.file("ppt/" + audioFile);
                    }
                    if (!audioFileEntry) {
                        // 如果仍然找不到，记录错误并跳过
                        console.error("Audio file not found:", audioFile);
                    } else {
                        uInt8ArrayAudio = audioFileEntry.asArrayBuffer();
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
            //console.log(node)
            //////////////////////////////////////////////////////////////////////////
            mimeType = PPTXUtils.getMimeType(imgFileExt);
            rtrnData = "<div class='block content' style='" +
                ((mediaProcess && audioPlayerFlag) ? getPosition(audioObjc, node, undefined, undefined) : getPosition(xfrmNode, node, undefined, undefined)) +
                ((mediaProcess && audioPlayerFlag) ? getSize(audioObjc, undefined, undefined) : getSize(xfrmNode, undefined, undefined)) +
                " z-index: " + order + ";" +
                "transform: rotate(" + rotate + "deg);'>";
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
                console.log("Founded supported media file but media process disabled (mediaProcess=false)");
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
            var readXmlFileFunc = PPTXParser && PPTXParser.readXmlFile ? PPTXParser.readXmlFile : function() { return null; };
            // processSpNode is defined inside processNodesInSlide handlers, use a placeholder here
            return PPTXDiagramUtils.genDiagram(node, warpObj, source, sType, readXmlFileFunc, getPosition, getSize, null);
        }



        function getBackground(warpObj, slideSize, index) {
            // 使用 PPTXBackgroundUtils 模块处理背景
            return PPTXBackgroundUtils.getBackground(warpObj, slideSize, index, processNodesInSlide);
        }
        function getSlideBackgroundFill(warpObj, index) {
            // 使用 PPTXBackgroundUtils 模块处理背景填充
            return PPTXBackgroundUtils.getSlideBackgroundFill(warpObj, index);
        }
        // getBgGradientFill, getBgPicFill 已移至 PPTXBackgroundUtils 模块



        return api;
    }

// Export for use in ES6 modules
export { pptxToHtml };

// Also export to global scope for backward compatibility
// window.pptxToHtml = pptxToHtml; // Removed for ES modules
