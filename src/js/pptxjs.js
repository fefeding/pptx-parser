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

(function () {

    var $ = window.jQuery;
    var PPTXUtils = window.PPTXUtils || {};
    var PPTXParser = window.PPTXParser || {};
    var PPTXHtml = window.PPTXHtml || {};
    
    // Wait for PPTXShapeUtils to be available
    function ensurePPTXShapeUtils() {
        if (window.PPTXShapeUtils && typeof window.PPTXShapeUtils.shapeSnipRoundRect === 'function') {
            return window.PPTXShapeUtils;
        }
        // If not available, check if it's defined but empty, or wait a bit
        if (window.PPTXShapeUtils === undefined) {
            window.PPTXShapeUtils = {};
        }
        return window.PPTXShapeUtils;
    }
    
    // Ensure PPTXShapeUtils is available
    var PPTXShapeUtils = ensurePPTXShapeUtils();

    // 确保所有依赖模块已加载
    var PPTXNodeUtils = window.PPTXNodeUtils || {};
    var PPTXBackgroundUtils = window.PPTXBackgroundUtils || {};
    var PPTXImageUtils = window.PPTXImageUtils || {};
    var PPTXDiagramUtils = window.PPTXDiagramUtils || {};
    var TextUtils = window.TextUtils || {};

        //var slideLayoutClrOvride = "";
        var defaultTextStyle = null;

        var chartID = 0;

        var _order = 1;

        var app_verssion ;


    // Main pptxToHtml function
    function pptxToHtml(element, options) {
        //var worker;
        var $result = $(element);
        var divId = element.id || element.getAttribute("id");

        var isDone = false;

        var MsgQueue = new Array();

        //var slideLayoutClrOvride = "";

        var defaultTextStyle = null;

        var chartID = 0;

        var _order = 1;

        var app_verssion ;

        var slideFactor = window.PPTXConstants.SLIDE_FACTOR;
        var fontSizeFactor = window.PPTXConstants.FONT_SIZE_FACTOR;
        ////////////////////// 
        var slideWidth = 0;
    var slideHeight = 0;
    var isSlideMode = false;

    // 计算元素位置和尺寸 - 使用 PPTXUtils 中的函数
    var getPosition = window.PPTXUtils ? window.PPTXUtils.getPosition : function() { return ""; };
    var getSize = window.PPTXUtils ? window.PPTXUtils.getSize : function() { return ""; };

    var processFullTheme = true;
        var styleTable = {};
        var settings = $.extend(true, {
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

        var container = $("#" + divId);
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
        }
        if (settings.slideMode) {
            if (!window.pptxjslideObj) {
                $.getScript('./js/divs2slides.js');
            }
        }
        if (settings.jsZipV2 !== false) {
            $.getScript(settings.jsZipV2);
            if (localStorage.getItem('isPPTXjsReLoaded') !== 'yes') {
                localStorage.setItem('isPPTXjsReLoaded', 'yes');
                location.reload();
            }
        }

        if (settings.keyBoardShortCut) {
            $(document).bind("keydown", function (event) {
                event.preventDefault();
                var key = event.keyCode;
                console.log(key, isDone)
                if (key == 116 && !isSlideMode) { //F5
                    isSlideMode = true;
                    initSlideMode(divId, settings);
                } else if (key == 116 && isSlideMode) { //F5 again - exit slide mode
                    isSlideMode = false;
                    exitSlideMode(divId);
                }
            });
        }
        FileReaderJS.setSync(false);
        if (settings.pptxFileUrl != "") {
            try{
                JSZipUtils.getBinaryContent(settings.pptxFileUrl, function (err, content) {
                    var blob = new Blob([content]);
                    var file_name = settings.pptxFileUrl;
                    var fArry = file_name.split(".");
                    fArry.pop();
                    blob.name = fArry[0];
                    FileReaderJS.setupBlob(blob, {
                        readAsDefault: "ArrayBuffer",
                        on: {
                            load: function (e, file) {
                                //console.log(e.target.result);
                                convertToHtml(e.target.result);
                            }
                        }
                    });
                });
            }catch(e){
                console.error("file url error (" + settings.pptxFileUrl+ "0)")
                window.PPTXUIUtils.removeLoadingMessage();
            }
        } else {
            window.PPTXUIUtils.removeLoadingMessage();
        }
        if (settings.fileInputId != "") {
            $("#" + settings.fileInputId).on("change", function (evt) {
                $result.html("");
                var file = evt.target.files[0];
                // var fileName = file[0].name;
                //var fileSize = file[0].size;
                var fileType = file.type;
                if (fileType == "application/vnd.openxmlformats-officedocument.presentationml.presentation") {
                    FileReaderJS.setupBlob(file, {
                        readAsDefault: "ArrayBuffer",
                        on: {
                            load: function (e, file) {
                                //console.log(e.target.result);
                                convertToHtml(e.target.result);
                            }
                        }
                    });
                } else {
                    alert("This is not pptx file");
                }
            });
        }

        function updateProgressBar(percent) {
            window.PPTXUIUtils.updateProgressBar(percent);
        }

        function convertToHtml(file) {
            //'use strict';
            //console.log("file", file, "size:", file.byteLength);
            if (file.byteLength < 10){
                console.error("file url error (" + settings.pptxFileUrl + "0)")
                window.PPTXUIUtils.removeLoadingMessage();
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

            var rslt_ary = processPPTX(zip);
            //s = readXmlFile(zip, 'ppt/tableStyles.xml');
            //var slidesHeight = $("#" + divId + " .slide").height();
            for (var i = 0; i < rslt_ary.length; i++) {
                switch (rslt_ary[i]["type"]) {
                    case "slide":
                        $result.append(rslt_ary[i]["data"]);
                        break;
                    case "pptx-thumb":
                        //$("#pptx-thumb").attr("src", "data:image/jpeg;base64," +rslt_ary[i]["data"]);
                        break;
                    case "slideSize":
                        slideWidth = rslt_ary[i]["data"].width;
                        slideHeight = rslt_ary[i]["data"].height;
                        /*
                        $("#"+divId).css({
                            'width': slideWidth + 80,
                            'height': slideHeight + 60
                        });
                        */
                        break;
                    case "globalCSS":
                        //console.log(rslt_ary[i]["data"])
                        $result.append("<style>" + rslt_ary[i]["data"] + "</style>");
                        break;
                    case "ExecutionTime":
                        // 生成并添加全局 CSS
                        if (typeof window.PPTXCSSUtils.genGlobalCSS === 'function') {
                            $result.append("<style>" + window.PPTXCSSUtils.genGlobalCSS(styleTable, settings, slideWidth) + "</style>");
                        }
                        PPTXHtml.processMsgQueue(MsgQueue);
                        PPTXHtml.setNumericBullets($(".block"));
                        PPTXHtml.setNumericBullets($("table td"));

                        isDone = true;

                        if (settings.slideMode && !isSlideMode) {
                            isSlideMode = true;
                            initSlideMode(divId, settings);
                        } else if (!settings.slideMode) {
                            window.PPTXUIUtils.removeLoadingMessage();
                        }
                        break;
                    case "progress-update":
                        //console.log(rslt_ary[i]["data"]); //update progress bar - TODO
                        updateProgressBar(rslt_ary[i]["data"])
                        break;
                    default:
                }
            }
            if (!settings.slideMode || (settings.slideMode && settings.slideType == "revealjs")) {
                window.PPTXUIUtils.getSlidesWrapper(divId);

                if (settings.slideMode && settings.slideType == "revealjs") {
                    window.PPTXUIUtils.addRevealClass(divId);
                }
            }

            window.PPTXUIUtils.updateWrapperHeight(divId, settings.slidesScale, false, settings.slideType, null);

            //}
        }

        function initSlideMode(divId, settings) {
            //console.log(settings.slideType)
            if (settings.slideType == "" || settings.slideType == "divs2slidesjs") {
                var slideElements = document.querySelectorAll("#" + divId + " .slide");
                var slidesHeight = slideElements.length > 0 ? slideElements[0].clientHeight : 0;
                // Hide all slide elements
                for (var i = 0; i < slideElements.length; i++) {
                    slideElements[i].style.display = "none";
                }
                setTimeout(function () {
                    var slideConf = settings.slideModeConfig;
                    var loadingMsg = document.querySelector(".slides-loadnig-msg");
                    if (loadingMsg) {
                        loadingMsg.remove();
                    }
                        pptxjslideObj.init({
                        divId: divId,
                        slides: $("#" + divId + " .slide"),
                        totalSlides: $("#" + divId + " .slide").length,
                        slideCount: 1,
                        prevSlide: -1,
                        isInit: false,
                        isSlideMode: true,
                        isEnbleNextBtn: true,
                        isEnblePrevBtn: false,
                        isAutoSlideMode: false,
                        isLoopMode: false,
                        loopIntrval: null,
                        first: slideConf.first,
                        nav: slideConf.nav,
                        showPlayPauseBtn: settings.showPlayPauseBtn,
                        showFullscreenBtn: settings.showFullscreenBtn,
                        navTxtColor: slideConf.navTxtColor,
                        keyBoardShortCut: slideConf.keyBoardShortCut,
                        showSlideNum: slideConf.showSlideNum,
                        showTotalSlideNum: slideConf.showTotalSlideNum,
                        autoSlide: slideConf.autoSlide,
                        timeBetweenSlides: slideConf.autoSlide,
                        randomAutoSlide: slideConf.randomAutoSlide,
                        loop: slideConf.loop,
                        background: slideConf.background,
                        slctdBgClr: slideConf.background,
                        transition: slideConf.transition,
                        transitionTime: slideConf.transitionTime
                    });

                    window.PPTXUIUtils.updateWrapperHeight(divId, settings.slidesScale, true, "divs2slidesjs", 1);

                }, 1500);
            } else if (settings.slideType == "revealjs") {
                window.PPTXUIUtils.removeLoadingMessage();
                var revealjsPath = "";
                if (settings.revealjsPath != "") {
                    revealjsPath = settings.revealjsPath;
                } else {
                    revealjsPath = "./revealjs/reveal.js";
                }
                $.getScript(revealjsPath, function (response, status) {
                    if (status == "success") {
                        // $("section").removeClass("slide");
                        Reveal.initialize(settings.revealjsConfig); //revealjsConfig - TODO
                    }
                });
            }



        }

        function processPPTX(zip) {
            // 使用 PPTXParser 模块处理 PPTX 解析
            return PPTXParser.processPPTX(zip);
        }

        // Process single slide function moved to PPTXParser module

        // readXmlFile, getContentTypes, getSlideSizeAndSetDefaultTextStyle, processSingleSlide, indexNodes 已移至 PPTXParser 模块
        function indexNodes(content) {
            // 使用 PPTXParser
            return PPTXParser.indexNodes(content);
        }

        function processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType) {
            // 使用 PPTXNodeUtils 模块处理节点
            var handlers = {
                processSpNode: processSpNode,
                processCxnSpNode: processCxnSpNode,
                processPicNode: processPicNode,
                processGraphicFrameNode: processGraphicFrameNode,
                processGroupSpNode: processGroupSpNode
            };
            return window.PPTXNodeUtils.processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType, handlers);
        }

        function processGroupSpNode(node, warpObj, source) {
            // 使用 PPTXNodeUtils 模块处理组节点
            return window.PPTXNodeUtils.processGroupSpNode(node, warpObj, source, slideFactor, processNodesInSlide);
        }

        function processSpNode(node, pNode, warpObj, source, sType) {

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

            var id = window.PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "id"]);
            var name = window.PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "name"]);
            var idx = (window.PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]) === undefined) ? undefined : window.PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]);
            var type = (window.PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]) === undefined) ? undefined : window.PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
            var order = window.PPTXUtils.getTextByPathList(node, ["attrs", "order"]);
            var isUserDrawnBg;
            if (source == "slideLayoutBg" || source == "slideMasterBg") {
                var userDrawn = window.PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "attrs", "userDrawn"]);
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
                txBoxVal = window.PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvSpPr", "attrs", "txBox"]);
                if (txBoxVal == "1") {
                    type = "textBox";
                }
            }
            if (type === undefined) {
                type = window.PPTXUtils.getTextByPathList(slideLayoutSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                if (type === undefined) {
                    //type = window.PPTXUtils.getTextByPathList(slideMasterSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    if (source == "diagramBg") {
                        type = "diagram";
                    } else {

                        type = "obj"; //default type
                    }
                }
            }
            //console.log("processSpNode type:", type, "idx:", idx);
            return genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, isUserDrawnBg, sType, source);
        }

        function processCxnSpNode(node, pNode, warpObj, source, sType) {

            var id = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
            var name = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
            var idx = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
            var type = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
            //<p:cNvCxnSpPr>(<p:cNvCxnSpPr>, <a:endCxn>)
            var order = node["attrs"]["order"];

            return genShape(node, pNode, undefined, undefined, id, name, idx, type, order, warpObj, undefined, sType, source);
        }

        /**
         * 生成形状的SVG HTML表示
         * 
         * @param {Object} node - 形状节点对象
         * @param {Object} pNode - 父节点对象
         * @param {Object} slideLayoutSpNode - 幻灯片布局中的形状节点
         * @param {Object} slideMasterSpNode - 幻灯片母版中的形状节点  
         * @param {string} id - 形状ID
         * @param {string} name - 形状名称
         * @param {string} idx - 形状索引
         * @param {string} type - 形状类型
         * @param {number} order - 显示顺序
         * @param {Object} warpObj - 包装对象，包含解析上下文
         * @param {boolean} isUserDrawnBg - 是否用户绘制的背景
         * @param {string} sType - 形状类型标识
         * @param {string} source - 来源标识
         * @returns {string} 形状的SVG HTML字符串
         */
        function genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, isUserDrawnBg, sType, source) {
            //var dltX = 0;
            //var dltY = 0;
            var xfrmList = ["p:spPr", "a:xfrm"];
            var slideXfrmNode = window.PPTXUtils.getTextByPathList(node, xfrmList);
            var slideLayoutXfrmNode = window.PPTXUtils.getTextByPathList(slideLayoutSpNode, xfrmList);
            var slideMasterXfrmNode = window.PPTXUtils.getTextByPathList(slideMasterSpNode, xfrmList);

            var result = "";
            var shpId = window.PPTXUtils.getTextByPathList(node, ["attrs", "order"]);
            //console.log("shpId: ",shpId)
            var shapType = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);

            //custGeom - Amir
            var custShapType = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:custGeom"]);

            var isFlipV = false;
            var isFlipH = false;
            var flip = "";
            if (window.PPTXUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipV"]) === "1") {
                isFlipV = true;
            }
            if (window.PPTXUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipH"]) === "1") {
                isFlipH = true;
            }
            if (isFlipH && !isFlipV) {
                flip = " scale(-1,1)"
            } else if (!isFlipH && isFlipV) {
                flip = " scale(1,-1)"
            } else if (isFlipH && isFlipV) {
                flip = " scale(-1,-1)"
            }
            /////////////////////////Amir////////////////////////
            //rotate
            var rotate = window.PPTXColorUtils.angleToDegrees(window.PPTXUtils.getTextByPathList(slideXfrmNode, ["attrs", "rot"]));

            //console.log("genShape rotate: " + rotate);
            var txtRotate;
            var txtXframeNode = window.PPTXUtils.getTextByPathList(node, ["p:txXfrm"]);
            if (txtXframeNode !== undefined) {
                var txtXframeRot = window.PPTXUtils.getTextByPathList(txtXframeNode, ["attrs", "rot"]);
                if (txtXframeRot !== undefined) {
                    txtRotate = window.PPTXColorUtils.angleToDegrees(txtXframeRot) + 90;
                }
            } else {
                txtRotate = rotate;
            }
            //////////////////////////////////////////////////
            if (shapType !== undefined || custShapType !== undefined /*&& slideXfrmNode !== undefined*/) {
                var off = window.PPTXUtils.getTextByPathList(slideXfrmNode, ["a:off", "attrs"]);
                var x = parseInt(off["x"]) * slideFactor;
                var y = parseInt(off["y"]) * slideFactor;

                var ext = window.PPTXUtils.getTextByPathList(slideXfrmNode, ["a:ext", "attrs"]);
                var w = parseInt(ext["cx"]) * slideFactor;
                var h = parseInt(ext["cy"]) * slideFactor;

                var svgCssName = "_svg_css_" + (Object.keys(styleTable).length + 1) + "_"  + Math.floor(Math.random() * 1001);
                //console.log("name:", name, "svgCssName: ", svgCssName)
                var effectsClassName = svgCssName + "_effects";
                result += "<svg class='drawing " + svgCssName + " " + effectsClassName + " ' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name + "'" +
                    "' style='" +
                    getPosition(slideXfrmNode, pNode, undefined, undefined, sType) +
                    getSize(slideXfrmNode, undefined, undefined) +
                    " z-index: " + order + ";" +
                    "transform: rotate(" + ((rotate !== undefined) ? rotate : 0) + "deg)" + flip + ";" +
                    "'>";
                result += '<defs>'
                // Fill Color
                var fillColor = window.PPTXShapeFillsUtils.getShapeFill(node, pNode, true, warpObj, source);
                //console.log("genShape: fillColor: ", fillColor)
                var grndFillFlg = false;
                var imgFillFlg = false;
                var clrFillType = window.PPTXColorUtils.getFillType(window.PPTXUtils.getTextByPathList(node, ["p:spPr"]));
                if (clrFillType == "GROUP_FILL") {
                    clrFillType = window.PPTXColorUtils.getFillType(window.PPTXUtils.getTextByPathList(pNode, ["p:grpSpPr"]));
                }
                // if (clrFillType == "") {
                //     var clrFillType = window.PPTXColorUtils.getFillType(window.PPTXUtils.getTextByPathList(node, ["p:style","a:fillRef"]));
                // }
                //console.log("genShape: fillColor: ", fillColor, ", clrFillType: ", clrFillType, ", node: ", node)
                /////////////////////////////////////////                    
                if (clrFillType == "GRADIENT_FILL") {
                    grndFillFlg = true;
                    var color_arry = fillColor.color;
                    var angl = fillColor.rot + 90;
                    var svgGrdnt = window.PPTXShapeFillsUtils.getSvgGradient(w, h, angl, color_arry, shpId);
                    //fill="url(#linGrd)"
                    //console.log("genShape: svgGrdnt: ", svgGrdnt)
                    result += svgGrdnt;

                } else if (clrFillType == "PIC_FILL") {
                    imgFillFlg = true;
                    // 提取图片 URL（fillColor 可能是对象或字符串）
                    var imgFill = typeof fillColor === 'object' && fillColor.img ? fillColor.img : fillColor;
                    var svgBgImg = window.PPTXShapeFillsUtils.getSvgImagePattern(node, imgFill, shpId, warpObj);
                    //fill="url(#imgPtrn)"
                    //console.log(svgBgImg)
                    result += svgBgImg;
                } else if (clrFillType == "PATTERN_FILL") {
                    var styleText = fillColor;
                    if (styleText in styleTable) {
                        styleText += "do-nothing: " + svgCssName +";";
                    }
                    styleTable[styleText] = {
                        "name": svgCssName,
                        "text": styleText
                    };
                    //}
                    fillColor = "none";
                } else {
                    if (clrFillType != "SOLID_FILL" && clrFillType != "PATTERN_FILL" &&
                        (shapType == "arc" ||
                            shapType == "bracketPair" ||
                            shapType == "bracePair" ||
                            shapType == "leftBracket" ||
                            shapType == "leftBrace" ||
                            shapType == "rightBrace" ||
                            shapType == "rightBracket")) { 
                        // 临时解决方案 - TODO: 对于弧形、括号等形状，当填充类型不是实心或图案时，设置为无填充
                        // 这是因为这些形状在某些情况下应该显示为轮廓，但需要更精确的处理
                        fillColor = "none";
                    }
                }
                // Border Color
                var border = window.PPTXShapeFillsUtils.getBorder(node, pNode, true, "shape", warpObj);

                var headEndNodeAttrs = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
                var tailEndNodeAttrs = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);
                // type: none, triangle, stealth, diamond, oval, arrow

                ////////////////////effects/////////////////////////////////////////////////////
                //p:spPr => a:effectLst =>
                //"a:blur"
                //"a:fillOverlay"
                //"a:glow"
                //"a:innerShdw"
                //"a:outerShdw"
                //"a:prstShdw"
                //"a:reflection"
                //"a:softEdge"
                //p:spPr => a:scene3d
                //"a:camera"
                //"a:lightRig"
                //"a:backdrop"
                //"a:extLst"?
                //p:spPr => a:sp3d
                //"a:bevelT"
                //"a:bevelB"
                //"a:extrusionClr"
                //"a:contourClr"
                //"a:extLst"?
                //////////////////////////////outerShdw///////////////////////////////////////////
                //not support sizing the shadow
                var outerShdwNode = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:effectLst", "a:outerShdw"]);
                var oShadowSvgUrlStr = ""
                if (outerShdwNode !== undefined) {
                    var chdwClrNode = window.PPTXColorUtils.getSolidFill(outerShdwNode, undefined, undefined, warpObj);
                    var outerShdwAttrs = outerShdwNode["attrs"];

                    //var algn = outerShdwAttrs["algn"];
                    var dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
                    var dist = parseInt(outerShdwAttrs["dist"]) * slideFactor;//(px) //* (3 / 4); //(pt)
                    //var rotWithShape = outerShdwAttrs["rotWithShape"];
                    var blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * slideFactor) : ""; //+ "px"
                    //var sx = (outerShdwAttrs["sx"]) ? (parseInt(outerShdwAttrs["sx"]) / 100000) : 1;
                    //var sy = (outerShdwAttrs["sy"]) ? (parseInt(outerShdwAttrs["sy"]) / 100000) : 1;
                    var vx = dist * Math.sin(dir * Math.PI / 180);
                    var hx = dist * Math.cos(dir * Math.PI / 180);
                    //SVG
                    //var oShadowId = "outerhadow_" + shpId;
                    //oShadowSvgUrlStr = "filter='url(#" + oShadowId+")'";
                    //var shadowFilterStr = '<filter id="' + oShadowId + '" x="0" y="0" width="' + w * (6 / 8) + '" height="' + h + '">';
                    //1:
                    //shadowFilterStr += '<feDropShadow dx="' + vx + '" dy="' + hx + '" stdDeviation="' + blurRad * (3 / 4) + '" flood-color="#' + chdwClrNode +'" flood-opacity="1" />'
                    //2:
                    //shadowFilterStr += '<feFlood result="floodColor" flood-color="red" flood-opacity="0.5"   width="' + w * (6 / 8) + '" height="' + h + '"  />'; //#' + chdwClrNode +'
                    //shadowFilterStr += '<feOffset result="offOut" in="SourceGraph ccfsdf-+ic"  dx="' + vx + '" dy="' + hx + '"/>'; //how much to offset
                    //shadowFilterStr += '<feGaussianBlur result="blurOut" in="offOut" stdDeviation="' + blurRad*(3/4) +'"/>'; //tdDeviation is how much to blur
                    //shadowFilterStr += '<feComponentTransfer><feFuncA type="linear" slope="0.5"/></feComponentTransfer>'; //slope is the opacity of the shadow
                    //shadowFilterStr += '<feBlend in="SourceGraphic" in2="blurOut"  mode="normal" />'; //this contains the element that the filter is applied to
                    //shadowFilterStr += '</filter>'; 
                    //result += shadowFilterStr;

                    //css:
                    var svg_css_shadow = "filter:drop-shadow(" + hx + "px " + vx + "px " + blurRad + "px #" + chdwClrNode + ");";

                    if (svg_css_shadow in styleTable) {
                        svg_css_shadow += "do-nothing: " + svgCssName + ";";
                    }

                    styleTable[svg_css_shadow] = {
                        "name": effectsClassName,
                        "text": svg_css_shadow
                    };

                } 
                ////////////////////////////////////////////////////////////////////////////////////////
                if ((headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) ||
                    (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow"))) {
                    var triangleMarker = "<marker id='markerTriangle_" + shpId + "' viewBox='0 0 10 10' refX='1' refY='5' markerWidth='5' markerHeight='5' stroke='" + border.color + "' fill='" + border.color +
                        "' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>";
                    result += triangleMarker;
                }
                result += '</defs>'
            }
            if (shapType !== undefined && custShapType === undefined) {
                //console.log("shapType: ", shapType)
                switch (shapType) {
                    case "rect":
                    case "flowChartProcess":
                    case "flowChartPredefinedProcess":
                    case "flowChartInternalStorage":
                    case "actionButtonBlank":
                        result += window.PPTXBasicShapes.genRectWithDecoration(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr, shapType);
                        break;
                    case "flowChartCollate":
                        result += window.PPTXFlowchartShapes.genFlowChartCollate(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "flowChartDocument":
                        result += window.PPTXFlowchartShapes.genFlowChartDocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "flowChartMultidocument":
                        result += window.PPTXFlowchartShapes.genFlowChartMultidocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonBackPrevious":
                        result += window.PPTXActionButtonShapes.genActionButtonBackPrevious(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonBeginning":
                        result += window.PPTXActionButtonShapes.genActionButtonBeginning(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonDocument":
                        result += window.PPTXActionButtonShapes.genActionButtonDocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonEnd":
                        result += window.PPTXActionButtonShapes.genActionButtonEnd(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonForwardNext":
                        result += window.PPTXActionButtonShapes.genActionButtonForwardNext(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonHelp":
                        result += window.PPTXActionButtonShapes.genActionButtonHelp(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonHome":
                        result += window.PPTXActionButtonShapes.genActionButtonHome(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonInformation":
                        result += window.PPTXActionButtonShapes.genActionButtonInformation(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonMovie":
                        result += window.PPTXActionButtonShapes.genActionButtonMovie(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
    case "actionButtonReturn":
        result += window.PPTXActionButtonShapes.genActionButtonReturn(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
        break;
                    case "actionButtonSound":
                        result += window.PPTXActionButtonShapes.genActionButtonSound(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "irregularSeal1":
                    case "irregularSeal2":
                        if (shapType == "irregularSeal1") {
                            var d = "M" + w * 10800 / 21600 + "," + h * 5800 / 21600 +
                                " L" + w * 14522 / 21600 + "," + 0 +
                                " L" + w * 14155 / 21600 + "," + h * 5325 / 21600 +
                                " L" + w * 18380 / 21600 + "," + h * 4457 / 21600 +
                                " L" + w * 16702 / 21600 + "," + h * 7315 / 21600 +
                                " L" + w * 21097 / 21600 + "," + h * 8137 / 21600 +
                                " L" + w * 17607 / 21600 + "," + h * 10475 / 21600 +
                                " L" + w + "," + h * 13290 / 21600 +
                                " L" + w * 16837 / 21600 + "," + h * 12942 / 21600 +
                                " L" + w * 18145 / 21600 + "," + h * 18095 / 21600 +
                                " L" + w * 14020 / 21600 + "," + h * 14457 / 21600 +
                                " L" + w * 13247 / 21600 + "," + h * 19737 / 21600 +
                                " L" + w * 10532 / 21600 + "," + h * 14935 / 21600 +
                                " L" + w * 8485 / 21600 + "," + h +
                                " L" + w * 7715 / 21600 + "," + h * 15627 / 21600 +
                                " L" + w * 4762 / 21600 + "," + h * 17617 / 21600 +
                                " L" + w * 5667 / 21600 + "," + h * 13937 / 21600 +
                                " L" + w * 135 / 21600 + "," + h * 14587 / 21600 +
                                " L" + w * 3722 / 21600 + "," + h * 11775 / 21600 +
                                " L" + 0 + "," + h * 8615 / 21600 +
                                " L" + w * 4627 / 21600 + "," + h * 7617 / 21600 +
                                " L" + w * 370 / 21600 + "," + h * 2295 / 21600 +
                                " L" + w * 7312 / 21600 + "," + h * 6320 / 21600 +
                                " L" + w * 8352 / 21600 + "," + h * 2295 / 21600 +
                                " z";
                        } else if (shapType == "irregularSeal2") {
                            var d = "M" + w * 11462 / 21600 + "," + h * 4342 / 21600 +
                                " L" + w * 14790 / 21600 + "," + 0 +
                                " L" + w * 14525 / 21600 + "," + h * 5777 / 21600 +
                                " L" + w * 18007 / 21600 + "," + h * 3172 / 21600 +
                                " L" + w * 16380 / 21600 + "," + h * 6532 / 21600 +
                                " L" + w + "," + h * 6645 / 21600 +
                                " L" + w * 16985 / 21600 + "," + h * 9402 / 21600 +
                                " L" + w * 18270 / 21600 + "," + h * 11290 / 21600 +
                                " L" + w * 16380 / 21600 + "," + h * 12310 / 21600 +
                                " L" + w * 18877 / 21600 + "," + h * 15632 / 21600 +
                                " L" + w * 14640 / 21600 + "," + h * 14350 / 21600 +
                                " L" + w * 14942 / 21600 + "," + h * 17370 / 21600 +
                                " L" + w * 12180 / 21600 + "," + h * 15935 / 21600 +
                                " L" + w * 11612 / 21600 + "," + h * 18842 / 21600 +
                                " L" + w * 9872 / 21600 + "," + h * 17370 / 21600 +
                                " L" + w * 8700 / 21600 + "," + h * 19712 / 21600 +
                                " L" + w * 7527 / 21600 + "," + h * 18125 / 21600 +
                                " L" + w * 4917 / 21600 + "," + h +
                                " L" + w * 4805 / 21600 + "," + h * 18240 / 21600 +
                                " L" + w * 1285 / 21600 + "," + h * 17825 / 21600 +
                                " L" + w * 3330 / 21600 + "," + h * 15370 / 21600 +
                                " L" + 0 + "," + h * 12877 / 21600 +
                                " L" + w * 3935 / 21600 + "," + h * 11592 / 21600 +
                                " L" + w * 1172 / 21600 + "," + h * 8270 / 21600 +
                                " L" + w * 5372 / 21600 + "," + h * 7817 / 21600 +
                                " L" + w * 4502 / 21600 + "," + h * 3625 / 21600 +
                                " L" + w * 8550 / 21600 + "," + h * 6382 / 21600 +
                                " L" + w * 9722 / 21600 + "," + h * 1887 / 21600 +
                                " z";
                        }
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartTerminator":
                        var x1, x2, y1, cd2 = 180, cd4 = 90, c3d4 = 270;
                        x1 = w * 3475 / 21600;
                        x2 = w * 18125 / 21600;
                        y1 = h * 10800 / 21600;
                        //path attrs: w = 21600; h = 21600; 
                        var d = "M" + x1 + "," + 0 +
                            " L" + x2 + "," + 0 +
                            window.PPTXShapeUtils.shapeArc(x2, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L") +
                            " L" + x1 + "," + h +
                            window.PPTXShapeUtils.shapeArc(x1, h / 2, x1, y1, cd4, cd4 + cd2, false).replace("M", "L") +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartPunchedTape":
                        var x1, x1, y1, y2, cd2 = 180;
                        x1 = w * 5 / 20;
                        y1 = h * 2 / 20;
                        y2 = h * 18 / 20;
                        var d = "M" + 0 + "," + y1 +
                            window.PPTXShapeUtils.shapeArc(x1, y1, x1, y1, cd2, 0, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w * (3 / 4), y1, x1, y1, cd2, 360, false).replace("M", "L") +
                            " L" + w + "," + y2 +
                            window.PPTXShapeUtils.shapeArc(w * (3 / 4), y2, x1, y1, 0, -cd2, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(x1, y2, x1, y1, 0, cd2, false).replace("M", "L") +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartOnlineStorage":
                        var x1, y1, c3d4 = 270, cd4 = 90;
                        x1 = w * 1 / 6;
                        y1 = h * 3 / 6;
                        var d = "M" + x1 + "," + 0 +
                            " L" + w + "," + 0 +
                            window.PPTXShapeUtils.shapeArc(w, h / 2, x1, y1, c3d4, 90, false).replace("M", "L") +
                            " L" + x1 + "," + h +
                            window.PPTXShapeUtils.shapeArc(x1, h / 2, x1, y1, cd4, 270, false).replace("M", "L") +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartDisplay":
                        var x1, x2, y1, c3d4 = 270, cd2 = 180;
                        x1 = w * 1 / 6;
                        x2 = w * 5 / 6;
                        y1 = h * 3 / 6;
                        //path attrs: w = 6; h = 6; 
                        var d = "M" + 0 + "," + y1 +
                            " L" + x1 + "," + 0 +
                            " L" + x2 + "," + 0 +
                            window.PPTXShapeUtils.shapeArc(w, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L") +
                            " L" + x1 + "," + h +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartDelay":
                        var wd2 = w / 2, hd2 = h / 2, cd2 = 180, c3d4 = 270, cd4 = 90;
                        var d = "M" + 0 + "," + 0 +
                            " L" + wd2 + "," + 0 +
                            window.PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, c3d4, c3d4 + cd2, false).replace("M", "L") +
                            " L" + 0 + "," + h +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartMagneticTape":
                        var wd2 = w / 2, hd2 = h / 2, cd2 = 180, c3d4 = 270, cd4 = 90;
                        var idy, ib, ang1;
                        idy = hd2 * Math.sin(Math.PI / 4);
                        ib = hd2 + idy;
                        ang1 = Math.atan(h / w);
                        var ang1Dg = ang1 * 180 / Math.PI;
                        var d = "M" + wd2 + "," + h +
                            window.PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, cd4, cd2, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, cd2, c3d4, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, c3d4, 360, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, 0, ang1Dg, false).replace("M", "L") +
                            " L" + w + "," + ib +
                            " L" + w + "," + h +
                            " z";
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "ellipse":
                    case "flowChartConnector":
                    case "flowChartSummingJunction":
                    case "flowChartOr":
                        result += window.PPTXBasicShapes.genEllipse(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        if (shapType == "flowChartOr") {
                            result += " <polyline points='" + w / 2 + " " + 0 + "," + w / 2 + " " + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                            result += " <polyline points='" + 0 + " " + h / 2 + "," + w + " " + h / 2 + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        } else if (shapType == "flowChartSummingJunction") {
                            var iDx, idy, il, ir, it, ib, hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
                            var angVal = Math.PI / 4;
                            iDx = wd2 * Math.cos(angVal);
                            idy = hd2 * Math.sin(angVal);
                            il = hc - iDx;
                            ir = hc + iDx;
                            it = vc - idy;
                            ib = vc + idy;
                            result += " <polyline points='" + il + " " + it + "," + ir + " " + ib + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                            result += " <polyline points='" + ir + " " + it + "," + il + " " + ib + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        }
                        break;
                    case "roundRect":
                    case "round1Rect":
                    case "round2DiagRect":
                    case "round2SameRect":
                    case "snip1Rect":
                    case "snip2DiagRect":
                    case "snip2SameRect":
                    case "flowChartAlternateProcess":
                    case "flowChartPunchedCard":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, sAdj1_val;// = 0.33334;
                        var sAdj2, sAdj2_val;// = 0.33334;
                        var shpTyp, adjTyp;
                        if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor === Array) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) / 50000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) / 50000;
                                }
                            }
                        } else if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor !== Array) {
                            var sAdj = window.PPTXUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                            sAdj1_val = parseInt(sAdj.substr(4)) / 50000;
                            sAdj2_val = 0;
                        }
                        //console.log("shapType: ",shapType,",node: ",node )
                        var tranglRott = "";
                        switch (shapType) {
                            case "roundRect":
                            case "flowChartAlternateProcess":
                                shpTyp = "round";
                                adjTyp = "cornrAll";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                break;
                            case "round1Rect":
                                shpTyp = "round";
                                adjTyp = "cornr1";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                break;
                            case "round2DiagRect":
                                shpTyp = "round";
                                adjTyp = "diag";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                            case "round2SameRect":
                                shpTyp = "round";
                                adjTyp = "cornr2";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                            case "snip1Rect":
                            case "flowChartPunchedCard":
                                shpTyp = "snip";
                                adjTyp = "cornr1";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                if (shapType == "flowChartPunchedCard") {
                                    tranglRott = "transform='translate(" + w + ",0) scale(-1,1)'";
                                }
                                break;
                            case "snip2DiagRect":
                                shpTyp = "snip";
                                adjTyp = "diag";
                                if (sAdj1_val === undefined) sAdj1_val = 0;
                                if (sAdj2_val === undefined) sAdj2_val = 0.33334;
                                break;
                            case "snip2SameRect":
                                shpTyp = "snip";
                                adjTyp = "cornr2";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                        }
                        var d_val = window.PPTXShapeUtils.shapeSnipRoundRect(w, h, sAdj1_val, sAdj2_val, shpTyp, adjTyp);
                        result += "<path " + tranglRott + "  d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "snipRoundRect":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, sAdj1_val = 0.33334;
                        var sAdj2, sAdj2_val = 0.33334;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) / 50000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) / 50000;
                                }
                            }
                        }
                        var d_val = "M0," + h + " L" + w + "," + h + " L" + w + "," + (h / 2) * sAdj2_val +
                            " L" + (w / 2 + (w / 2) * (1 - sAdj2_val)) + ",0 L" + (w / 2) * sAdj1_val + ",0 Q0,0 0," + (h / 2) * sAdj1_val + " z";

                        result += "<path   d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "bentConnector2":
                        var d = "";
                        // if (isFlipV) {
                        //     d = "M 0 " + w + " L " + h + " " + w + " L " + h + " 0";
                        // } else {
                        d = "M " + w + " 0 L " + w + " " + h + " L 0 " + h;
                        //}
                        result += "<path d='" + d + "' stroke='" + border.color +
                            "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' fill='none' ";
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                        }
                        result += "/>";
                        break;
                    case "rtTriangle":
                        result += " <polygon points='0 0,0 " + h + "," + w + " " + h + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "triangle":
                    case "flowChartExtract":
                    case "flowChartMerge":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var shapAdjst_val = 0.5;
                        if (shapAdjst !== undefined) {
                            shapAdjst_val = parseInt(shapAdjst.substr(4)) * slideFactor;
                            //console.log("w: "+w+"\nh: "+h+"\nshapAdjst: "+shapAdjst+"\nshapAdjst_val: "+shapAdjst_val);
                        }
                        var tranglRott = "";
                        if (shapType == "flowChartMerge") {
                            tranglRott = "transform='rotate(180 " + w / 2 + "," + h / 2 + ")'";
                        }
                        result += " <polygon " + tranglRott + " points='" + (w * shapAdjst_val) + " 0,0 " + h + "," + w + " " + h + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "diamond":
                    case "flowChartDecision":
                    case "flowChartSort":
                        result += " <polygon points='" + (w / 2) + " 0,0 " + (h / 2) + "," + (w / 2) + " " + h + "," + w + " " + (h / 2) + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        if (shapType == "flowChartSort") {
                            result += " <polyline points='0 " + h / 2 + "," + w + " " + h / 2 + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        }
                        break;
                    case "trapezoid":
                    case "flowChartManualOperation":
                    case "flowChartManualInput":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adjst_val = 0.2;
                        var max_adj_const = 0.7407;
                        if (shapAdjst !== undefined) {
                            var adjst = parseInt(shapAdjst.substr(4)) * slideFactor;
                            adjst_val = (adjst * 0.5) / max_adj_const;
                            // console.log("w: "+w+"\nh: "+h+"\nshapAdjst: "+shapAdjst+"\nadjst_val: "+adjst_val);
                        }
                        var cnstVal = 0;
                        var tranglRott = "";
                        if (shapType == "flowChartManualOperation") {
                            tranglRott = "transform='rotate(180 " + w / 2 + "," + h / 2 + ")'";
                        }
                        if (shapType == "flowChartManualInput") {
                            adjst_val = 0;
                            cnstVal = h / 5;
                        }
                        result += " <polygon " + tranglRott + " points='" + (w * adjst_val) + " " + cnstVal + ",0 " + h + "," + w + " " + h + "," + (1 - adjst_val) * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "parallelogram":
                    case "flowChartInputOutput":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adjst_val = 0.25;
                        var max_adj_const;
                        if (w > h) {
                            max_adj_const = w / h;
                        } else {
                            max_adj_const = h / w;
                        }
                        if (shapAdjst !== undefined) {
                            var adjst = parseInt(shapAdjst.substr(4)) / 100000;
                            adjst_val = adjst / max_adj_const;
                            //console.log("w: "+w+"\nh: "+h+"\nadjst: "+adjst_val+"\nmax_adj_const: "+max_adj_const);
                        }
                        result += " <polygon points='" + adjst_val * w + " 0,0 " + h + "," + (1 - adjst_val) * w + " " + h + "," + w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;

                        break;
                    case "pentagon":
                        result += " <polygon points='" + (0.5 * w) + " 0,0 " + (0.375 * h) + "," + (0.15 * w) + " " + h + "," + 0.85 * w + " " + h + "," + w + " " + 0.375 * h + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "hexagon":
                    case "flowChartPreparation":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 25000 * slideFactor;
                        var vf = 115470 * slideFactor;;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var angVal1 = 60 * Math.PI / 180;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var maxAdj, a, shd2, x1, x2, dy1, y1, y2, vc = h / 2, hd2 = h / 2;
                        var ss = Math.min(w, h);
                        maxAdj = cnstVal1 * w / ss;
                        a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
                        shd2 = hd2 * vf / cnstVal2;
                        x1 = ss * a / cnstVal2;
                        x2 = w - x1;
                        dy1 = shd2 * Math.sin(angVal1);
                        y1 = vc - dy1;
                        y2 = vc + dy1;

                        var d = "M" + 0 + "," + vc +
                            " L" + x1 + "," + y1 +
                            " L" + x2 + "," + y1 +
                            " L" + w + "," + vc +
                            " L" + x2 + "," + y2 +
                            " L" + x1 + "," + y2 +
                            " z";

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "heptagon":
                        result += " <polygon points='" + (0.5 * w) + " 0," + w / 8 + " " + h / 4 + ",0 " + (5 / 8) * h + "," + w / 4 + " " + h + "," + (3 / 4) * w + " " + h + "," +
                            w + " " + (5 / 8) * h + "," + (7 / 8) * w + " " + h / 4 + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "octagon":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj1 = 0.25;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) / 100000;

                        }
                        var adj2 = (1 - adj1);
                        //console.log("adj1: "+adj1+"\nadj2: "+adj2);
                        result += " <polygon points='" + adj1 * w + " 0,0 " + adj1 * h + ",0 " + adj2 * h + "," + adj1 * w + " " + h + "," + adj2 * w + " " + h + "," +
                            w + " " + adj2 * h + "," + w + " " + adj1 * h + "," + adj2 * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "decagon":
                        result += " <polygon points='" + (3 / 8) * w + " 0," + w / 8 + " " + h / 8 + ",0 " + h / 2 + "," + w / 8 + " " + (7 / 8) * h + "," + (3 / 8) * w + " " + h + "," +
                            (5 / 8) * w + " " + h + "," + (7 / 8) * w + " " + (7 / 8) * h + "," + w + " " + h / 2 + "," + (7 / 8) * w + " " + h / 8 + "," + (5 / 8) * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "dodecagon":
                        result += " <polygon points='" + (3 / 8) * w + " 0," + w / 8 + " " + h / 8 + ",0 " + (3 / 8) * h + ",0 " + (5 / 8) * h + "," + w / 8 + " " + (7 / 8) * h + "," + (3 / 8) * w + " " + h + "," +
                            (5 / 8) * w + " " + h + "," + (7 / 8) * w + " " + (7 / 8) * h + "," + w + " " + (5 / 8) * h + "," + w + " " + (3 / 8) * h + "," + (7 / 8) * w + " " + h / 8 + "," + (5 / 8) * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star4":
                        var d = window.PPTXStarShapes.genStar4(w, h, node, slideFactor);
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star5":
                        var d = window.PPTXStarShapes.genStar5(w, h, node, slideFactor);
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star6":
                        var d = window.PPTXStarShapes.genStar6(w, h, node, slideFactor);
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star7":
                        var d = window.PPTXStarShapes.genStar7(w, h, node, slideFactor);
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star8":
                        var d = window.PPTXStarShapes.genStar8(w, h, node, slideFactor);
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;

                    case "star10":
                        var d = window.PPTXStarShapes.genStar10(w, h, node, slideFactor);
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star12":
                        var d = window.PPTXStarShapes.genStar12(w, h, node, slideFactor);
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star16":
                        var d = window.PPTXStarShapes.genStar16(w, h, node, slideFactor);
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star24":
                        var d = window.PPTXStarShapes.genStar24(w, h, node, slideFactor);
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star32":
                        var d = window.PPTXStarShapes.genStar32(w, h, node, slideFactor);
                        result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;

                    case "pie":
                    case "pieWedge":
                    case "arc":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var adj1, adj2, H, shapAdjst1, shapAdjst2, isClose;
                        if (shapType == "pie") {
                            adj1 = 0;
                            adj2 = 270;
                            H = h;
                            isClose = true;
                        } else if (shapType == "pieWedge") {
                            adj1 = 180;
                            adj2 = 270;
                            H = 2 * h;
                            isClose = true;
                        } else if (shapType == "arc") {
                            adj1 = 270;
                            adj2 = 0;
                            H = h;
                            isClose = false;
                        }
                        if (shapAdjst !== undefined) {
                            shapAdjst1 = window.PPTXUtils.getTextByPathList(shapAdjst, ["attrs", "fmla"]);
                            shapAdjst2 = shapAdjst1;
                            if (shapAdjst1 === undefined) {
                                shapAdjst1 = shapAdjst[0]["attrs"]["fmla"];
                                shapAdjst2 = shapAdjst[1]["attrs"]["fmla"];
                            }
                            if (shapAdjst1 !== undefined) {
                                adj1 = parseInt(shapAdjst1.substr(4)) / 60000;
                            }
                            if (shapAdjst2 !== undefined) {
                                adj2 = parseInt(shapAdjst2.substr(4)) / 60000;
                            }
                        }
                        var pieVals = window.PPTXShapeUtils.shapePie(H, w, adj1, adj2, isClose);
                        //console.log("shapType: ",shapType,"\nimgFillFlg: ",imgFillFlg,"\ngrndFillFlg: ",grndFillFlg,"\nshpId: ",shpId,"\nfillColor: ",fillColor);
                        result += "<path   d='" + pieVals[0] + "' transform='" + pieVals[1] + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "chord":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, sAdj1_val = 45;
                        var sAdj2, sAdj2_val = 270;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) / 60000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) / 60000;
                                }
                            }
                        }
                        var hR = h / 2;
                        var wR = w / 2;
                        var d_val = window.PPTXShapeUtils.shapeArc(wR, hR, wR, hR, sAdj1_val, sAdj2_val, true);
                        //console.log("shapType: ",shapType,", sAdj1_val: ",sAdj1_val,", sAdj2_val: ",sAdj2_val)
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "frame":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj1 = 12500 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var a1, x1, x4, y4;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        x1 = Math.min(w, h) * a1 / cnstVal2;
                        x4 = w - x1;
                        y4 = h - x1;
                        var d = "M" + 0 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + h +
                            " L" + 0 + "," + h +
                            " z" +
                            "M" + x1 + "," + x1 +
                            " L" + x1 + "," + y4 +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + x1 +
                            " z";
                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "donut":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 25000 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var a, dr, iwd2, ihd2;
                        if (adj < 0) a = 0
                        else if (adj > cnstVal1) a = cnstVal1
                        else a = adj
                        dr = Math.min(w, h) * a / cnstVal2;
                        iwd2 = w / 2 - dr;
                        ihd2 = h / 2 - dr;
                        var d = "M" + 0 + "," + h / 2 +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") +
                            " z" +
                            "M" + dr + "," + h / 2 +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 180, 90, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 90, 0, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 0, -90, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 270, 180, false).replace("M", "L") +
                            " z";
                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "noSmoking":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 18750 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var a, dr, iwd2, ihd2, ang, ang2rad, ct, st, m, n, drd2, dang, dang2, swAng, t3, stAng1, stAng2;
                        if (adj < 0) a = 0
                        else if (adj > cnstVal1) a = cnstVal1
                        else a = adj
                        dr = Math.min(w, h) * a / cnstVal2;
                        iwd2 = w / 2 - dr;
                        ihd2 = h / 2 - dr;
                        ang = Math.atan(h / w);
                        //ang2rad = ang*Math.PI/180;
                        ct = ihd2 * Math.cos(ang);
                        st = iwd2 * Math.sin(ang);
                        m = Math.sqrt(ct * ct + st * st); //"mod ct st 0"
                        n = iwd2 * ihd2 / m;
                        drd2 = dr / 2;
                        dang = Math.atan(drd2 / n);
                        dang2 = dang * 2;
                        swAng = -Math.PI + dang2;
                        //t3 = Math.atan(h/w);
                        stAng1 = ang - dang;
                        stAng2 = stAng1 - Math.PI;
                        var ct1, st1, m1, n1, dx1, dy1, x1, y1, y1, y2;
                        ct1 = ihd2 * Math.cos(stAng1);
                        st1 = iwd2 * Math.sin(stAng1);
                        m1 = Math.sqrt(ct1 * ct1 + st1 * st1); //"mod ct1 st1 0"
                        n1 = iwd2 * ihd2 / m1;
                        dx1 = n1 * Math.cos(stAng1);
                        dy1 = n1 * Math.sin(stAng1);
                        x1 = w / 2 + dx1;
                        y1 = h / 2 + dy1;
                        x2 = w / 2 - dx1;
                        y2 = h / 2 - dy1;
                        var stAng1deg = stAng1 * 180 / Math.PI;
                        var stAng2deg = stAng2 * 180 / Math.PI;
                        var swAng2deg = swAng * 180 / Math.PI;
                        var d = "M" + 0 + "," + h / 2 +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") +
                            " z" +
                            "M" + x1 + "," + y1 +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng1deg, (stAng1deg + swAng2deg), false).replace("M", "L") +
                            " z" +
                            "M" + x2 + "," + y2 +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng2deg, (stAng2deg + swAng2deg), false).replace("M", "L") +
                            " z";
                        //console.log("adj: ",adj,"x1:",x1,",y1:",y1," x2:",x2,",y2:",y2,",stAng1:",stAng1,",stAng1deg:",stAng1deg,",stAng2:",stAng2,",stAng2deg:",stAng2deg,",swAng:",swAng,",swAng2deg:",swAng2deg)

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "halfFrame":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, sAdj1_val = 3.5;
                        var sAdj2, sAdj2_val = 3.5;
                        var cnsVal = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var minWH = Math.min(w, h);
                        var maxAdj2 = (cnsVal * w) / minWH;
                        var a1, a2;
                        if (sAdj2_val < 0) a2 = 0
                        else if (sAdj2_val > maxAdj2) a2 = maxAdj2
                        else a2 = sAdj2_val
                        var x1 = (minWH * a2) / cnsVal;
                        var g1 = h * x1 / w;
                        var g2 = h - g1;
                        var maxAdj1 = (cnsVal * g2) / minWH;
                        if (sAdj1_val < 0) a1 = 0
                        else if (sAdj1_val > maxAdj1) a1 = maxAdj1
                        else a1 = sAdj1_val
                        var y1 = minWH * a1 / cnsVal;
                        var dx2 = y1 * w / h;
                        var x2 = w - dx2;
                        var dy2 = x1 * h / w;
                        var y2 = h - dy2;
                        var d = "M0,0" +
                            " L" + w + "," + 0 +
                            " L" + x2 + "," + y1 +
                            " L" + x1 + "," + y1 +
                            " L" + x1 + "," + y2 +
                            " L0," + h + " z";

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        //console.log("w: ",w,", h: ",h,", sAdj1_val: ",sAdj1_val,", sAdj2_val: ",sAdj2_val,",maxAdj1: ",maxAdj1,",maxAdj2: ",maxAdj2)
                        break;
                    case "blockArc":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 180;
                        var sAdj2, adj2 = 0;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) / 60000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) / 60000;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }

                        var stAng, istAng, a3, sw11, sw12, swAng, iswAng;
                        var cd1 = 360;
                        if (adj1 < 0) stAng = 0
                        else if (adj1 > cd1) stAng = cd1
                        else stAng = adj1 //180

                        if (adj2 < 0) istAng = 0
                        else if (adj2 > cd1) istAng = cd1
                        else istAng = adj2 //0

                        if (adj3 < 0) a3 = 0
                        else if (adj3 > cnstVal1) a3 = cnstVal1
                        else a3 = adj3

                        sw11 = istAng - stAng; // -180
                        sw12 = sw11 + cd1; //180
                        swAng = (sw11 > 0) ? sw11 : sw12; //180
                        iswAng = -swAng; //-180

                        var endAng = stAng + swAng;
                        var iendAng = istAng + iswAng;

                        var wt1, ht1, dx1, dy1, x1, y1, stRd, istRd, wd2, hd2, hc, vc;
                        stRd = stAng * (Math.PI) / 180;
                        istRd = istAng * (Math.PI) / 180;
                        wd2 = w / 2;
                        hd2 = h / 2;
                        hc = w / 2;
                        vc = h / 2;
                        if (stAng > 90 && stAng < 270) {
                            wt1 = wd2 * (Math.sin((Math.PI) / 2 - stRd));
                            ht1 = hd2 * (Math.cos((Math.PI) / 2 - stRd));

                            dx1 = wd2 * (Math.cos(Math.atan(ht1 / wt1)));
                            dy1 = hd2 * (Math.sin(Math.atan(ht1 / wt1)));

                            x1 = hc - dx1;
                            y1 = vc - dy1;
                        } else {
                            wt1 = wd2 * (Math.sin(stRd));
                            ht1 = hd2 * (Math.cos(stRd));

                            dx1 = wd2 * (Math.cos(Math.atan(wt1 / ht1)));
                            dy1 = hd2 * (Math.sin(Math.atan(wt1 / ht1)));

                            x1 = hc + dx1;
                            y1 = vc + dy1;
                        }
                        var dr, iwd2, ihd2, wt2, ht2, dx2, dy2, x2, y2;
                        dr = Math.min(w, h) * a3 / cnstVal2;
                        iwd2 = wd2 - dr;
                        ihd2 = hd2 - dr;
                        //console.log("stAng: ",stAng," swAng: ",swAng ," endAng:",endAng)
                        if ((endAng <= 450 && endAng > 270) || ((endAng >= 630 && endAng < 720))) {
                            wt2 = iwd2 * (Math.sin(istRd));
                            ht2 = ihd2 * (Math.cos(istRd));
                            dx2 = iwd2 * (Math.cos(Math.atan(wt2 / ht2)));
                            dy2 = ihd2 * (Math.sin(Math.atan(wt2 / ht2)));
                            x2 = hc + dx2;
                            y2 = vc + dy2;
                        } else {
                            wt2 = iwd2 * (Math.sin((Math.PI) / 2 - istRd));
                            ht2 = ihd2 * (Math.cos((Math.PI) / 2 - istRd));

                            dx2 = iwd2 * (Math.cos(Math.atan(ht2 / wt2)));
                            dy2 = ihd2 * (Math.sin(Math.atan(ht2 / wt2)));
                            x2 = hc - dx2;
                            y2 = vc - dy2;
                        }
                        var d = "M" + x1 + "," + y1 +
                            window.PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, stAng, endAng, false).replace("M", "L") +
                            " L" + x2 + "," + y2 +
                            window.PPTXShapeUtils.shapeArc(wd2, hd2, iwd2, ihd2, istAng, iendAng, false).replace("M", "L") +
                            " z";
                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "bracePair":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 8333 * slideFactor;
                        var cnstVal1 = 25000 * slideFactor;
                        var cnstVal2 = 50000 * slideFactor;
                        var cnstVal3 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var vc = h / 2, cd = 360, cd2 = 180, cd4 = 90, c3d4 = 270, a, x1, x2, x3, x4, y2, y3, y4;
                        if (adj < 0) a = 0
                        else if (adj > cnstVal1) a = cnstVal1
                        else a = adj
                        var minWH = Math.min(w, h);
                        x1 = minWH * a / cnstVal3;
                        x2 = minWH * a / cnstVal2;
                        x3 = w - x2;
                        x4 = w - x1;
                        y2 = vc - x1;
                        y3 = vc + x1;
                        y4 = h - x1;
                        //console.log("w:",w," h:",h," x1:",x1," x2:",x2," x3:",x3," x4:",x4," y2:",y2," y3:",y3," y4:",y4)
                        var d = "M" + x2 + "," + h +
                            window.PPTXShapeUtils.shapeArc(x2, y4, x1, x1, cd4, cd2, false).replace("M", "L") +
                            " L" + x1 + "," + y3 +
                            window.PPTXShapeUtils.shapeArc(0, y3, x1, x1, 0, (-cd4), false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(0, y2, x1, x1, cd4, 0, false).replace("M", "L") +
                            " L" + x1 + "," + x1 +
                            window.PPTXShapeUtils.shapeArc(x2, x1, x1, x1, cd2, c3d4, false).replace("M", "L") +
                            " M" + x3 + "," + 0 +
                            window.PPTXShapeUtils.shapeArc(x3, x1, x1, x1, c3d4, cd, false).replace("M", "L") +
                            " L" + x4 + "," + y2 +
                            window.PPTXShapeUtils.shapeArc(w, y2, x1, x1, cd2, cd4, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w, y3, x1, x1, c3d4, cd2, false).replace("M", "L") +
                            " L" + x4 + "," + y4 +
                            window.PPTXShapeUtils.shapeArc(x3, y4, x1, x1, 0, cd4, false).replace("M", "L");

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "leftBrace":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 8333 * slideFactor;
                        var sAdj2, adj2 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, cd2 = 180, cd4 = 90, c3d4 = 270, a1, a2, q1, q2, q3, y1, y2, y3, y4;
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal2) a2 = cnstVal2
                        else a2 = adj2
                        var minWH = Math.min(w, h);
                        q1 = cnstVal2 - a2;
                        if (q1 < a2) q2 = q1
                        else q2 = a2
                        q3 = q2 / 2;
                        var maxAdj1 = q3 * h / minWH;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        y1 = minWH * a1 / cnstVal2;
                        y3 = h * a2 / cnstVal2;
                        y2 = y3 - y1;
                        y4 = y3 + y1;
                        //console.log("w:",w," h:",h," q1:",q1," q2:",q2," q3:",q3," y1:",y1," y3:",y3," y4:",y4," maxAdj1:",maxAdj1)
                        var d = "M" + w + "," + h +
                            window.PPTXShapeUtils.shapeArc(w, h - y1, w / 2, y1, cd4, cd2, false).replace("M", "L") +
                            " L" + w / 2 + "," + y4 +
                            window.PPTXShapeUtils.shapeArc(0, y4, w / 2, y1, 0, (-cd4), false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(0, y2, w / 2, y1, cd4, 0, false).replace("M", "L") +
                            " L" + w / 2 + "," + y1 +
                            window.PPTXShapeUtils.shapeArc(w, y1, w / 2, y1, cd2, c3d4, false).replace("M", "L");

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "rightBrace":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 8333 * slideFactor;
                        var sAdj2, adj2 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, cd = 360, cd2 = 180, cd4 = 90, c3d4 = 270, a1, a2, q1, q2, q3, y1, y2, y3, y4;
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal2) a2 = cnstVal2
                        else a2 = adj2
                        var minWH = Math.min(w, h);
                        q1 = cnstVal2 - a2;
                        if (q1 < a2) q2 = q1
                        else q2 = a2
                        q3 = q2 / 2;
                        var maxAdj1 = q3 * h / minWH;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        y1 = minWH * a1 / cnstVal2;
                        y3 = h * a2 / cnstVal2;
                        y2 = y3 - y1;
                        y4 = h - y1;
                        //console.log("w:",w," h:",h," q1:",q1," q2:",q2," q3:",q3," y1:",y1," y2:",y2," y3:",y3," y4:",y4," maxAdj1:",maxAdj1)
                        var d = "M" + 0 + "," + 0 +
                            window.PPTXShapeUtils.shapeArc(0, y1, w / 2, y1, c3d4, cd, false).replace("M", "L") +
                            " L" + w / 2 + "," + y2 +
                            window.PPTXShapeUtils.shapeArc(w, y2, w / 2, y1, cd2, cd4, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w, y3 + y1, w / 2, y1, c3d4, cd2, false).replace("M", "L") +
                            " L" + w / 2 + "," + y4 +
                            window.PPTXShapeUtils.shapeArc(0, y4, w / 2, y1, 0, cd4, false).replace("M", "L");

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "bracketPair":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 16667 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var r = w, b = h, cd2 = 180, cd4 = 90, c3d4 = 270, a, x1, x2, y2;
                        if (adj < 0) a = 0
                        else if (adj > cnstVal1) a = cnstVal1
                        else a = adj
                        x1 = Math.min(w, h) * a / cnstVal2;
                        x2 = r - x1;
                        y2 = b - x1;
                        //console.log("w:",w," h:",h," x1:",x1," x2:",x2," y2:",y2)
                        var d = window.PPTXShapeUtils.shapeArc(x1, x1, x1, x1, c3d4, cd2, false) +
                            window.PPTXShapeUtils.shapeArc(x1, y2, x1, x1, cd2, cd4, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(x2, x1, x1, x1, c3d4, (c3d4 + cd4), false) +
                            window.PPTXShapeUtils.shapeArc(x2, y2, x1, x1, 0, cd4, false).replace("M", "L");
                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "leftBracket":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 8333 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var maxAdj = cnstVal1 * h / Math.min(w, h);
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var r = w, b = h, cd2 = 180, cd4 = 90, c3d4 = 270, a, y1, y2;
                        if (adj < 0) a = 0
                        else if (adj > maxAdj) a = maxAdj
                        else a = adj
                        y1 = Math.min(w, h) * a / cnstVal2;
                        if (y1 > w) y1 = w;
                        y2 = b - y1;
                        var d = "M" + r + "," + b +
                            window.PPTXShapeUtils.shapeArc(y1, y2, y1, y1, cd4, cd2, false).replace("M", "L") +
                            " L" + 0 + "," + y1 +
                            window.PPTXShapeUtils.shapeArc(y1, y1, y1, y1, cd2, c3d4, false).replace("M", "L") +
                            " L" + r + "," + 0
                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "rightBracket":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 8333 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var maxAdj = cnstVal1 * h / Math.min(w, h);
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var cd = 360, cd2 = 180, cd4 = 90, c3d4 = 270, a, y1, y2, y3;
                        if (adj < 0) a = 0
                        else if (adj > maxAdj) a = maxAdj
                        else a = adj
                        y1 = Math.min(w, h) * a / cnstVal2;
                        y2 = h - y1;
                        y3 = w - y1;
                        //console.log("w:",w," h:",h," y1:",y1," y2:",y2," y3:",y3)
                        var d = "M" + 0 + "," + h +
                            window.PPTXShapeUtils.shapeArc(y3, y2, y1, y1, cd4, 0, false).replace("M", "L") +
                            //" L"+ r + "," + y2 +
                            " L" + w + "," + h / 2 +
                            window.PPTXShapeUtils.shapeArc(y3, y1, y1, y1, cd, c3d4, false).replace("M", "L") +
                            " L" + 0 + "," + 0
                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "moon":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 0.5;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) / 100000;//*96/914400;;
                        }
                        var hd2, cd2, cd4;

                        hd2 = h / 2;
                        cd2 = 180;
                        cd4 = 90;

                        var adj2 = (1 - adj) * w;
                        var d = "M" + w + "," + h +
                            window.PPTXShapeUtils.shapeArc(w, hd2, w, hd2, cd4, (cd4 + cd2), false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w, hd2, adj2, hd2, (cd4 + cd2), cd4, false).replace("M", "L") +
                            " z";
                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "corner":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, sAdj1_val = 50000 * slideFactor;
                        var sAdj2, sAdj2_val = 50000 * slideFactor;
                        var cnsVal = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var minWH = Math.min(w, h);
                        var maxAdj1 = cnsVal * h / minWH;
                        var maxAdj2 = cnsVal * w / minWH;
                        var a1, a2, x1, dy1, y1;
                        if (sAdj1_val < 0) a1 = 0
                        else if (sAdj1_val > maxAdj1) a1 = maxAdj1
                        else a1 = sAdj1_val

                        if (sAdj2_val < 0) a2 = 0
                        else if (sAdj2_val > maxAdj2) a2 = maxAdj2
                        else a2 = sAdj2_val
                        x1 = minWH * a2 / cnsVal;
                        dy1 = minWH * a1 / cnsVal;
                        y1 = h - dy1;

                        var d = "M0,0" +
                            " L" + x1 + "," + 0 +
                            " L" + x1 + "," + y1 +
                            " L" + w + "," + y1 +
                            " L" + w + "," + h +
                            " L0," + h + " z";

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "diagStripe":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var sAdj1_val = 50000 * slideFactor;
                        var cnsVal = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            sAdj1_val = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var a1, x2, y2;
                        if (sAdj1_val < 0) a1 = 0
                        else if (sAdj1_val > cnsVal) a1 = cnsVal
                        else a1 = sAdj1_val
                        x2 = w * a1 / cnsVal;
                        y2 = h * a1 / cnsVal;
                        var d = "M" + 0 + "," + y2 +
                            " L" + x2 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + 0 + "," + h + " z";

                        result += "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "gear6":
                    case "gear9":
                        txtRotate = 0;
                        var gearNum = shapType.substr(4), d;
                        if (gearNum == "6") {
                            d = window.PPTXShapeUtils.shapeGear(w, h / 3.5, parseInt(gearNum));
                        } else { //gearNum=="9"
                            d = window.PPTXShapeUtils.shapeGear(w, h / 3.5, parseInt(gearNum));
                        }
                        result += "<path   d='" + d + "' transform='rotate(20," + (3 / 7) * h + "," + (3 / 7) * h + ")' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "bentConnector3":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var shapAdjst_val = 0.5;
                        if (shapAdjst !== undefined) {
                            shapAdjst_val = parseInt(shapAdjst.substr(4)) / 100000;
                            // if (isFlipV) {
                            //     result += " <polyline points='" + w + " 0," + ((1 - shapAdjst_val) * w) + " 0," + ((1 - shapAdjst_val) * w) + " " + h + ",0 " + h + "' fill='transparent'" +
                            //         "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                            // } else {
                            result += " <polyline points='0 0," + (shapAdjst_val) * w + " 0," + (shapAdjst_val) * w + " " + h + "," + w + " " + h + "' fill='transparent'" +
                                "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                            //}
                            if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                                result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                            }
                            if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                                result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                            }
                            result += "/>";
                        }
                        break;
                    case "plus":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj1 = 0.25;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) / 100000;

                        }
                        var adj2 = (1 - adj1);
                        result += " <polygon points='" + adj1 * w + " 0," + adj1 * w + " " + adj1 * h + ",0 " + adj1 * h + ",0 " + adj2 * h + "," +
                            adj1 * w + " " + adj2 * h + "," + adj1 * w + " " + h + "," + adj2 * w + " " + h + "," + adj2 * w + " " + adj2 * h + "," + w + " " + adj2 * h + "," +
                            +w + " " + adj1 * h + "," + adj2 * w + " " + adj1 * h + "," + adj2 * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "teardrop":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj1 = 100000 * slideFactor;
                        var cnsVal1 = adj1;
                        var cnsVal2 = 200000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var a1, r2, tw, th, sw, sh, dx1, dy1, x1, y1, x2, y2, rd45;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnsVal2) a1 = cnsVal2
                        else a1 = adj1
                        r2 = Math.sqrt(2);
                        tw = r2 * (w / 2);
                        th = r2 * (h / 2);
                        sw = (tw * a1) / cnsVal1;
                        sh = (th * a1) / cnsVal1;
                        rd45 = (45 * (Math.PI) / 180);
                        dx1 = sw * (Math.cos(rd45));
                        dy1 = sh * (Math.cos(rd45));
                        x1 = (w / 2) + dx1;
                        y1 = (h / 2) - dy1;
                        x2 = ((w / 2) + x1) / 2;
                        y2 = ((h / 2) + y1) / 2;

                        var d_val = window.PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false) +
                            "Q " + x2 + ",0 " + x1 + "," + y1 +
                            "Q " + w + "," + y2 + " " + w + "," + h / 2 +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") + " z";
                        result += "<path   d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        // console.log("shapAdjst: ",shapAdjst,", adj1: ",adj1);
                        break;
                    case "plaque":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj1 = 16667 * slideFactor;
                        var cnsVal1 = 50000 * slideFactor;
                        var cnsVal2 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var a1, x1, x2, y2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnsVal1) a1 = cnsVal1
                        else a1 = adj1
                        x1 = a1 * (Math.min(w, h)) / cnsVal2;
                        x2 = w - x1;
                        y2 = h - x1;

                        var d_val = "M0," + x1 +
                            window.PPTXShapeUtils.shapeArc(0, 0, x1, x1, 90, 0, false).replace("M", "L") +
                            " L" + x2 + "," + 0 +
                            window.PPTXShapeUtils.shapeArc(w, 0, x1, x1, 180, 90, false).replace("M", "L") +
                            " L" + w + "," + y2 +
                            window.PPTXShapeUtils.shapeArc(w, h, x1, x1, 270, 180, false).replace("M", "L") +
                            " L" + x1 + "," + h +
                            window.PPTXShapeUtils.shapeArc(0, h, x1, x1, 0, -90, false).replace("M", "L") + " z";
                        result += "<path   d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "sun":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var refr = slideFactor;
                        var adj1 = 25000 * refr;
                        var cnstVal1 = 12500 * refr;
                        var cnstVal2 = 46875 * refr;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        var a1;
                        if (adj1 < cnstVal1) a1 = cnstVal1
                        else if (adj1 > cnstVal2) a1 = cnstVal2
                        else a1 = adj1

                        var cnstVa3 = 50000 * refr;
                        var cnstVa4 = 100000 * refr;
                        var g0 = cnstVa3 - a1,
                            g1 = g0 * (30274 * refr) / (32768 * refr),
                            g2 = g0 * (12540 * refr) / (32768 * refr),
                            g3 = g1 + cnstVa3,
                            g4 = g2 + cnstVa3,
                            g5 = cnstVa3 - g1,
                            g6 = cnstVa3 - g2,
                            g7 = g0 * (23170 * refr) / (32768 * refr),
                            g8 = cnstVa3 + g7,
                            g9 = cnstVa3 - g7,
                            g10 = g5 * 3 / 4,
                            g11 = g6 * 3 / 4,
                            g12 = g10 + 3662 * refr,
                            g13 = g11 + 36620 * refr,
                            g14 = g11 + 12500 * refr,
                            g15 = cnstVa4 - g10,
                            g16 = cnstVa4 - g12,
                            g17 = cnstVa4 - g13,
                            g18 = cnstVa4 - g14,
                            ox1 = w * (18436 * refr) / (21600 * refr),
                            oy1 = h * (3163 * refr) / (21600 * refr),
                            ox2 = w * (3163 * refr) / (21600 * refr),
                            oy2 = h * (18436 * refr) / (21600 * refr),
                            x8 = w * g8 / cnstVa4,
                            x9 = w * g9 / cnstVa4,
                            x10 = w * g10 / cnstVa4,
                            x12 = w * g12 / cnstVa4,
                            x13 = w * g13 / cnstVa4,
                            x14 = w * g14 / cnstVa4,
                            x15 = w * g15 / cnstVa4,
                            x16 = w * g16 / cnstVa4,
                            x17 = w * g17 / cnstVa4,
                            x18 = w * g18 / cnstVa4,
                            x19 = w * a1 / cnstVa4,
                            wR = w * g0 / cnstVa4,
                            hR = h * g0 / cnstVa4,
                            y8 = h * g8 / cnstVa4,
                            y9 = h * g9 / cnstVa4,
                            y10 = h * g10 / cnstVa4,
                            y12 = h * g12 / cnstVa4,
                            y13 = h * g13 / cnstVa4,
                            y14 = h * g14 / cnstVa4,
                            y15 = h * g15 / cnstVa4,
                            y16 = h * g16 / cnstVa4,
                            y17 = h * g17 / cnstVa4,
                            y18 = h * g18 / cnstVa4;

                        var d_val = "M" + w + "," + h / 2 +
                            " L" + x15 + "," + y18 +
                            " L" + x15 + "," + y14 +
                            "z" +
                            " M" + ox1 + "," + oy1 +
                            " L" + x16 + "," + y17 +
                            " L" + x13 + "," + y12 +
                            "z" +
                            " M" + w / 2 + "," + 0 +
                            " L" + x18 + "," + y10 +
                            " L" + x14 + "," + y10 +
                            "z" +
                            " M" + ox2 + "," + oy1 +
                            " L" + x17 + "," + y12 +
                            " L" + x12 + "," + y17 +
                            "z" +
                            " M" + 0 + "," + h / 2 +
                            " L" + x10 + "," + y14 +
                            " L" + x10 + "," + y18 +
                            "z" +
                            " M" + ox2 + "," + oy2 +
                            " L" + x12 + "," + y13 +
                            " L" + x17 + "," + y16 +
                            "z" +
                            " M" + w / 2 + "," + h +
                            " L" + x14 + "," + y15 +
                            " L" + x18 + "," + y15 +
                            "z" +
                            " M" + ox1 + "," + oy2 +
                            " L" + x13 + "," + y16 +
                            " L" + x16 + "," + y13 +
                            " z" +
                            " M" + x19 + "," + h / 2 +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, wR, hR, 180, 540, false).replace("M", "L") +
                            " z";
                        //console.log("adj1: ",adj1,d_val);
                        result += "<path   d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";


                        break;
                    case "heart":
                        var dx1, dx2, x1, x2, x3, x4, y1;
                        dx1 = w * 49 / 48;
                        dx2 = w * 10 / 48
                        x1 = w / 2 - dx1
                        x2 = w / 2 - dx2
                        x3 = w / 2 + dx2
                        x4 = w / 2 + dx1
                        y1 = -h / 3;
                        var d_val = "M" + w / 2 + "," + h / 4 +
                            "C" + x3 + "," + y1 + " " + x4 + "," + h / 4 + " " + w / 2 + "," + h +
                            "C" + x1 + "," + h / 4 + " " + x2 + "," + y1 + " " + w / 2 + "," + h / 4 + " z";

                        result += "<path   d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "lightningBolt":
                        var x1 = w * 5022 / 21600,
                            x2 = w * 11050 / 21600,
                            x3 = w * 8472 / 21600,
                            x4 = w * 8757 / 21600,
                            x5 = w * 10012 / 21600,
                            x6 = w * 14767 / 21600,
                            x7 = w * 12222 / 21600,
                            x8 = w * 12860 / 21600,
                            x9 = w * 13917 / 21600,
                            x10 = w * 7602 / 21600,
                            x11 = w * 16577 / 21600,
                            y1 = h * 3890 / 21600,
                            y2 = h * 6080 / 21600,
                            y3 = h * 6797 / 21600,
                            y4 = h * 7437 / 21600,
                            y5 = h * 12877 / 21600,
                            y6 = h * 9705 / 21600,
                            y7 = h * 12007 / 21600,
                            y8 = h * 13987 / 21600,
                            y9 = h * 8382 / 21600,
                            y10 = h * 14277 / 21600,
                            y11 = h * 14915 / 21600;

                        var d_val = "M" + x3 + "," + 0 +
                            " L" + x8 + "," + y2 +
                            " L" + x2 + "," + y3 +
                            " L" + x11 + "," + y7 +
                            " L" + x6 + "," + y5 +
                            " L" + w + "," + h +
                            " L" + x5 + "," + y11 +
                            " L" + x7 + "," + y8 +
                            " L" + x1 + "," + y6 +
                            " L" + x10 + "," + y9 +
                            " L" + 0 + "," + y1 + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "cube":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var refr = slideFactor;
                        var adj = 25000 * refr;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        var d_val;
                        var cnstVal2 = 100000 * refr;
                        var ss = Math.min(w, h);
                        var a, y1, y4, x4;
                        a = (adj < 0) ? 0 : (adj > cnstVal2) ? cnstVal2 : adj;
                        y1 = ss * a / cnstVal2;
                        y4 = h - y1;
                        x4 = w - y1;
                        d_val = "M" + 0 + "," + y1 +
                            " L" + y1 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + y4 +
                            " L" + x4 + "," + h +
                            " L" + 0 + "," + h +
                            " z" +
                            "M" + 0 + "," + y1 +
                            " L" + x4 + "," + y1 +
                            " M" + x4 + "," + y1 +
                            " L" + w + "," + 0 +
                            "M" + x4 + "," + y1 +
                            " L" + x4 + "," + h;

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "bevel":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var refr = slideFactor;
                        var adj = 12500 * refr;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        var d_val;
                        var cnstVal1 = 50000 * refr;
                        var cnstVal2 = 100000 * refr;
                        var ss = Math.min(w, h);
                        var a, x1, x2, y2;
                        a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                        x1 = ss * a / cnstVal2;
                        x2 = w - x1;
                        y2 = h - x1;
                        d_val = "M" + 0 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + h +
                            " L" + 0 + "," + h +
                            " z" +
                            " M" + x1 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + x2 + "," + y2 +
                            " L" + x1 + "," + y2 +
                            " z" +
                            " M" + 0 + "," + 0 +
                            " L" + x1 + "," + x1 +
                            " M" + 0 + "," + h +
                            " L" + x1 + "," + y2 +
                            " M" + w + "," + 0 +
                            " L" + x2 + "," + x1 +
                            " M" + w + "," + h +
                            " L" + x2 + "," + y2;

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "foldedCorner":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var refr = slideFactor;
                        var adj = 16667 * refr;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        var d_val;
                        var cnstVal1 = 50000 * refr;
                        var cnstVal2 = 100000 * refr;
                        var ss = Math.min(w, h);
                        var a, dy2, dy1, x1, x2, y2, y1;
                        a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                        dy2 = ss * a / cnstVal2;
                        dy1 = dy2 / 5;
                        x1 = w - dy2;
                        x2 = x1 + dy1;
                        y2 = h - dy2;
                        y1 = y2 + dy1;
                        d_val = "M" + x1 + "," + h +
                            " L" + x2 + "," + y1 +
                            " L" + w + "," + y2 +
                            " L" + x1 + "," + h +
                            " L" + 0 + "," + h +
                            " L" + 0 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + y2;

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "cloud":
                    case "cloudCallout":
                        var d1 = window.PPTXCalloutShapes.genCloudCallout(w, h, node, slideFactor, shapType);
                        result += "<path d='" + d1 + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "smileyFace":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var refr = slideFactor;
                        var adj = 4653 * refr;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        var d_val;
                        var cnstVal1 = 50000 * refr;
                        var cnstVal2 = 100000 * refr;
                        var cnstVal3 = 4653 * refr;
                        var ss = Math.min(w, h);
                        var a, x1, x2, x3, x4, y1, y3, dy2, y2, y4, dy3, y5, wR, hR, wd2, hd2;
                        wd2 = w / 2;
                        hd2 = h / 2;
                        a = (adj < -cnstVal3) ? -cnstVal3 : (adj > cnstVal3) ? cnstVal3 : adj;
                        x1 = w * 4969 / 21699;
                        x2 = w * 6215 / 21600;
                        x3 = w * 13135 / 21600;
                        x4 = w * 16640 / 21600;
                        y1 = h * 7570 / 21600;
                        y3 = h * 16515 / 21600;
                        dy2 = h * a / cnstVal2;
                        y2 = y3 - dy2;
                        y4 = y3 + dy2;
                        dy3 = h * a / cnstVal1;
                        y5 = y4 + dy3;
                        wR = w * 1125 / 21600;
                        hR = h * 1125 / 21600;
                        var cX1 = x2 - wR * Math.cos(Math.PI);
                        var cY1 = y1 - hR * Math.sin(Math.PI);
                        var cX2 = x3 - wR * Math.cos(Math.PI);
                        d_val = //eyes
                            window.PPTXShapeUtils.shapeArc(cX1, cY1, wR, hR, 180, 540, false) +
                            window.PPTXShapeUtils.shapeArc(cX2, cY1, wR, hR, 180, 540, false) +
                            //mouth
                            " M" + x1 + "," + y2 +
                            " Q" + wd2 + "," + y5 + " " + x4 + "," + y2 +
                            " Q" + wd2 + "," + y5 + " " + x1 + "," + y2 +
                            //head
                            " M" + 0 + "," + hd2 +
                            window.PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, 180, 540, false).replace("M", "L") +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "verticalScroll":
                    case "horizontalScroll":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var refr = slideFactor;
                        var adj = 12500 * refr;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        var d_val;
                        var cnstVal1 = 25000 * refr;
                        var cnstVal2 = 100000 * refr;
                        var ss = Math.min(w, h);
                        var t = 0, l = 0, b = h, r = w;
                        var a, ch, ch2, ch4;
                        a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                        ch = ss * a / cnstVal2;
                        ch2 = ch / 2;
                        ch4 = ch / 4;
                        if (shapType == "verticalScroll") {
                            var x3, x4, x6, x7, x5, y3, y4;
                            x3 = ch + ch2;
                            x4 = ch + ch;
                            x6 = r - ch;
                            x7 = r - ch2;
                            x5 = x6 - ch2;
                            y3 = b - ch;
                            y4 = b - ch2;

                            d_val = "M" + ch + "," + y3 +
                                " L" + ch + "," + ch2 +
                                window.PPTXShapeUtils.shapeArc(x3, ch2, ch2, ch2, 180, 270, false).replace("M", "L") +
                                " L" + x7 + "," + t +
                                window.PPTXShapeUtils.shapeArc(x7, ch2, ch2, ch2, 270, 450, false).replace("M", "L") +
                                " L" + x6 + "," + ch +
                                " L" + x6 + "," + y4 +
                                window.PPTXShapeUtils.shapeArc(x5, y4, ch2, ch2, 0, 90, false).replace("M", "L") +
                                " L" + ch2 + "," + b +
                                window.PPTXShapeUtils.shapeArc(ch2, y4, ch2, ch2, 90, 270, false).replace("M", "L") +
                                " z" +
                                " M" + x3 + "," + t +
                                window.PPTXShapeUtils.shapeArc(x3, ch2, ch2, ch2, 270, 450, false).replace("M", "L") +
                                window.PPTXShapeUtils.shapeArc(x3, x3 / 2, ch4, ch4, 90, 270, false).replace("M", "L") +
                                " L" + x4 + "," + ch2 +
                                " M" + x6 + "," + ch +
                                " L" + x3 + "," + ch +
                                " M" + ch + "," + y4 +
                                window.PPTXShapeUtils.shapeArc(ch2, y4, ch2, ch2, 0, 270, false).replace("M", "L") +
                                window.PPTXShapeUtils.shapeArc(ch2, (y4 + y3) / 2, ch4, ch4, 270, 450, false).replace("M", "L") +
                                " z" +
                                " M" + ch + "," + y4 +
                                " L" + ch + "," + y3;
                        } else if (shapType == "horizontalScroll") {
                            var y3, y4, y6, y7, y5, x3, x4;
                            y3 = ch + ch2;
                            y4 = ch + ch;
                            y6 = b - ch;
                            y7 = b - ch2;
                            y5 = y6 - ch2;
                            x3 = r - ch;
                            x4 = r - ch2;

                            d_val = "M" + l + "," + y3 +
                                window.PPTXShapeUtils.shapeArc(ch2, y3, ch2, ch2, 180, 270, false).replace("M", "L") +
                                " L" + x3 + "," + ch +
                                " L" + x3 + "," + ch2 +
                                window.PPTXShapeUtils.shapeArc(x4, ch2, ch2, ch2, 180, 360, false).replace("M", "L") +
                                " L" + r + "," + y5 +
                                window.PPTXShapeUtils.shapeArc(x4, y5, ch2, ch2, 0, 90, false).replace("M", "L") +
                                " L" + ch + "," + y6 +
                                " L" + ch + "," + y7 +
                                window.PPTXShapeUtils.shapeArc(ch2, y7, ch2, ch2, 0, 180, false).replace("M", "L") +
                                " z" +
                                "M" + x4 + "," + ch +
                                window.PPTXShapeUtils.shapeArc(x4, ch2, ch2, ch2, 90, -180, false).replace("M", "L") +
                                window.PPTXShapeUtils.shapeArc((x3 + x4) / 2, ch2, ch4, ch4, 180, 0, false).replace("M", "L") +
                                " z" +
                                " M" + x4 + "," + ch +
                                " L" + x3 + "," + ch +
                                " M" + ch2 + "," + y4 +
                                " L" + ch2 + "," + y3 +
                                window.PPTXShapeUtils.shapeArc(y3 / 2, y3, ch4, ch4, 180, 360, false).replace("M", "L") +
                                window.PPTXShapeUtils.shapeArc(ch2, y3, ch2, ch2, 0, 180, false).replace("M", "L") +
                                " M" + ch + "," + y3 +
                                " L" + ch + "," + y6;
                        }

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "wedgeEllipseCallout":
                        var d_val = window.PPTXCalloutShapes.genWedgeEllipseCallout(w, h, node, slideFactor);
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "wedgeRectCallout":
                        var d_val = window.PPTXCalloutShapes.genWedgeRectCallout(w, h, node, slideFactor);
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "wedgeRoundRectCallout":
                        var d_val = window.PPTXCalloutShapes.genWedgeRoundRectCallout(w, h, node, slideFactor);
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "accentBorderCallout1":
                    case "accentBorderCallout2":
                    case "accentBorderCallout3":
                    case "borderCallout1":
                    case "borderCallout2":
                    case "borderCallout3":
                    case "accentCallout1":
                    case "accentCallout2":
                    case "accentCallout3":
                    case "callout1":
                    case "callout2":
                    case "callout3":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var refr = slideFactor;
                        var sAdj1, adj1 = 18750 * refr;
                        var sAdj2, adj2 = -8333 * refr;
                        var sAdj3, adj3 = 18750 * refr;
                        var sAdj4, adj4 = -16667 * refr;
                        var sAdj5, adj5 = 100000 * refr;
                        var sAdj6, adj6 = -16667 * refr;
                        var sAdj7, adj7 = 112963 * refr;
                        var sAdj8, adj8 = -8333 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * refr;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * refr;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * refr;
                                } else if (sAdj_name == "adj6") {
                                    sAdj6 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj6 = parseInt(sAdj6.substr(4)) * refr;
                                } else if (sAdj_name == "adj7") {
                                    sAdj7 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj7 = parseInt(sAdj7.substr(4)) * refr;
                                } else if (sAdj_name == "adj8") {
                                    sAdj8 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj8 = parseInt(sAdj8.substr(4)) * refr;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal1 = 100000 * refr;
                        var isBorder = true;
                        switch (shapType) {
                            case "borderCallout1":
                            case "callout1":
                                if (shapType == "borderCallout1") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 112500 * refr;
                                    adj4 = -38333 * refr;
                                }
                                var y1, x1, y2, x2;
                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +
                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2;
                                break;
                            case "borderCallout2":
                            case "callout2":
                                if (shapType == "borderCallout2") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;

                                    adj5 = 112500 * refr;
                                    adj6 = -46667 * refr;
                                }
                                var y1, x1, y2, x2, y3, x3;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;

                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +

                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +

                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2;

                                break;
                            case "borderCallout3":
                            case "callout3":
                                if (shapType == "borderCallout3") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;

                                    adj5 = 100000 * refr;
                                    adj6 = -16667 * refr;

                                    adj7 = 112963 * refr;
                                    adj8 = -8333 * refr;
                                }
                                var y1, x1, y2, x2, y3, x3, y4, x4;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;

                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;

                                y4 = h * adj7 / cnstVal1;
                                x4 = w * adj8 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +

                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +

                                    " L" + x3 + "," + y3 +

                                    " L" + x4 + "," + y4 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2;
                                break;
                            case "accentBorderCallout1":
                            case "accentCallout1":
                                if (shapType == "accentBorderCallout1") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }

                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 112500 * refr;
                                    adj4 = -38333 * refr;
                                }
                                var y1, x1, y2, x2;
                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +

                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +

                                    " M" + x1 + "," + 0 +
                                    " L" + x1 + "," + h;
                                break;
                            case "accentBorderCallout2":
                            case "accentCallout2":
                                if (shapType == "accentBorderCallout2") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;
                                    adj5 = 112500 * refr;
                                    adj6 = -46667 * refr;
                                }
                                var y1, x1, y2, x2, y3, x3;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +

                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2 +

                                    " M" + x1 + "," + 0 +
                                    " L" + x1 + "," + h;

                                break;
                            case "accentBorderCallout3":
                            case "accentCallout3":
                                if (shapType == "accentBorderCallout3") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }
                                isBorder = true;
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;
                                    adj5 = 100000 * refr;
                                    adj6 = -16667 * refr;
                                    adj7 = 112963 * refr;
                                    adj8 = -8333 * refr;
                                }
                                var y1, x1, y2, x2, y3, x3, y4, x4;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                y4 = h * adj7 / cnstVal1;
                                x4 = w * adj8 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    " z" +

                                    " M" + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x4 + "," + y4 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2 +

                                    " M" + x1 + "," + 0 +
                                    " L" + x1 + "," + h;
                                break;
                        }

                        //console.log("shapType: ", shapType, ",isBorder:", isBorder)
                        //if(isBorder){
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        //}else{
                        //    result += "<path d='"+d_val+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                        //        "' stroke='none' />";

                        //}
                        break;
                    case "leftRightRibbon":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var refr = slideFactor;
                        var sAdj1, adj1 = 50000 * refr;
                        var sAdj2, adj2 = 50000 * refr;
                        var sAdj3, adj3 = 16667 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * refr;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal1 = 33333 * refr;
                        var cnstVal2 = 100000 * refr;
                        var cnstVal3 = 200000 * refr;
                        var cnstVal4 = 400000 * refr;
                        var ss = Math.min(w, h);
                        var a3, maxAdj1, a1, w1, maxAdj2, a2, x1, x4, dy1, dy2, ly1, ry4, ly2, ry3, ly4, ry1,
                            ly3, ry2, hR, x2, x3, y1, y2, wd32 = w / 32, vc = h / 2, hc = w / 2;

                        a3 = (adj3 < 0) ? 0 : (adj3 > cnstVal1) ? cnstVal1 : adj3;
                        maxAdj1 = cnstVal2 - a3;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        w1 = hc - wd32;
                        maxAdj2 = cnstVal2 * w1 / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        x1 = ss * a2 / cnstVal2;
                        x4 = w - x1;
                        dy1 = h * a1 / cnstVal3;
                        dy2 = h * a3 / -cnstVal3;
                        ly1 = vc + dy2 - dy1;
                        ry4 = vc + dy1 - dy2;
                        ly2 = ly1 + dy1;
                        ry3 = h - ly2;
                        ly4 = ly2 * 2;
                        ry1 = h - ly4;
                        ly3 = ly4 - ly1;
                        ry2 = h - ly3;
                        hR = a3 * ss / cnstVal4;
                        x2 = hc - wd32;
                        x3 = hc + wd32;
                        y1 = ly1 + hR;
                        y2 = ry2 - hR;

                        d_val = "M" + 0 + "," + ly2 +
                            "L" + x1 + "," + 0 +
                            "L" + x1 + "," + ly1 +
                            "L" + hc + "," + ly1 +
                            window.PPTXShapeUtils.shapeArc(hc, y1, wd32, hR, 270, 450, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(hc, y2, wd32, hR, 270, 90, false).replace("M", "L") +
                            "L" + x4 + "," + ry2 +
                            "L" + x4 + "," + ry1 +
                            "L" + w + "," + ry3 +
                            "L" + x4 + "," + h +
                            "L" + x4 + "," + ry4 +
                            "L" + hc + "," + ry4 +
                            window.PPTXShapeUtils.shapeArc(hc, ry4 - hR, wd32, hR, 90, 180, false).replace("M", "L") +
                            "L" + x2 + "," + ly3 +
                            "L" + x1 + "," + ly3 +
                            "L" + x1 + "," + ly4 +
                            " z" +
                            "M" + x3 + "," + y1 +
                            "L" + x3 + "," + ry2 +
                            "M" + x2 + "," + y2 +
                            "L" + x2 + "," + ly3;

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "ribbon":
                    case "ribbon2":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 16667 * slideFactor;
                        var sAdj2, adj2 = 50000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal1 = 25000 * slideFactor;
                        var cnstVal2 = 33333 * slideFactor;
                        var cnstVal3 = 75000 * slideFactor;
                        var cnstVal4 = 100000 * slideFactor;
                        var cnstVal5 = 200000 * slideFactor;
                        var cnstVal6 = 400000 * slideFactor;
                        var hc = w / 2, t = 0, l = 0, b = h, r = w, wd8 = w / 8, wd32 = w / 32;
                        var a1, a2, x10, dx2, x2, x9, x3, x8, x5, x6, x4, x7, y1, y2, y4, y3, hR, y6;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        a2 = (adj2 < cnstVal1) ? cnstVal1 : (adj2 > cnstVal3) ? cnstVal3 : adj2;
                        x10 = r - wd8;
                        dx2 = w * a2 / cnstVal5;
                        x2 = hc - dx2;
                        x9 = hc + dx2;
                        x3 = x2 + wd32;
                        x8 = x9 - wd32;
                        x5 = x2 + wd8;
                        x6 = x9 - wd8;
                        x4 = x5 - wd32;
                        x7 = x6 + wd32;
                        hR = h * a1 / cnstVal6;
                        if (shapType == "ribbon2") {
                            var dy1, dy2, y7;
                            dy1 = h * a1 / cnstVal5;
                            y1 = b - dy1;
                            dy2 = h * a1 / cnstVal4;
                            y2 = b - dy2;
                            y4 = t + dy2;
                            y3 = (y4 + b) / 2;
                            y6 = b - hR;///////////////////
                            y7 = y1 - hR;

                            d_val = "M" + l + "," + b +
                                " L" + wd8 + "," + y3 +
                                " L" + l + "," + y4 +
                                " L" + x2 + "," + y4 +
                                " L" + x2 + "," + hR +
                                window.PPTXShapeUtils.shapeArc(x3, hR, wd32, hR, 180, 270, false).replace("M", "L") +
                                " L" + x8 + "," + t +
                                window.PPTXShapeUtils.shapeArc(x8, hR, wd32, hR, 270, 360, false).replace("M", "L") +
                                " L" + x9 + "," + y4 +
                                " L" + x9 + "," + y4 +
                                " L" + r + "," + y4 +
                                " L" + x10 + "," + y3 +
                                " L" + r + "," + b +
                                " L" + x7 + "," + b +
                                window.PPTXShapeUtils.shapeArc(x7, y6, wd32, hR, 90, 270, false).replace("M", "L") +
                                " L" + x8 + "," + y1 +
                                window.PPTXShapeUtils.shapeArc(x8, y7, wd32, hR, 90, -90, false).replace("M", "L") +
                                " L" + x3 + "," + y2 +
                                window.PPTXShapeUtils.shapeArc(x3, y7, wd32, hR, 270, 90, false).replace("M", "L") +
                                " L" + x4 + "," + y1 +
                                window.PPTXShapeUtils.shapeArc(x4, y6, wd32, hR, 270, 450, false).replace("M", "L") +
                                " z" +
                                " M" + x5 + "," + y2 +
                                " L" + x5 + "," + y6 +
                                "M" + x6 + "," + y6 +
                                " L" + x6 + "," + y2 +
                                "M" + x2 + "," + y7 +
                                " L" + x2 + "," + y4 +
                                "M" + x9 + "," + y4 +
                                " L" + x9 + "," + y7;
                        } else if (shapType == "ribbon") {
                            var y5;
                            y1 = h * a1 / cnstVal5;
                            y2 = h * a1 / cnstVal4;
                            y4 = b - y2;
                            y3 = y4 / 2;
                            y5 = b - hR; ///////////////////////
                            y6 = y2 - hR;
                            d_val = "M" + l + "," + t +
                                " L" + x4 + "," + t +
                                window.PPTXShapeUtils.shapeArc(x4, hR, wd32, hR, 270, 450, false).replace("M", "L") +
                                " L" + x3 + "," + y1 +
                                window.PPTXShapeUtils.shapeArc(x3, y6, wd32, hR, 270, 90, false).replace("M", "L") +
                                " L" + x8 + "," + y2 +
                                window.PPTXShapeUtils.shapeArc(x8, y6, wd32, hR, 90, -90, false).replace("M", "L") +
                                " L" + x7 + "," + y1 +
                                window.PPTXShapeUtils.shapeArc(x7, hR, wd32, hR, 90, 270, false).replace("M", "L") +
                                " L" + r + "," + t +
                                " L" + x10 + "," + y3 +
                                " L" + r + "," + y4 +
                                " L" + x9 + "," + y4 +
                                " L" + x9 + "," + y5 +
                                window.PPTXShapeUtils.shapeArc(x8, y5, wd32, hR, 0, 90, false).replace("M", "L") +
                                " L" + x3 + "," + b +
                                window.PPTXShapeUtils.shapeArc(x3, y5, wd32, hR, 90, 180, false).replace("M", "L") +
                                " L" + x2 + "," + y4 +
                                " L" + l + "," + y4 +
                                " L" + wd8 + "," + y3 +
                                " z" +
                                " M" + x5 + "," + hR +
                                " L" + x5 + "," + y2 +
                                "M" + x6 + "," + y2 +
                                " L" + x6 + "," + hR +
                                "M" + x2 + "," + y4 +
                                " L" + x2 + "," + y6 +
                                "M" + x9 + "," + y6 +
                                " L" + x9 + "," + y4;
                        }
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "doubleWave":
                    case "wave":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = (shapType == "doubleWave") ? 6250 * slideFactor : 12500 * slideFactor;
                        var sAdj2, adj2 = 0;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal2 = -10000 * slideFactor;
                        var cnstVal3 = 50000 * slideFactor;
                        var cnstVal4 = 100000 * slideFactor;
                        var hc = w / 2, t = 0, l = 0, b = h, r = w, wd8 = w / 8, wd32 = w / 32;
                        if (shapType == "doubleWave") {
                            var cnstVal1 = 12500 * slideFactor;
                            var a1, a2, y1, dy2, y2, y3, y4, y5, y6, of2, dx2, x2, dx8, x8, dx3, x3, dx4, x4, x5, x6, x7, x9, x15, x10, x11, x12, x13, x14;
                            a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
                            a2 = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
                            y1 = h * a1 / cnstVal4;
                            dy2 = y1 * 10 / 3;
                            y2 = y1 - dy2;
                            y3 = y1 + dy2;
                            y4 = b - y1;
                            y5 = y4 - dy2;
                            y6 = y4 + dy2;
                            of2 = w * a2 / cnstVal3;
                            dx2 = (of2 > 0) ? 0 : of2;
                            x2 = l - dx2;
                            dx8 = (of2 > 0) ? of2 : 0;
                            x8 = r - dx8;
                            dx3 = (dx2 + x8) / 6;
                            x3 = x2 + dx3;
                            dx4 = (dx2 + x8) / 3;
                            x4 = x2 + dx4;
                            x5 = (x2 + x8) / 2;
                            x6 = x5 + dx3;
                            x7 = (x6 + x8) / 2;
                            x9 = l + dx8;
                            x15 = r + dx2;
                            x10 = x9 + dx3;
                            x11 = x9 + dx4;
                            x12 = (x9 + x15) / 2;
                            x13 = x12 + dx3;
                            x14 = (x13 + x15) / 2;

                            d_val = "M" + x2 + "," + y1 +
                                " C" + x3 + "," + y2 + " " + x4 + "," + y3 + " " + x5 + "," + y1 +
                                " C" + x6 + "," + y2 + " " + x7 + "," + y3 + " " + x8 + "," + y1 +
                                " L" + x15 + "," + y4 +
                                " C" + x14 + "," + y6 + " " + x13 + "," + y5 + " " + x12 + "," + y4 +
                                " C" + x11 + "," + y6 + " " + x10 + "," + y5 + " " + x9 + "," + y4 +
                                " z";
                        } else if (shapType == "wave") {
                            var cnstVal5 = 20000 * slideFactor;
                            var a1, a2, y1, dy2, y2, y3, y4, y5, y6, of2, dx2, x2, dx5, x5, dx3, x3, x4, x6, x10, x7, x8;
                            a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
                            a2 = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
                            y1 = h * a1 / cnstVal4;
                            dy2 = y1 * 10 / 3;
                            y2 = y1 - dy2;
                            y3 = y1 + dy2;
                            y4 = b - y1;
                            y5 = y4 - dy2;
                            y6 = y4 + dy2;
                            of2 = w * a2 / cnstVal3;
                            dx2 = (of2 > 0) ? 0 : of2;
                            x2 = l - dx2;
                            dx5 = (of2 > 0) ? of2 : 0;
                            x5 = r - dx5;
                            dx3 = (dx2 + x5) / 3;
                            x3 = x2 + dx3;
                            x4 = (x3 + x5) / 2;
                            x6 = l + dx5;
                            x10 = r + dx2;
                            x7 = x6 + dx3;
                            x8 = (x7 + x10) / 2;

                            d_val = "M" + x2 + "," + y1 +
                                " C" + x3 + "," + y2 + " " + x4 + "," + y3 + " " + x5 + "," + y1 +
                                " L" + x10 + "," + y4 +
                                " C" + x8 + "," + y6 + " " + x7 + "," + y5 + " " + x6 + "," + y4 +
                                " z";
                        }
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "ellipseRibbon":
                    case "ellipseRibbon2":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 50000 * slideFactor;
                        var sAdj3, adj3 = 12500 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var d_val;
                        var cnstVal1 = 25000 * slideFactor;
                        var cnstVal3 = 75000 * slideFactor;
                        var cnstVal4 = 100000 * slideFactor;
                        var cnstVal5 = 200000 * slideFactor;
                        var hc = w / 2, t = 0, l = 0, b = h, r = w, wd8 = w / 8;
                        var a1, a2, q10, q11, q12, minAdj3, a3, dx2, x2, x3, x4, x5, x6, dy1, f1, q1, q2,
                            cx1, cx2, q1, dy3, q3, q4, q5, rh, q8, cx4, q9, cx5;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal4) ? cnstVal4 : adj1;
                        a2 = (adj2 < cnstVal1) ? cnstVal1 : (adj2 > cnstVal3) ? cnstVal3 : adj2;
                        q10 = cnstVal4 - a1;
                        q11 = q10 / 2;
                        q12 = a1 - q11;
                        minAdj3 = (0 > q12) ? 0 : q12;
                        a3 = (adj3 < minAdj3) ? minAdj3 : (adj3 > a1) ? a1 : adj3;
                        dx2 = w * a2 / cnstVal5;
                        x2 = hc - dx2;
                        x3 = x2 + wd8;
                        x4 = r - x3;
                        x5 = r - x2;
                        x6 = r - wd8;
                        dy1 = h * a3 / cnstVal4;
                        f1 = 4 * dy1 / w;
                        q1 = x3 * x3 / w;
                        q2 = x3 - q1;
                        cx1 = x3 / 2;
                        cx2 = r - cx1;
                        q1 = h * a1 / cnstVal4;
                        dy3 = q1 - dy1;
                        q3 = x2 * x2 / w;
                        q4 = x2 - q3;
                        q5 = f1 * q4;
                        rh = b - q1;
                        q8 = dy1 * 14 / 16;
                        cx4 = x2 / 2;
                        q9 = f1 * cx4;
                        cx5 = r - cx4;
                        if (shapType == "ellipseRibbon") {
                            var y1, cy1, y3, q6, q7, cy3, y2, y5, y6,
                                cy4, cy6, y7, cy7, y8;
                            y1 = f1 * q2;
                            cy1 = f1 * cx1;
                            y3 = q5 + dy3;
                            q6 = dy1 + dy3 - y3;
                            q7 = q6 + dy1;
                            cy3 = q7 + dy3;
                            y2 = (q8 + rh) / 2;
                            y5 = q5 + rh;
                            y6 = y3 + rh;
                            cy4 = q9 + rh;
                            cy6 = cy3 + rh;
                            y7 = y1 + dy3;
                            cy7 = q1 + q1 - y7;
                            y8 = b - dy1;
                            //
                            d_val = "M" + l + "," + t +
                                " Q" + cx1 + "," + cy1 + " " + x3 + "," + y1 +
                                " L" + x2 + "," + y3 +
                                " Q" + hc + "," + cy3 + " " + x5 + "," + y3 +
                                " L" + x4 + "," + y1 +
                                " Q" + cx2 + "," + cy1 + " " + r + "," + t +
                                " L" + x6 + "," + y2 +
                                " L" + r + "," + rh +
                                " Q" + cx5 + "," + cy4 + " " + x5 + "," + y5 +
                                " L" + x5 + "," + y6 +
                                " Q" + hc + "," + cy6 + " " + x2 + "," + y6 +
                                " L" + x2 + "," + y5 +
                                " Q" + cx4 + "," + cy4 + " " + l + "," + rh +
                                " L" + wd8 + "," + y2 +
                                " z" +
                                "M" + x2 + "," + y5 +
                                " L" + x2 + "," + y3 +
                                "M" + x5 + "," + y3 +
                                " L" + x5 + "," + y5 +
                                "M" + x3 + "," + y1 +
                                " L" + x3 + "," + y7 +
                                "M" + x4 + "," + y7 +
                                " L" + x4 + "," + y1;
                        } else if (shapType == "ellipseRibbon2") {
                            var u1, y1, cu1, cy1, q3, q5, u3, y3, q6, q7, cu3, cy3, rh, q8, u2, y2,
                                u5, y5, u6, y6, cu4, cy4, cu6, cy6, u7, y7, cu7, cy7;
                            u1 = f1 * q2;
                            y1 = b - u1;
                            cu1 = f1 * cx1;
                            cy1 = b - cu1;
                            u3 = q5 + dy3;
                            y3 = b - u3;
                            q6 = dy1 + dy3 - u3;
                            q7 = q6 + dy1;
                            cu3 = q7 + dy3;
                            cy3 = b - cu3;
                            u2 = (q8 + rh) / 2;
                            y2 = b - u2;
                            u5 = q5 + rh;
                            y5 = b - u5;
                            u6 = u3 + rh;
                            y6 = b - u6;
                            cu4 = q9 + rh;
                            cy4 = b - cu4;
                            cu6 = cu3 + rh;
                            cy6 = b - cu6;
                            u7 = u1 + dy3;
                            y7 = b - u7;
                            cu7 = q1 + q1 - u7;
                            cy7 = b - cu7;
                            //
                            d_val = "M" + l + "," + b +
                                " L" + wd8 + "," + y2 +
                                " L" + l + "," + q1 +
                                " Q" + cx4 + "," + cy4 + " " + x2 + "," + y5 +
                                " L" + x2 + "," + y6 +
                                " Q" + hc + "," + cy6 + " " + x5 + "," + y6 +
                                " L" + x5 + "," + y5 +
                                " Q" + cx5 + "," + cy4 + " " + r + "," + q1 +
                                " L" + x6 + "," + y2 +
                                " L" + r + "," + b +
                                " Q" + cx2 + "," + cy1 + " " + x4 + "," + y1 +
                                " L" + x5 + "," + y3 +
                                " Q" + hc + "," + cy3 + " " + x2 + "," + y3 +
                                " L" + x3 + "," + y1 +
                                " Q" + cx1 + "," + cy1 + " " + l + "," + b +
                                " z" +
                                "M" + x2 + "," + y3 +
                                " L" + x2 + "," + y5 +
                                "M" + x5 + "," + y5 +
                                " L" + x5 + "," + y3 +
                                "M" + x3 + "," + y7 +
                                " L" + x3 + "," + y1 +
                                "M" + x4 + "," + y1 +
                                " L" + x4 + "," + y7;
                        }
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "line":
                    case "straightConnector1":
                    case "bentConnector4":
                    case "bentConnector5":
                    case "curvedConnector2":
                    case "curvedConnector3":
                    case "curvedConnector4":
                    case "curvedConnector5":
                        // if (isFlipV) {
                        //     result += "<line x1='" + w + "' y1='0' x2='0' y2='" + h + "' stroke='" + border.color +
                        //         "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                        // } else {
                        result += "<line x1='0' y1='0' x2='" + w + "' y2='" + h + "' stroke='" + border.color +
                            "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                        //}
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                        }
                        result += "/>";
                        break;
                    case "rightArrow":
                        var points = window.PPTXArrowShapes.genRightArrow(w, h, node, slideFactor).replace("polygon points='", "").replace("'", "");
                        result += " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "leftArrow":
                        var points = window.PPTXArrowShapes.genLeftArrow(w, h, node, slideFactor).replace("polygon points='", "").replace("'", "");
                        result += " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "downArrow":
                    case "flowChartOffpageConnector":
                        var points = window.PPTXArrowShapes.genDownArrow(w, h, node, slideFactor).replace("polygon points='", "").replace("'", "");
                        result += " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "upArrow":
                        var points = window.PPTXArrowShapes.genUpArrow(w, h, node, slideFactor).replace("polygon points='", "").replace("'", "");
                        result += " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "leftRightArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, sAdj1_val = 0.25;
                        var sAdj2, sAdj2_val = 0.25;
                        var max_sAdj2_const = w / h;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = 0.5 - (parseInt(sAdj1.substr(4)) / 200000);
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                                    sAdj2_val = (sAdj2_val2) / max_sAdj2_const;
                                }
                            }
                        }
                        //console.log("w: "+w+"\nh: "+h+"\nsAdj1: "+sAdj1_val+"\nsAdj2: "+sAdj2_val);

                        result += " <polygon points='0 " + h / 2 + "," + sAdj2_val * w + " " + h + "," + sAdj2_val * w + " " + (1 - sAdj1_val) * h + "," + (1 - sAdj2_val) * w + " " + (1 - sAdj1_val) * h +
                            "," + (1 - sAdj2_val) * w + " " + h + "," + w + " " + h / 2 + ", " + (1 - sAdj2_val) * w + " 0," + (1 - sAdj2_val) * w + " " + sAdj1_val * h + "," +
                            sAdj2_val * w + " " + sAdj1_val * h + "," + sAdj2_val * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "upDownArrow":
                        var points = window.PPTXArrowShapes.genUpDownArrow(w, h, node, slideFactor).replace("polygon points='", "").replace("'", "");
                        result += " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "quadArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 22500 * slideFactor;
                        var sAdj2, adj2 = 22500 * slideFactor;
                        var sAdj3, adj3 = 22500 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var cnstVal3 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, q1, x1, x2, dx2, x3, dx3, x4, x5, x6, y2, y3, y4, y5, y6, maxAdj1, maxAdj3;
                        var minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        q1 = cnstVal2 - maxAdj1;
                        maxAdj3 = q1 / 2;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        x1 = minWH * a3 / cnstVal2;
                        dx2 = minWH * a2 / cnstVal2;
                        x2 = hc - dx2;
                        x5 = hc + dx2;
                        dx3 = minWH * a1 / cnstVal3;
                        x3 = hc - dx3;
                        x4 = hc + dx3;
                        x6 = w - x1;
                        y2 = vc - dx2;
                        y5 = vc + dx2;
                        y3 = vc - dx3;
                        y4 = vc + dx3;
                        y6 = h - x1;
                        var d_val = "M" + 0 + "," + vc +
                            " L" + x1 + "," + y2 +
                            " L" + x1 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + hc + "," + 0 +
                            " L" + x5 + "," + x1 +
                            " L" + x4 + "," + x1 +
                            " L" + x4 + "," + y3 +
                            " L" + x6 + "," + y3 +
                            " L" + x6 + "," + y2 +
                            " L" + w + "," + vc +
                            " L" + x6 + "," + y5 +
                            " L" + x6 + "," + y4 +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + y6 +
                            " L" + x5 + "," + y6 +
                            " L" + hc + "," + h +
                            " L" + x2 + "," + y6 +
                            " L" + x3 + "," + y6 +
                            " L" + x3 + "," + y4 +
                            " L" + x1 + "," + y4 +
                            " L" + x1 + "," + y5 + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftRightUpArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 25000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var cnstVal3 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, q1, x1, x2, dx2, x3, dx3, x4, x5, x6, y2, dy2, y3, y4, y5, maxAdj1, maxAdj3;
                        var minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        q1 = cnstVal2 - maxAdj1;
                        maxAdj3 = q1 / 2;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        x1 = minWH * a3 / cnstVal2;
                        dx2 = minWH * a2 / cnstVal2;
                        x2 = hc - dx2;
                        x5 = hc + dx2;
                        dx3 = minWH * a1 / cnstVal3;
                        x3 = hc - dx3;
                        x4 = hc + dx3;
                        x6 = w - x1;
                        dy2 = minWH * a2 / cnstVal1;
                        y2 = h - dy2;
                        y4 = h - dx2;
                        y3 = y4 - dx3;
                        y5 = y4 + dx3;
                        var d_val = "M" + 0 + "," + y4 +
                            " L" + x1 + "," + y2 +
                            " L" + x1 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + hc + "," + 0 +
                            " L" + x5 + "," + x1 +
                            " L" + x4 + "," + x1 +
                            " L" + x4 + "," + y3 +
                            " L" + x6 + "," + y3 +
                            " L" + x6 + "," + y2 +
                            " L" + w + "," + y4 +
                            " L" + x6 + "," + h +
                            " L" + x6 + "," + y5 +
                            " L" + x1 + "," + y5 +
                            " L" + x1 + "," + h + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftUpArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 25000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var cnstVal3 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, x1, x2, dx4, dx3, x3, x4, x5, y2, y3, y4, y5, maxAdj1, maxAdj3;
                        var minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        maxAdj3 = cnstVal2 - maxAdj1;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        x1 = minWH * a3 / cnstVal2;
                        dx2 = minWH * a2 / cnstVal1;
                        x2 = w - dx2;
                        y2 = h - dx2;
                        dx4 = minWH * a2 / cnstVal2;
                        x4 = w - dx4;
                        y4 = h - dx4;
                        dx3 = minWH * a1 / cnstVal3;
                        x3 = x4 - dx3;
                        x5 = x4 + dx3;
                        y3 = y4 - dx3;
                        y5 = y4 + dx3;
                        var d_val = "M" + 0 + "," + y4 +
                            " L" + x1 + "," + y2 +
                            " L" + x1 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + x4 + "," + 0 +
                            " L" + w + "," + x1 +
                            " L" + x5 + "," + x1 +
                            " L" + x5 + "," + y5 +
                            " L" + x1 + "," + y5 +
                            " L" + x1 + "," + h + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "bentUpArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 25000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var cnstVal3 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, dx1, x1, dx2, x2, dx3, x3, x4, y1, y2, dy2;
                        var minWH = Math.min(w, h);
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        y1 = minWH * a3 / cnstVal2;
                        dx1 = minWH * a2 / cnstVal1;
                        x1 = w - dx1;
                        dx3 = minWH * a2 / cnstVal2;
                        x3 = w - dx3;
                        dx2 = minWH * a1 / cnstVal3;
                        x2 = x3 - dx2;
                        x4 = x3 + dx2;
                        dy2 = minWH * a1 / cnstVal2;
                        y2 = h - dy2;
                        var d_val = "M" + 0 + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + y1 +
                            " L" + x1 + "," + y1 +
                            " L" + x3 + "," + 0 +
                            " L" + w + "," + y1 +
                            " L" + x4 + "," + y1 +
                            " L" + x4 + "," + h +
                            " L" + 0 + "," + h + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "bentArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 25000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var sAdj4, adj4 = 43750 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var a1, a2, a3, a4, x3, x4, y3, y4, y5, y6, maxAdj1, maxAdj4;
                        var minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > cnstVal1) a3 = cnstVal1
                        else a3 = adj3
                        var th, aw2, th2, dh2, ah, bw, bh, bs, bd, bd3, bd2,
                            th = minWH * a1 / cnstVal2;
                        aw2 = minWH * a2 / cnstVal2;
                        th2 = th / 2;
                        dh2 = aw2 - th2;
                        ah = minWH * a3 / cnstVal2;
                        bw = w - ah;
                        bh = h - dh2;
                        bs = (bw < bh) ? bw : bh;
                        maxAdj4 = cnstVal2 * bs / minWH;
                        if (adj4 < 0) a4 = 0
                        else if (adj4 > maxAdj4) a4 = maxAdj4
                        else a4 = adj4
                        bd = minWH * a4 / cnstVal2;
                        bd3 = bd - th;
                        bd2 = (bd3 > 0) ? bd3 : 0;
                        x3 = th + bd2;
                        x4 = w - ah;
                        y3 = dh2 + th;
                        y4 = y3 + dh2;
                        y5 = dh2 + bd;
                        y6 = y3 + bd2;

                        var d_val = "M" + 0 + "," + h +
                            " L" + 0 + "," + y5 +
                            window.PPTXShapeUtils.shapeArc(bd, y5, bd, bd, 180, 270, false).replace("M", "L") +
                            " L" + x4 + "," + dh2 +
                            " L" + x4 + "," + 0 +
                            " L" + w + "," + aw2 +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            window.PPTXShapeUtils.shapeArc(x3, y6, bd2, bd2, 270, 180, false).replace("M", "L") +
                            " L" + th + "," + h + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "uturnArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 25000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var sAdj4, adj4 = 43750 * slideFactor;
                        var sAdj5, adj5 = 75000 * slideFactor;
                        var cnstVal1 = 25000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var a1, a2, a3, a4, a5, q1, q2, q3, x3, x4, x5, x6, x7, x8, x9, y4, y5, minAdj5, maxAdj1, maxAdj3, maxAdj4;
                        var minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        q2 = a1 * minWH / h;
                        q3 = cnstVal2 - q2;
                        maxAdj3 = q3 * h / minWH;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        q1 = a3 + a1;
                        minAdj5 = q1 * minWH / h;
                        if (adj5 < minAdj5) a5 = minAdj5
                        else if (adj5 > cnstVal2) a5 = cnstVal2
                        else a5 = adj5

                        var th, aw2, th2, dh2, ah, bw, bs, bd, bd3, bd2,
                            th = minWH * a1 / cnstVal2;
                        aw2 = minWH * a2 / cnstVal2;
                        th2 = th / 2;
                        dh2 = aw2 - th2;
                        y5 = h * a5 / cnstVal2;
                        ah = minWH * a3 / cnstVal2;
                        y4 = y5 - ah;
                        x9 = w - dh2;
                        bw = x9 / 2;
                        bs = (bw < y4) ? bw : y4;
                        maxAdj4 = cnstVal2 * bs / minWH;
                        if (adj4 < 0) a4 = 0
                        else if (adj4 > maxAdj4) a4 = maxAdj4
                        else a4 = adj4
                        bd = minWH * a4 / cnstVal2;
                        bd3 = bd - th;
                        bd2 = (bd3 > 0) ? bd3 : 0;
                        x3 = th + bd2;
                        x8 = w - aw2;
                        x6 = x8 - aw2;
                        x7 = x6 + dh2;
                        x4 = x9 - bd;
                        x5 = x7 - bd2;
                        cx = (th + x7) / 2
                        var cy = (y4 + th) / 2
                        var d_val = "M" + 0 + "," + h +
                            " L" + 0 + "," + bd +
                            window.PPTXShapeUtils.shapeArc(bd, bd, bd, bd, 180, 270, false).replace("M", "L") +
                            " L" + x4 + "," + 0 +
                            window.PPTXShapeUtils.shapeArc(x4, bd, bd, bd, 270, 360, false).replace("M", "L") +
                            " L" + x9 + "," + y4 +
                            " L" + w + "," + y4 +
                            " L" + x8 + "," + y5 +
                            " L" + x6 + "," + y4 +
                            " L" + x7 + "," + y4 +
                            " L" + x7 + "," + x3 +
                            window.PPTXShapeUtils.shapeArc(x5, x3, bd2, bd2, 0, -90, false).replace("M", "L") +
                            " L" + x3 + "," + th +
                            window.PPTXShapeUtils.shapeArc(x3, x3, bd2, bd2, 270, 180, false).replace("M", "L") +
                            " L" + th + "," + h + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "stripedRightArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 50000 * slideFactor;
                        var sAdj2, adj2 = 50000 * slideFactor;
                        var cnstVal1 = 100000 * slideFactor;
                        var cnstVal2 = 200000 * slideFactor;
                        var cnstVal3 = 84375 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var a1, a2, x4, x5, dx5, x6, dx6, y1, dy1, y2, maxAdj2, vc = h / 2;
                        var minWH = Math.min(w, h);
                        maxAdj2 = cnstVal3 * w / minWH;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > maxAdj2) a2 = maxAdj2
                        else a2 = adj2
                        x4 = minWH * 5 / 32;
                        dx5 = minWH * a2 / cnstVal1;
                        x5 = w - dx5;
                        dy1 = h * a1 / cnstVal2;
                        y1 = vc - dy1;
                        y2 = vc + dy1;
                        //dx6 = dy1*dx5/hd2;
                        //x6 = w-dx6;
                        var ssd8 = minWH / 8,
                            ssd16 = minWH / 16,
                            ssd32 = minWH / 32;
                        var d_val = "M" + 0 + "," + y1 +
                            " L" + ssd32 + "," + y1 +
                            " L" + ssd32 + "," + y2 +
                            " L" + 0 + "," + y2 + " z" +
                            " M" + ssd16 + "," + y1 +
                            " L" + ssd8 + "," + y1 +
                            " L" + ssd8 + "," + y2 +
                            " L" + ssd16 + "," + y2 + " z" +
                            " M" + x4 + "," + y1 +
                            " L" + x5 + "," + y1 +
                            " L" + x5 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x5 + "," + h +
                            " L" + x5 + "," + y2 +
                            " L" + x4 + "," + y2 + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "notchedRightArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 50000 * slideFactor;
                        var sAdj2, adj2 = 50000 * slideFactor;
                        var cnstVal1 = 100000 * slideFactor;
                        var cnstVal2 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var a1, a2, x1, x2, dx2, y1, dy1, y2, maxAdj2, vc = h / 2, hd2 = vc;
                        var minWH = Math.min(w, h);
                        maxAdj2 = cnstVal1 * w / minWH;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > maxAdj2) a2 = maxAdj2
                        else a2 = adj2
                        dx2 = minWH * a2 / cnstVal1;
                        x2 = w - dx2;
                        dy1 = h * a1 / cnstVal2;
                        y1 = vc - dy1;
                        y2 = vc + dy1;
                        x1 = dy1 * dx2 / hd2;
                        var d_val = "M" + 0 + "," + y1 +
                            " L" + x2 + "," + y1 +
                            " L" + x2 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x2 + "," + h +
                            " L" + x2 + "," + y2 +
                            " L" + 0 + "," + y2 +
                            " L" + x1 + "," + vc + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "homePlate":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 50000 * slideFactor;
                        var cnstVal1 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var a, x1, dx1, maxAdj, vc = h / 2;
                        var minWH = Math.min(w, h);
                        maxAdj = cnstVal1 * w / minWH;
                        if (adj < 0) a = 0
                        else if (adj > maxAdj) a = maxAdj
                        else a = adj
                        dx1 = minWH * a / cnstVal1;
                        x1 = w - dx1;
                        var d_val = "M" + 0 + "," + 0 +
                            " L" + x1 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x1 + "," + h +
                            " L" + 0 + "," + h + " z";

                        result += "<path  d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "chevron":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 50000 * slideFactor;
                        var cnstVal1 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var a, x1, dx1, x2, maxAdj, vc = h / 2;
                        var minWH = Math.min(w, h);
                        maxAdj = cnstVal1 * w / minWH;
                        if (adj < 0) a = 0
                        else if (adj > maxAdj) a = maxAdj
                        else a = adj
                        x1 = minWH * a / cnstVal1;
                        x2 = w - x1;
                        var d_val = "M" + 0 + "," + 0 +
                            " L" + x2 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x2 + "," + h +
                            " L" + 0 + "," + h +
                            " L" + x1 + "," + vc + " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";


                        break;
                    case "rightArrowCallout":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 25000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var sAdj4, adj4 = 64977 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var cnstVal3 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, dx3, x3, x2, x1;
                        var vc = h / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);
                        maxAdj2 = cnstVal1 * h / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * w / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / w;
                        maxAdj4 = cnstVal - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        dy2 = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        dx3 = ss * a3 / cnstVal2;
                        x3 = r - dx3;
                        x2 = w * a4 / cnstVal2;
                        x1 = x2 / 2;
                        var d_val = "M" + l + "," + t +
                            " L" + x2 + "," + t +
                            " L" + x2 + "," + y2 +
                            " L" + x3 + "," + y2 +
                            " L" + x3 + "," + y1 +
                            " L" + r + "," + vc +
                            " L" + x3 + "," + y4 +
                            " L" + x3 + "," + y3 +
                            " L" + x2 + "," + y3 +
                            " L" + x2 + "," + b +
                            " L" + l + "," + b +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "downArrowCallout":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 25000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var sAdj4, adj4 = 64977 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var cnstVal3 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx1, dx2, x1, x2, x3, x4, dy3, y3, y2, y1;
                        var hc = w / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);

                        maxAdj2 = cnstVal1 * w / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * h / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / h;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dx1 = ss * a2 / cnstVal2;
                        dx2 = ss * a1 / cnstVal3;
                        x1 = hc - dx1;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        x4 = hc + dx1;
                        dy3 = ss * a3 / cnstVal2;
                        y3 = b - dy3;
                        y2 = h * a4 / cnstVal2;
                        y1 = y2 / 2;
                        var d_val = "M" + l + "," + t +
                            " L" + r + "," + t +
                            " L" + r + "," + y2 +
                            " L" + x3 + "," + y2 +
                            " L" + x3 + "," + y3 +
                            " L" + x4 + "," + y3 +
                            " L" + hc + "," + b +
                            " L" + x1 + "," + y3 +
                            " L" + x2 + "," + y3 +
                            " L" + x2 + "," + y2 +
                            " L" + l + "," + y2 +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftArrowCallout":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 25000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var sAdj4, adj4 = 64977 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var cnstVal3 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, x1, dx2, x2, x3;
                        var vc = h / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);

                        maxAdj2 = cnstVal1 * h / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * w / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / w;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        dy2 = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        x1 = ss * a3 / cnstVal2;
                        dx2 = w * a4 / cnstVal2;
                        x2 = r - dx2;
                        x3 = (x2 + r) / 2;
                        var d_val = "M" + l + "," + vc +
                            " L" + x1 + "," + y1 +
                            " L" + x1 + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + t +
                            " L" + r + "," + t +
                            " L" + r + "," + b +
                            " L" + x2 + "," + b +
                            " L" + x2 + "," + y3 +
                            " L" + x1 + "," + y3 +
                            " L" + x1 + "," + y4 +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "upArrowCallout":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 25000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var sAdj4, adj4 = 64977 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var cnstVal3 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx1, dx2, x1, x2, x3, x4, y1, dy2, y2, y3;
                        var hc = w / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);
                        maxAdj2 = cnstVal1 * w / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * h / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / h;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dx1 = ss * a2 / cnstVal2;
                        dx2 = ss * a1 / cnstVal3;
                        x1 = hc - dx1;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        x4 = hc + dx1;
                        y1 = ss * a3 / cnstVal2;
                        dy2 = h * a4 / cnstVal2;
                        y2 = b - dy2;
                        y3 = (y2 + b) / 2;

                        var d_val = "M" + l + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + y1 +
                            " L" + x1 + "," + y1 +
                            " L" + hc + "," + t +
                            " L" + x4 + "," + y1 +
                            " L" + x3 + "," + y1 +
                            " L" + x3 + "," + y2 +
                            " L" + r + "," + y2 +
                            " L" + r + "," + b +
                            " L" + l + "," + b +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftRightArrowCallout":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 25000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var sAdj4, adj4 = 48123 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var cnstVal3 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, x1, x4, dx2, x2, x3;
                        var vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);
                        maxAdj2 = cnstVal1 * h / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal1 * w / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / wd2;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        dy2 = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        x1 = ss * a3 / cnstVal2;
                        x4 = r - x1;
                        dx2 = w * a4 / cnstVal3;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        var d_val = "M" + l + "," + vc +
                            " L" + x1 + "," + y1 +
                            " L" + x1 + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + t +
                            " L" + x3 + "," + t +
                            " L" + x3 + "," + y2 +
                            " L" + x4 + "," + y2 +
                            " L" + x4 + "," + y1 +
                            " L" + r + "," + vc +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + b +
                            " L" + x2 + "," + b +
                            " L" + x2 + "," + y3 +
                            " L" + x1 + "," + y3 +
                            " L" + x1 + "," + y4 +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "quadArrowCallout":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 18515 * slideFactor;
                        var sAdj2, adj2 = 18515 * slideFactor;
                        var sAdj3, adj3 = 18515 * slideFactor;
                        var sAdj4, adj4 = 48123 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var cnstVal3 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0;
                        var ss = Math.min(w, h);
                        var a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx2, dx3, ah, dx1, dy1, x8, x2, x7, x3, x6, x4, x5, y8, y2, y7, y3, y6, y4, y5;
                        a2 = (adj2 < 0) ? 0 : (adj2 > cnstVal1) ? cnstVal1 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal1 - a2;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * 2;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < a1) ? a1 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dx2 = ss * a2 / cnstVal2;
                        dx3 = ss * a1 / cnstVal3;
                        ah = ss * a3 / cnstVal2;
                        dx1 = w * a4 / cnstVal3;
                        dy1 = h * a4 / cnstVal3;
                        x8 = r - ah;
                        x2 = hc - dx1;
                        x7 = hc + dx1;
                        x3 = hc - dx2;
                        x6 = hc + dx2;
                        x4 = hc - dx3;
                        x5 = hc + dx3;
                        y8 = b - ah;
                        y2 = vc - dy1;
                        y7 = vc + dy1;
                        y3 = vc - dx2;
                        y6 = vc + dx2;
                        y4 = vc - dx3;
                        y5 = vc + dx3;
                        var d_val = "M" + l + "," + vc +
                            " L" + ah + "," + y3 +
                            " L" + ah + "," + y4 +
                            " L" + x2 + "," + y4 +
                            " L" + x2 + "," + y2 +
                            " L" + x4 + "," + y2 +
                            " L" + x4 + "," + ah +
                            " L" + x3 + "," + ah +
                            " L" + hc + "," + t +
                            " L" + x6 + "," + ah +
                            " L" + x5 + "," + ah +
                            " L" + x5 + "," + y2 +
                            " L" + x7 + "," + y2 +
                            " L" + x7 + "," + y4 +
                            " L" + x8 + "," + y4 +
                            " L" + x8 + "," + y3 +
                            " L" + r + "," + vc +
                            " L" + x8 + "," + y6 +
                            " L" + x8 + "," + y5 +
                            " L" + x7 + "," + y5 +
                            " L" + x7 + "," + y7 +
                            " L" + x5 + "," + y7 +
                            " L" + x5 + "," + y8 +
                            " L" + x6 + "," + y8 +
                            " L" + hc + "," + b +
                            " L" + x3 + "," + y8 +
                            " L" + x4 + "," + y8 +
                            " L" + x4 + "," + y7 +
                            " L" + x2 + "," + y7 +
                            " L" + x2 + "," + y5 +
                            " L" + ah + "," + y5 +
                            " L" + ah + "," + y6 +
                            " z";

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "curvedDownArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 50000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, wd2 = w / 2, r = w, b = h, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        var ss = Math.min(w, h);
                        var maxAdj2, a2, a1, th, aw, q1, wR, q7, q8, q9, q10, q11, idy, maxAdj3, a3, ah, x3, q2, q3, q4, q5, dx, x5, x7, q6, dh, x4, x8, aw2, x6, y1, swAng, mswAng, iy, ix, q12, dang2, stAng, stAng2, swAng2, swAng3;

                        maxAdj2 = cnstVal1 * w / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        wR = wd2 - q1;
                        q7 = wR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        idy = q11 * h / q7;
                        maxAdj3 = cnstVal2 * idy / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * adj3 / cnstVal2;
                        x3 = wR + th;
                        q2 = h * h;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dx = q5 * wR / h;
                        x5 = wR + dx;
                        x7 = x3 + dx;
                        q6 = aw - th;
                        dh = q6 / 2;
                        x4 = x5 - dh;
                        x8 = x7 + dh;
                        aw2 = aw / 2;
                        x6 = r - aw2;
                        y1 = b - ah;
                        swAng = Math.atan(dx / ah);
                        var swAngDeg = swAng * 180 / Math.PI;
                        mswAng = -swAngDeg;
                        iy = b - idy;
                        ix = (wR + x3) / 2;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / idy);
                        var dang2Deg = dang2 * 180 / Math.PI;
                        stAng = c3d4 + swAngDeg;
                        stAng2 = c3d4 - dang2Deg;
                        swAng2 = dang2Deg - cd4;
                        swAng3 = cd4 + dang2Deg;
                        //var cX = x5 - Math.cos(stAng*Math.PI/180) * wR;
                        //var cY = y1 - Math.sin(stAng*Math.PI/180) * h;

                        var d_val = "M" + x6 + "," + b +
                            " L" + x4 + "," + y1 +
                            " L" + x5 + "," + y1 +
                            window.PPTXShapeUtils.shapeArc(wR, h, wR, h, stAng, (stAng + mswAng), false).replace("M", "L") +
                            " L" + x3 + "," + t +
                            window.PPTXShapeUtils.shapeArc(x3, h, wR, h, c3d4, (c3d4 + swAngDeg), false).replace("M", "L") +
                            " L" + (x5 + th) + "," + y1 +
                            " L" + x8 + "," + y1 +
                            " z" +
                            "M" + x3 + "," + t +
                            window.PPTXShapeUtils.shapeArc(x3, h, wR, h, stAng2, (stAng2 + swAng2), false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(wR, h, wR, h, cd2, (cd2 + swAng3), false).replace("M", "L");

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "curvedLeftArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 50000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, hd2 = h / 2, r = w, b = h, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        var ss = Math.min(w, h);
                        var maxAdj2, a2, a1, th, aw, q1, hR, q7, q8, q9, q10, q11, iDx, maxAdj3, a3, ah, y3, q2, q3, q4, q5, dy, y5, y7, q6, dh, y4, y8, aw2, y6, x1, swAng, mswAng, ix, iy, q12, dang2, swAng2, swAng3, stAng3;

                        maxAdj2 = cnstVal1 * h / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > a2) ? a2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        hR = hd2 - q1;
                        q7 = hR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        iDx = q11 * w / q7;
                        maxAdj3 = cnstVal2 * iDx / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * a3 / cnstVal2;
                        y3 = hR + th;
                        q2 = w * w;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dy = q5 * hR / w;
                        y5 = hR + dy;
                        y7 = y3 + dy;
                        q6 = aw - th;
                        dh = q6 / 2;
                        y4 = y5 - dh;
                        y8 = y7 + dh;
                        aw2 = aw / 2;
                        y6 = b - aw2;
                        x1 = l + ah;
                        swAng = Math.atan(dy / ah);
                        mswAng = -swAng;
                        ix = l + iDx;
                        iy = (hR + y3) / 2;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / iDx);
                        swAng2 = dang2 - swAng;
                        swAng3 = swAng + dang2;
                        stAng3 = -dang2;
                        var swAngDg, swAng2Dg, swAng3Dg, stAng3dg;
                        swAngDg = swAng * 180 / Math.PI;
                        swAng2Dg = swAng2 * 180 / Math.PI;
                        swAng3Dg = swAng3 * 180 / Math.PI;
                        stAng3dg = stAng3 * 180 / Math.PI;

                        var d_val = "M" + r + "," + y3 +
                            window.PPTXShapeUtils.shapeArc(l, hR, w, hR, 0, -cd4, false).replace("M", "L") +
                            " L" + l + "," + t +
                            window.PPTXShapeUtils.shapeArc(l, y3, w, hR, c3d4, (c3d4 + cd4), false).replace("M", "L") +
                            " L" + r + "," + y3 +
                            window.PPTXShapeUtils.shapeArc(l, y3, w, hR, 0, swAngDg, false).replace("M", "L") +
                            " L" + x1 + "," + y7 +
                            " L" + x1 + "," + y8 +
                            " L" + l + "," + y6 +
                            " L" + x1 + "," + y4 +
                            " L" + x1 + "," + y5 +
                            window.PPTXShapeUtils.shapeArc(l, hR, w, hR, swAngDg, (swAngDg + swAng2Dg), false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(l, hR, w, hR, 0, -cd4, false).replace("M", "L") +
                            window.PPTXShapeUtils.shapeArc(l, y3, w, hR, c3d4, (c3d4 + cd4), false).replace("M", "L");

                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "curvedRightArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 50000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, hd2 = h / 2, r = w, b = h, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        var ss = Math.min(w, h);
                        var maxAdj2, a2, a1, th, aw, q1, hR, q7, q8, q9, q10, q11, iDx, maxAdj3, a3, ah, y3, q2, q3, q4, q5, dy,
                            y5, y7, q6, dh, y4, y8, aw2, y6, x1, swAng, stAng, mswAng, ix, iy, q12, dang2, swAng2, swAng3, stAng3;

                        maxAdj2 = cnstVal1 * h / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > a2) ? a2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        hR = hd2 - q1;
                        q7 = hR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        iDx = q11 * w / q7;
                        maxAdj3 = cnstVal2 * iDx / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * a3 / cnstVal2;
                        y3 = hR + th;
                        q2 = w * w;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dy = q5 * hR / w;
                        y5 = hR + dy;
                        y7 = y3 + dy;
                        q6 = aw - th;
                        dh = q6 / 2;
                        y4 = y5 - dh;
                        y8 = y7 + dh;
                        aw2 = aw / 2;
                        y6 = b - aw2;
                        x1 = r - ah;
                        swAng = Math.atan(dy / ah);
                        stAng = Math.PI + 0 - swAng;
                        mswAng = -swAng;
                        ix = r - iDx;
                        iy = (hR + y3) / 2;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / iDx);
                        swAng2 = dang2 - Math.PI / 2;
                        swAng3 = Math.PI / 2 + dang2;
                        stAng3 = Math.PI - dang2;

                        var stAngDg, mswAngDg, swAngDg, swAng2dg;
                        stAngDg = stAng * 180 / Math.PI;
                        mswAngDg = mswAng * 180 / Math.PI;
                        swAngDg = swAng * 180 / Math.PI;
                        swAng2dg = swAng2 * 180 / Math.PI;

                        var d_val = "M" + l + "," + hR +
                            window.PPTXShapeUtils.shapeArc(w, hR, w, hR, cd2, cd2 + mswAngDg, false).replace("M", "L") +
                            " L" + x1 + "," + y5 +
                            " L" + x1 + "," + y4 +
                            " L" + r + "," + y6 +
                            " L" + x1 + "," + y8 +
                            " L" + x1 + "," + y7 +
                            window.PPTXShapeUtils.shapeArc(w, y3, w, hR, stAngDg, stAngDg + swAngDg, false).replace("M", "L") +
                            " L" + l + "," + hR +
                            window.PPTXShapeUtils.shapeArc(w, hR, w, hR, cd2, cd2 + cd4, false).replace("M", "L") +
                            " L" + r + "," + th +
                            window.PPTXShapeUtils.shapeArc(w, y3, w, hR, c3d4, c3d4 + swAng2dg, false).replace("M", "L")
                        "";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "curvedUpArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 25000 * slideFactor;
                        var sAdj2, adj2 = 50000 * slideFactor;
                        var sAdj3, adj3 = 25000 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, wd2 = w / 2, r = w, b = h, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        var ss = Math.min(w, h);
                        var maxAdj2, a2, a1, th, aw, q1, wR, q7, q8, q9, q10, q11, idy, maxAdj3, a3, ah, x3, q2, q3, q4, q5, dx, x5, x7, q6, dh, x4, x8, aw2, x6, y1, swAng, mswAng, iy, ix, q12, dang2, swAng2, mswAng2, stAng3, swAng3, stAng2;

                        maxAdj2 = cnstVal1 * w / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        wR = wd2 - q1;
                        q7 = wR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        idy = q11 * h / q7;
                        maxAdj3 = cnstVal2 * idy / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * adj3 / cnstVal2;
                        x3 = wR + th;
                        q2 = h * h;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dx = q5 * wR / h;
                        x5 = wR + dx;
                        x7 = x3 + dx;
                        q6 = aw - th;
                        dh = q6 / 2;
                        x4 = x5 - dh;
                        x8 = x7 + dh;
                        aw2 = aw / 2;
                        x6 = r - aw2;
                        y1 = t + ah;
                        swAng = Math.atan(dx / ah);
                        mswAng = -swAng;
                        iy = t + idy;
                        ix = (wR + x3) / 2;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / idy);
                        swAng2 = dang2 - swAng;
                        mswAng2 = -swAng2;
                        stAng3 = Math.PI / 2 - swAng;
                        swAng3 = swAng + dang2;
                        stAng2 = Math.PI / 2 - dang2;

                        var stAng2dg, swAng2dg, swAngDg, swAng2dg;
                        stAng2dg = stAng2 * 180 / Math.PI;
                        swAng2dg = swAng2 * 180 / Math.PI;
                        stAng3dg = stAng3 * 180 / Math.PI;
                        swAngDg = swAng * 180 / Math.PI;

                        var d_val = //"M" + ix + "," +iy + 
                            window.PPTXShapeUtils.shapeArc(wR, 0, wR, h, stAng2dg, stAng2dg + swAng2dg, false) + //.replace("M","L") +
                            " L" + x5 + "," + y1 +
                            " L" + x4 + "," + y1 +
                            " L" + x6 + "," + t +
                            " L" + x8 + "," + y1 +
                            " L" + x7 + "," + y1 +
                            window.PPTXShapeUtils.shapeArc(x3, 0, wR, h, stAng3dg, stAng3dg + swAngDg, false).replace("M", "L") +
                            " L" + wR + "," + b +
                            window.PPTXShapeUtils.shapeArc(wR, 0, wR, h, cd4, cd2, false).replace("M", "L") +
                            " L" + th + "," + t +
                            window.PPTXShapeUtils.shapeArc(x3, 0, wR, h, cd2, cd4, false).replace("M", "L") +
                            "";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "mathDivide":
                    case "mathEqual":
                    case "mathMinus":
                    case "mathMultiply":
                    case "mathNotEqual":
                    case "mathPlus":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1;
                        var sAdj2, adj2;
                        var sAdj3, adj3;
                        if (shapAdjst_ary !== undefined) {
                            if (shapAdjst_ary.constructor === Array) {
                                for (var i = 0; i < shapAdjst_ary.length; i++) {
                                    var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                    if (sAdj_name == "adj1") {
                                        sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        adj1 = parseInt(sAdj1.substr(4));
                                    } else if (sAdj_name == "adj2") {
                                        sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        adj2 = parseInt(sAdj2.substr(4));
                                    } else if (sAdj_name == "adj3") {
                                        sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        adj3 = parseInt(sAdj3.substr(4));
                                    }
                                }
                            } else {
                                sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4));
                            }
                        }
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var cnstVal3 = 200000 * slideFactor;
                        var dVal;
                        var hc = w / 2, vc = h / 2, hd2 = h / 2;
                        if (shapType == "mathNotEqual") {
                            if (shapAdjst_ary === undefined) {
                                adj1 = 23520 * slideFactor;
                                adj2 = 110 * Math.PI / 180;
                                adj3 = 11760 * slideFactor;
                            } else {
                                adj1 = adj1 * slideFactor;
                                adj2 = (adj2 / 60000) * Math.PI / 180;
                                adj3 = adj3 * slideFactor;
                            }
                            var a1, crAng, a2a1, maxAdj3, a3, dy1, dy2, dx1, x1, x8, y2, y3, y1, y4,
                                cadj2, xadj2, len, bhw, bhw2, x7, dx67, x6, dx57, x5, dx47, x4, dx37,
                                x3, dx27, x2, rx7, rx6, rx5, rx4, rx3, rx2, dx7, rxt, lxt, rx, lx,
                                dy3, dy4, ry, ly, dlx, drx, dly, dry, xC1, xC2, yC1, yC2, yC3, yC4;
                            var angVal1 = 70 * Math.PI / 180, angVal2 = 110 * Math.PI / 180;
                            var cnstVal4 = 73490 * slideFactor;
                            //var cd4 = 90;
                            a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
                            crAng = (adj2 < angVal1) ? angVal1 : (adj2 > angVal2) ? angVal2 : adj2;
                            a2a1 = a1 * 2;
                            maxAdj3 = cnstVal2 - a2a1;
                            a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                            dy1 = h * a1 / cnstVal2;
                            dy2 = h * a3 / cnstVal3;
                            dx1 = w * cnstVal4 / cnstVal3;
                            x1 = hc - dx1;
                            x8 = hc + dx1;
                            y2 = vc - dy2;
                            y3 = vc + dy2;
                            y1 = y2 - dy1;
                            y4 = y3 + dy1;
                            cadj2 = crAng - Math.PI / 2;
                            xadj2 = hd2 * Math.tan(cadj2);
                            len = Math.sqrt(xadj2 * xadj2 + hd2 * hd2);
                            bhw = len * dy1 / hd2;
                            bhw2 = bhw / 2;
                            x7 = hc + xadj2 - bhw2;
                            dx67 = xadj2 * y1 / hd2;
                            x6 = x7 - dx67;
                            dx57 = xadj2 * y2 / hd2;
                            x5 = x7 - dx57;
                            dx47 = xadj2 * y3 / hd2;
                            x4 = x7 - dx47;
                            dx37 = xadj2 * y4 / hd2;
                            x3 = x7 - dx37;
                            dx27 = xadj2 * 2;
                            x2 = x7 - dx27;
                            rx7 = x7 + bhw;
                            rx6 = x6 + bhw;
                            rx5 = x5 + bhw;
                            rx4 = x4 + bhw;
                            rx3 = x3 + bhw;
                            rx2 = x2 + bhw;
                            dx7 = dy1 * hd2 / len;
                            rxt = x7 + dx7;
                            lxt = rx7 - dx7;
                            rx = (cadj2 > 0) ? rxt : rx7;
                            lx = (cadj2 > 0) ? x7 : lxt;
                            dy3 = dy1 * xadj2 / len;
                            dy4 = -dy3;
                            ry = (cadj2 > 0) ? dy3 : 0;
                            ly = (cadj2 > 0) ? 0 : dy4;
                            dlx = w - rx;
                            drx = w - lx;
                            dly = h - ry;
                            dry = h - ly;
                            xC1 = (rx + lx) / 2;
                            xC2 = (drx + dlx) / 2;
                            yC1 = (ry + ly) / 2;
                            yC2 = (y1 + y2) / 2;
                            yC3 = (y3 + y4) / 2;
                            yC4 = (dry + dly) / 2;

                            dVal = "M" + x1 + "," + y1 +
                                " L" + x6 + "," + y1 +
                                " L" + lx + "," + ly +
                                " L" + rx + "," + ry +
                                " L" + rx6 + "," + y1 +
                                " L" + x8 + "," + y1 +
                                " L" + x8 + "," + y2 +
                                " L" + rx5 + "," + y2 +
                                " L" + rx4 + "," + y3 +
                                " L" + x8 + "," + y3 +
                                " L" + x8 + "," + y4 +
                                " L" + rx3 + "," + y4 +
                                " L" + drx + "," + dry +
                                " L" + dlx + "," + dly +
                                " L" + x3 + "," + y4 +
                                " L" + x1 + "," + y4 +
                                " L" + x1 + "," + y3 +
                                " L" + x4 + "," + y3 +
                                " L" + x5 + "," + y2 +
                                " L" + x1 + "," + y2 +
                                " z";
                        } else if (shapType == "mathDivide") {
                            if (shapAdjst_ary === undefined) {
                                adj1 = 23520 * slideFactor;
                                adj2 = 5880 * slideFactor;
                                adj3 = 11760 * slideFactor;
                            } else {
                                adj1 = adj1 * slideFactor;
                                adj2 = adj2 * slideFactor;
                                adj3 = adj3 * slideFactor;
                            }
                            var a1, ma1, ma3h, ma3w, maxAdj3, a3, m4a3, maxAdj2, a2, dy1, yg, rad, dx1,
                                y3, y4, a, y2, y1, y5, x1, x3, x2;
                            var cnstVal4 = 1000 * slideFactor;
                            var cnstVal5 = 36745 * slideFactor;
                            var cnstVal6 = 73490 * slideFactor;
                            a1 = (adj1 < cnstVal4) ? cnstVal4 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
                            ma1 = -a1;
                            ma3h = (cnstVal6 + ma1) / 4;
                            ma3w = cnstVal5 * w / h;
                            maxAdj3 = (ma3h < ma3w) ? ma3h : ma3w;
                            a3 = (adj3 < cnstVal4) ? cnstVal4 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                            m4a3 = -4 * a3;
                            maxAdj2 = cnstVal6 + m4a3 - a1;
                            a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                            dy1 = h * a1 / cnstVal3;
                            yg = h * a2 / cnstVal2;
                            rad = h * a3 / cnstVal2;
                            dx1 = w * cnstVal6 / cnstVal3;
                            y3 = vc - dy1;
                            y4 = vc + dy1;
                            a = yg + rad;
                            y2 = y3 - a;
                            y1 = y2 - rad;
                            y5 = h - y1;
                            x1 = hc - dx1;
                            x3 = hc + dx1;
                            x2 = hc - rad;
                            var cd4 = 90, c3d4 = 270;
                            var cX1 = hc - Math.cos(c3d4 * Math.PI / 180) * rad;
                            var cY1 = y1 - Math.sin(c3d4 * Math.PI / 180) * rad;
                            var cX2 = hc - Math.cos(Math.PI / 2) * rad;
                            var cY2 = y5 - Math.sin(Math.PI / 2) * rad;
                            dVal = "M" + hc + "," + y1 +
                                window.PPTXShapeUtils.shapeArc(cX1, cY1, rad, rad, c3d4, c3d4 + 360, false).replace("M", "L") +
                                " z" +
                                " M" + hc + "," + y5 +
                                window.PPTXShapeUtils.shapeArc(cX2, cY2, rad, rad, cd4, cd4 + 360, false).replace("M", "L") +
                                " z" +
                                " M" + x1 + "," + y3 +
                                " L" + x3 + "," + y3 +
                                " L" + x3 + "," + y4 +
                                " L" + x1 + "," + y4 +
                                " z";
                        } else if (shapType == "mathEqual") {
                            dVal = window.PPTXMathShapes.genMathEqual(w, h, node, slideFactor);
                        } else if (shapType == "mathMinus") {
                            dVal = window.PPTXMathShapes.genMathMinus(w, h, node, slideFactor);
                        } else if (shapType == "mathMultiply") {
                            dVal = window.PPTXMathShapes.genMathMultiply(w, h, node, slideFactor);
                        } else if (shapType == "mathPlus") {
                            dVal = window.PPTXMathShapes.genMathPlus(w, h, node, slideFactor);
                        }
                        result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        //console.log(shapType);
                        break;
                    case "can":
                    case "flowChartMagneticDisk":
                    case "flowChartMagneticDrum":
                        var shapAdjst = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        var adj = 25000 * slideFactor;
                        var cnstVal1 = 50000 * slideFactor;
                        var cnstVal2 = 200000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        var ss = Math.min(w, h);
                        var maxAdj, a, y1, y2, y3, dVal;
                        if (shapType == "flowChartMagneticDisk" || shapType == "flowChartMagneticDrum") {
                            adj = 50000 * slideFactor;
                        }
                        maxAdj = cnstVal1 * h / ss;
                        a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
                        y1 = ss * a / cnstVal2;
                        y2 = y1 + y1;
                        y3 = h - y1;
                        var cd2 = 180, wd2 = w / 2;

                        var tranglRott = "";
                        if (shapType == "flowChartMagneticDrum") {
                            tranglRott = "transform='rotate(90 " + w / 2 + "," + h / 2 + ")'";
                        }
                        dVal = window.PPTXShapeUtils.shapeArc(wd2, y1, wd2, y1, 0, cd2, false) +
                            window.PPTXShapeUtils.shapeArc(wd2, y1, wd2, y1, cd2, cd2 + cd2, false).replace("M", "L") +
                            " L" + w + "," + y3 +
                            window.PPTXShapeUtils.shapeArc(wd2, y3, wd2, y1, 0, cd2, false).replace("M", "L") +
                            " L" + 0 + "," + y1;

                        result += "<path " + tranglRott + " d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "swooshArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var refr = slideFactor;
                        var sAdj1, adj1 = 25000 * refr;
                        var sAdj2, adj2 = 16667 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                }
                            }
                        }
                        var cnstVal1 = 1 * refr;
                        var cnstVal2 = 70000 * refr;
                        var cnstVal3 = 75000 * refr;
                        var cnstVal4 = 100000 * refr;
                        var ss = Math.min(w, h);
                        var ssd8 = ss / 8;
                        var hd6 = h / 6;

                        var a1, maxAdj2, a2, ad1, ad2, xB, yB, alfa, dx0, xC, dx1, yF, xF, xE, yE, dy2, dy22, dy3, yD, dy4, yP1, xP1, dy5, yP2, xP2;

                        a1 = (adj1 < cnstVal1) ? cnstVal1 : (adj1 > cnstVal3) ? cnstVal3 : adj1;
                        maxAdj2 = cnstVal2 * w / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        ad1 = h * a1 / cnstVal4;
                        ad2 = ss * a2 / cnstVal4;
                        xB = w - ad2;
                        yB = ssd8;
                        alfa = (Math.PI / 2) / 14;
                        dx0 = ssd8 * Math.tan(alfa);
                        xC = xB - dx0;
                        dx1 = ad1 * Math.tan(alfa);
                        yF = yB + ad1;
                        xF = xB + dx1;
                        xE = xF + dx0;
                        yE = yF + ssd8;
                        dy2 = yE - 0;
                        dy22 = dy2 / 2;
                        dy3 = h / 20;
                        yD = dy22 - dy3;
                        dy4 = hd6;
                        yP1 = hd6 + dy4;
                        xP1 = w / 6;
                        dy5 = hd6 / 2;
                        yP2 = yF + dy5;
                        xP2 = w / 4;

                        var dVal = "M" + 0 + "," + h +
                            " Q" + xP1 + "," + yP1 + " " + xB + "," + yB +
                            " L" + xC + "," + 0 +
                            " L" + w + "," + yD +
                            " L" + xE + "," + yE +
                            " L" + xF + "," + yF +
                            " Q" + xP2 + "," + yP2 + " " + 0 + "," + h +
                            " z";

                        result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "circularArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 12500 * slideFactor;
                        var sAdj2, adj2 = (1142319 / 60000) * Math.PI / 180;
                        var sAdj3, adj3 = (20457681 / 60000) * Math.PI / 180;
                        var sAdj4, adj4 = (10800000 / 60000) * Math.PI / 180;
                        var sAdj5, adj5 = 12500 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = (parseInt(sAdj3.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = (parseInt(sAdj4.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0, wd2 = w / 2, hd2 = h / 2;
                        var ss = Math.min(w, h);
                        var a5, maxAdj1, a1, enAng, stAng, th, thh, th2, rw1, rh1, rw2, rh2, rw3, rh3, wtH, htH, dxH,
                            dyH, xH, yH, rI, u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, u11, u12, u13, u14, u15, u16, u17,
                            u18, u19, u20, u21, maxAng, aAng, ptAng, wtA, htA, dxA, dyA, xA, yA, wtE, htE, dxE, dyE, xE, yE,
                            dxG, dyG, xG, yG, dxB, dyB, xB, yB, sx1, sy1, sx2, sy2, rO, x1O, y1O, x2O, y2O, dxO, dyO, dO,
                            q1, q2, DO, q3, q4, q5, q6, q7, q8, sdelO, ndyO, sdyO, q9, q10, q11, dxF1, q12, dxF2, adyO,
                            q13, q14, dyF1, q15, dyF2, q16, q17, q18, q19, q20, q21, q22, dxF, dyF, sdxF, sdyF, xF, yF,
                            x1I, y1I, x2I, y2I, dxI, dyI, dI, v1, v2, DI, v3, v4, v5, v6, v7, v8, sdelI, v9, v10, v11,
                            dxC1, v12, dxC2, adyI, v13, v14, dyC1, v15, dyC2, v16, v17, v18, v19, v20, v21, v22, dxC, dyC,
                            sdxC, sdyC, xC, yC, ist0, ist1, istAng, isw1, isw2, iswAng, p1, p2, p3, p4, p5, xGp, yGp,
                            xBp, yBp, en0, en1, en2, sw0, sw1, swAng;
                        var cnstVal1 = 25000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var rdAngVal1 = (1 / 60000) * Math.PI / 180;
                        var rdAngVal2 = (21599999 / 60000) * Math.PI / 180;
                        var rdAngVal3 = 2 * Math.PI;

                        a5 = (adj5 < 0) ? 0 : (adj5 > cnstVal1) ? cnstVal1 : adj5;
                        maxAdj1 = a5 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        enAng = (adj3 < rdAngVal1) ? rdAngVal1 : (adj3 > rdAngVal2) ? rdAngVal2 : adj3;
                        stAng = (adj4 < 0) ? 0 : (adj4 > rdAngVal2) ? rdAngVal2 : adj4; //////////////////////////////////////////
                        th = ss * a1 / cnstVal2;
                        thh = ss * a5 / cnstVal2;
                        th2 = th / 2;
                        rw1 = wd2 + th2 - thh;
                        rh1 = hd2 + th2 - thh;
                        rw2 = rw1 - th;
                        rh2 = rh1 - th;
                        rw3 = rw2 + th2;
                        rh3 = rh2 + th2;
                        wtH = rw3 * Math.sin(enAng);
                        htH = rh3 * Math.cos(enAng);

                        //dxH = rw3*Math.cos(Math.atan(wtH/htH));
                        //dyH = rh3*Math.sin(Math.atan(wtH/htH));
                        dxH = rw3 * Math.cos(Math.atan2(wtH, htH));
                        dyH = rh3 * Math.sin(Math.atan2(wtH, htH));

                        xH = hc + dxH;
                        yH = vc + dyH;
                        rI = (rw2 < rh2) ? rw2 : rh2;
                        u1 = dxH * dxH;
                        u2 = dyH * dyH;
                        u3 = rI * rI;
                        u4 = u1 - u3;
                        u5 = u2 - u3;
                        u6 = u4 * u5 / u1;
                        u7 = u6 / u2;
                        u8 = 1 - u7;
                        u9 = Math.sqrt(u8);
                        u10 = u4 / dxH;
                        u11 = u10 / dyH;
                        u12 = (1 + u9) / u11;

                        //u13 = Math.atan(u12/1);
                        u13 = Math.atan2(u12, 1);

                        u14 = u13 + rdAngVal3;
                        u15 = (u13 > 0) ? u13 : u14;
                        u16 = u15 - enAng;
                        u17 = u16 + rdAngVal3;
                        u18 = (u16 > 0) ? u16 : u17;
                        u19 = u18 - cd2;
                        u20 = u18 - rdAngVal3;
                        u21 = (u19 > 0) ? u20 : u18;
                        maxAng = Math.abs(u21);
                        aAng = (adj2 < 0) ? 0 : (adj2 > maxAng) ? maxAng : adj2;
                        ptAng = enAng + aAng;
                        wtA = rw3 * Math.sin(ptAng);
                        htA = rh3 * Math.cos(ptAng);
                        //dxA = rw3*Math.cos(Math.atan(wtA/htA));
                        //dyA = rh3*Math.sin(Math.atan(wtA/htA));
                        dxA = rw3 * Math.cos(Math.atan2(wtA, htA));
                        dyA = rh3 * Math.sin(Math.atan2(wtA, htA));

                        xA = hc + dxA;
                        yA = vc + dyA;
                        wtE = rw1 * Math.sin(stAng);
                        htE = rh1 * Math.cos(stAng);

                        //dxE = rw1*Math.cos(Math.atan(wtE/htE));
                        //dyE = rh1*Math.sin(Math.atan(wtE/htE));
                        dxE = rw1 * Math.cos(Math.atan2(wtE, htE));
                        dyE = rh1 * Math.sin(Math.atan2(wtE, htE));

                        xE = hc + dxE;
                        yE = vc + dyE;
                        dxG = thh * Math.cos(ptAng);
                        dyG = thh * Math.sin(ptAng);
                        xG = xH + dxG;
                        yG = yH + dyG;
                        dxB = thh * Math.cos(ptAng);
                        dyB = thh * Math.sin(ptAng);
                        xB = xH - dxB;
                        yB = yH - dyB;
                        sx1 = xB - hc;
                        sy1 = yB - vc;
                        sx2 = xG - hc;
                        sy2 = yG - vc;
                        rO = (rw1 < rh1) ? rw1 : rh1;
                        x1O = sx1 * rO / rw1;
                        y1O = sy1 * rO / rh1;
                        x2O = sx2 * rO / rw1;
                        y2O = sy2 * rO / rh1;
                        dxO = x2O - x1O;
                        dyO = y2O - y1O;
                        dO = Math.sqrt(dxO * dxO + dyO * dyO);
                        q1 = x1O * y2O;
                        q2 = x2O * y1O;
                        DO = q1 - q2;
                        q3 = rO * rO;
                        q4 = dO * dO;
                        q5 = q3 * q4;
                        q6 = DO * DO;
                        q7 = q5 - q6;
                        q8 = (q7 > 0) ? q7 : 0;
                        sdelO = Math.sqrt(q8);
                        ndyO = dyO * -1;
                        sdyO = (ndyO > 0) ? -1 : 1;
                        q9 = sdyO * dxO;
                        q10 = q9 * sdelO;
                        q11 = DO * dyO;
                        dxF1 = (q11 + q10) / q4;
                        q12 = q11 - q10;
                        dxF2 = q12 / q4;
                        adyO = Math.abs(dyO);
                        q13 = adyO * sdelO;
                        q14 = DO * dxO / -1;
                        dyF1 = (q14 + q13) / q4;
                        q15 = q14 - q13;
                        dyF2 = q15 / q4;
                        q16 = x2O - dxF1;
                        q17 = x2O - dxF2;
                        q18 = y2O - dyF1;
                        q19 = y2O - dyF2;
                        q20 = Math.sqrt(q16 * q16 + q18 * q18);
                        q21 = Math.sqrt(q17 * q17 + q19 * q19);
                        q22 = q21 - q20;
                        dxF = (q22 > 0) ? dxF1 : dxF2;
                        dyF = (q22 > 0) ? dyF1 : dyF2;
                        sdxF = dxF * rw1 / rO;
                        sdyF = dyF * rh1 / rO;
                        xF = hc + sdxF;
                        yF = vc + sdyF;
                        x1I = sx1 * rI / rw2;
                        y1I = sy1 * rI / rh2;
                        x2I = sx2 * rI / rw2;
                        y2I = sy2 * rI / rh2;
                        dxI = x2I - x1I;
                        dyI = y2I - y1I;
                        dI = Math.sqrt(dxI * dxI + dyI * dyI);
                        v1 = x1I * y2I;
                        v2 = x2I * y1I;
                        DI = v1 - v2;
                        v3 = rI * rI;
                        v4 = dI * dI;
                        v5 = v3 * v4;
                        v6 = DI * DI;
                        v7 = v5 - v6;
                        v8 = (v7 > 0) ? v7 : 0;
                        sdelI = Math.sqrt(v8);
                        v9 = sdyO * dxI;
                        v10 = v9 * sdelI;
                        v11 = DI * dyI;
                        dxC1 = (v11 + v10) / v4;
                        v12 = v11 - v10;
                        dxC2 = v12 / v4;
                        adyI = Math.abs(dyI);
                        v13 = adyI * sdelI;
                        v14 = DI * dxI / -1;
                        dyC1 = (v14 + v13) / v4;
                        v15 = v14 - v13;
                        dyC2 = v15 / v4;
                        v16 = x1I - dxC1;
                        v17 = x1I - dxC2;
                        v18 = y1I - dyC1;
                        v19 = y1I - dyC2;
                        v20 = Math.sqrt(v16 * v16 + v18 * v18);
                        v21 = Math.sqrt(v17 * v17 + v19 * v19);
                        v22 = v21 - v20;
                        dxC = (v22 > 0) ? dxC1 : dxC2;
                        dyC = (v22 > 0) ? dyC1 : dyC2;
                        sdxC = dxC * rw2 / rI;
                        sdyC = dyC * rh2 / rI;
                        xC = hc + sdxC;
                        yC = vc + sdyC;

                        //ist0 = Math.atan(sdyC/sdxC);
                        ist0 = Math.atan2(sdyC, sdxC);

                        ist1 = ist0 + rdAngVal3;
                        istAng = (ist0 > 0) ? ist0 : ist1;
                        isw1 = stAng - istAng;
                        isw2 = isw1 - rdAngVal3;
                        iswAng = (isw1 > 0) ? isw2 : isw1;
                        p1 = xF - xC;
                        p2 = yF - yC;
                        p3 = Math.sqrt(p1 * p1 + p2 * p2);
                        p4 = p3 / 2;
                        p5 = p4 - thh;
                        xGp = (p5 > 0) ? xF : xG;
                        yGp = (p5 > 0) ? yF : yG;
                        xBp = (p5 > 0) ? xC : xB;
                        yBp = (p5 > 0) ? yC : yB;

                        //en0 = Math.atan(sdyF/sdxF);
                        en0 = Math.atan2(sdyF, sdxF);

                        en1 = en0 + rdAngVal3;
                        en2 = (en0 > 0) ? en0 : en1;
                        sw0 = en2 - stAng;
                        sw1 = sw0 + rdAngVal3;
                        swAng = (sw0 > 0) ? sw0 : sw1;

                        var strtAng = stAng * 180 / Math.PI
                        var endAng = strtAng + (swAng * 180 / Math.PI);
                        var stiAng = istAng * 180 / Math.PI;
                        var swiAng = iswAng * 180 / Math.PI;
                        var ediAng = stiAng + swiAng;

                        var d_val = window.PPTXShapeUtils.shapeArc(w / 2, h / 2, rw1, rh1, strtAng, endAng, false) +
                            " L" + xGp + "," + yGp +
                            " L" + xA + "," + yA +
                            " L" + xBp + "," + yBp +
                            " L" + xC + "," + yC +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, rw2, rh2, stiAng, ediAng, false).replace("M", "L") +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftCircularArrow":
                        var shapAdjst_ary = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        var sAdj1, adj1 = 12500 * slideFactor;
                        var sAdj2, adj2 = (-1142319 / 60000) * Math.PI / 180;
                        var sAdj3, adj3 = (1142319 / 60000) * Math.PI / 180;
                        var sAdj4, adj4 = (10800000 / 60000) * Math.PI / 180;
                        var sAdj5, adj5 = 12500 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (var i = 0; i < shapAdjst_ary.length; i++) {
                                var sAdj_name = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = (parseInt(sAdj3.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = (parseInt(sAdj4.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = window.PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * slideFactor;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0, wd2 = w / 2, hd2 = h / 2;
                        var ss = Math.min(w, h);
                        var cnstVal1 = 25000 * slideFactor;
                        var cnstVal2 = 100000 * slideFactor;
                        var rdAngVal1 = (1 / 60000) * Math.PI / 180;
                        var rdAngVal2 = (21599999 / 60000) * Math.PI / 180;
                        var rdAngVal3 = 2 * Math.PI;
                        var a5, maxAdj1, a1, enAng, stAng, th, thh, th2, rw1, rh1, rw2, rh2, rw3, rh3, wtH, htH, dxH, dyH, xH, yH, rI,
                            u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, u11, u12, u13, u14, u15, u16, u17, u18, u19, u20, u21, u22,
                            minAng, u23, a2, aAng, ptAng, wtA, htA, dxA, dyA, xA, yA, wtE, htE, dxE, dyE, xE, yE, wtD, htD, dxD, dyD,
                            xD, yD, dxG, dyG, xG, yG, dxB, dyB, xB, yB, sx1, sy1, sx2, sy2, rO, x1O, y1O, x2O, y2O, dxO, dyO, dO,
                            q1, q2, DO, q3, q4, q5, q6, q7, q8, sdelO, ndyO, sdyO, q9, q10, q11, dxF1, q12, dxF2, adyO, q13, q14, dyF1,
                            q15, dyF2, q16, q17, q18, q19, q20, q21, q22, dxF, dyF, sdxF, sdyF, xF, yF, x1I, y1I, x2I, y2I, dxI, dyI, dI,
                            v1, v2, DI, v3, v4, v5, v6, v7, v8, sdelI, v9, v10, v11, dxC1, v12, dxC2, adyI, v13, v14, dyC1, v15, dyC2, v16,
                            v17, v18, v19, v20, v21, v22, dxC, dyC, sdxC, sdyC, xC, yC, ist0, ist1, istAng0, isw1, isw2, iswAng0, istAng,
                            iswAng, p1, p2, p3, p4, p5, xGp, yGp, xBp, yBp, en0, en1, en2, sw0, sw1, swAng, stAng0;

                        a5 = (adj5 < 0) ? 0 : (adj5 > cnstVal1) ? cnstVal1 : adj5;
                        maxAdj1 = a5 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        enAng = (adj3 < rdAngVal1) ? rdAngVal1 : (adj3 > rdAngVal2) ? rdAngVal2 : adj3;
                        stAng = (adj4 < 0) ? 0 : (adj4 > rdAngVal2) ? rdAngVal2 : adj4;
                        th = ss * a1 / cnstVal2;
                        thh = ss * a5 / cnstVal2;
                        th2 = th / 2;
                        rw1 = wd2 + th2 - thh;
                        rh1 = hd2 + th2 - thh;
                        rw2 = rw1 - th;
                        rh2 = rh1 - th;
                        rw3 = rw2 + th2;
                        rh3 = rh2 + th2;
                        wtH = rw3 * Math.sin(enAng);
                        htH = rh3 * Math.cos(enAng);
                        dxH = rw3 * Math.cos(Math.atan2(wtH, htH));
                        dyH = rh3 * Math.sin(Math.atan2(wtH, htH));
                        xH = hc + dxH;
                        yH = vc + dyH;
                        rI = (rw2 < rh2) ? rw2 : rh2;
                        u1 = dxH * dxH;
                        u2 = dyH * dyH;
                        u3 = rI * rI;
                        u4 = u1 - u3;
                        u5 = u2 - u3;
                        u6 = u4 * u5 / u1;
                        u7 = u6 / u2;
                        u8 = 1 - u7;
                        u9 = Math.sqrt(u8);
                        u10 = u4 / dxH;
                        u11 = u10 / dyH;
                        u12 = (1 + u9) / u11;
                        u13 = Math.atan2(u12, 1);
                        u14 = u13 + rdAngVal3;
                        u15 = (u13 > 0) ? u13 : u14;
                        u16 = u15 - enAng;
                        u17 = u16 + rdAngVal3;
                        u18 = (u16 > 0) ? u16 : u17;
                        u19 = u18 - cd2;
                        u20 = u18 - rdAngVal3;
                        u21 = (u19 > 0) ? u20 : u18;
                        u22 = Math.abs(u21);
                        minAng = u22 * -1;
                        u23 = Math.abs(adj2);
                        a2 = u23 * -1;
                        aAng = (a2 < minAng) ? minAng : (a2 > 0) ? 0 : a2;
                        ptAng = enAng + aAng;
                        wtA = rw3 * Math.sin(ptAng);
                        htA = rh3 * Math.cos(ptAng);
                        dxA = rw3 * Math.cos(Math.atan2(wtA, htA));
                        dyA = rh3 * Math.sin(Math.atan2(wtA, htA));
                        xA = hc + dxA;
                        yA = vc + dyA;
                        wtE = rw1 * Math.sin(stAng);
                        htE = rh1 * Math.cos(stAng);
                        dxE = rw1 * Math.cos(Math.atan2(wtE, htE));
                        dyE = rh1 * Math.sin(Math.atan2(wtE, htE));
                        xE = hc + dxE;
                        yE = vc + dyE;
                        wtD = rw2 * Math.sin(stAng);
                        htD = rh2 * Math.cos(stAng);
                        dxD = rw2 * Math.cos(Math.atan2(wtD, htD));
                        dyD = rh2 * Math.sin(Math.atan2(wtD, htD));
                        xD = hc + dxD;
                        yD = vc + dyD;
                        dxG = thh * Math.cos(ptAng);
                        dyG = thh * Math.sin(ptAng);
                        xG = xH + dxG;
                        yG = yH + dyG;
                        dxB = thh * Math.cos(ptAng);
                        dyB = thh * Math.sin(ptAng);
                        xB = xH - dxB;
                        yB = yH - dyB;
                        sx1 = xB - hc;
                        sy1 = yB - vc;
                        sx2 = xG - hc;
                        sy2 = yG - vc;
                        rO = (rw1 < rh1) ? rw1 : rh1;
                        x1O = sx1 * rO / rw1;
                        y1O = sy1 * rO / rh1;
                        x2O = sx2 * rO / rw1;
                        y2O = sy2 * rO / rh1;
                        dxO = x2O - x1O;
                        dyO = y2O - y1O;
                        dO = Math.sqrt(dxO * dxO + dyO * dyO);
                        q1 = x1O * y2O;
                        q2 = x2O * y1O;
                        DO = q1 - q2;
                        q3 = rO * rO;
                        q4 = dO * dO;
                        q5 = q3 * q4;
                        q6 = DO * DO;
                        q7 = q5 - q6;
                        q8 = (q7 > 0) ? q7 : 0;
                        sdelO = Math.sqrt(q8);
                        ndyO = dyO * -1;
                        sdyO = (ndyO > 0) ? -1 : 1;
                        q9 = sdyO * dxO;
                        q10 = q9 * sdelO;
                        q11 = DO * dyO;
                        dxF1 = (q11 + q10) / q4;
                        q12 = q11 - q10;
                        dxF2 = q12 / q4;
                        adyO = Math.abs(dyO);
                        q13 = adyO * sdelO;
                        q14 = DO * dxO / -1;
                        dyF1 = (q14 + q13) / q4;
                        q15 = q14 - q13;
                        dyF2 = q15 / q4;
                        q16 = x2O - dxF1;
                        q17 = x2O - dxF2;
                        q18 = y2O - dyF1;
                        q19 = y2O - dyF2;
                        q20 = Math.sqrt(q16 * q16 + q18 * q18);
                        q21 = Math.sqrt(q17 * q17 + q19 * q19);
                        q22 = q21 - q20;
                        dxF = (q22 > 0) ? dxF1 : dxF2;
                        dyF = (q22 > 0) ? dyF1 : dyF2;
                        sdxF = dxF * rw1 / rO;
                        sdyF = dyF * rh1 / rO;
                        xF = hc + sdxF;
                        yF = vc + sdyF;
                        x1I = sx1 * rI / rw2;
                        y1I = sy1 * rI / rh2;
                        x2I = sx2 * rI / rw2;
                        y2I = sy2 * rI / rh2;
                        dxI = x2I - x1I;
                        dyI = y2I - y1I;
                        dI = Math.sqrt(dxI * dxI + dyI * dyI);
                        v1 = x1I * y2I;
                        v2 = x2I * y1I;
                        DI = v1 - v2;
                        v3 = rI * rI;
                        v4 = dI * dI;
                        v5 = v3 * v4;
                        v6 = DI * DI;
                        v7 = v5 - v6;
                        v8 = (v7 > 0) ? v7 : 0;
                        sdelI = Math.sqrt(v8);
                        v9 = sdyO * dxI;
                        v10 = v9 * sdelI;
                        v11 = DI * dyI;
                        dxC1 = (v11 + v10) / v4;
                        v12 = v11 - v10;
                        dxC2 = v12 / v4;
                        adyI = Math.abs(dyI);
                        v13 = adyI * sdelI;
                        v14 = DI * dxI / -1;
                        dyC1 = (v14 + v13) / v4;
                        v15 = v14 - v13;
                        dyC2 = v15 / v4;
                        v16 = x1I - dxC1;
                        v17 = x1I - dxC2;
                        v18 = y1I - dyC1;
                        v19 = y1I - dyC2;
                        v20 = Math.sqrt(v16 * v16 + v18 * v18);
                        v21 = Math.sqrt(v17 * v17 + v19 * v19);
                        v22 = v21 - v20;
                        dxC = (v22 > 0) ? dxC1 : dxC2;
                        dyC = (v22 > 0) ? dyC1 : dyC2;
                        sdxC = dxC * rw2 / rI;
                        sdyC = dyC * rh2 / rI;
                        xC = hc + sdxC;
                        yC = vc + sdyC;
                        ist0 = Math.atan2(sdyC, sdxC);
                        ist1 = ist0 + rdAngVal3;
                        istAng0 = (ist0 > 0) ? ist0 : ist1;
                        isw1 = stAng - istAng0;
                        isw2 = isw1 + rdAngVal3;
                        iswAng0 = (isw1 > 0) ? isw1 : isw2;
                        istAng = istAng0 + iswAng0;
                        iswAng = -iswAng0;
                        p1 = xF - xC;
                        p2 = yF - yC;
                        p3 = Math.sqrt(p1 * p1 + p2 * p2);
                        p4 = p3 / 2;
                        p5 = p4 - thh;
                        xGp = (p5 > 0) ? xF : xG;
                        yGp = (p5 > 0) ? yF : yG;
                        xBp = (p5 > 0) ? xC : xB;
                        yBp = (p5 > 0) ? yC : yB;
                        en0 = Math.atan2(sdyF, sdxF);
                        en1 = en0 + rdAngVal3;
                        en2 = (en0 > 0) ? en0 : en1;
                        sw0 = en2 - stAng;
                        sw1 = sw0 - rdAngVal3;
                        swAng = (sw0 > 0) ? sw1 : sw0;
                        stAng0 = stAng + swAng;

                        var strtAng = stAng0 * 180 / Math.PI;
                        var endAng = stAng * 180 / Math.PI;
                        var stiAng = istAng * 180 / Math.PI;
                        var swiAng = iswAng * 180 / Math.PI;
                        var ediAng = stiAng + swiAng;

                        var d_val = "M" + xE + "," + yE +
                            " L" + xD + "," + yD +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, rw2, rh2, stiAng, ediAng, false).replace("M", "L") +
                            " L" + xBp + "," + yBp +
                            " L" + xA + "," + yA +
                            " L" + xGp + "," + yGp +
                            " L" + xF + "," + yF +
                            window.PPTXShapeUtils.shapeArc(w / 2, h / 2, rw1, rh1, strtAng, endAng, false).replace("M", "L") +
                            " z";
                        result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftRightCircularArrow":
                    case "chartPlus":
                    case "chartStar":
                    case "chartX":
                    case "cornerTabs":
                    case "flowChartOfflineStorage":
                    case "folderCorner":
                    case "funnel":
                    case "lineInv":
                    case "nonIsoscelesTrapezoid":
                    case "plaqueTabs":
                    case "squareTabs":
                    case "upDownArrowCallout":
                        console.log(shapType, " -unsupported shape type.");
                        break;
                    case undefined:
                    default:
                        console.warn("Undefine shape type.(" + shapType + ")");
                }

                result += "</svg>";

                result += "<div class='block " + window.PPTXTextStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) + //block content
                    " " + window.PPTXTextStyleUtils.getContentDir(node, type, warpObj) +
                    "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                    "' style='" +
                    getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                    getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
                    " z-index: " + order + ";" +
                    "transform: rotate(" + ((txtRotate !== undefined) ? txtRotate : 0) + "deg);" +
                    "'>";

                // TextBody
                if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    if (type != "diagram" && type != "textBox") {
                        type = "shape";
                    }
                    result += window.PPTXTextElementUtils.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, undefined, styleTable); //type='shape'
                }
                result += "</div>";
            } else if (custShapType !== undefined) {
                //custGeom here - Amir ///////////////////////////////////////////////////////
                //http://officeopenxml.com/drwSp-custGeom.php
                var pathLstNode = window.PPTXUtils.getTextByPathList(custShapType, ["a:pathLst"]);
                var pathNodes = window.PPTXUtils.getTextByPathList(pathLstNode, ["a:path"]);
                //var pathNode = window.PPTXUtils.getTextByPathList(pathLstNode, ["a:path", "attrs"]);
                var maxX = parseInt(pathNodes["attrs"]["w"]);// * slideFactor;
                var maxY = parseInt(pathNodes["attrs"]["h"]);// * slideFactor;
                var cX = (1 / maxX) * w;
                var cY = (1 / maxY) * h;
                //console.log("w = "+w+"\nh = "+h+"\nmaxX = "+maxX +"\nmaxY = " + maxY);
                //cheke if it is close shape

                //console.log("custShapType : ", custShapType, ", pathLstNode: ", pathLstNode, ", node: ", node);//, ", y:", y, ", w:", w, ", h:", h);

                var moveToNode = window.PPTXUtils.getTextByPathList(pathNodes, ["a:moveTo"]);
                var total_shapes = moveToNode.length;

                var lnToNodes = pathNodes["a:lnTo"]; //total a:pt : 1
                var cubicBezToNodes = pathNodes["a:cubicBezTo"]; //total a:pt : 3
                var arcToNodes = pathNodes["a:arcTo"]; //total a:pt : 0?1? ; attrs: ~4 ()
                var closeNode = window.PPTXUtils.getTextByPathList(pathNodes, ["a:close"]); //total a:pt : 0
                //quadBezTo //total a:pt : 2 - TODO
                //console.log("ia moveToNode array: ", Array.isArray(moveToNode))
                if (!Array.isArray(moveToNode)) {
                    moveToNode = [moveToNode];
                }
                //console.log("ia moveToNode array: ", Array.isArray(moveToNode))

                var multiSapeAry = [];
                if (moveToNode.length > 0) {
                    //a:moveTo
                    Object.keys(moveToNode).forEach(function (key) {
                        var moveToPtNode = moveToNode[key]["a:pt"];
                        if (moveToPtNode !== undefined) {
                            Object.keys(moveToPtNode).forEach(function (key2) {
                                var ptObj = {};
                                var moveToNoPt = moveToPtNode[key2];
                                var spX = moveToNoPt["attrs", "x"];//parseInt(moveToNoPt["attrs", "x"]) * slideFactor;
                                var spY = moveToNoPt["attrs", "y"];//parseInt(moveToNoPt["attrs", "y"]) * slideFactor;
                                var ptOrdr = moveToNoPt["attrs", "order"];
                                ptObj.type = "movto";
                                ptObj.order = ptOrdr;
                                ptObj.x = spX;
                                ptObj.y = spY;
                                multiSapeAry.push(ptObj);
                                //console.log(key2, lnToNoPt);

                            });
                        }
                    });
                    //a:lnTo
                    if (lnToNodes !== undefined) {
                        Object.keys(lnToNodes).forEach(function (key) {
                            var lnToPtNode = lnToNodes[key]["a:pt"];
                            if (lnToPtNode !== undefined) {
                                Object.keys(lnToPtNode).forEach(function (key2) {
                                    var ptObj = {};
                                    var lnToNoPt = lnToPtNode[key2];
                                    var ptX = lnToNoPt["attrs", "x"];
                                    var ptY = lnToNoPt["attrs", "y"];
                                    var ptOrdr = lnToNoPt["attrs", "order"];
                                    ptObj.type = "lnto";
                                    ptObj.order = ptOrdr;
                                    ptObj.x = ptX;
                                    ptObj.y = ptY;
                                    multiSapeAry.push(ptObj);
                                    //console.log(key2, lnToNoPt);
                                });
                            }
                        });
                    }
                    //a:cubicBezTo
                    if (cubicBezToNodes !== undefined) {

                        var cubicBezToPtNodesAry = [];
                        //console.log("cubicBezToNodes: ", cubicBezToNodes, ", is arry: ", Array.isArray(cubicBezToNodes))
                        if (!Array.isArray(cubicBezToNodes)) {
                            cubicBezToNodes = [cubicBezToNodes];
                        }
                        Object.keys(cubicBezToNodes).forEach(function (key) {
                            //console.log("cubicBezTo[" + key + "]:");
                            cubicBezToPtNodesAry.push(cubicBezToNodes[key]["a:pt"]);
                        });

                        //console.log("cubicBezToNodes: ", cubicBezToPtNodesAry)
                        cubicBezToPtNodesAry.forEach(function (key2) {
                            //console.log("cubicBezToPtNodesAry: key2 : ", key2)
                            var nodeObj = {};
                            nodeObj.type = "cubicBezTo";
                            nodeObj.order = key2[0]["attrs"]["order"];
                            var pts_ary = [];
                            key2.forEach(function (pt) {
                                var pt_obj = {
                                    x: pt["attrs"]["x"],
                                    y: pt["attrs"]["y"]
                                }
                                pts_ary.push(pt_obj)
                            })
                            nodeObj.cubBzPt = pts_ary;//key2;
                            multiSapeAry.push(nodeObj);
                        });
                    }
                    //a:arcTo
                    if (arcToNodes !== undefined) {
                        var arcToNodesAttrs = arcToNodes["attrs"];
                        var arcOrder = arcToNodesAttrs["order"];
                        var hR = arcToNodesAttrs["hR"];
                        var wR = arcToNodesAttrs["wR"];
                        var stAng = arcToNodesAttrs["stAng"];
                        var swAng = arcToNodesAttrs["swAng"];
                        var shftX = 0;
                        var shftY = 0;
                        var arcToPtNode = window.PPTXUtils.getTextByPathList(arcToNodes, ["a:pt", "attrs"]);
                        if (arcToPtNode !== undefined) {
                            shftX = arcToPtNode["x"];
                            shftY = arcToPtNode["y"];
                            //console.log("shftX: ",shftX," shftY: ",shftY)
                        }
                        var ptObj = {};
                        ptObj.type = "arcTo";
                        ptObj.order = arcOrder;
                        ptObj.hR = hR;
                        ptObj.wR = wR;
                        ptObj.stAng = stAng;
                        ptObj.swAng = swAng;
                        ptObj.shftX = shftX;
                        ptObj.shftY = shftY;
                        multiSapeAry.push(ptObj);

                    }
                    //a:quadBezTo - TODO

                    //a:close
                    if (closeNode !== undefined) {

                        if (!Array.isArray(closeNode)) {
                            closeNode = [closeNode];
                        }
                        // Object.keys(closeNode).forEach(function (key) {
                        //     //console.log("cubicBezTo[" + key + "]:");
                        //     cubicBezToPtNodesAry.push(closeNode[key]["a:pt"]);
                        // });
                        Object.keys(closeNode).forEach(function (key) {
                            //console.log("custShapType >> closeNode: key: ", key);
                            var clsAttrs = closeNode[key]["attrs"];
                            //var clsAttrs = closeNode["attrs"];
                            var clsOrder = clsAttrs["order"];
                            var ptObj = {};
                            ptObj.type = "close";
                            ptObj.order = clsOrder;
                            multiSapeAry.push(ptObj);

                        });

                    }

                    // console.log("custShapType >> multiSapeAry: ", multiSapeAry);

                    multiSapeAry.sort(function (a, b) {
                        return a.order - b.order;
                    });

                    //console.log("custShapType >>sorted  multiSapeAry: ");
                    //console.log(multiSapeAry);
                    var k = 0;
                    var isClose = false;
                    var d = "";
                    while (k < multiSapeAry.length) {

                        if (multiSapeAry[k].type == "movto") {
                            //start point
                            var spX = parseInt(multiSapeAry[k].x) * cX;//slideFactor;
                            var spY = parseInt(multiSapeAry[k].y) * cY;//slideFactor;
                            // if (d == "") {
                            //     d = "M" + spX + "," + spY;
                            // } else {
                            //     //shape without close : then close the shape and start new path
                            //     result += "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            //         "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' ";
                            //     result += "/>";

                            //     if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            //         result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                            //     }
                            //     if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            //         result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                            //     }
                            //     result += "/>";

                            //     d = "M" + spX + "," + spY;
                            //     isClose = true;
                            // }

                            d += " M" + spX + "," + spY;

                        } else if (multiSapeAry[k].type == "lnto") {
                            var Lx = parseInt(multiSapeAry[k].x) * cX;//slideFactor;
                            var Ly = parseInt(multiSapeAry[k].y) * cY;//slideFactor;
                            d += " L" + Lx + "," + Ly;

                        } else if (multiSapeAry[k].type == "cubicBezTo") {
                            var Cx1 = parseInt(multiSapeAry[k].cubBzPt[0].x) * cX;//slideFactor;
                            var Cy1 = parseInt(multiSapeAry[k].cubBzPt[0].y) * cY;//slideFactor;
                            var Cx2 = parseInt(multiSapeAry[k].cubBzPt[1].x) * cX;//slideFactor;
                            var Cy2 = parseInt(multiSapeAry[k].cubBzPt[1].y) * cY;//slideFactor;
                            var Cx3 = parseInt(multiSapeAry[k].cubBzPt[2].x) * cX;//slideFactor;
                            var Cy3 = parseInt(multiSapeAry[k].cubBzPt[2].y) * cY;//slideFactor;
                            d += " C" + Cx1 + "," + Cy1 + " " + Cx2 + "," + Cy2 + " " + Cx3 + "," + Cy3;
                        } else if (multiSapeAry[k].type == "arcTo") {
                            var hR = parseInt(multiSapeAry[k].hR) * cX;//slideFactor;
                            var wR = parseInt(multiSapeAry[k].wR) * cY;//slideFactor;
                            var stAng = parseInt(multiSapeAry[k].stAng) / 60000;
                            var swAng = parseInt(multiSapeAry[k].swAng) / 60000;
                            //var shftX = parseInt(multiSapeAry[k].shftX) * slideFactor;
                            //var shftY = parseInt(multiSapeAry[k].shftY) * slideFactor;
                            var endAng = stAng + swAng;

                            d += window.PPTXShapeUtils.shapeArc(wR, hR, wR, hR, stAng, endAng, false);
                        } else if (multiSapeAry[k].type == "quadBezTo") {
                            console.log("custShapType: quadBezTo - TODO")

                        } else if (multiSapeAry[k].type == "close") {
                            // result += "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            //     "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' ";
                            // result += "/>";
                            // d = "";
                            // isClose = true;

                            d += "z";
                        }
                        k++;
                    }
                    //if (!isClose) {
                    //only one "moveTo" and no "close"
                    result += "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                        "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' ";
                    result += "/>";
                    //console.log(result);
                }

                result += "</svg>";
                result += "<div class='block " + window.PPTXTextStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) + //block content 
                    " " + window.PPTXTextStyleUtils.getContentDir(node, type, warpObj) +
                    "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                    "' style='" +
                    getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                    getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
                    " z-index: " + order + ";" +
                    "transform: rotate(" + ((txtRotate !== undefined) ? txtRotate : 0) + "deg);" +
                    "'>";

                // TextBody
                if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    if (type != "diagram" && type != "textBox") {
                        type = "shape";
                    }
                    result += window.PPTXTextElementUtils.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, undefined, styleTable); //type=shape
                }
                result += "</div>";

                // result = "";
            } else {

                result += "<div class='block " + window.PPTXTextStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +//block content 
                    " " + window.PPTXTextStyleUtils.getContentDir(node, type, warpObj) +
                    "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                    "' style='" +
                    getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                    getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
                    window.PPTXShapeFillsUtils.getBorder(node, pNode, false, "shape", warpObj) +
                    window.PPTXShapeFillsUtils.getShapeFill(node, pNode, false, warpObj, source) +
                    " z-index: " + order + ";" +
                    "transform: rotate(" + ((txtRotate !== undefined) ? txtRotate : 0) + "deg);" +
                    "'>";

                // TextBody
                if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    result += window.PPTXTextElementUtils.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, undefined, styleTable);
                }
                result += "</div>";

            }
            //console.log("div block result:\n", result)
            return result;
        }






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
            var imgFileExt =window.PPTXUtils.extractFileExtension(imgName).toLowerCase();
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
                var idx = window.PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "idx"]);
                var type = window.PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "p:ph", "attrs", "type"]);
                if (idx !== undefined) {
                    xfrmNode = window.PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"], ["idxTable", idx, "p:spPr", "a:xfrm"]);
                }
            }
            ///////////////////////////////////////Amir//////////////////////////////
            var rotate = 0;
            var rotateNode = window.PPTXUtils.getTextByPathList(node, ["p:spPr", "a:xfrm", "attrs", "rot"]);
            if (rotateNode !== undefined) {
                rotate = window.PPTXColorUtils.angleToDegrees(rotateNode);
            }
            //video
            var vdoNode = window.PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:videoFile"]);
            var vdoRid, vdoFile, vdoFileExt, vdoMimeType, uInt8Array, blob, vdoBlob, mediaSupportFlag = false, isVdeoLink = false;
            var mediaProcess = settings.mediaProcess;
            if (vdoNode !== undefined & mediaProcess) {
                vdoRid = vdoNode["attrs"]["r:link"];
                vdoFile = resObj[vdoRid]["target"];
                var checkIfLink = window.PPTXUtils.IsVideoLink(vdoFile);
                if (checkIfLink) {
                    vdoFile = window.PPTXUtils.escapeHtml(vdoFile);
                    //vdoBlob = vdoFile;
                    isVdeoLink = true;
                    mediaSupportFlag = true;
                    mediaPicFlag = true;
                } else {
                    vdoFileExt =window.PPTXUtils.extractFileExtension(vdoFile).toLowerCase();
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
                            vdoMimeType = window.PPTXUtils.getMimeType(vdoFileExt);
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
            var audioNode = window.PPTXUtils.getTextByPathList(node, ["p:nvPicPr", "p:nvPr", "a:audioFile"]);
            var audioRid, audioFile, audioFileExt, audioMimeType, uInt8ArrayAudio, blobAudio, audioBlob;
            var audioPlayerFlag = false;
            var audioObjc;
            if (audioNode !== undefined & mediaProcess) {
                audioRid = audioNode["attrs"]["r:link"];
                audioFile = resObj[audioRid]["target"];
                audioFileExt =window.PPTXUtils.extractFileExtension(audioFile).toLowerCase();
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
            mimeType = window.PPTXUtils.getMimeType(imgFileExt);
            rtrnData = "<div class='block content' style='" +
                ((mediaProcess && audioPlayerFlag) ? getPosition(audioObjc, node, undefined, undefined) : getPosition(xfrmNode, node, undefined, undefined)) +
                ((mediaProcess && audioPlayerFlag) ? getSize(audioObjc, undefined, undefined) : getSize(xfrmNode, undefined, undefined)) +
                " z-index: " + order + ";" +
                "transform: rotate(" + rotate + "deg);'>";
            if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
                rtrnData += "<img src='data:" + mimeType + ";base64," + window.PPTXUtils.base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%; height: 100%'/>";
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
            return window.PPTXImageUtils.processGraphicFrameNode(node, warpObj, source, sType, genTableInternal, genDiagram, processGroupSpNode);
        }

        /**
         * 处理形状属性节点（p:spPr）
         * 
         * @param {Object} node - 形状属性节点
         * @param {Object} warpObj - 包装对象，包含解析上下文
         * @returns {Object} 解析后的形状属性对象（待实现）
         * 
         * TODO: 实现形状属性的完整解析，包括变换、几何、填充、线条、效果等
         * 当前仅返回空对象，需要根据XML架构定义实现
         */
        function processSpPrNode(node, warpObj) {
            // 临时实现：返回空对象
            // console.warn('processSpPrNode: 功能尚未实现，返回空对象');
            return {};
        }




        // genGlobalCSS 已移至 PPTXCSSUtils 模块

        function genTableInternal(node, warpObj) {
            return window.PPTXTableUtils.genTableInternal(node, warpObj, styleTable);
        }

        function getTableCellParams(tcNodes, getColsGrid , row_idx , col_idx , thisTblStyle, cellSource, warpObj) {
            return window.PPTXTableUtils.getTableCellParams(tcNodes, getColsGrid, row_idx, col_idx, thisTblStyle, cellSource, warpObj, styleTable);
        }


        function genDiagram(node, warpObj, source, sType) {
            var readXmlFileFunc = PPTXParser && PPTXParser.readXmlFile ? PPTXParser.readXmlFile : function() { return null; };
            return window.PPTXDiagramUtils.genDiagram(node, warpObj, source, sType, readXmlFileFunc, getPosition, getSize, processSpNode);
        }


        // getPregraphDir 已移至 PPTXTextStyleUtils 模块
        // getHorizontalAlign 已移至 PPTXTextStyleUtils 模块




        // getFontType 已移至 PPTXTextStyleUtils 模块




        // getFontBold, getFontItalic, getFontDecoration, getTextHorizontalAlign, getTextVerticalAlign 已移至 PPTXTextStyleUtils 模块
        // getTableBorders 已移至 PPTXTableUtils 模块
        // getBorder, getShapeFill, getSvgGradient, getSvgImagePattern 已移至 PPTXShapeFillsUtils 模块

        function getBackground(warpObj, slideSize, index) {
            // 使用 PPTXBackgroundUtils 模块处理背景
            return window.PPTXBackgroundUtils.getBackground(warpObj, slideSize, index, processNodesInSlide);
        }
        function getSlideBackgroundFill(warpObj, index) {
            // 使用 PPTXBackgroundUtils 模块处理背景填充
            return window.PPTXBackgroundUtils.getSlideBackgroundFill(warpObj, index);
        }
        // getBgGradientFill, getBgPicFill 已移至 PPTXBackgroundUtils 模块










        // Number formatting functions are now in pptx-utils.js
        // Ensure pptx-utils.js is loaded before this file

        // Export genTextBody for table module
        window._genTextBody = window.PPTXTextElementUtils.genTextBody;
    }


    // Export to global scope
    window.pptxToHtml = pptxToHtml;

})();


