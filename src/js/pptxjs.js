/**
 * pptxjs.js
 * Ver. : 1.21.1
 * last update: 16/11/2021
 * Author: meshesha , https://github.com/meshesha
 * LICENSE: MIT
 * url:https://pptx.js.org/
 * fix issues:
 * [#16](https://github.com/meshesha/PPTXjs/issues/16)
 */

(function ($) {
    $.fn.pptxToHtml = function (options) {
        //var worker;
        var $result = $(this);
        var divId = $result.attr("id");

        var isDone = false;

        //var slideLayoutClrOvride = "";

        var defaultTextStyle = null;

        var chartID = 0;

        var _order = 1;

        var app_verssion ;

        

        var slideFactor = 96 / 914400;
        var fontSizeFactor = 4 / 3.2;
        ////////////////////// 
        var slideWidth = 0;
        var slideHeight = 0;
        var isSlideMode = false;
        var processFullTheme = true;
        var styleTable = {};
        var settings = $.extend(true, {
            // These are the defaults.
            pptxFileUrl: "",
            fileInputId: "",
            slidesScale: "", //Change Slides scale by percent
            slideMode: false, /** true,false*/
            slideType: "divs2slidesjs",  /*'divs2slidesjs' (default) , 'revealjs'(https://revealjs.com)  -TODO*/
            revealjsPath: "", /*path to js file of revealjs - TODO*/
            keyBoardShortCut: false,  /** true,false ,condition: slideMode: true XXXXX - need to remove - this is doublcated*/
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
            revealjsConfig: {},
            styleTable,
        }, options);

        processFullTheme = settings.themeProcess;

        $("#" + divId).prepend(
            $("<div></div>").attr({
                "class": "slides-loadnig-msg",
                "style": "display:block; width:100%; color:white; background-color: #ddd;"
            })/*.html("Loading...")*/
                .html($("<div></div>").attr({
                    "class": "slides-loading-progress-bar",
                    "style": "width: 1%; background-color: #4775d1;"
                }).html("<span style='text-align: center;'>Loading... (1%)</span>"))
        );
        if (settings.slideMode) {
            if (!jQuery().divs2slides) {
                jQuery.getScript('./js/divs2slides.js');
            }
        }
        if (settings.jsZipV2 !== false) {
            jQuery.getScript(settings.jsZipV2);
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
                } else if (key == 116 && isSlideMode) {
                    //exit slide mode - TODO

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
                $(".slides-loadnig-msg").remove();
            }
        } else {
            $(".slides-loadnig-msg").remove()
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
            //console.log("percent: ", percent)
            var progressBarElemtnt = $(".slides-loading-progress-bar")
            progressBarElemtnt.width(percent + "%")
            progressBarElemtnt.html("<span style='text-align: center;'>Loading...(" + percent + "%)</span>");
        }

        function convertToHtml(file) {
            //'use strict';
            //console.log("file", file, "size:", file.byteLength);
            if (file.byteLength < 10){
                console.error("file url error (" + settings.pptxFileUrl + "0)")
                $(".slides-loadnig-msg").remove();
                return;
            }
            var MsgQueue = new Array();
            var zip = new JSZip(), s;
            //if (typeof file === 'string') { // Load
            zip = zip.load(file);  //zip.load(file, { base64: true });
            var rslt_ary = processPPTX(zip, MsgQueue);

            //s = PPTXXmlUtils.readXmlFile(zip, 'ppt/tableStyles.xml');
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
                        processMsgQueue(MsgQueue);
                        setNumericBullets($(".block"));
                        setNumericBullets($("table td"));

                        isDone = true;

                        if (settings.slideMode && !isSlideMode) {
                            isSlideMode = true;
                            initSlideMode(divId, settings);
                        } else if (!settings.slideMode) {
                            $(".slides-loadnig-msg").remove();
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

                if (document.getElementById("all_slides_warpper") === null) {
                    $("#" + divId + " .slide").wrapAll("<div id='all_slides_warpper' class='slides'></div>");
                    //$("#" + divId + " .slides").wrap("<div class='reveal'></div>");
                }

                if (settings.slideMode && settings.slideType == "revealjs") {
                    $("#" + divId).addClass("reveal")
                }
            }

            var sScale = settings.slidesScale;
            var trnsfrmScl = "";
            if (sScale != "") {
                var numsScale = parseInt(sScale);
                var scaleVal = numsScale / 100;
                if (settings.slideMode && settings.slideType != "revealjs") {
                    trnsfrmScl = 'transform:scale(' + scaleVal + '); transform-origin:top';
                }
            }

            var slidesHeight = $("#" + divId + " .slide").height();
            var numOfSlides = $("#" + divId + " .slide").length;
            var sScaleVal = (sScale != "") ? scaleVal : 1;
            //console.log("slidesHeight: " + slidesHeight + "\nnumOfSlides: " + numOfSlides + "\nScale: " + sScaleVal)

            $("#all_slides_warpper").attr({
                style: trnsfrmScl + ";height: " + (numOfSlides * slidesHeight * sScaleVal) + "px"
            })

            //}
        }

        function initSlideMode(divId, settings) {
            //console.log(settings.slideType)
            if (settings.slideType == "" || settings.slideType == "divs2slidesjs") {
                var slidesHeight = $("#" + divId + " .slide").height();
                $("#" + divId + " .slide").hide();
                setTimeout(function () {
                    var slideConf = settings.slideModeConfig;
                    $(".slides-loadnig-msg").remove();
                    $("#" + divId).divs2slides({
                        first: slideConf.first,
                        nav: slideConf.nav,
                        showPlayPauseBtn: settings.showPlayPauseBtn,
                        navTxtColor: slideConf.navTxtColor,
                        keyBoardShortCut: slideConf.keyBoardShortCut,
                        showSlideNum: slideConf.showSlideNum,
                        showTotalSlideNum: slideConf.showTotalSlideNum,
                        autoSlide: slideConf.autoSlide,
                        randomAutoSlide: slideConf.randomAutoSlide,
                        loop: slideConf.loop,
                        background: slideConf.background,
                        transition: slideConf.transition,
                        transitionTime: slideConf.transitionTime
                    });

                    var sScale = settings.slidesScale;
                    var trnsfrmScl = "";
                    if (sScale != "") {
                        var numsScale = parseInt(sScale);
                        var scaleVal = numsScale / 100;
                        trnsfrmScl = 'transform:scale(' + scaleVal + '); transform-origin:top';
                    }

                    var numOfSlides = 1;
                    var sScaleVal = (sScale != "") ? scaleVal : 1;
                    //console.log(slidesHeight);
                    $("#all_slides_warpper").attr({
                        style: trnsfrmScl + ";height: " + (numOfSlides * slidesHeight * sScaleVal) + "px"
                    })

                }, 1500);
            } else if (settings.slideType == "revealjs") {
                $(".slides-loadnig-msg").remove();
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

        function processPPTX(zip, MsgQueue) {
            var post_ary = [];
            var dateBefore = new Date();

            if (zip.file("docProps/thumbnail.jpeg") !== null) {
                var pptxThumbImg = PPTXXmlUtils.base64ArrayBuffer(zip.file("docProps/thumbnail.jpeg").asArrayBuffer());
                post_ary.push({
                    "type": "pptx-thumb",
                    "data": pptxThumbImg,
                    "slide_num": -1
                });
            }

            var filesInfo = PPTXXmlUtils.getContentTypes(zip);
            var slideSize = PPTXXmlUtils.getSlideSizeAndSetDefaultTextStyle(zip, settings);
            tableStyles = PPTXXmlUtils.readXmlFile(zip, "ppt/tableStyles.xml");
            //console.log("slideSize: ", slideSize)
            post_ary.push({
                "type": "slideSize",
                "data": slideSize,
                "slide_num": 0
            });

            var numOfSlides = filesInfo["slides"].length;
            for (var i = 0; i < numOfSlides; i++) {
                var filename = filesInfo["slides"][i];
                var filename_no_path = "";
                var filename_no_path_ary = [];
                if (filename.indexOf("/") != -1) {
                    filename_no_path_ary = filename.split("/");
                    filename_no_path = filename_no_path_ary.pop();
                } else {
                    filename_no_path = filename;
                }
                var filename_no_path_no_ext = "";
                if (filename_no_path.indexOf(".") != -1) {
                    var filename_no_path_no_ext_ary = filename_no_path.split(".");
                    var slide_ext = filename_no_path_no_ext_ary.pop();
                    filename_no_path_no_ext = filename_no_path_no_ext_ary.join(".");
                }
                var slide_number = 1;
                if (filename_no_path_no_ext != "" && filename_no_path.indexOf("slide") != -1) {
                    slide_number = Number(filename_no_path_no_ext.substr(5));
                }
                var slideHtml = processSingleSlide(zip, filename, i, slideSize, MsgQueue);
                post_ary.push({
                    "type": "slide",
                    "data": slideHtml,
                    "slide_num": slide_number,
                    "file_name": filename_no_path_no_ext
                });
                post_ary.push({
                    "type": "progress-update",
                    "slide_num": (numOfSlides + i + 1),
                    "data": (i + 1) * 100 / numOfSlides
                });
            }

            post_ary.sort(function (a, b) {
                return a.slide_num - b.slide_num;
            });

            post_ary.push({
                "type": "globalCSS",
                "data": genGlobalCSS()
            });

            var dateAfter = new Date();
            post_ary.push({
                "type": "ExecutionTime",
                "data": dateAfter - dateBefore
            });
            return post_ary;
        }

        function processSingleSlide(zip, sldFileName, index, slideSize, MsgQueue) {
            /*
            self.postMessage({
                "type": "INFO",
                "data": "Processing slide" + (index + 1)
            });
            */
            // =====< Step 1 >=====
            // Read relationship filename of the slide (Get slideLayoutXX.xml)
            // @sldFileName: ppt/slides/slide1.xml
            // @resName: ppt/slides/_rels/slide1.xml.rels
            var resName = sldFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
            var resContent = PPTXXmlUtils.readXmlFile(zip, resName);
            var RelationshipArray = resContent["Relationships"]["Relationship"];
            //console.log("RelationshipArray: " , RelationshipArray)
            var layoutFilename = "";
            var diagramFilename = "";
            var slideResObj = {};
            if (RelationshipArray.constructor === Array) {
                for (var i = 0; i < RelationshipArray.length; i++) {
                    switch (RelationshipArray[i]["attrs"]["Type"]) {
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                            layoutFilename = RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
                            break;
                        case "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing":
                            diagramFilename = RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
                            slideResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                                "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                "target": RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/")
                            };
                            break;
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart":
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":
                        default:
                            slideResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                                "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                "target": RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/")
                            };
                    }
                }
            } else {
                layoutFilename = RelationshipArray["attrs"]["Target"].replace("../", "ppt/");
            }
            //console.log(slideResObj);
            // Open slideLayoutXX.xml
            var slideLayoutContent = PPTXXmlUtils.readXmlFile(zip, layoutFilename);
            var slideLayoutTables = PPTXNodeUtils.indexNodes(slideLayoutContent);
            var sldLayoutClrOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);

            //console.log(slideLayoutClrOvride);
            if (sldLayoutClrOvr !== undefined) {
                slideLayoutClrOvride = sldLayoutClrOvr["attrs"];
            }
            // =====< Step 2 >=====
            // Read slide master filename of the slidelayout (Get slideMasterXX.xml)
            // @resName: ppt/slideLayouts/slideLayout1.xml
            // @masterName: ppt/slideLayouts/_rels/slideLayout1.xml.rels
            var slideLayoutResFilename = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
            var slideLayoutResContent = PPTXXmlUtils.readXmlFile(zip, slideLayoutResFilename);
            RelationshipArray = slideLayoutResContent["Relationships"]["Relationship"];
            var masterFilename = "";
            var layoutResObj = {};
            if (RelationshipArray.constructor === Array) {
                for (var i = 0; i < RelationshipArray.length; i++) {
                    switch (RelationshipArray[i]["attrs"]["Type"]) {
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
                            masterFilename = RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
                            break;
                        default:
                            layoutResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                                "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                "target": RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/")
                            };
                    }
                }
            } else {
                masterFilename = RelationshipArray["attrs"]["Target"].replace("../", "ppt/");
            }
            // Open slideMasterXX.xml
            var slideMasterContent = PPTXXmlUtils.readXmlFile(zip, masterFilename);
            var slideMasterTextStyles = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
            var slideMasterTables = PPTXNodeUtils.indexNodes(slideMasterContent);

            /////////////////Amir/////////////
            //Open slideMasterXX.xml.rels
            var slideMasterResFilename = masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster") + ".rels";
            var slideMasterResContent = PPTXXmlUtils.readXmlFile(zip, slideMasterResFilename);
            RelationshipArray = slideMasterResContent["Relationships"]["Relationship"];
            var themeFilename = "";
            var masterResObj = {};
            if (RelationshipArray.constructor === Array) {
                for (var i = 0; i < RelationshipArray.length; i++) {
                    switch (RelationshipArray[i]["attrs"]["Type"]) {
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                            themeFilename = RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
                            break;
                        default:
                            masterResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                                "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                "target": RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/")
                            };
                    }
                }
            } else {
                themeFilename = RelationshipArray["attrs"]["Target"].replace("../", "ppt/");
            }
            //console.log(themeFilename)
            //Load Theme file
            var themeResObj = {};
            if (themeFilename !== undefined) {
                var themeName = themeFilename.split("/").pop();
                var themeResFileName = themeFilename.replace(themeName, "_rels/" + themeName) + ".rels";
                //console.log("themeFilename: ", themeFilename, ", themeName: ", themeName, ", themeResFileName: ", themeResFileName)
                var themeContent = PPTXXmlUtils.readXmlFile(zip, themeFilename);
                var themeResContent = PPTXXmlUtils.readXmlFile(zip, themeResFileName);
                if (themeResContent !== null) {
                    var relationshipArray = themeResContent["Relationships"]["Relationship"];
                    if (relationshipArray !== undefined){
                        var themeFilename = "";
                        if (relationshipArray.constructor === Array) {
                            for (var i = 0; i < relationshipArray.length; i++) {
                                themeResObj[relationshipArray[i]["attrs"]["Id"]] = {
                                    "type": relationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                    "target": relationshipArray[i]["attrs"]["Target"].replace("../", "ppt/")
                                };
                            }
                        } else {
                            //console.log("theme relationshipArray : ", relationshipArray)
                            themeResObj[relationshipArray["attrs"]["Id"]] = {
                                "type": relationshipArray["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                "target": relationshipArray["attrs"]["Target"].replace("../", "ppt/")
                            };
                        }
                    }
                }
            }
            //Load diagram file
            var diagramResObj = {};
            var digramFileContent = {};
            if (diagramFilename !== undefined) {
                var diagName = diagramFilename.split("/").pop();
                var diagramResFileName = diagramFilename.replace(diagName, "_rels/" + diagName) + ".rels";
                //console.log("diagramFilename: ", diagramFilename, ", themeName: ", themeName, ", diagramResFileName: ", diagramResFileName)
                digramFileContent = PPTXXmlUtils.readXmlFile(zip, diagramFilename);
                if (digramFileContent !== null && digramFileContent !== undefined && digramFileContent != "") {
                    var digramFileContentObjToStr = JSON.stringify(digramFileContent);
                    digramFileContentObjToStr = digramFileContentObjToStr.replace(/dsp:/g, "p:");
                    digramFileContent = JSON.parse(digramFileContentObjToStr);
                }

                var digramResContent = PPTXXmlUtils.readXmlFile(zip, diagramResFileName);
                if (digramResContent !== null) {
                    var relationshipArray = digramResContent["Relationships"]["Relationship"];
                    var themeFilename = "";
                    if (relationshipArray.constructor === Array) {
                        for (var i = 0; i < relationshipArray.length; i++) {
                            diagramResObj[relationshipArray[i]["attrs"]["Id"]] = {
                                "type": relationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                "target": relationshipArray[i]["attrs"]["Target"].replace("../", "ppt/")
                            };
                        }
                    } else {
                        //console.log("theme relationshipArray : ", relationshipArray)
                        diagramResObj[relationshipArray["attrs"]["Id"]] = {
                            "type": relationshipArray["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            "target": relationshipArray["attrs"]["Target"].replace("../", "ppt/")
                        };
                    }
                }
            }
            //console.log("diagramResObj: " , diagramResObj)
            // =====< Step 3 >=====
            var slideContent = PPTXXmlUtils.readXmlFile(zip, sldFileName , true);
            var nodes = slideContent["p:sld"]["p:cSld"]["p:spTree"];
            var warpObj = {
                "zip": zip,
                "slideLayoutContent": slideLayoutContent,
                "slideLayoutTables": slideLayoutTables,
                "slideMasterContent": slideMasterContent,
                "slideMasterTables": slideMasterTables,
                "slideContent": slideContent,
                "slideResObj": slideResObj,
                "slideMasterTextStyles": slideMasterTextStyles,
                "layoutResObj": layoutResObj,
                "masterResObj": masterResObj,
                "themeContent": themeContent,
                "themeResObj": themeResObj,
                "digramFileContent": digramFileContent,
                "diagramResObj": diagramResObj,
                "defaultTextStyle": slideSize.defaultTextStyle || defaultTextStyle,
                styleTable,
                chartID,
                MsgQueue,
            };
            var bgResult = "";
            if (processFullTheme === true) {
                bgResult = PPTXStyleUtils.getBackground(warpObj, slideSize, index, settings);
            }

            var bgColor = "";
            if (processFullTheme == "colorsAndImageOnly") {
                bgColor = PPTXStyleUtils.getSlideBackgroundFill(warpObj, index);
            }

            if (settings.slideMode && settings.slideType == "revealjs") {
                var result = "<section class='slide' style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;" + bgColor + "'>"
            } else {
                var result = "<div class='slide' style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;" + bgColor + "'>"
            }
            result += bgResult;
            for (var nodeKey in nodes) {
                if (nodes[nodeKey].constructor === Array) {
                    for (var i = 0; i < nodes[nodeKey].length; i++) {
                        result += PPTXNodeUtils.processNodesInSlide(nodeKey, nodes[nodeKey][i], nodes, warpObj, "slide", 'group', settings);
                    }
                } else {
                    result += PPTXNodeUtils.processNodesInSlide(nodeKey, nodes[nodeKey], nodes, warpObj, "slide", 'group', settings);
                }
            }
            if (settings.slideMode && settings.slideType == "revealjs") {
                return result + "</div></section>";
            } else {
                return result + "</div></div>";
            }

        }

        


        function shapePie(H, w, adj1, adj2, isClose) {
            var pieVal = parseInt(adj2);
            var piAngle = parseInt(adj1);
            var size = parseInt(H),
                radius = (size / 2),
                value = pieVal - piAngle;
            if (value < 0) {
                value = 360 + value;
            }
            //console.log("value: ",value)      
            value = Math.min(Math.max(value, 0), 360);

            //calculate x,y coordinates of the point on the circle to draw the arc to. 
            var x = Math.cos((2 * Math.PI) / (360 / value));
            var y = Math.sin((2 * Math.PI) / (360 / value));


            //d is a string that describes the path of the slice.
            var longArc, d, rot;
            if (isClose) {
                longArc = (value <= 180) ? 0 : 1;
                d = "M" + radius + "," + radius + " L" + radius + "," + 0 + " A" + radius + "," + radius + " 0 " + longArc + ",1 " + (radius + y * radius) + "," + (radius - x * radius) + " z";
                rot = "rotate(" + (piAngle - 270) + ", " + radius + ", " + radius + ")";
            } else {
                longArc = (value <= 180) ? 0 : 1;
                var radius1 = radius;
                var radius2 = w / 2;
                d = "M" + radius1 + "," + 0 + " A" + radius2 + "," + radius1 + " 0 " + longArc + ",1 " + (radius2 + y * radius2) + "," + (radius1 - x * radius1);
                rot = "rotate(" + (piAngle + 90) + ", " + radius + ", " + radius + ")";
            }

            return [d, rot];
        }
        function shapeGear(w, h, points) {
            var innerRadius = h;//gear.innerRadius;
            var outerRadius = 1.5 * innerRadius;
            var cx = outerRadius;//Math.max(innerRadius, outerRadius),                   // center x
            cy = outerRadius;//Math.max(innerRadius, outerRadius),                    // center y
            notches = points,//gear.points,                      // num. of notches
                radiusO = outerRadius,                    // outer radius
                radiusI = innerRadius,                    // inner radius
                taperO = 50,                     // outer taper %
                taperI = 35,                     // inner taper %

                // pre-calculate values for loop

                pi2 = 2 * Math.PI,            // cache 2xPI (360deg)
                angle = pi2 / (notches * 2),    // angle between notches
                taperAI = angle * taperI * 0.005, // inner taper offset (100% = half notch)
                taperAO = angle * taperO * 0.005, // outer taper offset
                a = angle,                  // iterator (angle)
                toggle = false;
            // move to starting point
            var d = " M" + (cx + radiusO * Math.cos(taperAO)) + " " + (cy + radiusO * Math.sin(taperAO));

            // loop
            for (; a <= pi2 + angle; a += angle) {
                // draw inner to outer line
                if (toggle) {
                    d += " L" + (cx + radiusI * Math.cos(a - taperAI)) + "," + (cy + radiusI * Math.sin(a - taperAI));
                    d += " L" + (cx + radiusO * Math.cos(a + taperAO)) + "," + (cy + radiusO * Math.sin(a + taperAO));
                } else { // draw outer to inner line
                    d += " L" + (cx + radiusO * Math.cos(a - taperAO)) + "," + (cy + radiusO * Math.sin(a - taperAO)); // outer line
                    d += " L" + (cx + radiusI * Math.cos(a + taperAI)) + "," + (cy + radiusI * Math.sin(a + taperAI));// inner line

                }
                // switch level
                toggle = !toggle;
            }
            // close the final line
            d += " ";
            return d;
        }
        function shapeSnipRoundRect(w, h, adj1, adj2, shapeType, adjType) {
            /* 
            shapeType: snip,round
            adjType: cornr1,cornr2,cornrAll,diag
            */
            var adjA, adjB, adjC, adjD;
            if (adjType == "cornr1") {
                adjA = 0;
                adjB = 0;
                adjC = 0;
                adjD = adj1;
            } else if (adjType == "cornr2") {
                adjA = adj1;
                adjB = adj2;
                adjC = adj2;
                adjD = adj1;
            } else if (adjType == "cornrAll") {
                adjA = adj1;
                adjB = adj1;
                adjC = adj1;
                adjD = adj1;
            } else if (adjType == "diag") {
                adjA = adj1;
                adjB = adj2;
                adjC = adj1;
                adjD = adj2;
            }
            //d is a string that describes the path of the slice.
            var d;
            if (shapeType == "round") {
                d = "M0" + "," + (h / 2 + (1 - adjB) * (h / 2)) + " Q" + 0 + "," + h + " " + adjB * (w / 2) + "," + h + " L" + (w / 2 + (1 - adjC) * (w / 2)) + "," + h +
                    " Q" + w + "," + h + " " + w + "," + (h / 2 + (h / 2) * (1 - adjC)) + "L" + w + "," + (h / 2) * adjD +
                    " Q" + w + "," + 0 + " " + (w / 2 + (w / 2) * (1 - adjD)) + ",0 L" + (w / 2) * adjA + ",0" +
                    " Q" + 0 + "," + 0 + " 0," + (h / 2) * (adjA) + " z";
            } else if (shapeType == "snip") {
                d = "M0" + "," + adjA * (h / 2) + " L0" + "," + (h / 2 + (h / 2) * (1 - adjB)) + "L" + adjB * (w / 2) + "," + h +
                    " L" + (w / 2 + (w / 2) * (1 - adjC)) + "," + h + "L" + w + "," + (h / 2 + (h / 2) * (1 - adjC)) +
                    " L" + w + "," + adjD * (h / 2) + "L" + (w / 2 + (w / 2) * (1 - adjD)) + ",0 L" + ((w / 2) * adjA) + ",0 z";
            }
            return d;
        }
        /*
        function shapePolygon(sidesNum) {
            var sides  = sidesNum;
            var radius = 100;
            var angle  = 2 * Math.PI / sides;
            var points = []; 
            
            for (var i = 0; i < sides; i++) {
                points.push(radius + radius * Math.sin(i * angle));
                points.push(radius - radius * Math.cos(i * angle));
            }
            
            return points;
        }
        */
        



        
        

        

        function genGlobalCSS() {
            var cssText = "";
            //console.log("styleTable: ", styleTable)
            for (var key in styleTable) {
                var tagname = "";
                // if (settings.slideMode && settings.slideType == "revealjs") {
                //     tagname = "section";
                // } else {
                //     tagname = "div";
                // }
                //ADD suffix
                cssText += tagname + " ." + styleTable[key]["name"] +
                    ((styleTable[key]["suffix"]) ? styleTable[key]["suffix"] : "") +
                    "{" + styleTable[key]["text"] + "}\n"; //section > div
            }
            //cssText += " .slide{margin-bottom: 5px;}\n"; // TODO

            if (settings.slideMode && settings.slideType == "divs2slidesjs") {
                //divId
                //console.log("slideWidth: ", slideWidth)
                cssText += "#all_slides_warpper{margin-right: auto;margin-left: auto;padding-top:10px;width: " + slideWidth + "px;}\n"; // TODO
            }
            return cssText;
        }



        function processMsgQueue(queue) {
            for (var i = 0; i < queue.length; i++) {
                processSingleMsg(queue[i].data);
            }
        }

        function processSingleMsg(d) {

            var chartID = d.chartID;
            var chartType = d.chartType;
            var chartData = d.chartData;

            var data = [];

            var chart = null;
            switch (chartType) {
                case "lineChart":
                    data = chartData;
                    chart = nv.models.lineChart()
                        .useInteractiveGuideline(true);
                    chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
                    break;
                case "barChart":
                    data = chartData;
                    chart = nv.models.multiBarChart();
                    chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
                    break;
                case "pieChart":
                case "pie3DChart":
                    if (chartData.length > 0) {
                        data = chartData[0].values;
                    }
                    chart = nv.models.pieChart();
                    break;
                case "areaChart":
                    data = chartData;
                    chart = nv.models.stackedAreaChart()
                        .clipEdge(true)
                        .useInteractiveGuideline(true);
                    chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
                    break;
                case "scatterChart":

                    for (var i = 0; i < chartData.length; i++) {
                        var arr = [];
                        for (var j = 0; j < chartData[i].length; j++) {
                            arr.push({ x: j, y: chartData[i][j] });
                        }
                        data.push({ key: 'data' + (i + 1), values: arr });
                    }

                    //data = chartData;
                    chart = nv.models.scatterChart()
                        .showDistX(true)
                        .showDistY(true)
                        .color(d3.scale.category10().range());
                    chart.xAxis.axisLabel('X').tickFormat(d3.format('.02f'));
                    chart.yAxis.axisLabel('Y').tickFormat(d3.format('.02f'));
                    break;
                default:
            }

            if (chart !== null) {

                d3.select("#" + chartID)
                    .append("svg")
                    .datum(data)
                    .transition().duration(500)
                    .call(chart);

                nv.utils.windowResize(chart.update);
                isDone = true;
            }

        }

        function setNumericBullets(elem) {
            var prgrphs_arry = elem;
            for (var i = 0; i < prgrphs_arry.length; i++) {
                var buSpan = $(prgrphs_arry[i]).find('.numeric-bullet-style');
                if (buSpan.length > 0) {
                    //console.log("DIV-"+i+":");
                    var prevBultTyp = "";
                    var prevBultLvl = "";
                    var buletIndex = 0;
                    var tmpArry = new Array();
                    var tmpArryIndx = 0;
                    var buletTypSrry = new Array();
                    for (var j = 0; j < buSpan.length; j++) {
                        var bult_typ = $(buSpan[j]).data("bulltname");
                        var bult_lvl = $(buSpan[j]).data("bulltlvl");
                        //console.log(j+" - "+bult_typ+" lvl: "+bult_lvl );
                        if (buletIndex == 0) {
                            prevBultTyp = bult_typ;
                            prevBultLvl = bult_lvl;
                            tmpArry[tmpArryIndx] = buletIndex;
                            buletTypSrry[tmpArryIndx] = bult_typ;
                            buletIndex++;
                        } else {
                            if (bult_typ == prevBultTyp && bult_lvl == prevBultLvl) {
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                buletIndex++;
                                tmpArry[tmpArryIndx] = buletIndex;
                                buletTypSrry[tmpArryIndx] = bult_typ;
                            } else if (bult_typ != prevBultTyp && bult_lvl == prevBultLvl) {
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                tmpArryIndx++;
                                tmpArry[tmpArryIndx] = buletIndex;
                                buletTypSrry[tmpArryIndx] = bult_typ;
                                buletIndex = 1;
                            } else if (bult_typ != prevBultTyp && Number(bult_lvl) > Number(prevBultLvl)) {
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                tmpArryIndx++;
                                tmpArry[tmpArryIndx] = buletIndex;
                                buletTypSrry[tmpArryIndx] = bult_typ;
                                buletIndex = 1;
                            } else if (bult_typ != prevBultTyp && Number(bult_lvl) < Number(prevBultLvl)) {
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                tmpArryIndx--;
                                buletIndex = tmpArry[tmpArryIndx] + 1;
                            }
                        }
                        //console.log(buletTypSrry[tmpArryIndx]+" - "+buletIndex);
                        var numIdx = getNumTypeNum(buletTypSrry[tmpArryIndx], buletIndex);
                        $(buSpan[j]).html(numIdx);
                    }
                }
            }
        }
        function getNumTypeNum(numTyp, num) {
            var rtrnNum = "";
            switch (numTyp) {
                case "arabicPeriod":
                    rtrnNum = num + ". ";
                    break;
                case "arabicParenR":
                    rtrnNum = num + ") ";
                    break;
                case "alphaLcParenR":
                    rtrnNum = alphaNumeric(num, "lowerCase") + ") ";
                    break;
                case "alphaLcPeriod":
                    rtrnNum = alphaNumeric(num, "lowerCase") + ". ";
                    break;

                case "alphaUcParenR":
                    rtrnNum = alphaNumeric(num, "upperCase") + ") ";
                    break;
                case "alphaUcPeriod":
                    rtrnNum = alphaNumeric(num, "upperCase") + ". ";
                    break;

                case "romanUcPeriod":
                    rtrnNum = romanize(num) + ". ";
                    break;
                case "romanLcParenR":
                    rtrnNum = romanize(num) + ") ";
                    break;
                case "hebrew2Minus":
                    rtrnNum = hebrew2Minus.format(num) + "-";
                    break;
                default:
                    rtrnNum = num;
            }
            return rtrnNum;
        }
        function romanize(num) {
            if (!+num)
                return false;
            var digits = String(+num).split(""),
                key = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM",
                    "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC",
                    "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"],
                roman = "",
                i = 3;
            while (i--)
                roman = (key[+digits.pop() + (i * 10)] || "") + roman;
            return Array(+digits.join("") + 1).join("M") + roman;
        }
        var hebrew2Minus = archaicNumbers([
            [1000, ''],
            [400, 'ת'],
            [300, 'ש'],
            [200, 'ר'],
            [100, 'ק'],
            [90, 'צ'],
            [80, 'פ'],
            [70, 'ע'],
            [60, 'ס'],
            [50, 'נ'],
            [40, 'מ'],
            [30, 'ל'],
            [20, 'כ'],
            [10, 'י'],
            [9, 'ט'],
            [8, 'ח'],
            [7, 'ז'],
            [6, 'ו'],
            [5, 'ה'],
            [4, 'ד'],
            [3, 'ג'],
            [2, 'ב'],
            [1, 'א'],
            [/יה/, 'ט״ו'],
            [/יו/, 'ט״ז'],
            [/([א-ת])([א-ת])$/, '$1״$2'],
            [/^([א-ת])$/, "$1׳"]
        ]);
        function archaicNumbers(arr) {
            var arrParse = arr.slice().sort(function (a, b) { return b[1].length - a[1].length });
            return {
                format: function (n) {
                    var ret = '';
                    jQuery.each(arr, function () {
                        var num = this[0];
                        if (parseInt(num) > 0) {
                            for (; n >= num; n -= num) ret += this[1];
                        } else {
                            ret = ret.replace(num, this[1]);
                        }
                    });
                    return ret;
                }
            }
        }
        function alphaNumeric(num, upperLower) {
            num = Number(num) - 1;
            var aNum = "";
            if (upperLower == "upperCase") {
                aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toUpperCase();
            } else if (upperLower == "lowerCase") {
                aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toLowerCase();
            }
            return aNum;
        }

        
    }
}(jQuery));
