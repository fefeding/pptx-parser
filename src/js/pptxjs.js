/* 
*
 * pptxjs.js
 * Ver. : 1.21.1
 * last update: 16/11/2021
 * Author: meshesha , https://github.com/meshesha
 * LICENSE: MIT
 * url:https://pptx.js.org/
 * fix issues:
 * [#16](https://github.com/meshesha/PPTXjs/issues/16)
 */

function pptxToHtml(element, options) {
    var result = element;
    var divId = result.id;

    var isDone = false;

    var defaultTextStyle = null;

    var chartID = 0;

    var _order = 1;

    var app_verssion ;

    var slideFactor = 96 / 914400;
    var fontSizeFactor = 4 / 3.2;
    var slideWidth = 0;
    var slideHeight = 0;
    var isSlideMode = false;
    var processFullTheme = true;
    var styleTable = {};
    var settings = Object.assign({}, {
        pptxFileUrl: "",
        fileInputId: "",
        slidesScale: "",
        slideMode: false,
        slideType: "divs2slidesjs",
        revealjsPath: "",
        keyBoardShortCut: false,
        mediaProcess: true,
        jsZipV2: false,
        themeProcess: true,
        incSlide:{
            width: 0,
            height: 0
        },
        slideModeConfig: {
            first:1,
            nav: true,
            navTxtColor: "black",
            keyBoardShortCut: true,
            showSlideNum: true,
            showTotalSlideNum: true,
            autoSlide: true,
            randomAutoSlide: false,
            loop: false,
            background: false,
            transition: "default",
            transitionTime: 1
        },
        revealjsConfig: {},
        styleTable,
    }, options);

    processFullTheme = settings.themeProcess;

    var loadingMsgDiv = document.createElement("div");
    loadingMsgDiv.className = "slides-loadnig-msg";
    loadingMsgDiv.style.cssText = "display:block; width:100%; color:white; background-color: #ddd;";
    
    var progressBarDiv = document.createElement("div");
    progressBarDiv.className = "slides-loading-progress-bar";
    progressBarDiv.style.cssText = "width: 1%; background-color: #4775d1;";
    progressBarDiv.innerHTML = "<span style='text-align: center;'>Loading... (1%)</span>";
    
    loadingMsgDiv.appendChild(progressBarDiv);
    
    var targetDiv = document.getElementById(divId);
    if (targetDiv) {
        targetDiv.insertBefore(loadingMsgDiv, targetDiv.firstChild);
    }

    if (settings.slideMode) {
        if (typeof divs2slides === 'undefined') {
            var script = document.createElement('script');
            script.src = './js/divs2slides.js';
            document.head.appendChild(script);
        }
    }
    if (settings.jsZipV2 !== false) {
        var script = document.createElement('script');
        script.src = settings.jsZipV2;
        document.head.appendChild(script);
        if (localStorage.getItem('isPPTXjsReLoaded') !== 'yes') {
            localStorage.setItem('isPPTXjsReLoaded', 'yes');
            location.reload();
        }
    }

    if (settings.keyBoardShortCut) {
        document.addEventListener("keydown", function (event) {
            event.preventDefault();
            var key = event.keyCode;
            console.log(key, isDone)
            if (key == 116 && !isSlideMode) {
                isSlideMode = true;
                initSlideMode(divId, settings);
            } else if (key == 116 && isSlideMode) {
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
                            convertToHtml(e.target.result);
                        }
                    }
                });
            });
        }catch(e){ 
            console.error("file url error (" + settings.pptxFileUrl+ "0)")
            var loadingMsg = document.querySelector(".slides-loadnig-msg");
            if (loadingMsg) {
                loadingMsg.remove();
            }
        }
    } else {
        var loadingMsg = document.querySelector(".slides-loadnig-msg");
        if (loadingMsg) {
            loadingMsg.remove();
        }
    }
    if (settings.fileInputId != "") {
        var fileInput = document.getElementById(settings.fileInputId);
        if (fileInput) {
            fileInput.addEventListener("change", function (evt) {
                result.innerHTML = "";
                var file = evt.target.files[0];
                var fileType = file.type;
                if (fileType == "application/vnd.openxmlformats-officedocument.presentationml.presentation") {
                    FileReaderJS.setupBlob(file, {
                        readAsDefault: "ArrayBuffer",
                        on: {
                            load: function (e, file) {
                                convertToHtml(e.target.result);
                            }
                        }
                    });
                } else {
                    alert("This is not pptx file");
                }
            });
        }
    }

    function updateProgressBar(percent) {
        var progressBarElemtnt = document.querySelector(".slides-loading-progress-bar");
        if (progressBarElemtnt) {
            progressBarElemtnt.style.width = percent + "%";
            progressBarElemtnt.innerHTML = "<span style='text-align: center;'>Loading...(" + percent + "%)</span>";
        }
    }

    function convertToHtml(file) {
        if (file.byteLength < 10){
            console.error("file url error (" + settings.pptxFileUrl + "0)")
            var loadingMsg = document.querySelector(".slides-loadnig-msg");
            if (loadingMsg) {
                loadingMsg.remove();
            }
            return;
        }
        var MsgQueue = new Array();
        var zip = new JSZip(), s;
        zip = zip.load(file);
        var rslt_ary = processPPTX(zip, MsgQueue);

        for (var i = 0; i < rslt_ary.length; i++) {
            switch (rslt_ary[i]["type"]) {
                case "slide":
                    result.innerHTML += rslt_ary[i]["data"];
                    break;
                case "pptx-thumb":
                    break;
                case "slideSize":
                    slideWidth = rslt_ary[i]["data"].width;
                    slideHeight = rslt_ary[i]["data"].height;
                    break;
                case "globalCSS":
                    var style = document.createElement("style");
                    style.innerHTML = rslt_ary[i]["data"];
                    result.appendChild(style);
                    break;
                case "ExecutionTime":
                    processMsgQueue(MsgQueue);
                    setNumericBullets(document.querySelectorAll(".block"));
                    setNumericBullets(document.querySelectorAll("table td"));

                    isDone = true;

                    if (settings.slideMode && !isSlideMode) {
                        isSlideMode = true;
                        initSlideMode(divId, settings);
                    } else if (!settings.slideMode) {
                        var loadingMsg = document.querySelector(".slides-loadnig-msg");
                        if (loadingMsg) {
                            loadingMsg.remove();
                        }
                    }
                    break;
                case "progress-update":
                    updateProgressBar(rslt_ary[i]["data"])
                    break;
                default:
            }
        }
        if (!settings.slideMode || (settings.slideMode && settings.slideType == "revealjs")) {

            if (document.getElementById("all_slides_warpper") === null) {
                var slides = document.querySelectorAll("#" + divId + " .slide");
                var wrapper = document.createElement("div");
                wrapper.id = "all_slides_warpper";
                wrapper.className = "slides";
                var firstSlide = slides[0];
                if (firstSlide && firstSlide.parentNode) {
                    firstSlide.parentNode.insertBefore(wrapper, firstSlide);
                    for (var i = 0; i < slides.length; i++) {
                        wrapper.appendChild(slides[i]);
                    }
                }
            }

            if (settings.slideMode && settings.slideType == "revealjs") {
                document.getElementById(divId).classList.add("reveal");
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

        var slides = document.querySelectorAll("#" + divId + " .slide");
        var slidesHeight = slides.length > 0 ? slides[0].offsetHeight : 0;
        var numOfSlides = slides.length;
        var sScaleVal = (sScale != "") ? scaleVal : 1;

        var allSlidesWrapper = document.getElementById("all_slides_warpper");
        if (allSlidesWrapper) {
            allSlidesWrapper.style.cssText = trnsfrmScl + ";height: " + (numOfSlides * slidesHeight * sScaleVal) + "px";
        }
    }

        function initSlideMode(divId, settings) {
        if (settings.slideType == "" || settings.slideType == "divs2slidesjs") {
            var slides = document.querySelectorAll("#" + divId + " .slide");
            var slidesHeight = slides.length > 0 ? slides[0].offsetHeight : 0;
            for (var i = 0; i < slides.length; i++) {
                slides[i].style.display = "none";
            }
            setTimeout(function () {
                var slideConf = settings.slideModeConfig;
                var loadingMsg = document.querySelector(".slides-loadnig-msg");
                if (loadingMsg) {
                    loadingMsg.remove();
                }
                var resultDiv = document.getElementById(divId);
                if (resultDiv && typeof resultDiv.divs2slides === 'function') {
                    resultDiv.divs2slides({
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

                    var allSlidesWrapper = document.getElementById("all_slides_warpper");
                    if (allSlidesWrapper) {
                        allSlidesWrapper.style.cssText = trnsfrmScl + ";height: " + (numOfSlides * slidesHeight * sScaleVal) + "px";
                    }
                }
            }, 1500);
        } else if (settings.slideType == "revealjs") {
            var loadingMsg = document.querySelector(".slides-loadnig-msg");
            if (loadingMsg) {
                loadingMsg.remove();
            }
            var revealjsPath = "";
            if (settings.revealjsPath != "") {
                revealjsPath = settings.revealjsPath;
            } else {
                revealjsPath = "./revealjs/reveal.js";
            }
            var script = document.createElement('script');
            script.src = revealjsPath;
            script.onload = function() {
                var sections = document.querySelectorAll("section");
                for (var i = 0; i < sections.length; i++) {
                    sections[i].classList.remove("slide");
                }
                if (typeof Reveal !== 'undefined' && typeof Reveal.initialize === 'function') {
                    Reveal.initialize(settings.revealjsConfig);
                }
            };
            document.head.appendChild(script);
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
                var buSpan = prgrphs_arry[i].querySelectorAll('.numeric-bullet-style');
                if (buSpan.length > 0) {
                    var prevBultTyp = "";
                    var prevBultLvl = "";
                    var buletIndex = 0;
                    var tmpArry = new Array();
                    var tmpArryIndx = 0;
                    var buletTypSrry = new Array();
                    for (var j = 0; j < buSpan.length; j++) {
                        var bult_typ = buSpan[j].getAttribute("data-bulltname");
                        var bult_lvl = buSpan[j].getAttribute("data-bulltlvl");
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
                        var numIdx = PPTXTextUtils.getNumTypeNum(buletTypSrry[tmpArryIndx], buletIndex);
                        buSpan[j].innerHTML = numIdx;
                    }
                }
            }
        }
}

if (typeof window !== 'undefined') {
    window.pptxToHtml = pptxToHtml;
}
