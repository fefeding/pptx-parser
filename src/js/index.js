import { PPTXNodeUtils } from './utils/node.js';
import { PPTXXmlUtils } from './utils/xml.js';
import { PPTXStyleUtils } from './utils/style.js';
import { PPTXTextUtils } from './utils/text.js';
import { PPTXShapeUtils } from './shape/shape.js';
import { SLIDE_FACTOR, FONT_SIZE_FACTOR } from './core/constants.js';

function pptxToHtml(fileData, options) {
    var settings = Object.assign({}, {
        mediaProcess: true,
        themeProcess: true,
        incSlide:{
            width: 0,
            height: 0
        },
        styleTable: {},
    }, options);

    // Callback functions
    var callbacks = settings.callbacks || {};

    var defaultTextStyle = null;
    var chartID = 0;
    var _order = 1;
    var app_version;
    var slideWidth = 0;
    var slideHeight = 0;
    var processFullTheme = settings.themeProcess;
    var styleTable = settings.styleTable;
    var isDone = false;

    if (callbacks.onFileStart) {
        callbacks.onFileStart();
    }



    function convertToHtml(file) {
        if (file.byteLength < 10){
            console.error("Invalid file: file too small");
            if (callbacks.onError) {
                callbacks.onError({ type: "file_error", message: "Invalid file: file too small" });
            }
            return;
        }
        var MsgQueue = new Array();
        var zip = new JSZip();
        zip = zip.load(file);
        var rslt_ary = processPPTX(zip, MsgQueue);

        for (var i = 0; i < rslt_ary.length; i++) {
            switch (rslt_ary[i]["type"]) {
                case "slide":
                    if (callbacks.onSlide) {
                        callbacks.onSlide(rslt_ary[i]["data"], {
                            slide_num: rslt_ary[i]["slide_num"],
                            file_name: rslt_ary[i]["file_name"]
                        });
                    }
                    break;
                case "pptx-thumb":
                    if (callbacks.onThumbnail) {
                        callbacks.onThumbnail(rslt_ary[i]["data"]);
                    }
                    break;
                case "slideSize":
                    slideWidth = rslt_ary[i]["data"].width;
                    slideHeight = rslt_ary[i]["data"].height;
                    if (callbacks.onSlideSize) {
                        callbacks.onSlideSize(rslt_ary[i]["data"]);
                    }
                    break;
                case "globalCSS":
                    if (callbacks.onGlobalCSS) {
                        callbacks.onGlobalCSS(rslt_ary[i]["data"]);
                    }
                    break;
                case "ExecutionTime":
                    processMsgQueue(MsgQueue);
                    isDone = true;

                    if (callbacks.onComplete) {
                        callbacks.onComplete({
                            executionTime: rslt_ary[i]["data"],
                            slideWidth: slideWidth,
                            slideHeight: slideHeight,
                            styleTable: styleTable,
                            settings: settings
                        });
                    }
                    break;
                case "progress-update":
                    if (callbacks.onProgress) {
                        callbacks.onProgress(rslt_ary[i]["data"]);
                    }
                    break;
                default:
            }
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
            const tableStyles = PPTXXmlUtils.readXmlFile(zip, "ppt/tableStyles.xml");

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
                tableStyles,
                styleTable,
                chartID,
                MsgQueue,
            };
            var bgResult = "";
            if (processFullTheme === true) {
                bgResult = PPTXNodeUtils.getBackground(warpObj, slideSize, index, settings, PPTXStyleUtils);
            }

            var bgColor = "";
            if (processFullTheme == "colorsAndImageOnly") {
                bgColor = PPTXStyleUtils.getSlideBackgroundFill(warpObj, index);
            }

            var result = "<section class='slide' style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;" + bgColor + "'>"
            
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
           
            return result + "</div></section>";
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
        for (var key in styleTable) {
            var tagname = "";
            cssText += tagname + " ." + styleTable[key]["name"] +
                ((styleTable[key]["suffix"]) ? styleTable[key]["suffix"] : "") +
                "{" + styleTable[key]["text"] + "}\n";
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

        // Validate chart data
        if (!chartData || !Array.isArray(chartData) || chartData.length === 0) {
            console.warn("Invalid chart data for chart ID: " + chartID);
            return;
        }

        switch (chartType) {
            case "lineChart":
                data = chartData;
                chart = nv.models.lineChart()
                    .useInteractiveGuideline(true);
                if (chartData[0] && chartData[0].xlabels) {
                    chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
                }
                break;
            case "barChart":
                data = chartData;
                chart = nv.models.multiBarChart();
                if (chartData[0] && chartData[0].xlabels) {
                    chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
                }
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
                if (chartData[0] && chartData[0].xlabels) {
                    chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
                }
                break;
            case "scatterChart":
                for (var i = 0; i < chartData.length; i++) {
                    var arr = [];
                    if (Array.isArray(chartData[i])) {
                        for (var j = 0; j < chartData[i].length; j++) {
                            arr.push({ x: j, y: chartData[i][j] });
                        }
                    }
                    data.push({ key: 'data' + (i + 1), values: arr });
                }
                chart = nv.models.scatterChart()
                    .showDistX(true)
                    .showDistY(true)
                    .color(d3.scale.category10().range());
                chart.xAxis.axisLabel('X').tickFormat(d3.format('.02f'));
                chart.yAxis.axisLabel('Y').tickFormat(d3.format('.02f'));
                break;
            default:
                console.warn("Unknown chart type: " + chartType);
                break;
        }

        if (chart !== null && callbacks.onChartReady) {
            callbacks.onChartReady({
                chartID: chartID,
                chart: chart,
                data: data
            });
        }
    }

    // Process the file data
    if (fileData) {
        convertToHtml(fileData);
    }
}

export default pptxToHtml;


