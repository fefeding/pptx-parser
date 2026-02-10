import { PPTXNodeUtils } from './utils/node.js';
import { PPTXXmlUtils } from './utils/xml.js';
import { PPTXStyleUtils } from './utils/style.js';
import { PPTXTextUtils } from './utils/text.js';
import { PPTXShapeUtils } from './shape/shape.js';
import { SLIDE_FACTOR, FONT_SIZE_FACTOR } from './core/constants.js';

/**
 * PPTX to HTML converter
 * @param {ArrayBuffer} fileData - The PPTX file data
 * @param {Object} options - Conversion options
 * @returns {void}
 */
function pptxToHtml(fileData, options) {
    // Merge default settings with user options
    const settings = {
        mediaProcess: true,
        themeProcess: true,
        incSlide: {
            width: 0,
            height: 0
        },
        styleTable: {},
        ...options
    };

    // Callback functions
    const callbacks = settings.callbacks || {};

    // State variables
    let defaultTextStyle = null;
    let chartId = 0;
    let order = 1;
    let appVersion;
    let slideWidth = 0;
    let slideHeight = 0;
    const processFullTheme = settings.themeProcess;
    const styleTable = settings.styleTable;
    let isDone = false;

    // Trigger file start callback
    if (callbacks.onFileStart) {
        callbacks.onFileStart();
    }

    /**
     * Convert PPTX file to HTML
     * @param {ArrayBuffer} file - The PPTX file data
     */
    function convertToHtml(file) {
        if (file.byteLength < 10) {
            console.error("Invalid file: file too small");
            if (callbacks.onError) {
                callbacks.onError({ type: "file_error", message: "Invalid file: file too small" });
            }
            return;
        }

        const msgQueue = [];
        const zip = new JSZip().load(file);
        const resultArray = processPPTX(zip, msgQueue);

        for (const result of resultArray) {
            switch (result.type) {
                case "slide":
                    if (callbacks.onSlide) {
                        callbacks.onSlide(result.data, {
                            slideNum: result.slideNum,
                            fileName: result.fileName
                        });
                    }
                    break;
                case "pptx-thumb":
                    if (callbacks.onThumbnail) {
                        callbacks.onThumbnail(result.data);
                    }
                    break;
                case "slideSize":
                    slideWidth = result.data.width;
                    slideHeight = result.data.height;
                    if (callbacks.onSlideSize) {
                        callbacks.onSlideSize(result.data);
                    }
                    break;
                case "globalCSS":
                    if (callbacks.onGlobalCSS) {
                        callbacks.onGlobalCSS(result.data);
                    }
                    break;
                case "ExecutionTime":
                    processMsgQueue(msgQueue);
                    isDone = true;

                    if (callbacks.onComplete) {
                        callbacks.onComplete({
                            executionTime: result.data,
                            slideWidth,
                            slideHeight,
                            styleTable,
                            settings
                        });
                    }
                    break;
                case "progress-update":
                    if (callbacks.onProgress) {
                        callbacks.onProgress(result.data);
                    }
                    break;
                default:
                    // Unknown type, ignore
            }
        }
    }

    /**
     * Process PPTX zip file
     * @param {JSZip} zip - The JSZip instance
     * @param {Array} msgQueue - Message queue for charts
     * @returns {Array} Result array
     */
    function processPPTX(zip, msgQueue) {
        const postArray = [];
        const dateBefore = new Date();

        // Extract thumbnail if exists
        const thumbFile = zip.file("docProps/thumbnail.jpeg");
        if (thumbFile !== null) {
            const pptxThumbImg = PPTXXmlUtils.base64ArrayBuffer(thumbFile.asArrayBuffer());
            postArray.push({
                type: "pptx-thumb",
                data: pptxThumbImg,
                slideNum: -1
            });
        }

        const filesInfo = PPTXXmlUtils.getContentTypes(zip);
        const slideSize = PPTXXmlUtils.getSlideSizeAndSetDefaultTextStyle(zip, settings);

        postArray.push({
            type: "slideSize",
            data: slideSize,
            slideNum: 0
        });

        const numOfSlides = filesInfo.slides.length;
        for (let i = 0; i < numOfSlides; i++) {
            const filename = filesInfo.slides[i];
            let fileNameNoPath = "";

            if (filename.includes("/")) {
                const pathParts = filename.split("/");
                fileNameNoPath = pathParts.pop();
            } else {
                fileNameNoPath = filename;
            }

            let fileNameNoExt = "";
            if (fileNameNoPath.includes(".")) {
                const nameParts = fileNameNoPath.split(".");
                nameParts.pop(); // Remove extension
                fileNameNoExt = nameParts.join(".");
            }

            let slideNumber = 1;
            if (fileNameNoExt !== "" && fileNameNoPath.includes("slide")) {
                slideNumber = Number(fileNameNoExt.substring(5));
            }

            const slideHtml = processSingleSlide(zip, filename, i, slideSize, msgQueue);
            postArray.push({
                type: "slide",
                data: slideHtml,
                slideNum: slideNumber,
                fileName: fileNameNoExt
            });

            postArray.push({
                type: "progress-update",
                slideNum: numOfSlides + i + 1,
                data: (i + 1) * 100 / numOfSlides
            });
        }

        postArray.sort((a, b) => a.slideNum - b.slideNum);

        postArray.push({
            type: "globalCSS",
            data: genGlobalCSS()
        });

        const dateAfter = new Date();
        postArray.push({
            type: "ExecutionTime",
            data: dateAfter - dateBefore
        });

        return postArray;
    }

    /**
     * Process a single slide
     * @param {JSZip} zip - The JSZip instance
     * @param {string} slideFileName - Slide filename
     * @param {number} index - Slide index
     * @param {Object} slideSize - Slide size info
     * @param {Array} msgQueue - Message queue
     * @returns {string} Slide HTML
     */
    function processSingleSlide(zip, slideFileName, index, slideSize, msgQueue) {
        // Read relationship file of the slide
        const resName = slideFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
        const resContent = PPTXXmlUtils.readXmlFile(zip, resName);
        const relationshipArray = resContent.Relationships.Relationship;

        let layoutFilename = "";
        let diagramFilename = "";
        const slideResObj = {};

        if (Array.isArray(relationshipArray)) {
            for (const rel of relationshipArray) {
                const relType = rel.attrs.Type;
                const target = rel.attrs.Target.replace("../", "ppt/");

                switch (relType) {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                        layoutFilename = target;
                        break;
                    case "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing":
                        diagramFilename = target;
                        slideResObj[rel.attrs.Id] = {
                            type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target
                        };
                        break;
                    default:
                        slideResObj[rel.attrs.Id] = {
                            type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target
                        };
                }
            }
        } else {
            layoutFilename = relationshipArray.attrs.Target.replace("../", "ppt/");
        }

        // Open slide layout
        const slideLayoutContent = PPTXXmlUtils.readXmlFile(zip, layoutFilename);
        const slideLayoutTables = PPTXNodeUtils.indexNodes(slideLayoutContent);
        const layoutColorOverride = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);

        if (layoutColorOverride !== undefined) {
            slideLayoutClrOvride = layoutColorOverride.attrs;
        }

        // Read slide master
        const slideLayoutResFilename = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
        const slideLayoutResContent = PPTXXmlUtils.readXmlFile(zip, slideLayoutResFilename);
        const layoutRelArray = slideLayoutResContent.Relationships.Relationship;

        let masterFilename = "";
        const layoutResObj = {};

        if (Array.isArray(layoutRelArray)) {
            for (const rel of layoutRelArray) {
                const relType = rel.attrs.Type;
                const target = rel.attrs.Target.replace("../", "ppt/");

                if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster") {
                    masterFilename = target;
                } else {
                    layoutResObj[rel.attrs.Id] = {
                        type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target
                    };
                }
            }
        } else {
            masterFilename = layoutRelArray.attrs.Target.replace("../", "ppt/");
        }

        // Open slide master
        const slideMasterContent = PPTXXmlUtils.readXmlFile(zip, masterFilename);
        const slideMasterTextStyles = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
        const slideMasterTables = PPTXNodeUtils.indexNodes(slideMasterContent);

        // Read slide master relationships
        const slideMasterResFilename = masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster") + ".rels";
        const slideMasterResContent = PPTXXmlUtils.readXmlFile(zip, slideMasterResFilename);
        const masterRelArray = slideMasterResContent.Relationships.Relationship;

        let themeFilename = "";
        const masterResObj = {};

        if (Array.isArray(masterRelArray)) {
            for (const rel of masterRelArray) {
                const relType = rel.attrs.Type;
                const target = rel.attrs.Target.replace("../", "ppt/");

                if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
                    themeFilename = target;
                } else {
                    masterResObj[rel.attrs.Id] = {
                        type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target
                    };
                }
            }
        } else {
            themeFilename = masterRelArray.attrs.Target.replace("../", "ppt/");
        }

        // Load theme file
        let themeContent;
        const themeResObj = {};

        if (themeFilename !== undefined) {
            const themeName = themeFilename.split("/").pop();
            const themeResFileName = themeFilename.replace(themeName, `_rels/${themeName}`) + ".rels";

            themeContent = PPTXXmlUtils.readXmlFile(zip, themeFilename);
            const themeResContent = PPTXXmlUtils.readXmlFile(zip, themeResFileName);

            if (themeResContent !== null) {
                const themeRelArray = themeResContent.Relationships.Relationship;
                if (themeRelArray !== undefined) {
                    if (Array.isArray(themeRelArray)) {
                        for (const rel of themeRelArray) {
                            themeResObj[rel.attrs.Id] = {
                                type: rel.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                                target: rel.attrs.Target.replace("../", "ppt/")
                            };
                        }
                    } else {
                        themeResObj[themeRelArray.attrs.Id] = {
                            type: themeRelArray.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target: themeRelArray.attrs.Target.replace("../", "ppt/")
                        };
                    }
                }
            }
        }

        // Load diagram file
        let diagramContent = {};
        const diagramResObj = {};

        if (diagramFilename !== undefined) {
            const diagramName = diagramFilename.split("/").pop();
            const diagramResFileName = diagramFilename.replace(diagramName, `_rels/${diagramName}`) + ".rels";

            diagramContent = PPTXXmlUtils.readXmlFile(zip, diagramFilename);
            console.log("DEBUG processSingleSlide: diagramFilename:", diagramFilename, "diagramContent type:", typeof diagramContent, "is null:", diagramContent === null);
            if (diagramContent !== null && diagramContent !== undefined && diagramContent !== "") {
                const diagramJson = JSON.stringify(diagramContent);
                console.log("DEBUG processSingleSlide: diagram JSON before replace, length:", diagramJson.length);
                const cleanedJson = diagramJson.replace(/dsp:/g, "p:");
                console.log("DEBUG processSingleSlide: diagram JSON after replace, length:", cleanedJson.length, "diff:", diagramJson.length - cleanedJson.length);
                console.log("DEBUG processSingleSlide: diagram JSON replace, original keys:", Object.keys(diagramContent || {}));
                diagramContent = JSON.parse(cleanedJson);
                console.log("DEBUG processSingleSlide: diagram JSON replace, new keys:", Object.keys(diagramContent || {}));
            } else {
                console.log("DEBUG processSingleSlide: diagramContent is null/undefined/empty, skipping replacement");
            }

            const diagramResContent = PPTXXmlUtils.readXmlFile(zip, diagramResFileName);
            if (diagramResContent !== null) {
                const diagramRelArray = diagramResContent.Relationships.Relationship;
                if (Array.isArray(diagramRelArray)) {
                    for (const rel of diagramRelArray) {
                        diagramResObj[rel.attrs.Id] = {
                            type: rel.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target: rel.attrs.Target.replace("../", "ppt/")
                        };
                    }
                } else {
                    diagramResObj[diagramRelArray.attrs.Id] = {
                        type: diagramRelArray.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target: diagramRelArray.attrs.Target.replace("../", "ppt/")
                    };
                }
            }
        }

        // Load table styles
        const tableStyles = PPTXXmlUtils.readXmlFile(zip, "ppt/tableStyles.xml");

        // Read slide content
        const slideContent = PPTXXmlUtils.readXmlFile(zip, slideFileName, true);
        const nodes = slideContent["p:sld"]["p:cSld"]["p:spTree"];

        const warpObj = {
            zip,
            slideLayoutContent,
            slideLayoutTables,
            slideMasterContent,
            slideMasterTables,
            slideContent,
            slideResObj,
            slideMasterTextStyles,
            layoutResObj,
            masterResObj,
            themeContent,
            themeResObj,
            diagramContent,
            diagramResObj,
            defaultTextStyle: slideSize.defaultTextStyle || defaultTextStyle,
            tableStyles,
            styleTable,
            chartId,
            msgQueue
        };

        let bgResult = "";
        if (processFullTheme === true) {
            bgResult = PPTXNodeUtils.getBackground(warpObj, slideSize, index, settings, PPTXStyleUtils);
        }

        let bgColor = "";
        if (processFullTheme === "colorsAndImageOnly") {
            bgColor = PPTXStyleUtils.getSlideBackgroundFill(warpObj, index);
        }

        let result = `<section class='slide' style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;
        result += bgResult;

        for (const nodeKey in nodes) {
            if (Array.isArray(nodes[nodeKey])) {
                for (const node of nodes[nodeKey]) {
                    result += PPTXNodeUtils.processNodesInSlide(nodeKey, node, nodes, warpObj, "slide", "group", settings);
                }
            } else {
                result += PPTXNodeUtils.processNodesInSlide(nodeKey, nodes[nodeKey], nodes, warpObj, "slide", "group", settings);
            }
        }

        return `${result}</div></section>`;
    }

    /**
     * Generate pie shape path
     * @param {number} height - Shape height
     * @param {number} width - Shape width
     * @param {number} startAngle - Start angle
     * @param {number} endAngle - End angle
     * @param {boolean} isClosed - Whether shape is closed
     * @returns {Array} Path and rotation
     */
    function shapePie(height, width, startAngle, endAngle, isClosed) {
        const pieValue = parseInt(endAngle);
        const piAngle = parseInt(startAngle);
        const size = parseInt(height);
        const radius = size / 2;
        let value = pieValue - piAngle;

        if (value < 0) {
            value = 360 + value;
        }

        value = Math.min(Math.max(value, 0), 360);

        // Calculate x, y coordinates of the point on the circle
        const x = Math.cos((2 * Math.PI) / (360 / value));
        const y = Math.sin((2 * Math.PI) / (360 / value));

        const longArc = value <= 180 ? 0 : 1;
        let path, rotation;

        if (isClosed) {
            path = `M${radius},${radius} L${radius},0 A${radius},${radius} 0 ${longArc},1 ${radius + y * radius},${radius - x * radius} z`;
            rotation = `rotate(${piAngle - 270}, ${radius}, ${radius})`;
        } else {
            const radiusX = width / 2;
            path = `M${radius},0 A${radiusX},${radius} 0 ${longArc},1 ${radiusX + y * radiusX},${radius - x * radius}`;
            rotation = `rotate(${piAngle + 90}, ${radius}, ${radius})`;
        }

        return [path, rotation];
    }

    /**
     * Generate gear shape path
     * @param {number} width - Shape width
     * @param {number} height - Shape height
     * @param {number} points - Number of points
     * @returns {string} Path string
     */
    function shapeGear(width, height, points) {
        const innerRadius = height;
        const outerRadius = 1.5 * innerRadius;
        const centerX = outerRadius;
        const centerY = outerRadius;
        const notches = points;
        const radiusOuter = outerRadius;
        const radiusInner = innerRadius;
        const taperOuter = 50;
        const taperInner = 35;

        // Pre-calculate values for loop
        const pi2 = 2 * Math.PI;
        const angle = pi2 / (notches * 2);
        const taperAngleInner = angle * taperInner * 0.005;
        const taperAngleOuter = angle * taperOuter * 0.005;

        let currentAngle = angle;
        let toggle = false;

        // Move to starting point
        let path = ` M${centerX + radiusOuter * Math.cos(taperAngleOuter)} ${centerY + radiusOuter * Math.sin(taperAngleOuter)}`;

        // Loop
        for (; currentAngle <= pi2 + angle; currentAngle += angle) {
            if (toggle) {
                // Inner to outer line
                path += ` L${centerX + radiusInner * Math.cos(currentAngle - taperAngleInner)},${centerY + radiusInner * Math.sin(currentAngle - taperAngleInner)}`;
                path += ` L${centerX + radiusOuter * Math.cos(currentAngle + taperAngleOuter)},${centerY + radiusOuter * Math.sin(currentAngle + taperAngleOuter)}`;
            } else {
                // Outer to inner line
                path += ` L${centerX + radiusOuter * Math.cos(currentAngle - taperAngleOuter)},${centerY + radiusOuter * Math.sin(currentAngle - taperAngleOuter)}`;
                path += ` L${centerX + radiusInner * Math.cos(currentAngle + taperAngleInner)},${centerY + radiusInner * Math.sin(currentAngle + taperAngleInner)}`;
            }
            toggle = !toggle;
        }

        return `${path} `;
    }

    /**
     * Generate snip/round rectangle shape path
     * @param {number} width - Shape width
     * @param {number} height - Shape height
     * @param {number} adj1 - Adjustment 1
     * @param {number} adj2 - Adjustment 2
     * @param {string} shapeType - Shape type (snip/round)
     * @param {string} adjType - Adjustment type
     * @returns {string} Path string
     */
    function shapeSnipRoundRect(width, height, adj1, adj2, shapeType, adjType) {
        let adjA, adjB, adjC, adjD;

        switch (adjType) {
            case "cornr1":
                adjA = 0;
                adjB = 0;
                adjC = 0;
                adjD = adj1;
                break;
            case "cornr2":
                adjA = adj1;
                adjB = adj2;
                adjC = adj2;
                adjD = adj1;
                break;
            case "cornrAll":
                adjA = adjB = adjC = adjD = adj1;
                break;
            case "diag":
                adjA = adj1;
                adjB = adj2;
                adjC = adj1;
                adjD = adj2;
                break;
        }

        let path;

        if (shapeType === "round") {
            const halfH = height / 2;
            const halfW = width / 2;
            path = `M0,${halfH + (1 - adjB) * halfH} Q0,${height} ${adjB * halfW},${height}`;
            path += ` L${halfW + (1 - adjC) * halfW},${height} Q${width},${height} ${width},${halfH + halfH * (1 - adjC)}`;
            path += `L${width},${halfH * adjD} Q${width},0 ${halfW + halfW * (1 - adjD)},0 L${halfW * adjA},0`;
            path += ` Q0,0 0,${halfH * adjA} z`;
        } else if (shapeType === "snip") {
            const halfH = height / 2;
            const halfW = width / 2;
            path = `M0,${adjA * halfH} L0,${halfH + halfH * (1 - adjB)}L${adjB * halfW},${height}`;
            path += ` L${halfW + halfW * (1 - adjC)},${height} L${width},${halfH + halfH * (1 - adjC)}`;
            path += ` L${width},${adjD * halfH} L${halfW + halfW * (1 - adjD)},0 L${halfW * adjA},0 z`;
        }

        return path;
    }

    /**
     * Generate global CSS
     * @returns {string} CSS text
     */
    function genGlobalCSS() {
        let cssText = "";
        for (const key in styleTable) {
            const suffix = styleTable[key].suffix || "";
            cssText += ` .${styleTable[key].name}${suffix}{${styleTable[key].text}}\n`;
        }
        return cssText;
    }

    /**
     * Process message queue for charts
     * @param {Array} queue - Message queue
     */
    function processMsgQueue(queue) {
        for (const msg of queue) {
            processSingleMsg(msg.data);
        }
    }

    /**
     * Process single chart message
     * @param {Object} data - Chart data
     */
    function processSingleMsg(data) {
        const { chartId, chartType, chartData } = data;
        let chartDataArray = [];
        let chart = null;

        // Validate chart data
        if (!chartData || !Array.isArray(chartData) || chartData.length === 0) {
            console.warn(`Invalid chart data for chart ID: ${chartId}`);
            return;
        }

        switch (chartType) {
            case "lineChart":
                chartDataArray = chartData;
                chart = nv.models.lineChart().useInteractiveGuideline(true);
                if (chartData[0]?.xlabels) {
                    chart.xAxis.tickFormat(d => chartData[0].xlabels[d] || d);
                }
                break;

            case "barChart":
                chartDataArray = chartData;
                chart = nv.models.multiBarChart();
                if (chartData[0]?.xlabels) {
                    chart.xAxis.tickFormat(d => chartData[0].xlabels[d] || d);
                }
                break;

            case "pieChart":
            case "pie3DChart":
                chartDataArray = chartData[0]?.values || [];
                chart = nv.models.pieChart();
                break;

            case "areaChart":
                chartDataArray = chartData;
                chart = nv.models.stackedAreaChart()
                    .clipEdge(true)
                    .useInteractiveGuideline(true);
                if (chartData[0]?.xlabels) {
                    chart.xAxis.tickFormat(d => chartData[0].xlabels[d] || d);
                }
                break;

            case "scatterChart":
                for (let i = 0; i < chartData.length; i++) {
                    const arr = [];
                    if (Array.isArray(chartData[i])) {
                        for (let j = 0; j < chartData[i].length; j++) {
                            arr.push({ x: j, y: chartData[i][j] });
                        }
                    }
                    chartDataArray.push({ key: `data${i + 1}`, values: arr });
                }
                chart = nv.models.scatterChart()
                    .showDistX(true)
                    .showDistY(true)
                    .color(d3.scale.category10().range());
                chart.xAxis.axisLabel('X').tickFormat(d3.format('.02f'));
                chart.yAxis.axisLabel('Y').tickFormat(d3.format('.02f'));
                break;

            default:
                console.warn(`Unknown chart type: ${chartType}`);
        }

        if (chart !== null && callbacks.onChartReady) {
            callbacks.onChartReady({
                chartId,
                chart,
                data: chartDataArray
            });
        }
    }

    // Process the file data
    if (fileData) {
        convertToHtml(fileData);
    }
}

export default pptxToHtml;
