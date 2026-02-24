import { PPTXNodeUtils } from './utils/node.js';
import { PPTXXmlUtils } from './utils/xml.js';
import { PPTXStyleUtils } from './utils/style.js';
import { PPTXTextUtils } from './utils/text.js';
import { PPTXShapeUtils } from './shape/shape.js';
import { processMsgQueue, processSingleMsg } from './utils/chart.js';
import { SLIDE_FACTOR, FONT_SIZE_FACTOR } from './core/constants.js';

/**
 * Parse PPTX to structured data (internal function)
 * @param {JSZip} zip - The JSZip instance
 * @param {Array} msgQueue - Message queue for charts
 * @param {Object} settings - Conversion settings
 * @param {Object} chartId - Chart ID tracker
 * @param {Object} styleTable - Style table
 * @param {*} defaultTextStyle - Default text style
 * @returns {Promise<Object>} Structured PPTX data
 */
async function parsePPTXInternal(zip, msgQueue, settings, chartId, styleTable, defaultTextStyle) {
    const postArray = [];
    const dateBefore = new Date();

    // Extract thumbnail if exists
    const thumbFile = zip.file("docProps/thumbnail.jpeg");
    let thumbnail = null;
    if (thumbFile !== null) {
        const pptxThumbImg = PPTXXmlUtils.base64ArrayBuffer(thumbFile.asArrayBuffer());
        thumbnail = pptxThumbImg;
    }

    // Extract metadata from core.xml
    let metadata = {};
    try {
        const coreFile = zip.file("docProps/core.xml");
        if (coreFile !== null) {
            const coreContent = await PPTXXmlUtils.readXmlFile(zip, "docProps/core.xml");
            if (coreContent !== null) {
                const coreProperties = coreContent["cp:coreProperties"];
                if (coreProperties) {
                    // Extract common metadata fields
                    metadata = {
                        title: coreProperties["dc:title"] || undefined,
                        subject: coreProperties["dc:subject"] || undefined,
                        author: coreProperties["dc:creator"] || undefined,
                        keywords: coreProperties["cp:keywords"] || undefined,
                        description: coreProperties["dc:description"] || undefined,
                        lastModifiedBy: coreProperties["cp:lastModifiedBy"] || undefined,
                        created: coreProperties["dcterms:created"] || undefined,
                        modified: coreProperties["dcterms:modified"] || undefined,
                        category: coreProperties["cp:category"] || undefined,
                        status: coreProperties["cp:contentStatus"] || undefined,
                        contentType: coreProperties["dc:type"] || undefined,
                        language: coreProperties["dc:language"] || undefined
                    };
                }
            }
        }
    } catch (error) {
        console.error("Error parsing metadata:", error);
        // If error, return empty metadata object
        metadata = {};
    }

    const filesInfo = await PPTXXmlUtils.getContentTypes(zip);
    const slideSize = await PPTXXmlUtils.getSlideSizeAndSetDefaultTextStyle(zip, settings);

    const slides = [];
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
            nameParts.pop();
            fileNameNoExt = nameParts.join(".");
        }

        let slideNumber = 1;
        if (fileNameNoExt !== "" && fileNameNoPath.includes("slide")) {
            slideNumber = Number(fileNameNoExt.substring(5));
        }

        // Process slide and get structured data
        const slideData = await processSingleSlideStructured(zip, filename, i, slideSize, msgQueue, settings, chartId, styleTable, defaultTextStyle);
        
        slides.push({
            slideNum: slideNumber,
            fileName: fileNameNoExt,
            data: slideData
        });
    }

    // Sort slides by slideNum to ensure correct order
    slides.sort((a, b) => a.slideNum - b.slideNum);

    const dateAfter = new Date();

    return {
        slides,
        slideSize,
        thumbnail,
        metadata,
        executionTime: dateAfter - dateBefore
    };
}

/**
 * Process a single slide and extract structured data
 * @param {JSZip} zip - The JSZip instance
 * @param {string} slideFileName - Slide file name
 * @param {number} index - Slide index
 * @param {Object} slideSize - Slide size info
 * @param {Array} msgQueue - Message queue
 * @param {Object} settings - Conversion settings
 * @param {Object} chartId - Chart ID tracker
 * @param {Object} styleTable - Style table
 * @param {*} defaultTextStyle - Default text style
 * @returns {Promise<Object>} Structured slide data
 */
async function processSingleSlideStructured(zip, slideFileName, index, slideSize, msgQueue, settings, chartId, styleTable, defaultTextStyle) {
    // Read relationship file of the slide
    const resName = slideFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
    const resContent = await PPTXXmlUtils.readXmlFile(zip, resName);
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
    const slideLayoutContent = await PPTXXmlUtils.readXmlFile(zip, layoutFilename);
    const slideLayoutTables = PPTXNodeUtils.indexNodes(slideLayoutContent);
    const layoutColorOverride = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);

    let slideLayoutClrOvride = {};
    if (layoutColorOverride !== undefined) {
        slideLayoutClrOvride = layoutColorOverride.attrs;
    }

    // Read slide master
    const slideLayoutResFilename = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
    const slideLayoutResContent = await PPTXXmlUtils.readXmlFile(zip, slideLayoutResFilename);
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
    const slideMasterContent = await PPTXXmlUtils.readXmlFile(zip, masterFilename);
    const slideMasterTextStyles = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
    const slideMasterTables = PPTXNodeUtils.indexNodes(slideMasterContent);

    // Read slide master relationships
    const slideMasterResFilename = masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster") + ".rels";
    const slideMasterResContent = await PPTXXmlUtils.readXmlFile(zip, slideMasterResFilename);
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

        themeContent = await PPTXXmlUtils.readXmlFile(zip, themeFilename);
        const themeResContent = await PPTXXmlUtils.readXmlFile(zip, themeResFileName);

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

        diagramContent = await PPTXXmlUtils.readXmlFile(zip, diagramFilename);
        if (diagramContent !== null && diagramContent !== undefined && diagramContent !== "") {
            const diagramJson = JSON.stringify(diagramContent);
            const cleanedJson = diagramJson.replace(/dsp:/g, "p:");
            diagramContent = JSON.parse(cleanedJson);
        }

        const diagramResContent = await PPTXXmlUtils.readXmlFile(zip, diagramResFileName);
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
    const tableStyles = await PPTXXmlUtils.readXmlFile(zip, "ppt/tableStyles.xml");

    // Read slide content
    const slideContent = await PPTXXmlUtils.readXmlFile(zip, slideFileName, true);
    const nodes = slideContent["p:sld"]["p:cSld"]["p:spTree"];

    const processFullTheme = settings.themeProcess;

    // Return structured slide data
    return {
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
        msgQueue,
        bulletCounter: {},
        slideSize,
        index
    };
}

/**
 * Convert structured slide data to HTML
 * @param {Object} slideData - Structured slide data
 * @param {Object} slideSize - Slide size info
 * @param {Object} settings - Conversion settings
 * @param {JSZip} zip - The JSZip instance
 * @returns {Promise<string>} Slide HTML
 */
async function convertSlideDataToHtml(slideData, slideSize, settings, zip) {
    const warpObj = {
        slideLayoutContent: slideData.slideLayoutContent,
        slideLayoutTables: slideData.slideLayoutTables,
        slideMasterContent: slideData.slideMasterContent,
        slideMasterTables: slideData.slideMasterTables,
        slideContent: slideData.slideContent,
        slideResObj: slideData.slideResObj,
        slideMasterTextStyles: slideData.slideMasterTextStyles,
        layoutResObj: slideData.layoutResObj,
        masterResObj: slideData.masterResObj,
        themeContent: slideData.themeContent,
        themeResObj: slideData.themeResObj,
        diagramContent: slideData.diagramContent,
        diagramResObj: slideData.diagramResObj,
        defaultTextStyle: slideData.defaultTextStyle,
        tableStyles: slideData.tableStyles,
        styleTable: slideData.styleTable,
        chartId: slideData.chartId,
        msgQueue: slideData.msgQueue,
        bulletCounter: slideData.bulletCounter,
        zip: zip
    };

    const processFullTheme = settings.themeProcess;
    let bgResult = "";
    if (processFullTheme === true) {
        bgResult = await PPTXNodeUtils.getBackground(warpObj, slideSize, slideData.index, settings, PPTXStyleUtils);
    }

    let bgColor = "";
    if (processFullTheme === "colorsAndImageOnly") {
        bgColor = PPTXStyleUtils.getSlideBackgroundFill(warpObj, slideData.index);
    }

    let result = `<section class='slide' style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;
    result += bgResult;

    const nodes = slideData.slideContent["p:sld"]["p:cSld"]["p:spTree"];
    for (const nodeKey in nodes) {
        if (Array.isArray(nodes[nodeKey])) {
            for (const node of nodes[nodeKey]) {
                result += await PPTXNodeUtils.processNodesInSlide(nodeKey, node, nodes, warpObj, "slide", "group", settings);
            }
        } else {
            result += await PPTXNodeUtils.processNodesInSlide(nodeKey, nodes[nodeKey], nodes, warpObj, "slide", "group", settings);
        }
    }

    return `${result}</div></section>`;
}

/**
 * Generate global CSS
 * @param {Object} styleTable - Style table
 * @returns {string} CSS text
 */
function genGlobalCSS(styleTable) {
    let cssText = "";
    for (const key in styleTable) {
        const suffix = styleTable[key].suffix || "";
        cssText += ` .${styleTable[key].name}${suffix}{${styleTable[key].text}}\n`;
    }
    return cssText;
}

/**
 * PPTX to HTML converter
 * @param {ArrayBuffer} fileData - The PPTX file data
 * @param {Object} options - Conversion options
 * @returns {Promise<Object>} Parsed result
 */
async function pptxToHtml(fileData, options) {
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
    const chartId = { value: 0 };
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
     * @returns {Promise<Object>} Parsed result
     */
    async function convertToHtml(file) {
        if (file.byteLength < 10) {
            console.error("Invalid file: file too small");
            if (callbacks.onError) {
                callbacks.onError({ type: "file_error", message: "Invalid file: file too small" });
            }
            throw new Error("Invalid file: file too small");
        }

        const msgQueue = [];
        const zip = new JSZip().load(file);
        
        // Parse PPTX to structured data
        const parsedData = await parsePPTXInternal(zip, msgQueue, settings, chartId, styleTable, defaultTextStyle);

        // Convert structured data to HTML result
        const result = {
            slides: [],
            slideSize: parsedData.slideSize,
            thumbnail: parsedData.thumbnail,
            styles: {
                global: ""
            },
            metadata: parsedData.metadata,
            charts: []
        };

        // Process slides and convert to HTML
        for (const slideData of parsedData.slides) {
            const slideHtml = await convertSlideDataToHtml(slideData.data, parsedData.slideSize, settings, zip);
            result.slides.push({
                html: slideHtml,
                slideNum: slideData.slideNum,
                fileName: slideData.fileName
            });
            
            if (callbacks.onSlide) {
                callbacks.onSlide(slideHtml, {
                    slideNum: slideData.slideNum,
                    fileName: slideData.fileName
                });
            }
        }

        // Generate global CSS after all slides are processed (styleTable is populated during slide conversion)
        result.styles.global = genGlobalCSS(styleTable);

        // Trigger other callbacks
        if (parsedData.thumbnail && callbacks.onThumbnail) {
            callbacks.onThumbnail(parsedData.thumbnail);
        }
        
        if (parsedData.slideSize && callbacks.onSlideSize) {
            callbacks.onSlideSize(parsedData.slideSize);
        }
        
        if (callbacks.onGlobalCSS) {
            callbacks.onGlobalCSS(result.styles.global);
        }

        // Process message queue for charts
        processMsgQueue(msgQueue, result);
        isDone = true;

        if (callbacks.onComplete) {
            callbacks.onComplete({
                executionTime: parsedData.executionTime,
                slideWidth: parsedData.slideSize?.width || 0,
                slideHeight: parsedData.slideSize?.height || 0,
                styleTable,
                settings
            });
        }

        return result;
    }

    // Process the file data
    if (fileData) {
        return convertToHtml(fileData);
    }
    return null;
}

/**
 * PPTX to JSON converter
 * @param {ArrayBuffer} fileData - The PPTX file data
 * @param {Object} options - Conversion options
 * @returns {Promise<Object>} Parsed result
 */
async function pptxToJson(fileData, options) {
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
    const chartId = { value: 0 };
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
     * Convert PPTX file to JSON
     * @param {ArrayBuffer} file - The PPTX file data
     * @returns {Promise<Object>} Parsed result
     */
    async function convertToJson(file) {
        if (file.byteLength < 10) {
            console.error("Invalid file: file too small");
            if (callbacks.onError) {
                callbacks.onError({ type: "file_error", message: "Invalid file: file too small" });
            }
            throw new Error("Invalid file: file too small");
        }

        const msgQueue = [];
        const zip = new JSZip().load(file);
        
        // Parse PPTX to structured data
        const parsedData = await parsePPTXInternal(zip, msgQueue, settings, chartId, styleTable, defaultTextStyle);

        // Convert structured data to JSON result
        const result = {
            slides: [],
            slideSize: parsedData.slideSize,
            thumbnail: parsedData.thumbnail,
            styles: {
                global: genGlobalCSS(styleTable)
            },
            metadata: parsedData.metadata,
            charts: []
        };

        // Process slides and keep as structured data
        for (const slideData of parsedData.slides) {
            result.slides.push({
                data: slideData.data,
                slideNum: slideData.slideNum,
                fileName: slideData.fileName
            });
            
            if (callbacks.onSlide) {
                callbacks.onSlide(slideData.data, {
                    slideNum: slideData.slideNum,
                    fileName: slideData.fileName
                });
            }
        }

        // Trigger other callbacks
        if (parsedData.thumbnail && callbacks.onThumbnail) {
            callbacks.onThumbnail(parsedData.thumbnail);
        }
        
        if (parsedData.slideSize && callbacks.onSlideSize) {
            callbacks.onSlideSize(parsedData.slideSize);
        }
        
        if (callbacks.onGlobalCSS) {
            callbacks.onGlobalCSS(result.styles.global);
        }

        // Process message queue for charts
        processMsgQueue(msgQueue, result);
        isDone = true;

        if (callbacks.onComplete) {
            callbacks.onComplete({
                executionTime: parsedData.executionTime,
                slideWidth: parsedData.slideSize?.width || 0,
                slideHeight: parsedData.slideSize?.height || 0,
                styleTable,
                settings
            });
        }

        return result;
    }

    // Process the file data
    if (fileData) {
        return convertToJson(fileData);
    }
    return null;
}

// Export functions
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        pptxToHtml,
        pptxToJson
    };
} else if (typeof window !== 'undefined') {
    window.pptxParser = {
        pptxToHtml,
        pptxToJson
    };
}

export default pptxToHtml;
export { pptxToJson, pptxToHtml };
