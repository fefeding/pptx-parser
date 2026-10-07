import JSZip from 'jszip';
import { PPTXNodeUtils } from './utils/node';
import { PPTXXmlUtils } from './utils/xml';
import { PPTXStyleUtils } from './utils/style';
import { PPTXTextUtils } from './utils/text';
import { PPTXShapeUtils } from './shape/shape';
import { processMsgQueue, processSingleMsg, extractChartsFromSlide } from './utils/chart';
import { SLIDE_FACTOR, FONT_SIZE_FACTOR } from './core/constants';
import { jsonToPptx, editPptx } from './serializer/json-to-pptx';
import { PPTXComposer } from './serializer/composer';
import { buildStandardDocument, type StandardExtractOptions } from './serializer/json-from-pptx';
import type { PptxDocument } from './types/pptx-document';
import type { XmlNode, ParseSettings, ParseCallbacks } from './core/types';
import baseLayoutCss from './css/pptxjs.css';

/** 图表消息队列条目（由解析阶段产生，交给 processMsgQueue 消费） */
interface ChartQueueItem {
    type: string;
    data: Record<string, unknown>;
}

/** styleTable 条目 */
interface StyleTableEntry {
    name: string;
    suffix?: string;
    text: string;
}
/** 样式表：CSS 文本 -> 条目 */
export type StyleTable = Record<string, StyleTableEntry>;

/** 资源对应关系：rId -> { target, type } */
type ResourceMap = Record<string, Record<string, string>>;

/** 幻灯片元数据（docProps/core.xml） */
interface PptxMetadata {
    title?: string;
    subject?: string;
    author?: string;
    keywords?: string;
    description?: string;
    lastModifiedBy?: string;
    created?: string;
    modified?: string;
    category?: string;
    status?: string;
    contentType?: string;
    language?: string;
}

/** 单页批注（来自 ppt/comments/commentsN.xml） */
interface CommentItem {
    author: string;
    text: string;
    dt?: string;
    x?: number;
    y?: number;
}

/** 解析 ppt/comments/commentsN.xml 为批注数组（authorId 映射为作者名） */
function parseComments(xml: any, authors: Record<string, string>): CommentItem[] {
    if (!xml || !xml['p:cmLst']) return [];
    const cms = xml['p:cmLst']['p:cm'];
    const arr = Array.isArray(cms) ? cms : (cms ? [cms] : []);
    return arr.map((cm: any) => {
        const attrs = cm.attrs || {};
        const pos = cm['p:pos'];
        const text = cm['p:text'];
        return {
            author: authors[String(attrs.authorId)] || 'Author',
            text: typeof text === 'string' ? text : (text && text['#text']) || '',
            dt: attrs.dt,
            x: pos && pos.attrs && pos.attrs.x ? Number(pos.attrs.x) : undefined,
            y: pos && pos.attrs && pos.attrs.y ? Number(pos.attrs.y) : undefined
        };
    });
}

/** 单页解析数据（承载 processNodesInSlide 所需的 warp 上下文字段） */
interface SlideDataRecord {
    index: number;
    comments?: CommentItem[];
    slideContent: XmlNode;
    slideLayoutContent?: XmlNode;
    slideMasterContent?: XmlNode;
    themeContent?: XmlNode;
    diagramContent?: XmlNode | string | null;
    slideLayoutTables?: XmlNode;
    slideMasterTables?: XmlNode;
    slideMasterTextStyles?: unknown;
    tableStyles?: XmlNode;
    slideResObj?: ResourceMap;
    layoutResObj?: ResourceMap;
    masterResObj?: ResourceMap;
    themeResObj?: ResourceMap;
    diagramResObj?: ResourceMap;
    styleTable?: StyleTable;
    chartId?: { value: number };
    msgQueue?: ChartQueueItem[];
    bulletCounter?: unknown;
    defaultTextStyle?: XmlNode | null;
    [key: string]: unknown;
}

/** pptxToHtml 单页输出 */
interface HtmlSlideResult {
    html: string;
    data: SlideDataRecord | undefined;
    slideNum: number | undefined;
    fileName: string | undefined;
    /** 隐藏幻灯片（p:sld show="0"） */
    hidden: boolean;
}

/** pptxToJson 单页输出（保留结构化数据，不含 html） */
interface JsonSlideResult {
    data: SlideDataRecord | undefined;
    slideNum: number | undefined;
    fileName: string | undefined;
}

/** zip 文件索引条目 */
interface FileIndexEntry {
    name: string;
    dir: boolean;
    size: number;
}

/** 可被 JSZip.loadAsync / 旧版 load 接受的 PPTX 文件数据 */
type PptxFileData = ArrayBuffer | Uint8Array | string;

/** 幻灯片尺寸（含解析端回传的默认文本样式） */
export interface SlideSize {
    width: number;
    height: number;
    /** getSlideSizeAndSetDefaultTextStyle 一并返回，可能为 undefined */
    defaultTextStyle?: XmlNode;
}

/**
 * Parse PPTX file to structured JSON data (internal function)
 * @param {ArrayBuffer} file - The PPTX file data
 * @param {Object} settings - Conversion settings
 * @param {Object} callbacks - Callback functions
 * @param {Object} chartId - Chart ID tracker
 * @param {Object} styleTable - Style table
 * @param {*} defaultTextStyle - Default text style
 * @returns {Promise<Object>} Parsed result with structured data
 */
async function processToJson(file: PptxFileData, settings: ParseSettings, callbacks: ParseCallbacks, chartId: { value: number }, styleTable: StyleTable, defaultTextStyle: XmlNode | null) {
    if ((typeof file === 'string' ? file.length : file.byteLength) < 10) {
        if (callbacks.onError) {
            callbacks.onError({ type: "file_error", message: "Invalid file: file too small" });
        }
        throw new Error("Invalid file: file too small");
    }

    const msgQueue: ChartQueueItem[] = [];
    const zip: JSZip = JSZip.loadAsync ? await JSZip.loadAsync(file) : (new JSZip() as unknown as { load: (f: PptxFileData) => JSZip }).load(file);

    // Parse PPTX to structured data
    const parsedData = await parsePPTXInternal(zip, msgQueue, settings, chartId, styleTable, defaultTextStyle);

    // Return structured data (styleTable is populated during parsing)
    return {
        parsedData,
        msgQueue,
        zip,
        slideSize: parsedData.slideSize,
        thumbnail: parsedData.thumbnail,
        metadata: parsedData.metadata,
        executionTime: parsedData.executionTime
    };
}

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
async function parsePPTXInternal(zip: JSZip, msgQueue: ChartQueueItem[], settings: ParseSettings, chartId: { value: number }, styleTable: StyleTable, defaultTextStyle: XmlNode | null) {
    const postArray = [];
    const dateBefore = new Date();

    // Extract thumbnail if exists
    const thumbFile = zip.file("docProps/thumbnail.jpeg");
    let thumbnail = null;
    if (thumbFile !== null) {
        const pptxThumbImg = PPTXXmlUtils.base64ArrayBuffer(await thumbFile.async("arraybuffer"));
        thumbnail = pptxThumbImg;
    }

    // Extract metadata from core.xml
    let metadata: PptxMetadata = {};
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
        // If error, return empty metadata object
        metadata = {};
    }

    // 解析文档自定义属性（docProps/custom.xml）
    let customProps: Record<string, string> = {};
    try {
        const customContent = await PPTXXmlUtils.readXmlFile(zip, "docProps/custom.xml");
        if (customContent && customContent["Properties"]) {
            const props = customContent["Properties"]["property"];
            const arr = Array.isArray(props) ? props : (props ? [props] : []);
            for (const p of arr) {
                const name = p && p.attrs && p.attrs.name;
                if (!name) continue;
                const valKey = p ? Object.keys(p).find(k => k !== 'attrs' && k !== 'order') : undefined;
                const raw = valKey ? p[valKey] : undefined;
                customProps[name] = (typeof raw === 'string' ? raw : (raw && raw['#text'])) || '';
            }
        }
    } catch {
        customProps = {};
    }

    // 解析批注作者表（ppt/commentAuthors.xml）：id -> name
    const commentAuthors: Record<string, string> = {};
    try {
        const authorsContent = await PPTXXmlUtils.readXmlFile(zip, "ppt/commentAuthors.xml");
        if (authorsContent && authorsContent["p:commentAuthors"]) {
            const authors = authorsContent["p:commentAuthors"]["p:cmAuthor"];
            const arr = Array.isArray(authors) ? authors : (authors ? [authors] : []);
            for (const a of arr) {
                const id = a && a.attrs && a.attrs.id;
                const name = a && a.attrs && a.attrs.name;
                if (id !== undefined) commentAuthors[String(id)] = name || 'Author';
            }
        }
    } catch {
        // ignore
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

        // 解析本页批注（通过 slide rels 中的 comments 关系定位 commentsN.xml）
        try {
            const resObj: Record<string, any> = (slideData as any).slideResObj || {};
            const commentEntry = Object.values(resObj).find((e: any) => e && e.type === 'comments');
            if (commentEntry) {
                const cXml = await PPTXXmlUtils.readXmlFile(zip, commentEntry.target);
                (slideData as any).comments = parseComments(cXml, commentAuthors);
            }
        } catch {
            // ignore
        }

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
        customProps,
        executionTime: dateAfter.getTime() - dateBefore.getTime()
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

/**
 * 关系 Target → 包内绝对路径（ppt/...）。
 *
 * OOXML 的相对 Target 是相对于「源部件所在目录」解析的（rels 文件位于 dir/_rels/，
 * 源文件位于 dir/）。旧逻辑 `target.replace("../","ppt/")` 只能处理 "../media/x.png"
 * 这种写法，无法处理部分生成器（如 WPS）把图片放在 ppt/slides/media/、rels Target
 * 写成 "media/image1.png"（无 ../ 前缀）的情况，导致这些图片被定位到不存在的
 * ppt/media/ 而全部丢失。这里基于 rels 文件路径推导源目录，完整处理相对路径上溯。
 */
function resolveRelTargetFromRels(relsPath: string, target: string): string {
    if (!target) return target;
    if (/^https?:/i.test(target)) return target;
    if (target.startsWith('/')) return 'ppt' + target;
    if (target.startsWith('ppt/')) return target;
    // ppt/slides/_rels/slide1.xml.rels → ppt/slides/（源部件所在目录）
    const baseDir = relsPath.replace(/_rels\/[^/]*\.rels$/, '').replace(/[^/]+$/, '');
    const stack: string[] = [];
    for (const seg of (baseDir + target).split('/')) {
        if (seg === '' || seg === '.') continue;
        if (seg === '..') stack.pop();
        else stack.push(seg);
    }
    return stack.join('/');
}

async function processSingleSlideStructured(zip: JSZip, slideFileName: string, index: number, slideSize: SlideSize, msgQueue: ChartQueueItem[], settings: ParseSettings, chartId: { value: number }, styleTable: StyleTable, defaultTextStyle: XmlNode | null) {
    // Read relationship file of the slide
    const resName = `${slideFileName.replace("slides/slide", "slides/_rels/slide")}.rels`;
    const resContent = await PPTXXmlUtils.readXmlFile(zip, resName);
    const relationshipArray = resContent.Relationships.Relationship;

    let layoutFilename = "";
    let diagramFilename = "";
    let notesFilename = ""; // 添加备注文件名
    const slideResObj: ResourceMap = {};

    if (Array.isArray(relationshipArray)) {
        for (const rel of relationshipArray) {
            const relType = rel.attrs.Type;
            const target = resolveRelTargetFromRels(resName, rel.attrs.Target);

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
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
                    // 处理备注关系
                    notesFilename = target;
                    break;
                default:
                    slideResObj[rel.attrs.Id] = {
                        type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target
                    };
            }
        }
    } else {
        const relType = relationshipArray.attrs.Type;
        const target = resolveRelTargetFromRels(resName, relationshipArray.attrs.Target);

        if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout") {
            layoutFilename = target;
        } else if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide") {
            notesFilename = target;
        } else if (relType === "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing") {
            diagramFilename = target;
            slideResObj[relationshipArray.attrs.Id] = {
                type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                target: target
            };
        } else {
            // 默认处理为slideLayout
            layoutFilename = target;
        }
    }

    // Open slide layout
    const slideLayoutContent = await PPTXXmlUtils.readXmlFile(zip, layoutFilename);
    const slideLayoutTables = PPTXNodeUtils.indexNodes(slideLayoutContent);
    const layoutColorOverride = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);

    let slideLayoutClrOvride: Record<string, unknown> = {};
    if (layoutColorOverride !== undefined) {
        slideLayoutClrOvride = layoutColorOverride.attrs;
    }

    // Read slide master
    const slideLayoutResFilename = `${layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout")}.rels`;
    const slideLayoutResContent = await PPTXXmlUtils.readXmlFile(zip, slideLayoutResFilename);
    const layoutRelArray = slideLayoutResContent.Relationships.Relationship;

    let masterFilename = "";
    const layoutResObj: ResourceMap = {};

    if (Array.isArray(layoutRelArray)) {
        for (const rel of layoutRelArray) {
            const relType = rel.attrs.Type;
            const target = resolveRelTargetFromRels(slideLayoutResFilename, rel.attrs.Target);

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
        masterFilename = resolveRelTargetFromRels(slideLayoutResFilename, layoutRelArray.attrs.Target);
    }

    // Open slide master
    const slideMasterContent = await PPTXXmlUtils.readXmlFile(zip, masterFilename);
    const slideMasterTextStyles = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
    const slideMasterTables = PPTXNodeUtils.indexNodes(slideMasterContent);

    // Read slide master relationships
    const slideMasterResFilename = `${masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster")}.rels`;
    const slideMasterResContent = await PPTXXmlUtils.readXmlFile(zip, slideMasterResFilename);
    const masterRelArray = slideMasterResContent.Relationships.Relationship;

    let themeFilename = "";
    const masterResObj: ResourceMap = {};

    if (Array.isArray(masterRelArray)) {
        for (const rel of masterRelArray) {
            const relType = rel.attrs.Type;
            const target = resolveRelTargetFromRels(slideMasterResFilename, rel.attrs.Target);

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
        themeFilename = resolveRelTargetFromRels(slideMasterResFilename, masterRelArray.attrs.Target);
    }

    // Load theme file
    let themeContent;
    const themeResObj: ResourceMap = {};

    if (themeFilename !== undefined) {
        const themeName = themeFilename.split("/").pop();
        const themeResFileName = `${themeFilename.replace(String(themeName), `_rels/${themeName}`)}.rels`;

        themeContent = await PPTXXmlUtils.readXmlFile(zip, themeFilename);
        const themeResContent = await PPTXXmlUtils.readXmlFile(zip, themeResFileName);

        if (themeResContent !== null) {
            const themeRelArray = themeResContent.Relationships.Relationship;
            if (themeRelArray !== undefined) {
                if (Array.isArray(themeRelArray)) {
                    for (const rel of themeRelArray) {
                        themeResObj[rel.attrs.Id] = {
                            type: rel.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target: resolveRelTargetFromRels(themeResFileName, rel.attrs.Target)
                        };
                    }
                } else {
                    themeResObj[themeRelArray.attrs.Id] = {
                        type: themeRelArray.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target: resolveRelTargetFromRels(themeResFileName, themeRelArray.attrs.Target)
                    };
                }
            }
        }
    }

    // Load diagram file
    let diagramContent: XmlNode | string | null = {};
    const diagramResObj: ResourceMap = {};

    if (diagramFilename !== undefined) {
        const diagramName = diagramFilename.split("/").pop();
        const diagramResFileName = `${diagramFilename.replace(String(diagramName), `_rels/${diagramName}`)}.rels`;

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
                        target: resolveRelTargetFromRels(diagramResFileName, rel.attrs.Target)
                    };
                }
            } else {
                diagramResObj[diagramRelArray.attrs.Id] = {
                    type: diagramRelArray.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                    target: resolveRelTargetFromRels(diagramResFileName, diagramRelArray.attrs.Target)
                };
            }
        }
    }

    // Load table styles
    const tableStyles = await PPTXXmlUtils.readXmlFile(zip, "ppt/tableStyles.xml");

    // Read slide content
    const slideContent = await PPTXXmlUtils.readXmlFile(zip, slideFileName, true);
    
    // Read notes content if available
    let notesContent = null;
    if (notesFilename) {
        notesContent = await PPTXXmlUtils.readXmlFile(zip, notesFilename);
    }

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
        notesContent, // 添加备注内容
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
async function convertSlideDataToHtml(slideData: SlideDataRecord, slideSize: SlideSize, settings: ParseSettings, zip: JSZip, slideNum: number | undefined, docCustomProps?: Record<string, string>) {
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

    const processFullTheme: boolean | string | undefined = settings.themeProcess;
    let bgResult = "";
    if (processFullTheme === true) {
        bgResult = await PPTXNodeUtils.getBackground(warpObj, slideSize, slideData.index, settings);
    }

    let bgColor: string | undefined = "";
    if (processFullTheme === "colorsAndImageOnly") {
        bgColor = await PPTXStyleUtils.getSlideBackgroundFill(warpObj, slideData.index);
    }
    
    // 检测幻灯片过渡效果
    let transitionClass = "";
    const transitionData = extractSlideTransition(slideData.slideContent);
    if (transitionData) {
        transitionClass = ` data-transition='${JSON.stringify(transitionData)}'`;
    }

    // 隐藏幻灯片（p:sld show="0"）
    let hiddenClass = "";
    const sldAttrs = (slideData.slideContent as any)["p:sld"] && (slideData.slideContent as any)["p:sld"].attrs;
    if (sldAttrs && String(sldAttrs.show) === "0") {
        hiddenClass = ` data-hidden='true'`;
    }

    // 计时：自动播放 + 元素动画（p:timing）
    let timingClass = "";
    const timingData = extractSlideTiming(slideData.slideContent);
    if (timingData) {
        const payload: Record<string, unknown> = {};
        if (timingData.advanceTime != null) payload.autoplay = timingData.advanceTime;
        if (timingData.animations.length) payload.animations = timingData.animations;
        timingClass = ` data-timing='${JSON.stringify(payload)}'`;
    }

    const slideIdAttr = slideNum ? ` id="slide-${slideNum}"` : "";
    let result = `<section class='slide'${slideIdAttr}${transitionClass}${hiddenClass}${timingClass} style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;
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

    const esc = (s: string) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

    // ===== 批注可视化（来自 ppt/comments/commentsN.xml）=====
    const comments = (slideData as any).comments as CommentItem[] | undefined;
    if (comments && comments.length) {
        const emuToPx = (e?: number) => (e == null ? 0 : e / 9525);
        const noteHtml = comments.map((c, i) => {
            const left = c.x != null ? emuToPx(c.x) : (slideSize.width - 220);
            const top = c.y != null ? emuToPx(c.y) : (8 + i * 76);
            return `<div class="pptx-comment" style="position:absolute;left:${left}px;top:${top}px;width:200px;box-sizing:border-box;background:#fff7cc;border:1px solid #f1d27a;border-radius:8px;padding:8px 10px;box-shadow:0 2px 6px rgba(0,0,0,.15);font:12px/1.4 sans-serif;color:#3b2f00;z-index:60;">` +
                `<div style="font-weight:700;margin-bottom:2px;">${esc(c.author)}</div>` +
                `<div style="white-space:pre-wrap;">${esc(c.text)}</div>` +
                (c.dt ? `<div style="margin-top:4px;font-size:10px;color:#8a7a3a;">${esc(String(c.dt).slice(0, 10))}</div>` : '') +
                `</div>`;
        }).join('');
        result += `<div class="pptx-comments">${noteHtml}</div>`;
    }

    // ===== 文档自定义属性可视化（docProps/custom.xml）=====
    // 自定义属性是文档级数据，仅在首页渲染一次，避免每页重复
    if (docCustomProps && slideNum === 1 && Object.keys(docCustomProps).length) {
        const rows = Object.entries(docCustomProps)
            .map(([k, v]) => `<div class="pptx-custom-prop"><span class="pptx-custom-prop-name">${esc(k)}</span><span class="pptx-custom-prop-value">${esc(v)}</span></div>`)
            .join('');
        result += `<div class="pptx-custom-props" style="position:absolute;left:0;bottom:0;max-width:60%;box-sizing:border-box;background:rgba(255,255,255,.92);border:1px solid #cbd5e1;border-radius:6px;padding:6px 8px;font:11px/1.5 sans-serif;color:#334155;z-index:55;">${rows}</div>`;
    }

    return `${result}</div></section>`;
}

/**
 * Generate global CSS
 * @param {Object} styleTable - Style table
 * @returns {string} CSS text
 */
function genGlobalCSS(styleTable: StyleTable) {
    // 基础布局骨架（.block/.content 等）随 global 一起返回，消费方无需再单独引入 pptxjs.css
    let cssText = baseLayoutCss + "\n";
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
async function pptxToHtml(fileData: PptxFileData, options: Partial<ParseSettings>) {
    // Merge default settings with user options
    const settings: ParseSettings = {
        mediaProcess: true,
        mediaMuted: false,
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
    let defaultTextStyle: any = null;
    const chartId = { value: 0 };
    const styleTable = settings.styleTable as StyleTable;
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
    async function convertToHtml(file: PptxFileData) {
        // Step 1: Parse PPTX to structured JSON data
        const { parsedData, msgQueue, zip, slideSize, thumbnail, metadata, executionTime } = 
            await processToJson(file, settings, callbacks, chartId, styleTable, defaultTextStyle);

        // Step 2: Convert structured data to HTML result
        const result = {
            slides: [] as HtmlSlideResult[],
            slideSize,
            thumbnail,
            styles: {
                global: ""
            },
            metadata,
            customProps: parsedData.customProps,
            charts: [] as Array<Record<string, unknown>>
        };

        // Step 3: Process slides and convert to HTML
        for (const slideData of parsedData.slides) {
            const slideHtml = await convertSlideDataToHtml(slideData.data, slideSize, settings, zip, slideData.slideNum, parsedData.customProps);
            const sldAttrs = (slideData.data as any).slideContent && (slideData.data as any).slideContent["p:sld"] && (slideData.data as any).slideContent["p:sld"].attrs;
            const hidden = !!(sldAttrs && String(sldAttrs.show) === "0");
            result.slides.push({
                html: slideHtml,
                data: slideData.data,  // Keep structured data for potential reuse
                slideNum: slideData.slideNum,
                fileName: slideData.fileName,
                hidden
            });

            if (callbacks.onSlide) {
                callbacks.onSlide(slideHtml, {
                    slideNum: slideData.slideNum,
                    fileName: slideData.fileName
                });
            }
        }

        // Step 4: Generate global CSS after all slides are processed
        result.styles.global = genGlobalCSS(styleTable);

        // Step 5: Trigger other callbacks
        if (thumbnail && callbacks.onThumbnail) {
            callbacks.onThumbnail(thumbnail);
        }

        if (slideSize && callbacks.onSlideSize) {
            callbacks.onSlideSize(slideSize);
        }

        if (callbacks.onGlobalCSS) {
            callbacks.onGlobalCSS(result.styles.global);
        }

        // Step 6: Process message queue for charts
        processMsgQueue(msgQueue, result);
        isDone = true;

        if (callbacks.onComplete) {
            callbacks.onComplete({
                executionTime,
                slideWidth: slideSize?.width || 0,
                slideHeight: slideSize?.height || 0,
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
async function pptxToJson(fileData: PptxFileData, options: Partial<ParseSettings> & { mode?: 'raw' | 'semantic' } & StandardExtractOptions = {}) {
    // Merge default settings with user options
    const settings: ParseSettings = {
        mediaProcess: true,
        mediaMuted: false,
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
    let defaultTextStyle: any = null;
    const chartId = { value: 0 };
    const styleTable = settings.styleTable as StyleTable;
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
    async function convertToJson(file: PptxFileData) {
        // Step 1: Parse PPTX to structured JSON data
        const { parsedData, msgQueue, zip, slideSize, thumbnail, metadata, executionTime } = 
            await processToJson(file, settings, callbacks, chartId, styleTable, defaultTextStyle);

        // Step 2: Convert structured data to JSON result
        const result = {
            slides: [] as JsonSlideResult[],
            slideSize,
            thumbnail,
            styles: {
                global: genGlobalCSS(styleTable)
            },
            metadata,
            customProps: parsedData.customProps,
            charts: [] as Array<Record<string, unknown>>
        };

        // Step 3: Process slides and keep as structured data
        for (const slideData of parsedData.slides) {
            // 本链路不做 HTML 转换，需独立扫描一次图表部件，
            // 否则 msgQueue 始终为空、结果 charts 恒为 []
            await extractChartsFromSlide(slideData.data, zip);

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

        // Step 4: Trigger other callbacks
        if (thumbnail && callbacks.onThumbnail) {
            callbacks.onThumbnail(thumbnail);
        }

        if (slideSize && callbacks.onSlideSize) {
            callbacks.onSlideSize(slideSize);
        }

        if (callbacks.onGlobalCSS) {
            callbacks.onGlobalCSS(result.styles.global);
        }

        // Step 5: Process message queue for charts
        processMsgQueue(msgQueue, result);
        isDone = true;

        // semantic 模式：附加标准 PptxDocument（与 jsonToPptx 同源）
        if (options.mode === 'semantic') {
            (result as any).document = await buildStandardDocument(parsedData, zip, { rawDeps: options.rawDeps || 'auto' });
        }

        if (callbacks.onComplete) {
            callbacks.onComplete({
                executionTime,
                slideWidth: slideSize?.width || 0,
                slideHeight: slideSize?.height || 0,
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

/**
 * PPTX → 标准 JSON（统一契约 PptxDocument）
 *
 * 直接产出 types/pptx-document.ts 定义的 PptxDocument，与 jsonToPptx 输入同源，
 * 可立即交回 jsonToPptx 还原，实现 JSON 级 round-trip。等价于 pptxToJson(..., {mode:'semantic'}).document。
 * @param {ArrayBuffer} fileData - The PPTX file data
 * @param {Object} options - 同 pptxToJson 的解析选项
 * @returns {Promise<PptxDocument>} 标准 PPTX JSON
 */
async function pptxToStandard(fileData: PptxFileData, options: Partial<ParseSettings> & StandardExtractOptions = {}) {
    // Merge default settings with user options
    const settings: ParseSettings = {
        mediaProcess: true,
        mediaMuted: false,
        themeProcess: true,
        incSlide: {
            width: 0,
            height: 0
        },
        styleTable: {},
        ...options
    };

    const callbacks = settings.callbacks || {};

    let defaultTextStyle: any = null;
    const chartId = { value: 0 };
    const styleTable = settings.styleTable as StyleTable;

    if (callbacks.onFileStart) {
        callbacks.onFileStart();
    }

    const { parsedData, zip } = await processToJson(fileData, settings, callbacks, chartId, styleTable, defaultTextStyle);

    const doc = await buildStandardDocument(parsedData, zip, { rawDeps: options.rawDeps || 'auto' });

    if (callbacks.onComplete) {
        callbacks.onComplete({
            executionTime: parsedData.executionTime,
            slideWidth: parsedData.slideSize?.width || 0,
            slideHeight: parsedData.slideSize?.height || 0,
            styleTable,
            settings
        });
    }

    return doc;
}

/**
 * PPTX to File Index and Content converter
 * @param {ArrayBuffer} fileData - The PPTX file data
 * @returns {Promise<Object>} File index and content result
 */
async function pptxToFiles(fileData: PptxFileData) {
    if ((typeof fileData === 'string' ? fileData.length : fileData.byteLength) < 10) {
        throw new Error("Invalid file: file too small");
    }

    const zip: JSZip = JSZip.loadAsync ? await JSZip.loadAsync(fileData) : (new JSZip() as unknown as { load: (f: PptxFileData) => JSZip }).load(fileData);

    const result: { files: FileIndexEntry[]; content: Record<string, unknown> } = {
        files: [],
        content: {}
    };

    // Iterate through all files in the zip
    const promises: Promise<void>[] = [];
    zip.forEach((relativePath: string, zipEntry: JSZip.JSZipObject) => {
        result.files.push({
            name: relativePath,
            dir: zipEntry.dir,
            size: (zipEntry as unknown as { _data: { uncompressedSize: number } })._data.uncompressedSize
        });

        // Read file content based on type
        const promise = (async () => {
            try {
                if (zipEntry.dir) {
                    return;
                }

                const ext = (relativePath.split('.').pop() ?? '').toLowerCase();

                // For XML files, read as text
                if (ext === 'xml' || ext === 'rels') {
                    const content = await zipEntry.async('text');
                    result.content[relativePath] = {
                        type: 'text',
                        content: content
                    };
                }
                // For image files, read as base64
                else if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg'].includes(ext)) {
                    const base64 = await zipEntry.async('base64');
                    result.content[relativePath] = {
                        type: 'image',
                        format: ext,
                        base64: base64,
                        dataUrl: `data:image/${ext === 'jpg' ? 'jpeg' : ext};base64,${base64}`
                    };
                }
                // For other binary files, read as base64
                else {
                    const base64 = await zipEntry.async('base64');
                    result.content[relativePath] = {
                        type: 'binary',
                        base64: base64
                    };
                }
            } catch (error: unknown) {
                result.content[relativePath] = {
                    type: 'error',
                    error: error instanceof Error ? error.message : String(error)
                };
            }
        })();

        promises.push(promise);
    });

    await Promise.all(promises);

    return result;
}

// Export functions
/**
 * 提取幻灯片过渡效果数据
 * @param {Object} slideContent - 幻灯片内容
 * @returns {Object|null} 过渡效果数据或null
 */
function extractSlideTransition(slideContent: XmlNode) {
    // 检查slide中是否有transition元素
    const sld = slideContent["p:sld"];
    if (!sld) return null;
    
    const transition = PPTXXmlUtils.getTextByPathList(sld, ["p:transition"]);
    if (!transition) return null;
    
    // 解析过渡效果类型
    let transitionType = "fade"; // 默认淡入淡出
    let duration = 1000; // 默认1秒
    
    // 检查具体的过渡类型
    const transitionTypes = [
        "p:blinds", "p:checker", "p:circle", "p:comb", 
        "p:cover", "p:dissolve", "p:fade", "p:push",
        "p:random", "p:split", "p:strips", "p:wipe"
    ];
    
    for (const type of transitionTypes) {
        if (transition[type]) {
            transitionType = type.replace("p:", "");
            break;
        }
    }
    
    // 获取持续时间
    if (transition.attrs && transition.attrs["spd"]) {
        // spd值: slow=3, med=2, fast=1
        const speedMap: Record<string, number> = { "1": 500, "2": 1000, "3": 2000 };
        duration = speedMap[transition.attrs["spd"]] || 1000;
    }
    
    return {
        type: transitionType,
        duration: duration
    };
}

/**
 * 提取幻灯片计时（自动播放 + 元素动画），解析自 p:timing
 * @param {Object} slideContent - 幻灯片内容
 * @returns {Object|null} { advanceTime, animations:[{spid,type,duration}] } 或 null
 */
function extractSlideTiming(slideContent: XmlNode) {
    const sld: any = (slideContent as any)["p:sld"];
    const timing = sld && sld["p:timing"];
    if (!timing) return null;

    // 动画条目除必需字段外，还透传 presetClass/presetId/delay/repeat 等可选属性
    const animations: Array<{ spid: number; type: string; duration: number } & Record<string, unknown>> = [];
    let advanceTime: number | undefined;

    // 递归收集所有 p:cTn
    const allCtn: any[] = [];
    (function walk(n: any) {
        if (!n || typeof n !== "object") return;
        for (const k of Object.keys(n)) {
            if (k === "p:cTn") {
                const arr = Array.isArray(n[k]) ? n[k] : [n[k]];
                for (const c of arr) { allCtn.push(c); walk(c); }
            } else {
                const v = n[k];
                if (Array.isArray(v)) v.forEach(walk);
                else if (v && typeof v === "object") walk(v);
            }
        }
    })(timing);

    for (const c of allCtn) {
        // 自动播放：afterTime cond
        const stCondLst = c["p:stCondLst"];
        if (stCondLst) {
            const conds = Array.isArray(stCondLst["p:cond"]) ? stCondLst["p:cond"] : [stCondLst["p:cond"]];
            for (const cd of conds) {
                if (cd && cd.attrs && cd.attrs.type === "afterTime" && cd.attrs.val != null) {
                    advanceTime = Number(cd.attrs.val);
                }
            }
        }
        // 元素动画：tgtEl > spTgt spid + childTnLst > cTn preset/dur
        const tgtEl = c["p:tgtEl"];
        const spTgt = tgtEl && tgtEl["p:spTgt"];
        const spid = spTgt && spTgt.attrs && spTgt.attrs.spid;
        if (spid != null) {
            let preset = "fade";
            let dur: number | undefined;
            let presetClass: string | undefined;
            let presetId: number | undefined;
            let delay: number | undefined;
            let repeat: number | 'indefinite' | undefined;
            const childTnLst = c["p:childTnLst"];
            if (childTnLst) {
                const inner = Array.isArray(childTnLst["p:cTn"]) ? childTnLst["p:cTn"] : [childTnLst["p:cTn"]];
                for (const ic of inner) {
                    if (ic && ic.attrs) {
                        // preset 透传真实名称（此前仅保留 4 种，其余一律退化为 fade）
                        if (ic.attrs.preset) preset = String(ic.attrs.preset);
                        if (ic.attrs.dur != null) dur = Number(ic.attrs.dur) / 1000;
                        if (ic.attrs.presetClass) presetClass = String(ic.attrs.presetClass);
                        if (ic.attrs.presetId != null) presetId = Number(ic.attrs.presetId);
                        if (ic.attrs.delay != null && ic.attrs.delay !== 'indefinite') delay = Number(ic.attrs.delay) / 1000;
                        if (ic.attrs.repeatCount != null) {
                            repeat = ic.attrs.repeatCount === 'indefinite' ? 'indefinite' : Number(ic.attrs.repeatCount) / 1000;
                        }
                    }
                }
            }
            const anim: Record<string, unknown> = {
                spid: Number(spid),
                type: preset,
                duration: dur ?? 1
            };
            if (presetClass) anim.presetClass = presetClass;
            if (presetId != null) anim.presetId = presetId;
            if (delay != null) anim.delay = delay;
            if (repeat != null) anim.repeat = repeat;
            animations.push(anim as { spid: number; type: string; duration: number });
        }
    }
    if (advanceTime == null && animations.length === 0) return null;
    return { advanceTime, animations };
}

export default pptxToHtml;
export { pptxToJson, pptxToHtml, pptxToFiles, jsonToPptx, editPptx, PPTXComposer, pptxToStandard };
export { PPTXShapeUtils } from './shape/shape';
export { presetShapePath } from './shape/preset-shape';
export * from './utils/color';
export * from './utils/geometry';
export * from './utils/units';
export * from './utils/format';
export * from './types/pptx-document';
export * from './compatibility-types';