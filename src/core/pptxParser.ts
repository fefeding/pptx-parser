import JSZip from "jszip";
import { PptxElementType, MediaInfo } from "../types/index";
import { XMLParser } from "fast-xml-parser";

const slideFactor = 96 / 914400;
const fontSizeFactor = 4 / 3.2;

/** WarpObj - 在解析过程中传递的上下文对象 */
interface WarpObj {
  zip: JSZip;
  slideContent: any;
  slideLayoutContent: any;
  slideLayoutTables?: any;
  slideMasterContent: any;
  slideMasterTables?: any;
  slideResObj: Record<string, any>;
  layoutResObj: Record<string, any>;
  masterResObj: Record<string, any>;
  themeContent: any;
  themeResObj: Record<string, any>;
  defaultTextStyle: any;
  slideMasterTextStyles?: any;
}

interface TextRun {
  text: string;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  link?: string;
}

interface TextBlock {
  runs: TextRun[];
  alignment?: "left" | "center" | "right" | "justify";
  verticalAlign?: "top" | "middle" | "bottom";
  x: number;
  y: number;
  width: number;
  height: number;
}

interface Shape {
  type: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  fillColor?: string;
  borderColor?: string;
  borderWidth?: number;
  textBlocks?: TextBlock[];
}

interface Slide {
  id: number;
  width: number;
  height: number;
  backgroundColor?: string;
  shapes: Shape[];
  images: MediaInfo[];
  tables: any[];
  graphs: any[];
}

interface ParsedData {
  slides: Record<number, Slide>;
  totalSlides: number;
  slideWidth: number;
  slideHeight: number;
  themeColors: Record<string, string>;
}

interface XmlElement {
  [key: string]: any;
  attrs?: Record<string, string>;
}

/** PPTX解析核心类 - 基于PPTXjs的解析逻辑重构 */
export class PptxParser {
  private zip!: JSZip;
  private slides: Record<number, any> = {};
  private totalSlides = 0;
  private slideWidth = 0;
  private slideHeight = 0;
  private themeColors: Record<string, string> = {};
  private defaultTextStyle: any = null;
  private slideLayoutContent: Record<string, any> = {};
  private slideMasterContent: Record<string, any> = {};
  private themeContent: any = null;
  private layoutResObj: Record<string, any> = {};
  private masterResObj: Record<string, any> = {};
  private themeResObj: Record<string, any> = {};
  private imageCache: Map<string, string> = new Map(); // 图片缓存

  constructor(private buffer: Uint8Array) {}

  /** 初始化解析（加载JSZip并解析PPTX结构） */
  async init(): Promise<void> {
    this.zip = await JSZip.loadAsync(this.buffer);
    await this.parseTheme();
    await this.parseSlideSize();
    await this.parseSlides();
    await this.parseMedia();
  }

  /** 解析幻灯片尺寸和默认文本样式 */
  private async parseSlideSize(): Promise<void> {
    try {
      const content = await this.readXmlFile("ppt/presentation.xml");
      if (!content) return;

      const sldSz = this.getTextByPathList(content, ["p:presentation", "p:sldSz"]);
      const sldSzAttrs = this.getAttrs(sldSz);
      if (sldSzAttrs) {
        const cx = parseInt(sldSzAttrs.cx);
        const cy = parseInt(sldSzAttrs.cy);
        this.slideWidth = Math.round(cx * slideFactor);
        this.slideHeight = Math.round(cy * slideFactor);
      }

      this.defaultTextStyle = this.getTextByPathList(content, ["p:presentation", "p:defaultTextStyle"]);
    } catch (e) {
      console.error("Failed to parse slide size:", e);
    }
  }

  /** 解析主题颜色 */
  private async parseTheme(): Promise<void> {
    // Theme会在parseSingleSlide中处理，这里保留空实现保持兼容性
  }

  /** 解析主题颜色（从themeContent中提取） */
  private parseThemeColors(themeContent: any): void {
    const colorScheme = this.getTextByPathList(themeContent, ["a:theme", "a:themeElements", "a:clrScheme"]);
    if (colorScheme) {
      Object.keys(colorScheme).forEach((key: string) => {
        if (key.startsWith("a:")) {
          const colorObj = colorScheme[key];
          const colorValue = this.getColorValue(colorObj);
          if (colorValue) {
            this.themeColors[key.substring(2)] = colorValue;
          }
        }
      });
    }
  }

  /** 索引节点（建立快速查找表） */
  private indexNodes(content: any): any {
    try {
      const rootKey = Object.keys(content).find(key => key.includes("sld"));
      if (!rootKey) return { idTable: {}, idxTable: {}, typeTable: {} };

      const cSld = content[rootKey]["p:cSld"];
      if (!cSld) return { idTable: {}, idxTable: {}, typeTable: {} };

      const spTree = cSld["p:spTree"];
      if (!spTree) return { idTable: {}, idxTable: {}, typeTable: {} };

      const idTable: Record<string, any> = {};
      const idxTable: Record<string, any> = {};
      const typeTable: Record<string, any> = {};

      for (const key in spTree) {
        if (key === "p:nvGrpSpPr" || key === "p:grpSpPr") continue;

        const targetNode = spTree[key];
        const nodeArray = Array.isArray(targetNode) ? targetNode : [targetNode];

        nodeArray.forEach((node: any) => {
          const nvSpPrNode = node["p:nvSpPr"];
          if (nvSpPrNode) {
            const cNvPr = nvSpPrNode["p:cNvPr"];
            const nvPr = nvSpPrNode["p:nvPr"];
            const ph = nvPr?.["p:ph"];

            const id = this.getAttr(cNvPr, "id");
            const idx = this.getAttr(ph, "idx");
            const type = this.getAttr(ph, "type");

            if (id !== undefined) idTable[id] = node;
            if (idx !== undefined) idxTable[idx] = node;
            if (type !== undefined) typeTable[type] = node;
          }
        });
      }

      return { idTable, idxTable, typeTable };
    } catch (e) {
      console.error("Failed to index nodes:", e);
      return { idTable: {}, idxTable: {}, typeTable: {} };
    }
  }

  /** 解析幻灯片列表 */
  private async parseSlides(): Promise<void> {
    try {
      const contentTypes = await this.readXmlFile("[Content_Types].xml");
      if (!contentTypes) return;

      const overrides = this.getTextByPathList(contentTypes, ["Types", "Override"]);
      if (!Array.isArray(overrides)) return;

      const slideFiles: string[] = [];
      overrides.forEach((o: any) => {
        const attrs = this.getAttrs(o);
        if (attrs?.ContentType === "application/vnd.openxmlformats-officedocument.presentationml.slide+xml") {
          slideFiles.push(attrs.PartName.substring(1));
        }
      });

      this.totalSlides = slideFiles.length;

      for (const slideFile of slideFiles) {
        const match = slideFile.match(/slide(\d+)\.xml$/);
        if (match) {
          const slideId = parseInt(match[1]);
          this.slides[slideId] = await this.parseSingleSlide(slideFile, slideId);
        }
      }
    } catch (e) {
      console.error("Failed to parse slides:", e);
    }
  }

  /** 解析layout或master中的shapes */
  private parseLayoutShapes(layoutContent: any, originalWarpObj: WarpObj): any[] {
    const shapes: any[] = [];
    
    // 获取layout或master的根节点
    const rootKey = Object.keys(layoutContent).find(key => key.includes("sld"));
    if (!rootKey) return shapes;

    const cSld = layoutContent[rootKey]["p:cSld"];
    if (!cSld) return shapes;

    const spTree = cSld["p:spTree"];
    if (!spTree) return shapes;

    // 创建专门的warpObj用于layout/master解析
    const layoutWarpObj: WarpObj = {
      ...originalWarpObj,
      slideContent: layoutContent, // 使用layout/master作为当前内容
    };

    // 解析所有子节点
    Object.keys(spTree).forEach((nodeKey: string) => {
      if (nodeKey === "p:nvGrpSpPr" || nodeKey === "p:grpSpPr") return;

      const nodes = spTree[nodeKey];
      const nodeArray = Array.isArray(nodes) ? nodes : [nodes];

      nodeArray.forEach((node: any) => {
        const result = this.processNode(nodeKey, node, layoutWarpObj);
        if (result && result.type === "shape") {
          // 标记这些形状来自layout或master
          result.isLayoutShape = true;
          shapes.push(result);
        }
      });
    });

    return shapes;
  }

  /** 解析单个幻灯片 */
  private async parseSingleSlide(slideFile: string, slideId: number): Promise<any> {
    try {
      const slideContent = await this.readXmlFile(slideFile);
      if (!slideContent) return null;

      // Step 1: 读取slide的relationship文件，获取slideLayout文件名
      const resFile = slideFile.replace("slides/slide", "slides/_rels/slide") + ".rels";
      const resContent = await this.readXmlFile(resFile);
      if (!resContent) {
        console.warn(`No relationship file found for slide ${slideId}`);
        return null;
      }

      const relationships = resContent["Relationships"]?.["Relationship"];
      if (!relationships) {
        console.warn(`No relationships found in ${resFile}`);
        return null;
      }

      const relsArray = Array.isArray(relationships) ? relationships : [relationships];
      const slideResObj: Record<string, any> = {};
      let layoutFilename = "";

      relsArray.forEach((r: any) => {
        const rAttrs = this.getAttrs(r);
        if (rAttrs) {
          slideResObj[rAttrs.Id] = { target: rAttrs.Target, type: rAttrs.Type };
          if (rAttrs.Type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout") {
            layoutFilename = rAttrs.Target.replace("../", "ppt/");
          }
        }
      });

      if (!layoutFilename) {
        console.warn(`No slideLayout found for slide ${slideId}`);
        return null;
      }

      // Step 2: 读取slideLayout文件
      const slideLayoutContent = await this.readXmlFile(layoutFilename);
      if (!slideLayoutContent) {
        console.warn(`Failed to read slideLayout ${layoutFilename}`);
        return null;
      }

      // 缓存slideLayout内容
      this.slideLayoutContent[layoutFilename] = slideLayoutContent;
      const slideLayoutTables = this.indexNodes(slideLayoutContent);

      // Step 3: 读取slideLayout的relationship文件，获取slideMaster文件名
      const layoutResFile = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
      const layoutResContent = await this.readXmlFile(layoutResFile);
      if (!layoutResContent) {
        console.warn(`No relationship file found for slideLayout ${layoutFilename}`);
        return null;
      }

      const layoutRelationships = layoutResContent["Relationships"]?.["Relationship"];
      if (!layoutRelationships) return null;

      const layoutRelsArray = Array.isArray(layoutRelationships) ? layoutRelationships : [layoutRelationships];
      const layoutResObj: Record<string, any> = {};
      let masterFilename = "";

      layoutRelsArray.forEach((r: any) => {
        const rAttrs = this.getAttrs(r);
        if (rAttrs) {
          layoutResObj[rAttrs.Id] = { target: rAttrs.Target, type: rAttrs.Type };
          if (rAttrs.Type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster") {
            masterFilename = rAttrs.Target.replace("../", "ppt/");
          }
        }
      });

      if (!masterFilename) {
        console.warn(`No slideMaster found for slideLayout ${layoutFilename}`);
        return null;
      }

      // Step 4: 读取slideMaster文件
      const slideMasterContent = await this.readXmlFile(masterFilename);
      if (!slideMasterContent) {
        console.warn(`Failed to read slideMaster ${masterFilename}`);
        return null;
      }

      // 缓存slideMaster内容
      this.slideMasterContent[masterFilename] = slideMasterContent;
      const slideMasterTables = this.indexNodes(slideMasterContent);

      // Step 5: 读取slideMaster的relationship文件，获取theme文件名
      const masterResFile = masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster") + ".rels";
      const masterResContent = await this.readXmlFile(masterResFile);
      if (!masterResContent) {
        console.warn(`No relationship file found for slideMaster ${masterFilename}`);
        return null;
      }

      const masterRelationships = masterResContent["Relationships"]?.["Relationship"];
      if (!masterRelationships) return null;

      const masterRelsArray = Array.isArray(masterRelationships) ? masterRelationships : [masterRelationships];
      const masterResObj: Record<string, any> = {};
      let themeFilename = "";

      masterRelsArray.forEach((r: any) => {
        const rAttrs = this.getAttrs(r);
        if (rAttrs) {
          masterResObj[rAttrs.Id] = { target: rAttrs.Target, type: rAttrs.Type };
          if (rAttrs.Type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
            themeFilename = rAttrs.Target.replace("../", "ppt/");
          }
        }
      });

      // Step 6: 读取theme文件
      let themeContent: any = null;
      if (themeFilename) {
        themeContent = await this.readXmlFile(themeFilename);
        if (themeContent) {
          this.themeContent = themeContent;
          this.parseThemeColors(themeContent);
        }
      }

      // 构建warpObj上下文对象
      const warpObj: WarpObj = {
        zip: this.zip,
        slideContent,
        slideLayoutContent,
        slideLayoutTables,
        slideMasterContent,
        slideMasterTables,
        slideResObj,
        layoutResObj,
        masterResObj,
        themeContent: this.themeContent,
        themeResObj: {},
        defaultTextStyle: this.defaultTextStyle,
      };

      const bgInfo = this.getSlideBackgroundInfo(warpObj);

      // 获取颜色映射覆盖
      const clrMapOvr = this.getColorMapOverride(slideContent, slideLayoutContent, slideMasterContent);

      // 解析layout中的shapes
      const layoutShapes = this.parseLayoutShapes(slideLayoutContent, warpObj);
      
      // 解析master中的shapes
      const masterShapes = this.parseLayoutShapes(slideMasterContent, warpObj);

      const slideData: any = {
        id: slideId,
        width: this.slideWidth,
        height: this.slideHeight,
        backgroundColor: bgInfo.color,
        backgroundImage: bgInfo.image,
        shapes: [],
        images: [],
        tables: [],
        graphs: [],
        // 保存layout和master信息
        layout: {
          filename: layoutFilename,
          content: slideLayoutContent,
          tables: slideLayoutTables,
          colorMapOvr: clrMapOvr.layout,
          shapes: layoutShapes,
        },
        master: {
          filename: masterFilename,
          content: slideMasterContent,
          tables: slideMasterTables,
          colorMapOvr: clrMapOvr.master,
          shapes: masterShapes,
        },
        theme: {
          filename: themeFilename,
          content: themeContent,
        },
        warpObj: warpObj, // 保留完整的warpObj引用以便后续使用
      };

      const spTree = this.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:spTree"]);
      if (!spTree) return slideData;

      // 解析所有子节点
      Object.keys(spTree).forEach((nodeKey: string) => {
        if (nodeKey === "p:nvGrpSpPr" || nodeKey === "p:grpSpPr") return;

        const nodes = spTree[nodeKey];
        const nodeArray = Array.isArray(nodes) ? nodes : [nodes];

        nodeArray.forEach((node: any) => {
          const result = this.processNode(nodeKey, node, warpObj);
          if (result) {
            if (result.type === "shape") {
              slideData.shapes.push(result);
            } else if (result.type === "image") {
              slideData.images.push(result);
            } else if (result.type === "table") {
              slideData.tables.push(result);
            } else if (result.type === "chart") {
              slideData.graphs.push(result);
            }
          }
        });
      });

      return slideData;
    } catch (e) {
      console.error(`Failed to parse slide ${slideFile}:`, e);
      return null;
    }
  }

  /** 处理节点 */
  private processNode(nodeKey: string, node: any, warpObj: WarpObj): any {
    switch (nodeKey) {
      case "p:sp":
        return this.processShapeNode(node, warpObj);
      case "p:pic":
        return this.processPictureNode(node, warpObj);
      case "p:graphicFrame":
        return this.processGraphicFrameNode(node, warpObj);
      case "p:grpSp":
        return this.processGroupNode(node, warpObj);
      default:
        return null;
    }
  }

  /** 处理形状节点（包括文本） */
  private processShapeNode(node: any, warpObj: WarpObj): any {
    const shape: any = { type: "shape" };

    // 获取节点ID和名称
    const nvSpPr = node["p:nvSpPr"];
    const cNvPr = nvSpPr?.["p:cNvPr"];
    const cNvPrAttrs = this.getAttrs(cNvPr);
    if (cNvPrAttrs) {
      shape.id = cNvPrAttrs.id;
      shape.name = cNvPrAttrs.name;
    }

    // 获取位置和尺寸
    let xfrm = this.getTextByPathList(node, ["p:spPr", "a:xfrm"]);
    if (xfrm) {
      const offAttrs = this.getAttrs(xfrm["a:off"]);
      const extAttrs = this.getAttrs(xfrm["a:ext"]);
      if (offAttrs) {
        shape.x = Math.round(parseInt(offAttrs.x) * slideFactor);
        shape.y = Math.round(parseInt(offAttrs.y) * slideFactor);
      }
      if (extAttrs) {
        shape.width = Math.round(parseInt(extAttrs.cx) * slideFactor);
        shape.height = Math.round(parseInt(extAttrs.cy) * slideFactor);
      }

      // 旋转
      const xfrmAttrs = this.getAttrs(xfrm);
      if (xfrmAttrs?.rot) {
        shape.rotation = Math.round((parseInt(xfrmAttrs.rot) / 60000));
      }
    } else {
      // 如果没有xfrm，尝试从占位符获取位置信息
      if (nvSpPr) {
        const nvPr = nvSpPr["p:nvPr"];
        const ph = nvPr?.["p:ph"];
        if (ph) {
          const phAttrs = this.getAttrs(ph);
          if (phAttrs) {
            const idx = phAttrs.idx;
            const type = phAttrs.type;

            // 从slideLayout或slideMaster中查找占位符
            let placeholderNode = null;
          if (warpObj.slideLayoutTables) {
              placeholderNode = idx !== undefined
                ? warpObj.slideLayoutTables.idxTable[idx]
                : type !== undefined
                ? warpObj.slideLayoutTables.typeTable[type]
                : null;
          }

          if (!placeholderNode && warpObj.slideMasterTables) {
              placeholderNode = idx !== undefined
                ? warpObj.slideMasterTables.idxTable[idx]
                : type !== undefined
                ? warpObj.slideMasterTables.typeTable[type]
                : null;
            }

            if (placeholderNode) {
              const phXfrm = this.getTextByPathList(placeholderNode, ["p:spPr", "a:xfrm"]);
              if (phXfrm) {
                const offAttrs = this.getAttrs(phXfrm["a:off"]);
                const extAttrs = this.getAttrs(phXfrm["a:ext"]);
                if (offAttrs) {
                  shape.x = Math.round(parseInt(offAttrs.x) * slideFactor);
                  shape.y = Math.round(parseInt(offAttrs.y) * slideFactor);
                }
                if (extAttrs) {
                  shape.width = Math.round(parseInt(extAttrs.cx) * slideFactor);
                  shape.height = Math.round(parseInt(extAttrs.cy) * slideFactor);
                }
              }
            }
          }
        }
      }
    }

    // 获取填充色
    shape.fillColor = this.parseShapeFill(node, warpObj);

    // 获取边框
    shape.borderColor = this.parseShapeBorder(node, warpObj);

    // 获取占位符信息
    const nvPr = nvSpPr?.["p:nvPr"];
    const ph = nvPr?.["p:ph"];
    const phAttrs = this.getAttrs(ph);
    if (phAttrs) {
      shape.placeholder = {
        idx: phAttrs.idx,
        type: phAttrs.type,
        orient: phAttrs.orient,
        sz: phAttrs.sz,
        hasCustomPrompt: phAttrs.hasCustomPrompt,
      };
    }

    // 保存原始节点引用
    shape.node = node;

    // 获取文本
    const txBody = this.getTextByPathList(node, ["p:txBody"]);
    if (txBody) {
      shape.textBlocks = this.parseTextBody(txBody, warpObj, shape.x || 0, shape.y || 0, shape.width || 0, shape.height || 0);
    } else if (shape.placeholder) {
      // 如果当前节点没有文本，尝试从占位符继承文本
      shape.textBlocks = this.getPlaceholderText(shape.placeholder, warpObj, shape.x || 0, shape.y || 0, shape.width || 0, shape.height || 0);
    }

    // 获取形状类型
    const prstGeom = this.getTextByPathList(node, ["p:spPr", "a:prstGeom"]);
    const prstGeomAttrs = this.getAttrs(prstGeom);
    if (prstGeomAttrs?.prst) {
      shape.shapeType = prstGeomAttrs.prst;
    }

    // 保存warpObj引用以便后续使用
    shape.warpObj = warpObj;

    return shape;
  }

  /** 处理图片节点 */
  private processPictureNode(node: any, warpObj: WarpObj): any {
    const blip = this.getTextByPathList(node, ["p:blipFill", "a:blip"]);
    const blipAttrs = this.getAttrs(blip);
    if (!blipAttrs) return null;

    const rEmbed = blipAttrs["r:embed"];
    if (!rEmbed) return null;

    const xfrm = this.getTextByPathList(node, ["p:spPr", "a:xfrm"]);
    const offAttrs = this.getAttrs(xfrm?.["a:off"]);
    const extAttrs = this.getAttrs(xfrm?.["a:ext"]);

    return {
      type: "image",
      rId: rEmbed,
      x: offAttrs ? Math.round(parseInt(offAttrs.x) * slideFactor) : 0,
      y: offAttrs ? Math.round(parseInt(offAttrs.y) * slideFactor) : 0,
      width: extAttrs ? Math.round(parseInt(extAttrs.cx) * slideFactor) : 0,
      height: extAttrs ? Math.round(parseInt(extAttrs.cy) * slideFactor) : 0,
    };
  }

  /** 处理图形框架节点（表格、图表） */
  private processGraphicFrameNode(node: any, warpObj: WarpObj): any {
    const graphic = this.getTextByPathList(node, ["a:graphic"]);
    if (!graphic) return null;

    // 检查是否是表格
    const table = this.getTextByPathList(graphic, ["a:graphicData", "a:tbl"]);
    if (table) {
      return this.parseTable(table, node, warpObj);
    }

    // 检查是否是图表
    const chart = this.getTextByPathList(graphic, ["a:graphicData", "c:chart"]);
    if (chart) {
      return this.parseChart(chart, node);
    }

    return null;
  }

  /** 处理组节点 */
  private processGroupNode(node: any, warpObj: WarpObj): any {
    const group: any = { type: "group", children: [] };

    const xfrm = this.getTextByPathList(node, ["p:grpSpPr", "a:xfrm"]);
    if (xfrm) {
      const offAttrs = this.getAttrs(xfrm["a:off"]);
      const extAttrs = this.getAttrs(xfrm["a:ext"]);
      if (offAttrs) {
        group.x = Math.round(parseInt(offAttrs.x) * slideFactor);
        group.y = Math.round(parseInt(offAttrs.y) * slideFactor);
      }
      if (extAttrs) {
        group.width = Math.round(parseInt(extAttrs.cx) * slideFactor);
        group.height = Math.round(parseInt(extAttrs.cy) * slideFactor);
      }
    }

    // 处理组内所有子节点
    Object.keys(node).forEach((nodeKey: string) => {
      if (nodeKey === "p:nvGrpSpPr" || nodeKey === "p:grpSpPr") return;

      const nodes = node[nodeKey];
      const nodeArray = Array.isArray(nodes) ? nodes : [nodes];

      nodeArray.forEach((childNode: any) => {
        const result = this.processNode(nodeKey, childNode, warpObj);
        if (result) {
          group.children.push(result);
        }
      });
    });

    return group;
  }

  /** 解析文本主体 */
  private parseTextBody(txBody: any, warpObj: WarpObj, x: number, y: number, width: number, height: number): TextBlock[] {
    const blocks: TextBlock[] = [];
    const paragraphs = txBody["a:p"];

    if (!paragraphs) return blocks;

    const paragraphArray = Array.isArray(paragraphs) ? paragraphs : [paragraphs];

    paragraphArray.forEach((p: any, idx: number) => {
      const block: TextBlock = {
        runs: [],
        alignment: this.parseTextAlignment(p),
        verticalAlign: this.parseVerticalAlignment(p),
        x,
        y,
        width,
        height,
      };

      const textRuns = p["a:r"];
      if (!textRuns) return;

      const runsArray = Array.isArray(textRuns) ? textRuns : [textRuns];

      runsArray.forEach((r: any) => {
        const run: TextRun = this.parseTextRunWithLink(r, warpObj);
        if (run) {
          block.runs.push(run);
        }
      });

      if (block.runs.length > 0) {
        blocks.push(block);
      }
    });

    return blocks;
  }

  /** 解析文本运行（包含链接） */
  private parseTextRunWithLink(r: any, warpObj: WarpObj): TextRun {
    const run = this.parseTextRun(r);

    // 检查是否有超链接
    const hlinkClick = r["a:rPr"]?.["a:hlinkClick"];
    if (hlinkClick) {
      const linkId = this.getAttr(hlinkClick, "r:id");
      if (linkId && warpObj.slideResObj[linkId]) {
        run.link = warpObj.slideResObj[linkId].target;
      }
    }

    return run;
  }

  /** 解析单个文本运行 */
  private parseTextRun(r: any): TextRun {
    // 获取文本内容
    let text = r["a:t"];
    if (typeof text !== "string") {
      // 尝试从fld元素获取
      text = this.getTextByPathList(r, ["a:fld", "a:t"]);
      if (typeof text !== "string") {
        text = "";
      }
    }

    // 处理特殊字符
    if (text) {
      text = text.replace(/\t/g, "    ").replace(/\n/g, " ");
    }

    const run: TextRun = {
      text: text || "",
    };

    const rPr = r["a:rPr"];
    const rPrAttrs = this.getAttrs(rPr);
    if (rPrAttrs) {
      // PPTX字体大小单位是百分之一磅，转换为像素
      // pptxjs 使用：fontSize = parseInt(sz) / 100 * fontSizeFactor (其中fontSizeFactor = 4 / 3.2 = 1.25)
      if (rPrAttrs.sz) {
        const szValue = parseInt(rPrAttrs.sz);
        if (!isNaN(szValue)) {
          run.fontSize = szValue / 100 * fontSizeFactor;
        }
      }
      if (rPrAttrs.b === "1") run.bold = true;
      if (rPrAttrs.i === "1") run.italic = true;
      if (rPrAttrs.u === "1") run.underline = true;
      if (rPrAttrs.strike === "1") run.strike = true;
    }

    // 获取字体族 - 支持多种字体类型
    if (rPr) {
      // 优先级1: ea (东亚字体)
      const eaAttrs = this.getAttrs(rPr["a:ea"]);
      if (eaAttrs?.typeface) {
        run.fontFamily = eaAttrs.typeface;
      }

      // 优先级2: latin (拉丁字体)
      const latinAttrs = this.getAttrs(rPr["a:latin"]);
      if (latinAttrs?.typeface && !run.fontFamily) {
        run.fontFamily = latinAttrs.typeface;
      }

      // 优先级3: cs (复杂脚本字体)
      const csAttrs = this.getAttrs(rPr["a:cs"]);
      if (csAttrs?.typeface && !run.fontFamily) {
        run.fontFamily = csAttrs.typeface;
      }

      // 备用方案：尝试type属性
      const eaTypeAttrs = this.getAttrs(rPr["a:ea"]);
      if (eaTypeAttrs?.type && !run.fontFamily) {
        run.fontFamily = eaTypeAttrs.type;
      }

      const latinTypeAttrs = this.getAttrs(rPr["a:latin"]);
      if (latinTypeAttrs?.type && !run.fontFamily) {
        run.fontFamily = latinTypeAttrs.type;
      }

      const csTypeAttrs = this.getAttrs(rPr["a:cs"]);
      if (csTypeAttrs?.type && !run.fontFamily) {
        run.fontFamily = csTypeAttrs.type;
      }
    }

    if (rPr?.["a:solidFill"]) {
      run.color = this.getColorValue(rPr["a:solidFill"]);
    }

    return run;
  }

  /** 解析文本对齐 */
  private parseTextAlignment(p: any): "left" | "center" | "right" | "justify" {
    const pPr = p["a:pPr"];
    const pPrAttrs = this.getAttrs(pPr);
    const algn = pPrAttrs?.algn;
    switch (algn) {
      case "ctr": return "center";
      case "r": return "right";
      case "just": return "justify";
      default: return "left";
    }
  }

  /** 解析垂直对齐 */
  private parseVerticalAlignment(p: any): "top" | "middle" | "bottom" {
    const pPr = p["a:pPr"];
    const pPrAttrs = this.getAttrs(pPr);
    const anchor = pPrAttrs?.anchor;
    switch (anchor) {
      case "ctr": return "middle";
      case "b": return "bottom";
      default: return "top";
    }
  }

  /** 解析形状填充 */
  private parseShapeFill(node: any, warpObj?: WarpObj): string | undefined {
    const solidFill = this.getTextByPathList(node, ["p:spPr", "a:solidFill"]);
    if (solidFill && warpObj) {
      const color = this.getSolidFill(solidFill, warpObj);
      return color ? `#${color}` : undefined;
    }
    return undefined;
  }

  /** 解析形状边框 */
  private parseShapeBorder(node: any, warpObj?: WarpObj): string | undefined {
    const ln = this.getTextByPathList(node, ["p:spPr", "a:ln"]);
    if (ln?.["a:solidFill"] && warpObj) {
      const color = this.getSolidFill(ln["a:solidFill"], warpObj);
      return color ? `#${color}` : undefined;
    }
    return undefined;
  }

  /** 解析幻灯片背景 */
  private parseSlideBackground(slideContent: any): string | undefined {
    const bg = this.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg"]);
    if (!bg) return undefined;

    const solidFill = this.getTextByPathList(bg, ["p:bgPr", "a:solidFill"]);
    if (solidFill) {
      return this.getColorValue(solidFill);
    }

    return undefined;
  }

  /** 获取幻灯片背景填充（支持layout和master继承） */
  private getSlideBackgroundFill(warpObj: WarpObj): string | undefined {
    const { slideContent, slideLayoutContent, slideMasterContent, themeContent } = warpObj;

    // 优先级1: Slide级别的背景
    let bg = this.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg"]);
    if (bg) {
      const bgPr = bg["p:bgPr"];
      const bgRef = bg["p:bgRef"];
      if (bgPr) {
        return this.getBgFill(bgPr, warpObj);
      } else if (bgRef) {
        return this.getBgRef(bgRef, warpObj);
      }
    }

    // 优先级2: SlideLayout级别的背景
    bg = this.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg"]);
    if (bg) {
      const bgPr = bg["p:bgPr"];
      const bgRef = bg["p:bgRef"];
      if (bgPr) {
        return this.getBgFill(bgPr, warpObj);
      } else if (bgRef) {
        return this.getBgRef(bgRef, warpObj);
      }
    }

    // 优先级3: SlideMaster级别的背景
    bg = this.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg"]);
    if (bg) {
      const bgPr = bg["p:bgPr"];
      const bgRef = bg["p:bgRef"];
      if (bgPr) {
        return this.getBgFill(bgPr, warpObj);
      } else if (bgRef) {
        return this.getBgRef(bgRef, warpObj);
      }
    }

    return undefined;
  }

  /** 获取幻灯片背景信息（包含颜色和图片） */
  private getSlideBackgroundInfo(warpObj: WarpObj): { color?: string; image?: string } {
    const { slideContent, slideLayoutContent, slideMasterContent } = warpObj;
    let result: { color?: string; image?: string } = {};

    // 优先级1: Slide级别的背景
    let bg = this.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg"]);
    if (bg) {
      const bgPr = bg["p:bgPr"];
      const bgRef = bg["p:bgRef"];
      if (bgPr) {
        result.color = this.getBgColor(bgPr, warpObj);
        result.image = this.getBgImage(bgPr, warpObj);
        if (result.color || result.image) return result;
      } else if (bgRef) {
        result = this.getBgRefInfo(bgRef, warpObj);
        if (result.color || result.image) return result;
      }
    }

    // 优先级2: SlideLayout级别的背景
    bg = this.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg"]);
    if (bg) {
      const bgPr = bg["p:bgPr"];
      const bgRef = bg["p:bgRef"];
      if (bgPr) {
        result.color = this.getBgColor(bgPr, warpObj);
        result.image = this.getBgImage(bgPr, warpObj);
        if (result.color || result.image) return result;
      } else if (bgRef) {
        result = this.getBgRefInfo(bgRef, warpObj);
        if (result.color || result.image) return result;
      }
    }

    // 优先级3: SlideMaster级别的背景
    bg = this.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg"]);
    if (bg) {
      const bgPr = bg["p:bgPr"];
      const bgRef = bg["p:bgRef"];
      if (bgPr) {
        result.color = this.getBgColor(bgPr, warpObj);
        result.image = this.getBgImage(bgPr, warpObj);
        if (result.color || result.image) return result;
      } else if (bgRef) {
        result = this.getBgRefInfo(bgRef, warpObj);
        if (result.color || result.image) return result;
      }
    }

    return result;
  }

  /** 获取背景填充 */
  private getBgFill(bgPr: any, warpObj: WarpObj): string | undefined {
    // 检查纯色填充
    const solidFill = bgPr["a:solidFill"];
    if (solidFill) {
      const color = this.getSolidFill(solidFill, warpObj);
      return color ? `#${color}` : undefined;
    }

    // 检查图片填充
    const blipFill = bgPr["a:blipFill"];
    if (blipFill) {
      return this.getBlipFill(blipFill, warpObj);
    }

    // 检查渐变填充
    const gradFill = bgPr["a:gradFill"];
    if (gradFill) {
      return this.getGradFill(gradFill, warpObj);
    }

    return undefined;
  }

  /** 获取背景颜色 */
  private getBgColor(bgPr: any, warpObj: WarpObj): string | undefined {
    const solidFill = bgPr["a:solidFill"];
    if (solidFill) {
      const color = this.getSolidFill(solidFill, warpObj);
      return color ? `#${color}` : undefined;
    }

    const gradFill = bgPr["a:gradFill"];
    if (gradFill) {
      const gsLst = gradFill["a:gsLst"];
      if (gsLst) {
        const gs = gsLst["a:gs"];
        if (gs) {
          const gsArray = Array.isArray(gs) ? gs : [gs];
          if (gsArray[0]) {
            const solidFill = gsArray[0]["a:solidFill"];
            if (solidFill) {
              const color = this.getSolidFill(solidFill, warpObj);
              return color ? `#${color}` : undefined;
            }
          }
        }
      }
    }

    return undefined;
  }

  /** 获取背景图片 */
  private getBgImage(bgPr: any, warpObj: WarpObj): string | undefined {
    const blipFill = bgPr["a:blipFill"];
    if (blipFill) {
      return this.getBlipFill(blipFill, warpObj);
    }
    return undefined;
  }

  /** 获取图片填充 */
  private getBlipFill(blipFill: any, warpObj: WarpObj): string | undefined {
    const blip = blipFill["a:blip"];
    if (blip) {
      const rEmbed = this.getAttr(blip, "r:embed");
      if (rEmbed) {
        // 查找图片路径
        if (warpObj.layoutResObj[rEmbed]) {
          return warpObj.layoutResObj[rEmbed].target.replace("../", "ppt/");
        } else if (warpObj.masterResObj[rEmbed]) {
          return warpObj.masterResObj[rEmbed].target.replace("../", "ppt/");
        } else if (warpObj.themeResObj[rEmbed]) {
          return warpObj.themeResObj[rEmbed].target.replace("../", "ppt/");
        }
      }
    }
    return undefined;
  }

  /** 获取渐变填充 */
  private getGradFill(gradFill: any, warpObj: WarpObj): string | undefined {
    // 简化处理：使用第一个颜色作为背景
    const gsLst = gradFill["a:gsLst"];
    if (gsLst) {
      const gs = gsLst["a:gs"];
      if (gs) {
        const gsArray = Array.isArray(gs) ? gs : [gs];
        if (gsArray[0]) {
          const solidFill = gsArray[0]["a:solidFill"];
          if (solidFill) {
            const color = this.getSolidFill(solidFill, warpObj);
            return color ? `#${color}` : undefined;
          }
        }
      }
    }
    return undefined;
  }

  /** 获取背景引用信息 */
  private getBgRefInfo(bgRef: any, warpObj: WarpObj): { color?: string; image?: string } {
    const result: { color?: string; image?: string } = {};
    const idx = this.getAttr(bgRef, "idx");
    if (!idx) return result;

    const idxNum = parseInt(idx);
    if (idxNum > 1000) {
      // 从theme的bgFillStyleLst获取
      const trueIdx = idxNum - 1000;
      const bgFillLst = this.getTextByPathList(warpObj.themeContent, ["a:theme", "a:themeElements", "a:fmtScheme", "a:bgFillStyleLst"]);
      if (bgFillLst) {
        const bgFillKeys = Object.keys(bgFillLst).filter(key => key.startsWith("a:"));
        if (bgFillKeys[trueIdx - 1]) {
          const bgFill = bgFillLst[bgFillKeys[trueIdx - 1]];

          // 检查纯色填充
          if (bgFill["a:solidFill"]) {
            const color = this.getSolidFill(bgFill["a:solidFill"], warpObj);
            result.color = color ? `#${color}` : undefined;
          }

          // 检查图片填充
          if (bgFill["a:blipFill"]) {
            result.image = this.getBlipFill(bgFill["a:blipFill"], warpObj);
          }

          // 检查渐变填充
          if (bgFill["a:gradFill"]) {
            const color = this.getGradFill(bgFill["a:gradFill"], warpObj);
            result.color = color;
          }
        }
      }
    } else {
      // idx <= 1000: 从layout或master的背景图片获取
      // 这里返回路径，稍后会被转换为base64
      result.image = this.getBgImageFromRef(idxNum, warpObj);
    }
    return result;
  }

  /** 从引用获取背景图片路径 */
  private getBgImageFromRef(idx: number, warpObj: WarpObj): string | undefined {
    // 根据idx从layout或master的背景图片列表中获取
    // 这是一个简化实现，实际PPTX的背景图片引用可能更复杂
    return undefined;
  }

  /** 获取背景引用 */
  private getBgRef(bgRef: any, warpObj: WarpObj): string | undefined {
    const idx = this.getAttr(bgRef, "idx");
    if (!idx) return undefined;

    const idxNum = parseInt(idx);
    if (idxNum > 1000) {
      // 从theme的bgFillStyleLst获取
      const trueIdx = idxNum - 1000;
      const bgFillLst = this.getTextByPathList(warpObj.themeContent, ["a:theme", "a:themeElements", "a:fmtScheme", "a:bgFillStyleLst"]);
      if (bgFillLst) {
        const bgFillKeys = Object.keys(bgFillLst).filter(key => key.startsWith("a:"));
        if (bgFillKeys[trueIdx - 1]) {
          const bgFill = bgFillLst[bgFillKeys[trueIdx - 1]];

          // 检查纯色填充
          if (bgFill["a:solidFill"]) {
            const color = this.getSolidFill(bgFill["a:solidFill"], warpObj);
            return color ? `#${color}` : undefined;
          }

          // 检查图片填充
          if (bgFill["a:blipFill"]) {
            return this.getBlipFill(bgFill["a:blipFill"], warpObj);
          }

          // 检查渐变填充
          if (bgFill["a:gradFill"]) {
            return this.getGradFill(bgFill["a:gradFill"], warpObj);
          }
        }
      }
    }
    return undefined;
  }

  /** 获取纯色填充 */
  private getSolidFill(solidFill: any, warpObj: WarpObj): string | undefined {
    if (solidFill["a:srgbClr"]) {
      const srgbAttrs = this.getAttrs(solidFill["a:srgbClr"]);
      return srgbAttrs?.val;
    } else if (solidFill["a:schemeClr"]) {
      const schemeClr = solidFill["a:schemeClr"];
      const schemeClrAttrs = this.getAttrs(schemeClr);
      if (schemeClrAttrs?.val) {
        return this.getSchemeColorFromTheme(schemeClrAttrs.val, warpObj);
      }
    } else if (solidFill["a:prstClr"]) {
      const prstClr = solidFill["a:prstClr"];
      const prstClrAttrs = this.getAttrs(prstClr);
      if (prstClrAttrs?.val) {
        return this.getPrstColorValue(prstClrAttrs.val);
      }
    } else if (solidFill["a:sysClr"]) {
      const sysClr = solidFill["a:sysClr"];
      const sysClrAttrs = this.getAttrs(sysClr);
      return sysClrAttrs?.lastClr;
    }
    return undefined;
  }

  /** 从主题获取方案颜色 */
  private getSchemeColorFromTheme(schemeClrName: string, warpObj: WarpObj): string | undefined {
    // 首先尝试直接从缓存的主题颜色获取
    if (this.themeColors[schemeClrName]) {
      return this.themeColors[schemeClrName];
    }

    // 从themeContent中查找
    const colorScheme = this.getTextByPathList(warpObj.themeContent, ["a:theme", "a:themeElements", "a:clrScheme"]);
    if (colorScheme) {
      const schemeClr = colorScheme[`a:${schemeClrName}`];
      if (schemeClr) {
        return this.getColorValue(schemeClr);
      }
    }

    return undefined;
  }

  /** 获取占位符文本 */
  private getPlaceholderText(placeholder: any, warpObj: WarpObj, x: number, y: number, width: number, height: number): TextBlock[] {
    const idx = placeholder.idx;
    const type = placeholder.type;

    // 从slideLayout中查找占位符文本
    let placeholderNode = null;
    if (warpObj.slideLayoutTables) {
      placeholderNode = idx !== undefined
        ? warpObj.slideLayoutTables.idxTable[idx]
        : type !== undefined
        ? warpObj.slideLayoutTables.typeTable[type]
        : null;
    }

    if (!placeholderNode && warpObj.slideMasterTables) {
      placeholderNode = idx !== undefined
        ? warpObj.slideMasterTables.idxTable[idx]
        : type !== undefined
        ? warpObj.slideMasterTables.typeTable[type]
        : null;
    }

    if (placeholderNode) {
      const txBody = this.getTextByPathList(placeholderNode, ["p:txBody"]);
      if (txBody) {
        return this.parseTextBody(txBody, warpObj, x, y, width, height);
      }
    }

    return [];
  }

  /** 获取颜色映射覆盖 */
  private getColorMapOverride(slideContent: any, slideLayoutContent: any, slideMasterContent: any): {
    slide?: Record<string, string>;
    layout?: Record<string, string>;
    master?: Record<string, string>;
  } {
    const result: { slide?: Record<string, string>; layout?: Record<string, string>; master?: Record<string, string> } = {};

    // Slide级别的颜色映射
    const slideClrMapOvr = this.getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping"]);
    if (slideClrMapOvr) {
      const attrs = this.getAttrs(slideClrMapOvr);
      result.slide = attrs;
    }

    // Layout级别的颜色映射
    const layoutClrMapOvr = this.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);
    if (layoutClrMapOvr) {
      const attrs = this.getAttrs(layoutClrMapOvr);
      result.layout = attrs;
    }

    // Master级别的颜色映射
    const masterClrMap = this.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "a:clrMapping"]);
    if (masterClrMap) {
      const attrs = this.getAttrs(masterClrMap);
      result.master = attrs;
    }

    return result;
  }

  /** 解析表格 */
  private parseTable(table: any, node: any, warpObj: WarpObj): any {
    const tableData: any = {
      type: "table",
      rows: [],
      style: {},
    };

    const gridCol = table["a:tblGrid"]?.["a:gridCol"];
    if (gridCol) {
      const colArray = Array.isArray(gridCol) ? gridCol : [gridCol];
      tableData.columnWidths = colArray.map((col: any) => {
        const colAttrs = this.getAttrs(col);
        return Math.round(parseInt(colAttrs?.w || "0") * slideFactor);
      });
    }

    const tr = table["a:tr"];
    if (!tr) return tableData;

    const rowArray = Array.isArray(tr) ? tr : [tr];

    rowArray.forEach((row: any) => {
      const rowData: any = { cells: [] };
      const tc = row["a:tc"];
      if (!tc) return;

      const cellArray = Array.isArray(tc) ? tc : [tc];

      cellArray.forEach((cell: any) => {
        const cellData: any = { textBlocks: [] };
        const txBody = cell["a:txBody"];
        if (txBody) {
          cellData.textBlocks = this.parseTextBody(txBody, warpObj, 0, 0, 0, 0);
        }
        rowData.cells.push(cellData);
      });

      tableData.rows.push(rowData);
    });

    return tableData;
  }

  /** 解析图表 */
  private parseChart(chart: any, node: any): any {
    return {
      type: "chart",
      chartType: this.getTextByPathList(chart, ["c:chart", "attrs", "r:id"]) || "unknown",
    };
  }

  /** 解析媒体文件（图片、音频、视频） */
  private async parseMedia(): Promise<void> {
    // 解析幻灯片的关系文件，获取媒体文件的引用
    for (const slideId in this.slides) {
      const slideData = this.slides[slideId];

      // 处理背景图片
      if (slideData.backgroundImage) {
        slideData.backgroundImageData = await this.extractImage(slideData.backgroundImage);
      }

      if (!slideData.images || slideData.images.length === 0) continue;

      const relsFile = `ppt/slides/_rels/slide${slideId}.xml.rels`;
      const relsContent = await this.readXmlFile(relsFile);
      if (!relsContent) continue;

      const relationships = relsContent["Relationships"]?.["Relationship"];
      if (!relationships) continue;

      const relsArray = Array.isArray(relationships) ? relationships : [relationships];
      const relMap = new Map(relsArray.map((r: any) => {
        const rAttrs = this.getAttrs(r);
        return [rAttrs?.Id, rAttrs?.Target];
      }));

      for (const image of slideData.images) {
        const target = relMap.get(image.rId);
        if (target) {
          image.path = target.replace("../", "ppt/");
          // 提取图片数据
          image.data = await this.extractImage(image.path);
        }
      }
    }
  }

  /** 从zip中提取图片并转换为base64 */
  private async extractImage(path: string): Promise<string> {
    // 检查缓存
    if (this.imageCache.has(path)) {
      return this.imageCache.get(path)!;
    }

    try {
      const file = this.zip.file(path);
      if (!file) return "";

      // 读取为base64
      const base64 = await file.async("base64");

      // 根据文件扩展名确定MIME类型
      const ext = path.split('.').pop()?.toLowerCase() || '';
      const mimeTypes: Record<string, string> = {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif': 'image/gif',
        'bmp': 'image/bmp',
        'svg': 'image/svg+xml',
        'emf': 'image/x-emf',
        'wmf': 'image/x-wmf',
      };

      const mimeType = mimeTypes[ext] || 'image/png';
      const dataUrl = `data:${mimeType};base64,${base64}`;

      // 缓存结果
      this.imageCache.set(path, dataUrl);
      return dataUrl;
    } catch (e) {
      console.error(`Failed to extract image ${path}:`, e);
      return "";
    }
  }

  /** 获取颜色值 */
  private getColorValue(colorObj: any): string | undefined {
    if (!colorObj) return undefined;

    // 预设颜色
    if (colorObj["a:prstClr"]) {
      return this.getPrstColorValue(colorObj["a:prstClr"]);
    }

    // 方案颜色
    if (colorObj["a:schemeClr"]) {
      const schemeClr = colorObj["a:schemeClr"];
      const schemeClrAttrs = this.getAttrs(schemeClr);
      if (schemeClrAttrs?.val) {
        return this.themeColors[schemeClrAttrs.val];
      }
    }

    // RGB 颜色
    if (colorObj["a:srgbClr"]) {
      const srgbClr = colorObj["a:srgbClr"];
      const srgbAttrs = this.getAttrs(srgbClr);
      return `#${srgbAttrs?.val || "000000"}`;
    }

    // 系统颜色
    if (colorObj["a:sysClr"]) {
      const sysClr = colorObj["a:sysClr"];
      const sysClrAttrs = this.getAttrs(sysClr);
      return sysClrAttrs?.lastClr || undefined;
    }

    return undefined;
  }

  /** 获取预设颜色值 */
  private getPrstColorValue(prstClr: any): string | undefined {
    const prstClrAttrs = this.getAttrs(prstClr);
    if (!prstClrAttrs?.val) return undefined;
    const val = prstClrAttrs.val;
    const prstColors: Record<string, string> = {
      black: "000000", white: "FFFFFF", red: "FF0000", green: "00FF00", blue: "0000FF",
      yellow: "FFFF00", cyan: "00FFFF", magenta: "FF00FF", gray: "808080",
    };
    return prstColors[val] ? `#${prstColors[val]}` : undefined;
  }

  /** 读取XML文件 */
  private async readXmlFile(filename: string): Promise<any> {
    try {
      const file = this.zip.file(filename);
      if (!file) return null;

      // 使用Uint8Array读取，然后用TextDecoder处理UTF-8编码
      const buffer = await file.async("uint8array");
      const decoder = new TextDecoder("utf-8");
      const content = decoder.decode(buffer);

      const parser = new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: "@_",
        textNodeName: "#text",
        ignoreDeclaration: true,
        ignorePiTags: true,
        isArray: (name, jpath, isLeafNode, isAttribute) => {
          // 对于某些标签，始终保持为数组格式
          if (["a:gridCol", "a:tr", "a:tc", "a:p", "a:r", "Relationship"].includes(name)) {
            return true;
          }
          return false;
        },
      });
      return parser.parse(content);
    } catch (e) {
      console.error(`Failed to read XML file ${filename}:`, e);
      return null;
    }
  }

  /** 根据路径获取文本 */
  private getTextByPathList(obj: any, pathList: string[]): any {
    if (!obj) return undefined;

    let current = obj;
    for (const key of pathList) {
      if (current && current[key] !== undefined) {
        current = current[key];
      } else {
        return undefined;
      }
    }
    return current;
  }

  /** 获取对象属性（兼容 fast-xml-parser 的 @_ 前缀） */
  private getAttrs(obj: any): Record<string, string> | undefined {
    if (!obj) return undefined;

    // fast-xml-parser 使用 @_ 前缀
    const attrs: Record<string, string> = {};
    for (const key in obj) {
      if (key.startsWith("@_")) {
        attrs[key.substring(2)] = obj[key];
      }
    }
    return Object.keys(attrs).length > 0 ? attrs : undefined;
  }

  /** 获取单个属性值（兼容 fast-xml-parser 的 @_ 前缀） */
  private getAttr(obj: any, name: string): string | undefined {
    if (!obj) return undefined;
    return obj[`@_${name}`];
  }

  /** 解析文本元素 */
  private parseTextElements(xml: string): any[] {
    return [];
  }

  /** 解析文本块 */
  private parseTextBlocks(xml: string): any[] {
    return [];
  }

  /** 解析形状 */
  private parseShapes(xml: string): any[] {
    return [];
  }

  /** 解析媒体元素（图片/音频/视频，适配各浏览器格式） */
  private parseMediaElements(xml: string): MediaInfo[] {
    return [];
  }

  /** 解析图表 */
  private parseGraphs(xml: string): any[] {
    return [];
  }

  /** 解析表格 */
  private parseTables(xml: string): any[] {
    return [];
  }

  /** 解析SmartArt */
  private parseSmartArt(xml: string): any[] {
    return [];
  }

  /** 解析公式 */
  private parseEquations(xml: string): any[] {
    return [];
  }

  /** 获取解析后的幻灯片数据 */
  getSlides(): Record<number, any> {
    return this.slides;
  }

  /** 获取总页数 */
  getTotalSlides(): number {
    return this.totalSlides;
  }

  /** 获取幻灯片宽度 */
  getSlideWidth(): number {
    return this.slideWidth;
  }

  /** 获取幻灯片高度 */
  getSlideHeight(): number {
    return this.slideHeight;
  }

  /** 获取指定幻灯片的warpObj对象（用于渲染） */
  getWarpObjForSlide(slideId: number): WarpObj | null {
    const slideData = this.slides[slideId];
    if (!slideData) return null;

    // 从slideData中提取layout、master和theme信息
    const slideLayoutContent = slideData.layout?.content;
    const slideLayoutTables = slideData.layout?.tables;
    const slideMasterContent = slideData.master?.content;
    const slideMasterTables = slideData.master?.tables;
    const themeContent = slideData.theme?.content;
    const slideMasterTextStyles = this.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);

    return {
      zip: this.zip,
      slideContent: slideData,
      slideLayoutContent,
      slideLayoutTables,
      slideMasterContent,
      slideMasterTables,
      slideResObj: {},
      layoutResObj: {},
      masterResObj: {},
      themeContent,
      themeResObj: {},
      defaultTextStyle: this.defaultTextStyle,
      slideMasterTextStyles: slideMasterTextStyles
    };
  }
}