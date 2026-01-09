/**
 * PPTXjs主解析器 - TypeScript转译版
 * 完整转译PPTXjs.js的核心解析功能
 * 原始版本: PPTXjs.js v1.21.1
 */

import JSZip from 'jszip';
import {
  PptxjsCoreParser,
  WarpObj,
  ContentTypes,
  SlideSize,
  IndexTable,
  angleToDegrees,
} from './pptxjs-core-parser';
import {
  getColorValue,
  applyColorMap,
  parseColorMapOverride,
  generateCssColor,
  parseColorFill,
  getTextByPathList,
} from './pptxjs-color-utils';
import {
  parseTextBoxContent,
  generateTextBoxHtml,
  mergeTextStyles,
  getDefaultTextStyle,
} from './pptxjs-text-utils';
import { base64ArrayBuffer, getImageBase64, getImageMimeType, generateDataUrl } from './pptxjs-utils';

// 使用统一的getTextByPathList
const _getTextByPathList = getTextByPathList;

/**
 * PPTXjs解析器配置
 */
export interface PptxjsParserOptions {
  processFullTheme?: boolean;
  incSlideWidth?: number;
  incSlideHeight?: number;
  slideMode?: boolean;
  slideType?: 'div' | 'section' | 'divs2slidesjs' | 'revealjs';
  slidesScale?: string;
}

/**
 * 幻灯片数据结构
 */
export interface SlideData {
  id: number;
  fileName: string;
  width: number;
  height: number;
  bgColor?: string;
  bgFill?: any;
  shapes: any[];
  images: any[];
  tables: any[];
  charts: any[];
  layout?: {
    fileName: string;
    content: any;
    tables: IndexTable;
    colorMapOvr?: any;
  };
  master?: {
    fileName: string;
    content: any;
    tables: IndexTable;
    colorMapOvr?: any;
  };
  theme?: {
    fileName: string;
    content: any;
  };
  warpObj: WarpObj;
}

/**
 * PPTXjs主解析器类
 */
export class PptxjsParser {
  private zip: JSZip;
  private coreParser: PptxjsCoreParser;
  private options: Required<PptxjsParserOptions>;
  private tableStyles: any = null;

  constructor(zip: JSZip, options: PptxjsParserOptions = {}) {
    this.zip = zip;
    this.options = {
      processFullTheme: options.processFullTheme !== false,
      incSlideWidth: options.incSlideWidth || 0,
      incSlideHeight: options.incSlideHeight || 0,
      slideMode: options.slideMode || false,
      slideType: options.slideType || 'div',
      slidesScale: options.slidesScale || '',
    };

    this.coreParser = new PptxjsCoreParser(zip, {
      processFullTheme: this.options.processFullTheme,
      incSlideWidth: this.options.incSlideWidth,
      incSlideHeight: this.options.incSlideHeight,
    });
  }

  /**
   * 解析整个PPTX文件 - 对齐PPTXjs的processPPTX函数（第321-394行）
   */
  async parse(): Promise<{
    slides: SlideData[];
    size: SlideSize;
    thumb?: string;
    globalCSS: string;
  }> {
    const dateBefore = new Date();

    // 1. 获取缩略图（如果存在）
    let pptxThumbImg: string | undefined;
    const thumbFile = this.zip.file('docProps/thumbnail.jpeg');
    if (thumbFile) {
      const arrayBuffer = thumbFile.asArrayBuffer();
      pptxThumbImg = base64ArrayBuffer(arrayBuffer);
    }

    // 2. 获取内容类型
    const contentTypes = this.coreParser.getContentTypes();

    // 3. 获取幻灯片尺寸
    const slideSize = this.coreParser.getSlideSizeAndSetDefaultTextStyle();

    // 4. 读取表格样式
    this.tableStyles = this.coreParser.readXmlFile('ppt/tableStyles.xml');

    // 5. 解析所有幻灯片
    const slides: SlideData[] = [];
    const slidesFiles = contentTypes.slides;

    for (let i = 0; i < slidesFiles.length; i++) {
      const slideFile = slidesFiles[i];
      const slideId = i + 1;
      const slideData = await this.processSingleSlide(slideFile, slideId, slideSize);
      slides.push(slideData);
    }

    // 6. 生成全局CSS
    const globalCSS = this.generateGlobalCSS();

    const dateAfter = new Date();
    const executionTime = dateAfter.getTime() - dateBefore.getTime();
    console.log(`PPTX parsing completed in ${executionTime}ms`);

    return {
      slides,
      size: slideSize,
      thumb: pptxThumbImg,
      globalCSS,
    };
  }

  /**
   * 解析单个幻灯片 - 对齐PPTXjs的processSingleSlide函数（第499-723行）
   */
  private async processSingleSlide(
    sldFileName: string,
    slideId: number,
    slideSize: SlideSize
  ): Promise<SlideData> {
    // =====< Step 1 >===== 
    // 读取幻灯片关系文件，获取layout文件名
    const resName = sldFileName.replace('slides/slide', 'slides/_rels/slide') + '.rels';
    const resContent = this.coreParser.readXmlFile(resName);

    let layoutFilename = '';
    let diagramFilename = '';
    const slideResObj: Record<string, { type: string; target: string }> = {};

    const relationshipArray = resContent?.Relationships?.Relationship;
    const relArray = Array.isArray(relationshipArray) ? relationshipArray : [relationshipArray];

    for (const rel of relArray) {
      const attrs = rel?.attrs;
      if (!attrs) continue;

      switch (attrs.Type) {
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout':
          layoutFilename = attrs.Target.replace('../', 'ppt/');
          break;
        case 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing':
          diagramFilename = attrs.Target.replace('../', 'ppt/');
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide':
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image':
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart':
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink':
        default:
          slideResObj[attrs.Id] = {
            type: attrs.Type.replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: attrs.Target.replace('../', 'ppt/'),
          };
          break;
      }
    }

    // =====< Step 2 >===== 
    // 读取slideLayout文件
    const slideLayoutContent = this.coreParser.readXmlFile(layoutFilename);
    const slideLayoutTables = this.coreParser.indexNodes(slideLayoutContent);

    // 读取slideLayout关系文件
    const slideLayoutResFilename = layoutFilename.replace('slideLayouts/slideLayout', 'slideLayouts/_rels/slideLayout') + '.rels';
    const slideLayoutResContent = this.coreParser.readXmlFile(slideLayoutResFilename);

    let masterFilename = '';
    const layoutResObj: Record<string, { type: string; target: string }> = {};

    const layoutRelArray = slideLayoutResContent?.Relationships?.Relationship;
    const layoutRels = Array.isArray(layoutRelArray) ? layoutRelArray : [layoutRelArray];

    for (const rel of layoutRels) {
      const attrs = rel?.attrs;
      if (!attrs) continue;

      switch (attrs.Type) {
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster':
          masterFilename = attrs.Target.replace('../', 'ppt/');
          break;
        default:
          layoutResObj[attrs.Id] = {
            type: attrs.Type.replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: attrs.Target.replace('../', 'ppt/'),
          };
          break;
      }
    }

    // =====< Step 3 >===== 
    // 读取slideMaster文件
    const slideMasterContent = this.coreParser.readXmlFile(masterFilename);
    const slideMasterTextStyles = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:txStyles']);
    const slideMasterTables = this.coreParser.indexNodes(slideMasterContent);

    // 读取slideMaster关系文件
    const slideMasterResFilename = masterFilename.replace('slideMasters/slideMaster', 'slideMasters/_rels/slideMaster') + '.rels';
    const slideMasterResContent = this.coreParser.readXmlFile(slideMasterResFilename);

    let themeFilename = '';
    const masterResObj: Record<string, { type: string; target: string }> = {};

    const masterRelArray = slideMasterResContent?.Relationships?.Relationship;
    const masterRels = Array.isArray(masterRelArray) ? masterRelArray : [masterRelArray];

    for (const rel of masterRels) {
      const attrs = rel?.attrs;
      if (!attrs) continue;

      switch (attrs.Type) {
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme':
          themeFilename = attrs.Target.replace('../', 'ppt/');
          break;
        default:
          masterResObj[attrs.Id] = {
            type: attrs.Type.replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: attrs.Target.replace('../', 'ppt/'),
          };
          break;
      }
    }

    // =====< Step 4 >===== 
    // 读取theme文件
    const themeContent = this.coreParser.readXmlFile(themeFilename);
    const themeResObj: Record<string, { type: string; target: string }> = {};

    if (themeFilename) {
      const themeName = themeFilename.split('/').pop();
      const themeResFileName = themeFilename.replace(themeName, '_rels/' + themeName) + '.rels';
      const themeResContent = this.coreParser.readXmlFile(themeResFileName);

      const themeRelArray = themeResContent?.Relationships?.Relationship;
      const themeRels = Array.isArray(themeRelArray) ? themeRelArray : [themeRelArray];

      for (const rel of themeRels) {
        const attrs = rel?.attrs;
        if (!attrs) continue;

        themeResObj[attrs.Id] = {
          type: attrs.Type.replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
          target: attrs.Target.replace('../', 'ppt/'),
        };
      }
    }

    // =====< Step 5 >===== 
    // 处理diagram文件（如果存在）
    let diagramFileContent: any = null;
    const diagramResObj: Record<string, { type: string; target: string }> = {};

    if (diagramFilename) {
      diagramFileContent = this.coreParser.readXmlFile(diagramFilename);
      if (diagramFileContent) {
        const diagramObjStr = JSON.stringify(diagramFileContent);
        diagramFileContent = JSON.parse(diagramObjStr.replace(/dsp:/g, 'p:'));

        const diagName = diagramFilename.split('/').pop();
        const diagramResFileName = diagramFilename.replace(diagName, '_rels/' + diagName) + '.rels';
        const diagramResContent = this.coreParser.readXmlFile(diagramResFileName);

        const diagRelArray = diagramResContent?.Relationships?.Relationship;
        const diagRels = Array.isArray(diagRelArray) ? diagRelArray : [diagRelArray];

        for (const rel of diagRels) {
          const attrs = rel?.attrs;
          if (!attrs) continue;

          diagramResObj[attrs.Id] = {
            type: attrs.Type.replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: attrs.Target.replace('../', 'ppt/'),
          };
        }
      }
    }

    // =====< Step 6 >===== 
    // 读取幻灯片内容
    const slideContent = this.coreParser.readXmlFile(sldFileName, true);
    const nodes = _getTextByPathList(slideContent, ['p:sld', 'p:cSld', 'p:spTree']);

    // 构建warpObj上下文对象（对齐PPTXjs第675-691行）
    const warpObj: WarpObj = {
      zip: this.zip,
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
      diagramFileContent,
      diagramResObj,
      defaultTextStyle: this.coreParser.getDefaultTextStyle(),
    };

    // =====< Step 7 >===== 
    // 获取背景信息
    let bgResult = '';
    let bgColor = '';
    
    if (this.options.processFullTheme) {
      bgResult = this.getBackground(warpObj, slideId);
    } else if (this.options.processFullTheme === 'colorsAndImageOnly') {
      bgColor = this.getSlideBackgroundFill(warpObj, slideId);
    }

    // =====< Step 8 >===== 
    // 解析形状和元素
    const shapes: any[] = [];
    const images: any[] = [];
    const tables: any[] = [];
    const charts: any[] = [];

    if (nodes) {
      for (const nodeKey in nodes) {
        const nodeArray = Array.isArray(nodes[nodeKey]) ? nodes[nodeKey] : [nodes[nodeKey]];
        
        for (const node of nodeArray) {
          const result = this.processNodesInSlide(nodeKey, node, nodes, warpObj, 'slide');
          
          if (result) {
            // 根据结果类型分类
            if (result.type === 'shape') {
              shapes.push(result);
            } else if (result.type === 'image') {
              images.push(result);
            } else if (result.type === 'table') {
              tables.push(result);
            } else if (result.type === 'chart') {
              charts.push(result);
            }
          }
        }
      }
    }

    // =====< Step 9 >===== 
    // 获取颜色映射覆盖
    const clrMapOvr = parseColorMapOverride(slideContent, slideLayoutContent, slideMasterContent);

    // 构建幻灯片数据对象
    const slideData: SlideData = {
      id: slideId,
      fileName: sldFileName,
      width: slideSize.width,
      height: slideSize.height,
      bgColor,
      bgFill: bgResult,
      shapes,
      images,
      tables,
      charts,
      layout: {
        fileName: layoutFilename,
        content: slideLayoutContent,
        tables: slideLayoutTables,
        colorMapOvr: clrMapOvr.layout,
      },
      master: {
        fileName: masterFilename,
        content: slideMasterContent,
        tables: slideMasterTables,
        colorMapOvr: clrMapOvr.master,
      },
      theme: {
        fileName: themeFilename,
        content: themeContent,
      },
      warpObj,
    };

    return slideData;
  }

  /**
   * 处理幻灯片中的节点 - 对齐PPTXjs的processNodesInSlide函数（第781-811行）
   */
  private processNodesInSlide(
    nodeKey: string,
    nodeValue: any,
    nodes: any,
    warpObj: WarpObj,
    source: string,
    sType?: string
  ): any {
    switch (nodeKey) {
      case 'p:sp': // 形状、文本
        return this.processSpNode(nodeValue, nodes, warpObj, source, sType);
      case 'p:cxnSp': // 连接形状、文本
        return this.processCxnSpNode(nodeValue, nodes, warpObj, source, sType);
      case 'p:pic': // 图片
        return this.processPicNode(nodeValue, warpObj, source, sType);
      case 'p:graphicFrame': // 图表、图形框架、表格
        return this.processGraphicFrameNode(nodeValue, warpObj, source, sType);
      case 'p:grpSp': // 组形状
        return this.processGroupSpNode(nodeValue, warpObj, source);
      case 'mc:AlternateContent': // 公式和公式作为图片
        const mcFallbackNode = _getTextByPathList(nodeValue, ['mc:Fallback']);
        return this.processGroupSpNode(mcFallbackNode, warpObj, source);
      default:
        return null;
    }
  }

  /**
   * 处理形状节点 - 对齐PPTXjs的processSpNode函数（第891-956行）
   */
  private processSpNode(
    node: any,
    pNode: any,
    warpObj: WarpObj,
    source: string,
    sType?: string
  ): any {
    // 获取ID、名称、索引、类型
    const id = _getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'id']);
    const name = _getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'name']);
    const idx = _getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx']);
    const type = _getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
    const order = _getTextByPathList(node, ['attrs', 'order']);

    // 查找layout和master中的对应节点
    let slideLayoutSpNode: any = undefined;
    let slideMasterSpNode: any = undefined;

    if (idx !== undefined) {
      slideLayoutSpNode = warpObj.slideLayoutTables.idxTable[idx];
      if (type !== undefined) {
        slideMasterSpNode = warpObj.slideMasterTables.typeTable[type];
      } else {
        slideMasterSpNode = warpObj.slideMasterTables.idxTable[idx];
      }
    } else {
      if (type !== undefined) {
        slideLayoutSpNode = warpObj.slideLayoutTables.typeTable[type];
        slideMasterSpNode = warpObj.slideMasterTables.typeTable[type];
      }
    }

    // 检查是否为文本框
    let finalType = type;
    if (finalType === undefined) {
      const txBoxVal = _getTextByPathList(node, ['p:nvSpPr', 'p:cNvSpPr', 'attrs', 'txBox']);
      if (txBoxVal === '1') {
        finalType = 'textBox';
      }
    }

    if (finalType === undefined) {
      finalType = _getTextByPathList(slideLayoutSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
      if (finalType === undefined) {
        if (source === 'diagramBg') {
          finalType = 'diagram';
        } else {
          finalType = 'obj';
        }
      }
    }

    // 生成形状（将在后续实现中完成）
    return {
      type: 'shape',
      id,
      name,
      idx,
      shapeType: finalType,
      order,
      // 其他属性将在后续添加
    };
  }

  /**
   * 处理连接形状节点 - 对齐PPTXjs的processCxnSpNode函数（第958-968行）
   */
  private processCxnSpNode(
    node: any,
    pNode: any,
    warpObj: WarpObj,
    source: string,
    sType?: string
  ): any {
    const id = _getTextByPathList(node, ['p:nvCxnSpPr', 'p:cNvPr', 'attrs', 'id']);
    const name = _getTextByPathList(node, ['p:nvCxnSpPr', 'p:cNvPr', 'attrs', 'name']);
    const idx = _getTextByPathList(node, ['p:nvCxnSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx']);
    const type = _getTextByPathList(node, ['p:nvCxnSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
    const order = _getTextByPathList(node, ['attrs', 'order']);

    return {
      type: 'shape',
      id,
      name,
      idx,
      shapeType: type,
      order,
      isConnector: true,
    };
  }

  /**
   * 处理图片节点 - 对齐PPTXjs的processPicNode函数
   */
  private processPicNode(
    node: any,
    warpObj: WarpObj,
    source: string,
    sType?: string
  ): any {
    const id = _getTextByPathList(node, ['p:nvPicPr', 'p:cNvPr', 'attrs', 'id']);
    const name = _getTextByPathList(node, ['p:nvPicPr', 'p:cNvPr', 'attrs', 'name']);
    const order = _getTextByPathList(node, ['attrs', 'order']);

    // 获取图片关系ID
    const blipEmbed = _getTextByPathList(node, ['p:blipFill', 'a:blip', 'attrs', 'r:embed']);

    // 从关系对象中获取图片路径
    let imagePath = '';
    if (blipEmbed && warpObj.slideResObj[blipEmbed]) {
      imagePath = warpObj.slideResObj[blipEmbed].target;
    }

    // 获取图片尺寸和位置
    const xfrm = getTextByPathList(node, ['p:spPr', 'a:xfrm']);
    let position = { x: 0, y: 0 };
    let size = { width: 0, height: 0 };

    if (xfrm) {
      const off = xfrm['a:off']?.attrs;
      const ext = xfrm['a:ext']?.attrs;
      
      if (off) {
        position = {
          x: parseInt(off.x),
          y: parseInt(off.y),
        };
      }
      if (ext) {
        size = {
          width: parseInt(ext.cx),
          height: parseInt(ext.cy),
        };
      }
    }

    // 读取图片数据并转为base64
    let imageData = '';
    if (imagePath) {
      imageData = getImageBase64(warpObj.zip, imagePath) || '';
    }

    const mimeType = getImageMimeType(imagePath);

    return {
      type: 'image',
      id,
      name,
      order,
      position,
      size,
      imagePath,
      data: imageData,
      mimeType,
      src: imageData ? generateDataUrl(imageData, mimeType) : '',
    };
  }

  /**
   * 处理图形框架节点 - 对齐PPTXjs的processGraphicFrameNode函数
   */
  private processGraphicFrameNode(
    node: any,
    warpObj: WarpObj,
    source: string,
    sType?: string
  ): any {
    const graphic = node['a:graphic'];
    if (!graphic) return null;

    // 检查是否为表格
    const table = graphic['a:graphicData']['a:tbl'];
    if (table) {
      return this.processTableNode(node, table, warpObj, source, sType);
    }

    // 检查是否为图表
    const chart = graphic['a:graphicData']['c:chart'];
    if (chart) {
      return this.processChartNode(node, chart, warpObj, source, sType);
    }

    return null;
  }

  /**
   * 处理表格节点
   */
  private processTableNode(
    node: any,
    table: any,
    warpObj: WarpObj,
    source: string,
    sType?: string
  ): any {
    const id = _getTextByPathList(node, ['p:nvGraphicFramePr', 'p:cNvPr', 'attrs', 'id']);
    const name = _getTextByPathList(node, ['p:nvGraphicFramePr', 'p:cNvPr', 'attrs', 'name']);
    const order = _getTextByPathList(node, ['attrs', 'order']);

    // 获取表格尺寸和位置
    const xfrm = getTextByPathList(node, ['p:xfrm', 'a:xfrm']);
    let position = { x: 0, y: 0 };
    let size = { width: 0, height: 0 };

    if (xfrm) {
      const off = xfrm['a:off']?.attrs;
      const ext = xfrm['a:ext']?.attrs;
      
      if (off) {
        position = {
          x: parseInt(off.x),
          y: parseInt(off.y),
        };
      }
      if (ext) {
        size = {
          width: parseInt(ext.cx),
          height: parseInt(ext.cy),
        };
      }
    }

    // 解析表格数据（将在后续实现中完成）
    const rows: any[] = [];
    const trNodes = table['a:tr'];
    const trArray = Array.isArray(trNodes) ? trNodes : [trNodes];

    for (const tr of trArray) {
      const row: any = { cells: [] };
      const tcNodes = tr['a:tc'];
      const tcArray = Array.isArray(tcNodes) ? tcNodes : [tcNodes];

      for (const tc of tcArray) {
        const cell: any = {};
        // 解析单元格内容
        const txBody = tc['a:txBody'];
        if (txBody) {
          const paragraphs = parseTextBoxContent(txBody);
          cell.content = paragraphs;
        }
        row.cells.push(cell);
      }
      rows.push(row);
    }

    return {
      type: 'table',
      id,
      name,
      order,
      position,
      size,
      rows,
    };
  }

  /**
   * 处理图表节点
   */
  private processChartNode(
    node: any,
    chart: any,
    warpObj: WarpObj,
    source: string,
    sType?: string
  ): any {
    const id = _getTextByPathList(node, ['p:nvGraphicFramePr', 'p:cNvPr', 'attrs', 'id']);
    const name = _getTextByPathList(node, ['p:nvGraphicFramePr', 'p:cNvPr', 'attrs', 'name']);
    const order = _getTextByPathList(node, ['attrs', 'order']);

    // 获取图表尺寸和位置
    const xfrm = getTextByPathList(node, ['p:xfrm', 'a:xfrm']);
    let position = { x: 0, y: 0 };
    let size = { width: 0, height: 0 };

    if (xfrm) {
      const off = xfrm['a:off']?.attrs;
      const ext = xfrm['a:ext']?.attrs;
      
      if (off) {
        position = {
          x: parseInt(off.x),
          y: parseInt(off.y),
        };
      }
      if (ext) {
        size = {
          width: parseInt(ext.cx),
          height: parseInt(ext.cy),
        };
      }
    }

    return {
      type: 'chart',
      id,
      name,
      order,
      position,
      size,
      chartData: chart,
    };
  }

  /**
   * 处理组形状节点 - 对齐PPTXjs的processGroupSpNode函数（第813-889行）
   */
  private processGroupSpNode(
    node: any,
    warpObj: WarpObj,
    source: string
  ): any {
    // 获取组形状的变换信息
    const xfrmNode = _getTextByPathList(node, ['p:grpSpPr', 'a:xfrm']);
    
    if (!xfrmNode) return null;

    const x = parseInt(xfrmNode['a:off']?.attrs.x || '0');
    const y = parseInt(xfrmNode['a:off']?.attrs.y || '0');
    const chx = parseInt(xfrmNode['a:chOff']?.attrs.x || '0');
    const chy = parseInt(xfrmNode['a:chOff']?.attrs.y || '0');
    const cx = parseInt(xfrmNode['a:ext']?.attrs.cx || '0');
    const cy = parseInt(xfrmNode['a:ext']?.attrs.cy || '0');
    const chcx = parseInt(xfrmNode['a:chExt']?.attrs.cx || '0');
    const chcy = parseInt(xfrmNode['a:chExt']?.attrs.cy || '0');

    let rotate = parseInt(xfrmNode.attrs?.rot || '0');
    rotate = angleToDegrees(rotate);

    let top = y - chy;
    let left = x - chx;
    let width = cx - chcx;
    let height = cy - chcy;

    // 处理旋转
    let sType = 'group';
    if (rotate !== 0) {
      top = y;
      left = x;
      width = cx;
      height = cy;
      sType = 'group-rotate';
    }

    const order = node.attrs?.order;

    const group: any = {
      type: 'group',
      order,
      position: { x: left, y: top },
      size: { width, height },
      rotate,
      sType,
      children: [],
    };

    // 处理子节点
    for (const nodeKey in node) {
      if (nodeKey === 'p:nvGrpSpPr' || nodeKey === 'p:grpSpPr') continue;

      const nodeArray = Array.isArray(node[nodeKey]) ? node[nodeKey] : [node[nodeKey]];
      
      for (const childNode of nodeArray) {
        const child = this.processNodesInSlide(nodeKey, childNode, node, warpObj, source, sType);
        if (child) {
          group.children.push(child);
        }
      }
    }

    return group;
  }

  /**
   * 获取背景信息 - 对齐PPTXjs的getBackground函数
   */
  private getBackground(warpObj: WarpObj, slideId: number): string {
    // 简化实现，将在后续完整实现
    return '';
  }

  /**
   * 获取幻灯片背景填充
   */
  private getSlideBackgroundFill(warpObj: WarpObj, slideId: number): string {
    // 简化实现，将在后续完整实现
    return '';
  }

  /**
   * 生成全局CSS - 对齐PPTXjs的genGlobalCSS函数
   */
  private generateGlobalCSS(): string {
    return `
.slide {
  position: relative;
  overflow: hidden;
  margin: 0;
  padding: 0;
}

.block {
  position: absolute;
}

.text-content {
  display: flex;
  word-wrap: break-word;
}
    `.trim();
  }
}

// 模块导出由 index.ts 处理
