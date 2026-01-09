/**
 * PPTXjs核心解析器 - TypeScript转译版
 * 原始版本: PPTXjs.js v1.21.1
 * 作者: meshesha
 * 许可: MIT
 * 
 * 完全转译PPTXjs的核心解析逻辑，包括：
 * - XML解析和zip处理
 * - 内容类型解析
 * - 幻灯片尺寸计算
 * - 单位转换系统
 */

import JSZip from 'jszip';

/**
 * PPTXjs核心常量 - 对齐原始PPTXjs实现
 */
export const PPTXJS_CONSTANTS = {
  // PPTXjs核心转换因子（完全对齐原始实现）
  slideFactor: 96 / 914400,      // EMU → PX转换因子
  fontSizeFactor: 4 / 3.2,       // 字体大小转换因子
  
  // RTL语言数组
  rtlLangs: ['he-IL', 'ar-AE', 'ar-SA', 'dv-MV', 'fa-IR', 'ur-PK'],
  
  // 标准尺寸（EMU单位）
  standardHeight: 6858000,        // 标准高度
  standardWidth: 9144000,         // 标准宽度
} as const;

/**
 * WarpObj上下文对象 - PPTXjs核心上下文
 * 对齐PPTXjs第675-691行的warpObj结构
 */
export interface WarpObj {
  zip: JSZip;
  slideLayoutContent: any;
  slideLayoutTables: any;
  slideMasterContent: any;
  slideMasterTables: any;
  slideContent: any;
  slideResObj: any;
  slideMasterTextStyles: any;
  layoutResObj: any;
  masterResObj: any;
  themeContent: any;
  themeResObj: any;
  digramFileContent?: any;
  diagramResObj?: any;
  defaultTextStyle: any;
}

/**
 * 关系对象 - PPTXjs关系结构
 */
export interface Relationship {
  id: string;
  type: string;
  target: string;
}

/**
 * 内容类型信息 - PPTXjs内容类型结构
 */
export interface ContentTypes {
  slides: string[];
  slideLayouts: string[];
}

/**
 * 幻灯片尺寸信息
 */
export interface SlideSize {
  width: number;
  height: number;
}

/**
 * 索引表 - PPTXjs的节点索引系统
 * 对齐PPTXjs第725-779行的indexNodes函数
 */
export interface IndexTable {
  idTable: Record<string, any>;
  idxTable: Record<string, any>;
  typeTable: Record<string, any>;
}

/**
 * PPTXjs核心解析器类
 * 转译自PPTXjs.js的核心解析逻辑
 */
export class PptxjsCoreParser {
  private zip: JSZip;
  private slideWidth: number = 0;
  private slideHeight: number = 0;
  private slideFactor: number = PPTXJS_CONSTANTS.slideFactor;
  private fontSizeFactor: number = PPTXJS_CONSTANTS.fontSizeFactor;
  private defaultTextStyle: any = null;
  private appVersion: number = 0;
  private processFullTheme: boolean = true;
  private incSlide: { width: number; height: number } = { width: 0, height: 0 };

  constructor(zip: JSZip, options: {
    processFullTheme?: boolean;
    incSlideWidth?: number;
    incSlideHeight?: number;
  } = {}) {
    this.zip = zip;
    this.processFullTheme = options.processFullTheme !== false;
    this.incSlide = {
      width: options.incSlideWidth || 0,
      height: options.incSlideHeight || 0,
    };
  }

  /**
   * 读取XML文件 - 对齐PPTXjs第396-415行
   * @param filename XML文件路径
   * @param isSlideContent 是否为幻灯片内容（处理CDATA）
   */
  readXmlFile(filename: string, isSlideContent = false): any {
    try {
      const file = this.zip.file(filename);
      if (!file) {
        return null;
      }

      let fileContent = file.asText();
      
      // 处理CDATA标签（Office 2007及以下版本）
      if (isSlideContent && this.appVersion <= 12) {
        fileContent = fileContent.replace(/<!\[CDATA\[(.*?)\]\]>/g, '$1');
      }

      // 使用tXml解析XML
      // 注意：这里需要确保项目已安装@trivago/txtml或类似库
      // 暂时返回简化后的JSON结构
      const xmlData = this.parseXml(fileContent);
      
      if (xmlData && xmlData['?xml'] !== undefined) {
        return xmlData['?xml'];
      }
      return xmlData;
    } catch (e) {
      console.error(`Error reading XML file '${filename}':`, e);
      return null;
    }
  }

  /**
   * 简单的XML解析器（替代tXml）
   */
  private parseXml(xmlString: string): any {
    // 这里是一个简化的XML解析器
    // 在实际项目中应该使用完整的XML解析库
    try {
      // 如果项目中有xml2js或类似库，使用它
      // 否则返回一个简化的结构
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlString, 'text/xml');
      return this.domToJson(xmlDoc);
    } catch (e) {
      console.error('XML parsing error:', e);
      return null;
    }
  }

  /**
   * DOM转JSON - 辅助方法
   */
  private domToJson(node: Node): any {
    if (node.nodeType === Node.TEXT_NODE) {
      return node.textContent;
    }

    const obj: any = {};
    
    // 处理属性
    if (node.nodeType === Node.ELEMENT_NODE && node instanceof Element) {
      const attrs: Record<string, string> = {};
      for (let i = 0; i < node.attributes.length; i++) {
        attrs[node.attributes[i].name] = node.attributes[i].value;
      }
      if (Object.keys(attrs).length > 0) {
        obj['attrs'] = attrs;
      }
    }

    // 处理子节点
    const children = node.childNodes;
    const childMap: Record<string, any[]> = {};

    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      if (child.nodeType === Node.TEXT_NODE) {
        const text = child.textContent?.trim();
        if (text) {
          // 文本内容直接添加到对象
          obj['#text'] = text;
        }
        continue;
      }

      if (child.nodeType === Node.ELEMENT_NODE) {
        const childJson = this.domToJson(child);
        const childName = child.nodeName;
        
        if (!childMap[childName]) {
          childMap[childName] = [];
        }
        childMap[childName].push(childJson);
      }
    }

    // 合并子节点
    for (const [key, values] of Object.entries(childMap)) {
      if (values.length === 1) {
        obj[key] = values[0];
      } else {
        obj[key] = values;
      }
    }

    return obj;
  }

  /**
   * 获取内容类型 - 对齐PPTXjs第416-437行
   */
  getContentTypes(): ContentTypes {
    const contentTypesJson = this.readXmlFile('[Content_Types].xml');
    
    if (!contentTypesJson || !contentTypesJson.Types) {
      return { slides: [], slideLayouts: [] };
    }

    const overrides = contentTypesJson.Types.Override;
    const slidesLocArray: string[] = [];
    const slideLayoutsLocArray: string[] = [];

    const overrideArray = Array.isArray(overrides) ? overrides : [overrides];

    for (const override of overrideArray) {
      const attrs = override.attrs;
      if (!attrs) continue;

      switch (attrs.ContentType) {
        case 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml':
          slidesLocArray.push(attrs.PartName.substring(1));
          break;
        case 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml':
          slideLayoutsLocArray.push(attrs.PartName.substring(1));
          break;
      }
    }

    return {
      slides: slidesLocArray,
      slideLayouts: slideLayoutsLocArray,
    };
  }

  /**
   * 获取幻灯片尺寸并设置默认文本样式 - 对齐PPTXjs第439-498行
   */
  getSlideSizeAndSetDefaultTextStyle(): SlideSize {
    // 获取应用版本
    const app = this.readXmlFile('docProps/app.xml');
    if (app && app.Properties && app.Properties.AppVersion) {
      this.appVersion = parseInt(app.Properties.AppVersion);
      console.log(`Created by Office PowerPoint app version: ${app.Properties.AppVersion}`);
    }

    // 获取幻灯片尺寸
    const content = this.readXmlFile('ppt/presentation.xml');
    if (!content || !content['p:presentation']) {
      return { width: 960, height: 720 };
    }

    const sldSzAttrs = content['p:presentation']['p:sldSz']?.attrs;
    if (!sldSzAttrs) {
      return { width: 960, height: 720 };
    }

    const sldSzWidth = parseInt(sldSzAttrs.cx);
    const sldSzHeight = parseInt(sldSzAttrs.cy);
    const sldSzType = sldSzAttrs.type;
    
    console.log(`Presentation size type: ${sldSzType}`);

    // 1英寸 = 96px = 2.54cm
    // 1 EMU = 1 / 914400 英寸
    // Pixel = EMUs * Resolution / 914400 (Resolution = 96)
    
    // 获取默认文本样式
    this.defaultTextStyle = content['p:presentation']['p:defaultTextStyle'];

    // 计算幻灯片宽高（对齐PPTXjs第491-492行）
    this.slideWidth = Math.floor(sldSzWidth * this.slideFactor + this.incSlide.width);
    this.slideHeight = Math.floor(sldSzHeight * this.slideFactor + this.incSlide.height);

    return {
      width: this.slideWidth,
      height: this.slideHeight,
    };
  }

  /**
   * 索引节点 - 对齐PPTXjs第725-779行
   * 创建idTable、idxTable、typeTable用于形状查找
   */
  indexNodes(content: any): IndexTable {
    if (!content) {
      return { idTable: {}, idxTable: {}, typeTable: {} };
    }

    const keys = Object.keys(content);
    const rootKey = keys[0];
    
    if (!content[rootKey]) {
      return { idTable: {}, idxTable: {}, typeTable: {} };
    }

    const cSld = content[rootKey]['p:cSld'];
    if (!cSld) {
      return { idTable: {}, idxTable: {}, typeTable: {} };
    }

    const spTreeNode = cSld['p:spTree'];
    if (!spTreeNode) {
      return { idTable: {}, idxTable: {}, typeTable: {} };
    }

    const idTable: Record<string, any> = {};
    const idxTable: Record<string, any> = {};
    const typeTable: Record<string, any> = {};

    for (const key in spTreeNode) {
      if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') {
        continue;
      }

      const targetNode = spTreeNode[key];
      const nodeArray = Array.isArray(targetNode) ? targetNode : [targetNode];

      for (const node of nodeArray) {
        const nvSpPrNode = node['p:nvSpPr'];
        if (!nvSpPrNode) continue;

        const cNvPr = nvSpPrNode['p:cNvPr'];
        const nvPr = nvSpPrNode['p:nvPr'];
        const ph = nvPr?.['p:ph'];

        const id = cNvPr?.attrs?.id;
        const idx = ph?.attrs?.idx;
        const type = ph?.attrs?.type;

        if (id !== undefined) {
          idTable[id] = node;
        }
        if (idx !== undefined) {
          idxTable[idx] = node;
        }
        if (type !== undefined) {
          typeTable[type] = node;
        }
      }
    }

    return { idTable, idxTable, typeTable };
  }

  /**
   * 获取路径文本值 - 对齐PPTXjs的getTextByPathList函数
   */
  getTextByPathList(obj: any, pathList: string[]): any {
    if (!obj || !pathList || pathList.length === 0) {
      return undefined;
    }

    let current = obj;
    for (const path of pathList) {
      if (current === undefined || current === null) {
        return undefined;
      }
      current = current[path];
    }

    return current;
  }

  /**
   * 获取幻灯片宽度
   */
  getSlideWidth(): number {
    return this.slideWidth;
  }

  /**
   * 获取幻灯片高度
   */
  getSlideHeight(): number {
    return this.slideHeight;
  }

  /**
   * 获取slide因子
   */
  getSlideFactor(): number {
    return this.slideFactor;
  }

  /**
   * 获取字体大小因子
   */
  getFontSizeFactor(): number {
    return this.fontSizeFactor;
  }

  /**
   * 获取默认文本样式
   */
  getDefaultTextStyle(): any {
    return this.defaultTextStyle;
  }

  /**
   * 获取应用版本
   */
  getAppVersion(): number {
    return this.appVersion;
  }

  /**
   * 是否处理完整主题
   */
  getProcessFullTheme(): boolean {
    return this.processFullTheme;
  }
}

/**
 * 角度转度数 - 对齐PPTXjs的角度转换逻辑
 */
export function angleToDegrees(angle: number | undefined): number {
  if (angle === undefined || angle === null) {
    return 0;
  }
  // PPTX中的角度单位是1/60000度
  return angle / 60000;
}

/**
 * 度数转弧度
 */
export function degreesToRadians(degrees: number): number {
  return degrees * (Math.PI / 180);
}

/**
 * 导出所有核心工具函数
 */
export * from './pptxjs-utils';
