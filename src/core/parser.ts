/**
 * PPTX 统一解析器
 * 合并 parsePptx 和 parsePptxEnhanced 功能
 * 支持基础解析和完整解析（包含元数据、图片、关系等）
 */

import JSZip from 'jszip';


import { generateId, log, emu2px, px2emu } from '../utils/index';
import { parseCoreProperties, parseSlideLayoutSize, inferPageSize } from './metadata-parser';
import { parseAllSlides } from './slide-parser';
import { parseGlobalRels, getSlideLayoutRef } from './relationships-parser';
import { parseImages } from './image-parser';
import { parseTheme, resolveSchemeColor } from './theme-parser';
import { parseAllMasterSlides } from './master-slide-parser';
import { parseAllSlideLayouts, mergeBackgrounds } from './layout-parser';
import { parseAllNotesMasters, parseAllNotesSlides, linkNotesToMasters } from './notes-parser';
import { parseAllCharts, parseAllDiagrams } from './drawings-parser';
import { parseAllSlideTags } from './tags-parser';
import { applyStyleInheritance } from './style-inheritance';
import type { PptxParseResult, ParseOptions, SlideLayoutResult } from './types';
import { unescape } from 'html-escaper';
import type { PptDocument } from '../types';

/**
 * PPT解析工具函数
 */
export const PptParseUtils = {
  /** 生成唯一ID */
  generateId,

  /** XML文本节点解析 - 处理XML转义字符 */
  parseXmlText: (text: string): string => unescape(text || '').trim(),

  /** 反向转换：前端PX转PPT的EMU单位（序列化时用） */
  px2emu,

  /** EMU转PX */
  emu2px,

  /** 解析XML字符串为树结构 */
  parseXmlToTree: (xmlStr: string): any => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlStr, 'application/xml');
    const root = doc.documentElement;
    const buildTree = (node: Element): any => {
      const children: any[] = [];
      Array.from(node.childNodes).forEach(child => {
        if (child.nodeType === 1) children.push(buildTree(child as Element));
      });
      return {
        tag: node.tagName,
        attrs: PptParseUtils.parseXmlAttrs(node.attributes),
        children,
        text: PptParseUtils.parseXmlText(node.textContent || '')
      };
    };
    return buildTree(root);
  },

  /** 解析XML属性 */
  parseXmlAttrs: (attrs: NamedNodeMap): Record<string, string> => {
    const result: Record<string, string> = {};
    Array.from(attrs).forEach(attr => {
      result[attr.nodeName] = attr.nodeValue || '';
    });
    return result;
  },

  /** 解析矩形属性 */
  parseXmlRect: (attrs: Record<string, string>): { x: number; y: number; width: number; height: number } => {
    return {
      x: emu2px(attrs['x'] || '0'),
      y: emu2px(attrs['y'] || '0'),
      width: emu2px(attrs['cx'] || '0'),
      height: emu2px(attrs['cy'] || '0'),
    };
  },

  /** 解析样式属性 */
  parseXmlStyle: (attrs: Record<string, string>): any => {
    const style: any = {};
    
    // 字体大小（可能以1/100 pt为单位）
    if (attrs.fontSize) {
      style.fontSize = parseInt(attrs.fontSize) / 100;
    } else {
      style.fontSize = 14;
    }
    
    // 颜色（fill属性）
    if (attrs.fill) {
      style.color = attrs.fill;
    } else {
      style.color = '#333333';
    }
    
    // 字体粗细
    if (attrs.bold === '1' || attrs.bold === 'true') {
      style.fontWeight = 'bold';
    } else {
      style.fontWeight = 'normal';
    }
    
    // 文本对齐
    const align = attrs.align;
    if (align === 'left' || align === 'center' || align === 'right' || align === 'justify') {
      style.textAlign = align;
    } else {
      style.textAlign = 'left';
    }
    
    // 背景色
    if (attrs.bgFill) {
      style.backgroundColor = attrs.bgFill;
    } else {
      style.backgroundColor = 'transparent';
    }
    
    // 边框颜色
    if (attrs.border) {
      style.borderColor = attrs.border;
    } else {
      style.borderColor = '#000000';
    }
    
    // 边框宽度
    if (attrs.borderWidth) {
      style.borderWidth = parseInt(attrs.borderWidth);
    } else {
      style.borderWidth = 1;
    }
    
    return style;
  },

  /** 十六进制颜色转RGB */
  hexToRgb: (hex: string): { r: number; g: number; b: number } => {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
      r: parseInt(result[1], 16),
      g: parseInt(result[2], 16),
      b: parseInt(result[3], 16)
    } : { r: 0, g: 0, b: 0 };
  },

  /** RGB转十六进制颜色 */
  rgbToHex: (r: number, g: number, b: number): string => {
    return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
  },

  /** 解析颜色值 */
  parseColor: (color: string): string => {
    // 简单实现：如果是十六进制格式则返回，否则尝试解析
    if (color.startsWith('#')) {
      return color;
    }
    // 其他颜色格式（如rgb(), named colors）可以在这里扩展
    return color;
  }
};

/**
 * 统一的PPTX解析函数
 * @param file PPTX文件（File | Blob | ArrayBuffer）
 * @param options 解析选项
 * @returns 解析结果对象
 */
export async function parsePptx(
  file: File | Blob | ArrayBuffer,
  options?: ParseOptions & { returnFormat?: 'enhanced' | 'simple' }
): Promise<PptxParseResult | PptDocument> {
  const opts = {
    parseImages: true,
    keepRawXml: false,
    verbose: false,
    returnFormat: 'enhanced' as const,
    ...options
  };

  const zip = await JSZip.loadAsync(file);

  // 简单模式（返回 PptDocument）
  if (opts.returnFormat === 'simple') {
    return parseSimpleMode(zip, opts);
  }

  // 增强模式（返回 PptxParseResult，默认）
  return parseEnhancedMode(zip, opts);
}

/**
 * 简单解析模式 - 返回 PptDocument
 */
async function parseSimpleMode(
  zip: JSZip,
  opts: ParseOptions
): Promise<PptDocument> {
  const pptProps = await parsePptProperties(zip);
  const xmlSlides = await parseXmlSlides(zip);
  const slides = await Promise.all(xmlSlides.map((slide, index) => parseSlide(zip, slide, index)));

  return {
    id: PptParseUtils.generateId('ppt-doc'),
    title: pptProps.title || '未命名PPT',
    slides,
    props: {
      width: pptProps.width || 1280,
      height: pptProps.height || 720,
      ratio: (pptProps.width || 1280) / (pptProps.height || 720),
    },
  };
}

/**
 * 增强解析模式 - 返回 PptxParseResult（包含完整信息）
 */
async function parseEnhancedMode(
  zip: JSZip,
  opts: ParseOptions
): Promise<PptxParseResult> {
  try {
    log('info', 'Starting PPTX parsing...');

    // 解析元数据（docProps/core.xml）
    const metadata = await parseCoreProperties(zip);

    // 解析幻灯片尺寸
    const slideSize = await parseSlideLayoutSize(zip);

    // 解析主题文件（第一步：PPTXjs 中 theme 是基础）
    const theme = await parseTheme(zip);
    log('info', theme ? 'Theme parsed successfully' : 'Theme not found');

    // 解析所有幻灯片母版（第二步：master 引用 theme）
    const masterSlides = await parseAllMasterSlides(zip);
    log('info', `Parsed ${masterSlides.length} master slides`);

    // 建立从 masterId 到 master 对象的映射（关键：使用 masterId 而不是 themeRef）
    const masterMap = new Map<string, any>();
    masterSlides.forEach(master => {
      if (master.masterId) {
        masterMap.set(master.masterId, master);
        log('info', `Master map: ${master.masterId} -> ${master.id}`);
      }
    });

    // 解析所有幻灯片布局（第三步：layout 引用 master）
    const slideLayouts = await parseAllSlideLayouts(zip);
    log('info', `Parsed ${Object.keys(slideLayouts).length} slide layouts`);

    // 将 master 对象附加到 layout 上
    Object.entries(slideLayouts).forEach(([layoutId, layout]) => {
      if (layout.masterRef && masterMap.has(layout.masterRef)) {
        const layoutAny = layout as any;
        layoutAny.master = masterMap.get(layout.masterRef);
        log('info', `Layout ${layoutId} references master: ${layout.masterRef}`);
      }
    });

    // 解析所有幻灯片（第四步：slide 引用 layout）
    const slides = await parseAllSlides(zip, opts);

    // 合并背景和样式信息：slide > layout > master（核心：PPTXjs 样式优先级）
    slides.forEach((slide) => {
      // 获取slide的关联关系
      const relsMap = slide.relsMap || {};

      // 使用工具函数获取布局引用ID
      const layoutId = getSlideLayoutRef(relsMap);

      if (layoutId && slideLayouts[layoutId]) {
        const layout = slideLayouts[layoutId];
        const layoutAny = layout as any;

        // 合并背景
        const mergedBg = mergeBackgrounds(
          slide.background as any,
          layout.background,
          layoutAny.master?.background
        );
        slide.background = mergedBg;

        // 将布局ID和占位符信息附加到幻灯片对象（供渲染使用）
        (slide as any).layoutId = layoutId;
        (slide as any).layout = layout;
        (slide as any).master = layoutAny.master;

        // 应用样式继承到所有元素
        applyStyleInheritance(slide, layout, layoutAny.master, theme);

        const masterRef = (layout as any).masterRef;
        log('info', `Slide relationship chain: slide -> layout (${layoutId}) -> master (${masterRef})`);
      } else if (layoutId) {
        log('warn', `Layout ${layoutId} referenced by slide not found in parsed layouts`);
      }
    });

    // 解析全局关联关系
    const globalRels = await parseGlobalRels(zip);

    log('info', `Parsed ${slides.length} slides successfully`);

    // 计算页面比例
    const ratio = slideSize.width / slideSize.height;
    const pageSize = inferPageSize(ratio);

    // 解析背景：解析主题颜色到实际颜色
    if (theme && theme.colors) {
      // 解析 master 背景的 schemeRef
      masterSlides.forEach(master => {
        if (master.background && typeof master.background === 'object') {
          const bg = master.background as any;
          if (bg.schemeRef) {
            const actualColor = resolveSchemeColor(
              bg.schemeRef,
              theme.colors,
              master.colorMap || {}
            );
            bg.value = actualColor;
            delete bg.schemeRef;
          }
        }
      });

      // 解析 layout 背景的 schemeRef
      Object.values(slideLayouts).forEach(layout => {
        if (layout.background && typeof layout.background === 'object') {
          const bg = layout.background as any;
          if (bg.schemeRef) {
            const actualColor = resolveSchemeColor(
              bg.schemeRef,
              theme.colors,
              layout.colorMap || {}
            );
            bg.value = actualColor;
            delete bg.schemeRef;
          }
        }
      });

      // 解析 slide 背景的 schemeRef
      slides.forEach((slide, index) => {
        if (slide.background && typeof slide.background === 'object') {
          const bg = slide.background as any;
          // 如果是方案颜色引用，解析实际颜色
          if (bg.schemeRef) {
            const actualColor = resolveSchemeColor(
              bg.schemeRef,
              theme.colors,
              {} // 可以添加slide的colorMap override
            );
            bg.value = actualColor;
            delete bg.schemeRef;
          }
        }
      });
    }

    // 解析备注母版和备注页
    const notesMasters = await parseAllNotesMasters(zip);
    const notesSlides = await parseAllNotesSlides(zip);
    linkNotesToMasters(notesSlides, notesMasters);

    // 解析图表和SmartArt
    const charts = await parseAllCharts(zip);
    const diagrams = await parseAllDiagrams(zip);

    // 解析幻灯片标签和扩展属性
    const tags = await parseAllSlideTags(zip, slides.length);

    const result: PptxParseResult = {
      id: generateId('ppt-doc'),
      title: metadata.title || '未命名PPT',
      author: metadata.author,
      subject: metadata.subject,
      keywords: metadata.keywords,
      description: metadata.description,
      created: metadata.created,
      modified: metadata.modified,
      slides,
      props: {
        width: slideSize.width,
        height: slideSize.height,
        ratio,
        pageSize
      },
      globalRelsMap: globalRels,
      theme: theme || undefined,
      masterSlides,
      slideLayouts,
      notesMasters: notesMasters.length > 0 ? notesMasters : undefined,
      notesSlides: notesSlides.length > 0 ? notesSlides : undefined,
      charts: charts.length > 0 ? charts : undefined,
      diagrams: diagrams.length > 0 ? diagrams : undefined,
      tags: tags.length > 0 ? tags : undefined
    };

    // 解析图片（如果需要）
    if (opts.parseImages) {
      await parseImages(zip, result);
    }

    return result;
  } catch (error) {
    log('error', 'PPTX parsing failed', error);
    throw new Error(`PPTX解析失败: ${error instanceof Error ? error.message : String(error)}`);
  }
}

/**
 * 解析PPT属性
 */
async function parsePptProperties(zip: JSZip): Promise<{ width: number; height: number; title: string }> {
  const props = { width: 1280, height: 720, title: '未命名PPT' };
  try {
    const slideLayoutXml = await zip.file('ppt/slideLayouts/slideLayout1.xml')?.async('string');
    if (slideLayoutXml) {
      const xmlTree = PptParseUtils.parseXmlToTree(slideLayoutXml);
    const cSld = xmlTree.children.find((node: any) => node.tag === 'cSld');
    const sldSz = cSld?.children.find((node: any) => node.tag === 'sldSz');
    if (sldSz) {
      const rect = PptParseUtils.parseXmlRect(sldSz.attrs);
      props.width = rect.width;
      props.height = rect.height;
    }
    }
    const coreProps = await zip.file('docProps/core.xml')?.async('string');
    if (coreProps) {
      const xmlTree = PptParseUtils.parseXmlToTree(coreProps);
        props.title = xmlTree.children.find((node: any) => node.tag.includes('title'))?.text || '未命名PPT';
    }
  } catch (e) {
    console.warn('PPT配置解析失败，使用默认配置', e);
  }
  return props;
}

/**
 * 序列化函数（从原 core.ts 迁移）
 */
export async function serializePptx(pptDoc: PptDocument): Promise<Blob> {
  const zip = new JSZip();
  createPptxStructure(zip);
  await writePptProperties(zip, pptDoc);
  await writeSlides(zip, pptDoc.slides);
  return await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });

  function createPptxStructure(zip: JSZip): void {
    zip.folder('ppt')?.folder('slides');
    zip.folder('ppt')?.folder('slideLayouts');
    zip.folder('ppt')?.folder('_rels');
    zip.folder('docProps');
    zip.folder('_rels');
  }

  async function writePptProperties(zip: JSZip, pptDoc: PptDocument): Promise<void> {
    const { width, height } = pptDoc.props;
    const emuWidth = PptParseUtils.px2emu(width);
    const emuHeight = PptParseUtils.px2emu(height);

    const slideLayoutXml = `<sldLayout xmlns="http://schemas.openxmlformats.org/presentationml/2006/main"><cSld><sldSz cx="${emuWidth}" cy="${emuHeight}" type="screen"/></cSld></sldLayout>`;
    zip.file('ppt/slideLayouts/slideLayout1.xml', slideLayoutXml);

    const corePropsXml = `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><dc:title>${pptDoc.title}</dc:title></cp:coreProperties>`;
    zip.file('docProps/core.xml', corePropsXml);
  }
}

// 序列化辅助函数移到外部
export async function writeSlides(zip: JSZip, slides: any[]): Promise<void> {
    slides.forEach(async (slide, index) => {
      const slideXml = slideToXml(slide);
      zip.file(`ppt/slides/slide${index + 1}.xml`, slideXml);
    });
  }

  function slideToXml(slide: any): string {
    const { elements, bgColor, props } = slide;
    const emuWidth = PptParseUtils.px2emu(props.width);
    const emuHeight = PptParseUtils.px2emu(props.height);
    const elementsXml = elements.map((el: any) => elementToXml(el)).join('\n');

    return `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:grpSp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emuWidth}" cy="${emuHeight}"/></a:xfrm></p:spPr>${elementsXml}</p:grpSp></p:spTree></p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bgColor.replace('#', '')}"/></a:solidFill></p:bgPr></p:bg></p:sld>`
      .replace(/\s+/g, ' ').trim();
  }

  function elementToXml(el: any): string {
    const { type, rect, style, content } = el;
    const { x, y, width, height } = rect;
    const emuX = PptParseUtils.px2emu(x);
    const emuY = PptParseUtils.px2emu(y);
    const emuW = PptParseUtils.px2emu(width);
    const emuH = PptParseUtils.px2emu(height);

    switch (type) {
      case 'text':
        return `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="${(style.fontSize ||14)*100}" b="${style.fontWeight==='bold'?1:0}"/><a:t>${content}</a:t></a:r></a:p></p:txBody><p:spPr><a:xfrm><a:off x="${emuX}" y="${emuY}"/><a:ext cx="${emuW}" cy="${emuH}"/></a:xfrm></p:spPr>`;
      case 'shape':
        return `<p:sp><p:spPr><a:xfrm><a:off x="${emuX}" y="${emuY}"/><a:ext cx="${emuW}" cy="${emuH}"/></a:xfrm><a:solidFill><a:srgbClr val="${style.backgroundColor?.replace('#','') || 'ffffff'}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${style.borderColor?.replace('#','') || '000000'}"/></a:solidFill></a:ln></p:spPr></p:sp>`;
      case 'image':
        return `<p:pic><p:blipFill><a:blip r:embed="rId1"/></p:blipFill><p:spPr><a:xfrm><a:off x="${emuX}" y="${emuY}"/><a:ext cx="${emuW}" cy="${emuH}"/></a:xfrm></p:spPr></p:pic>`;
      case 'table':
        const rows = content as string[][];
        const tableRows = rows.map(row => `<a:tr><a:tc><a:txBody><a:p><a:r><a:t>${row.join('')}</a:t></a:r></a:p></a:txBody></a:tc></a:tr>`).join('');
        return `<p:tbl><p:tblPr><a:tblW w="${emuW}" type="dxa"/></p:tblPr>${tableRows}</p:tbl>`;
      default:
        return '';
    }
  }

/**
 * 扩展工具函数
 */






// 简单模式需要的辅助函数
async function parseXmlSlides(zip: JSZip): Promise<any[]> {
  const slides: any[] = [];
  const slideFiles = Object.keys(zip.files).filter(path => path.startsWith('ppt/slides/slide') && path.endsWith('.xml'));
  for (const path of slideFiles) {
    const xml = await zip.file(path)!.async('string');
    slides.push({
      xml,
      slideId: PptParseUtils.generateId('slide'),
      rId: path.replace('ppt/slides/slide', '').replace('.xml', ''),
      layout: 'normal',
    });
  }
  return slides;
}

async function parseSlide(zip: JSZip, xmlSlide: any, slideIndex: number): Promise<any> {
  const { xml, layout } = xmlSlide;
  const xmlTree = PptParseUtils.parseXmlToTree(xml);

  const cSld = xmlTree.children.find((node: any) => node.tag === 'p:cSld' || node.tag === 'cSld');
  const spTree = cSld?.children.find((node: any) => node.tag === 'p:spTree' || node.tag === 'spTree');

  const elements: any[] = [];

  if (spTree) {
    spTree.children.forEach((node: any) => {
      const element = parseXmlElement(node);
      if (element) {
        elements.push(element);
      }
    });
  }

  const bgColor = parseSlideBgColor(xmlTree);
  const title = parseSlideTitle(xmlTree, slideIndex + 1);

  return {
    id: xmlSlide.slideId,
    title,
    bgColor,
    elements,
    props: {
      width: 1280,
      height: 720,
      slideLayout: layout,
    },
  };
}

function parseXmlElement(xmlNode: any): any | null {
  const { tag } = xmlNode;

  if (tag === 'p:sp' || tag === 'sp') {
    return parseShape(xmlNode);
  }

  if (tag === 'p:pic' || tag === 'pic') {
    return parsePicture(xmlNode);
  }

  if (tag === 'p:graphicFrame' || tag === 'graphicFrame') {
    return parseGraphicFrame(xmlNode);
  }

  return null;
}

function parseShape(xmlNode: any): any | null {
  const id = PptParseUtils.generateId('shape');
  const children = xmlNode.children;

  const spPr = children.find((node: any) => node.tag === 'p:spPr' || node.tag === 'spPr');
  const txBody = children.find((node: any) => node.tag === 'p:txBody' || node.tag === 'txBody');

  let rect = { x: 0, y: 0, width: 0, height: 0 };
  if (spPr) {
    const xfrm = spPr.children.find((node: any) => node.tag === 'a:xfrm');
    if (xfrm) {
      const off = xfrm.children.find((node: any) => node.tag === 'a:off');
        const ext = xfrm.children.find((node: any) => node.tag === 'a:ext');
        if (off && ext) {
          rect = PptParseUtils.parseXmlRect({
            x: off.attrs['x'] || '0',
            y: off.attrs['y'] || '0',
            cx: ext.attrs['cx'] || '0',
            cy: ext.attrs['cy'] || '0'
          });
        }
    }
  }

  let content: any = '';
  if (txBody) {
    const paragraphs = txBody.children.filter((node: any) => node.tag === 'a:p' || node.tag === 'p');
    const textRuns: any[] = [];

    paragraphs.forEach((p: any) => {
      const runs = p.children.filter((node: any) => node.tag === 'a:r' || node.tag === 'r');
        runs.forEach((r: any) => {
          const t = r.children.find((node: any) => node.tag === 'a:t' || node.tag === 't');
          if (t && t.text) {
            const rPr = r.children.find((node: any) => node.tag === 'a:rPr' || node.tag === 'rPr');
            textRuns.push({
              text: t.text,
              fontSize: rPr?.attrs['sz'] ? parseInt(rPr.attrs['sz']) / 100 : 14,
              fontColor: rPr?.attrs['fill'] || '#333333',
              bold: rPr?.attrs['b'] === '1',
              italic: rPr?.attrs['i'] === '1'
            });
          }
        });
    });

    content = {
      paragraphs: [{
        runs: textRuns
      }]
    };
  }

  return {
    id,
    type: 'text',
    rect,
    style: {
      fontSize: 14,
      color: '#333333',
      fontWeight: 'normal',
      textAlign: 'left' as const,
      backgroundColor: 'transparent',
      borderColor: '#000000',
      borderWidth: 1
    },
    content,
  };
}

function parsePicture(xmlNode: any): any | null {
  const id = PptParseUtils.generateId('image');
  const children = xmlNode.children;

  const spPr = children.find((node: any) => node.tag === 'p:spPr' || node.tag === 'spPr');
  const blipFill = children.find((node: any) => node.tag === 'p:blipFill' || node.tag === 'blipFill');

  let rect = { x: 0, y: 0, width: 0, height: 0 };
  if (spPr) {
    const xfrm = spPr.children.find((node: any) => node.tag === 'a:xfrm');
    if (xfrm) {
      const off = xfrm.children.find((node: any) => node.tag === 'a:off');
        const ext = xfrm.children.find((node: any) => node.tag === 'a:ext');
        if (off && ext) {
          rect = PptParseUtils.parseXmlRect({
            x: off.attrs['x'] || '0',
            y: off.attrs['y'] || '0',
            cx: ext.attrs['cx'] || '0',
            cy: ext.attrs['cy'] || '0'
          });
        }
    }
  }

  let imgId = '';
  if (blipFill) {
    const blip = blipFill.children.find((node: any) => node.tag === 'a:blip' || node.tag === 'blip');
    imgId = blip?.attrs['r:embed'] || blip?.attrs['embed'] || '';
  }

  return {
    id,
    type: 'image',
    rect,
    style: {
      fontSize: 14,
      color: '#333333',
      fontWeight: 'normal',
      textAlign: 'left' as const,
      backgroundColor: 'transparent',
      borderColor: '#000000',
      borderWidth: 1
    },
    content: {
      url: '',
      imgId
    },
    props: {
      imgId,
      alt: '图片'
    }
  };
}

function parseGraphicFrame(xmlNode: any): any | null {
  const id = PptParseUtils.generateId('graphic');
  const children = xmlNode.children;

  const xfrm = children.find((node: any) => node.tag === 'p:xfrm' || node.tag === 'xfrm');
  let rect = { x: 0, y: 0, width: 0, height: 0 };
  if (xfrm) {
    const off = xfrm.children.find((node: any) => node.tag === 'a:off');
    const ext = xfrm.children.find((node: any) => node.tag === 'a:ext');
    if (off && ext) {
      rect = PptParseUtils.parseXmlRect({
        x: off.attrs['x'] || '0',
        y: off.attrs['y'] || '0',
        cx: ext.attrs['cx'] || '0',
        cy: ext.attrs['cy'] || '0'
      });
    }
  }

  const graphic = children.find((node: any) => node.tag === 'a:graphic' || node.tag === 'graphic');
  const graphicData = graphic?.children.find((node: any) => node.tag === 'a:graphicData' || node.tag === 'graphicData');

  if (!graphicData) return null;

  const uri = graphicData.attrs['uri'] || '';

  if (uri.includes('table')) {
    return {
      id,
      type: 'table',
      rect,
      style: {
        fontSize: 14,
        color: '#333333',
        fontWeight: 'normal',
        textAlign: 'left' as const,
        backgroundColor: 'transparent',
        borderColor: '#cccccc',
        borderWidth: 1
      },
      content: {
        rows: []
      },
      props: {
        rowCount: 0,
        colCount: 0
      }
    };
  }

  if (uri.includes('chart')) {
    return {
      id,
      type: 'chart',
      rect,
      style: {
        fontSize: 14,
        color: '#333333',
        fontWeight: 'normal',
        textAlign: 'left' as const,
        backgroundColor: 'transparent',
        borderColor: '#000000',
        borderWidth: 1
      },
      content: {
        type: 'unknown',
        data: []
      },
      props: {
        legend: true,
        grid: true
      }
    };
  }

  return null;
}

function parseSlideTitle(xmlTree: any, defaultIndex: number): string {
  const cSld = xmlTree.children.find((node: any) => node.tag === 'p:cSld' || node.tag === 'cSld');
  const spTree = cSld?.children.find((node: any) => node.tag === 'p:spTree' || node.tag === 'spTree');

  if (!spTree) return `幻灯片${defaultIndex}`;

  for (const node of spTree.children) {
    if (node.tag === 'p:sp' || node.tag === 'sp') {
      const nvSpPr = node.children.find((n: any) => n.tag === 'p:nvSpPr' || n.tag === 'nvSpPr');
      const nvPr = nvSpPr?.children.find((n: any) => n.tag === 'p:nvPr' || n.tag === 'nvPr');
      const ph = nvPr?.children.find((n: any) => n.tag === 'p:ph' || n.tag === 'ph');

      if (ph?.attrs['type'] === 'title') {
        const txBody = node.children.find((n: any) => n.tag === 'p:txBody' || n.tag === 'txBody');
        const p = txBody?.children.find((n: any) => n.tag === 'a:p' || n.tag === 'p');
        const r = p?.children.find((n: any) => n.tag === 'a:r' || n.tag === 'r');
        const t = r?.children.find((n: any) => n.tag === 'a:t' || n.tag === 't');
        if (t?.text) return t.text;
      }
    }
  }

  return `幻灯片${defaultIndex}`;
}

function parseSlideBgColor(xmlTree: any): string {
  const bgNode = xmlTree.children.find((node: any) => node.tag === 'bg');
  if (!bgNode) return '#ffffff';
  const fillNode = bgNode.children.find((node: any) => node.tag === 'fill');
  return fillNode?.attrs['color'] || '#ffffff';
}
