import JSZip from 'jszip';
import { unescape } from 'html-escaper';
import type { PptDocument, PptSlide, PptElement, XmlSlide, XmlElement, PptRect, PptStyle } from './types'; // ✅ 去掉.ts后缀

/**
 * PPT解析核心工具函数 - 纯TS，无副作用，无耦合
 */
export const PptParseUtils = {
  /** 生成唯一ID - PPT元素/幻灯片ID */
  generateId: (prefix = 'ppt-node'): string => `${prefix}-${Date.now()}-${Math.floor(Math.random() * 10000)}`,

  /** XML文本节点解析 - 处理XML转义字符 */
  parseXmlText: (text: string): string => unescape(text || '').trim(),

  /** XML属性解析 - 提取XML节点的属性键值对 */
  parseXmlAttrs: (attrs: NamedNodeMap): Record<string, string> => {
    const result: Record<string, string> = {};
    Array.from(attrs).forEach(attr => {
      result[attr.nodeName] = attr.nodeValue || '';
    });
    return result;
  },

  /** XML字符串转结构化XML节点树 - 核心XML解析器 */
  parseXmlToTree: (xmlStr: string): XmlElement => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlStr, 'application/xml');
    const root = doc.documentElement;
    const buildTree = (node: Element): XmlElement => {
      const children: XmlElement[] = [];
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

  /** 从XML中提取坐标尺寸 - PPT核心转换逻辑 (EMU单位转PX) */
  parseXmlRect: (attrs: Record<string, string>): PptRect => {
    const emu2px = (emu: string): number => Math.round(parseInt(emu || '0') / 914400 * 96);
    return {
      x: emu2px(attrs['x']),
      y: emu2px(attrs['y']),
      width: emu2px(attrs['cx']),
      height: emu2px(attrs['cy']),
    };
  },

  /** 解析PPT样式 - XML样式转前端标准样式 ✅ 已修复TS2322 类型断言，零报错 */
  parseXmlStyle: (attrs: Record<string, string>): PptStyle => {
    // 合法的对齐方式列表
    const validAligns = ['left', 'center', 'right'] as const;
    // 校验+类型断言，确保值合法
    const textAlign = validAligns.includes(attrs['align'] as any) ? (attrs['align'] as 'left'|'center'|'right') : 'left';
    
    return {
      fontSize: attrs['fontSize'] ? parseInt(attrs['fontSize']) / 100 : 14,
      color: attrs['fill'] || '#333333',
      fontWeight: attrs['bold'] === '1' ? 'bold' : 'normal',
      textAlign: textAlign, // ✅ 类型匹配，零报错
      backgroundColor: attrs['bgFill'] || 'transparent',
      borderColor: attrs['border'] || '#000000',
      borderWidth: attrs['borderWidth'] ? parseInt(attrs['borderWidth']) : 1
    };
  },

  /** 反向转换：前端PX转PPT的EMU单位（序列化时用） */
  px2emu: (px: number): number => Math.round(px * 914400 / 96),

  /** EMU转PX */
  emu2px: (emu: number): number => Math.round(emu / 914400 * 96)
};

/**
 * 【核心】PPT解析算法 - 纯TS编写
 * 能力：解析标准 .pptx 文件 -> 转为前端结构化的 PptDocument 对象
 * @param file File|Blob|string 上传的ppt文件/二进制/blob/base64
 */
export async function parsePptx(file: File | Blob | string): Promise<PptDocument> {
  const zip = await JSZip.loadAsync(file);
  const pptProps = await parsePptProperties(zip);
  const xmlSlides = await parseXmlSlides(zip);
  const slides = await Promise.all(xmlSlides.map(slide => parseSlide(zip, slide)));

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

  async function parsePptProperties(zip: JSZip): Promise<{ width: number; height: number; title: string }> {
    const props = { width: 1280, height: 720, title: '未命名PPT' };
    try {
      const slideLayoutXml = await zip.file('ppt/slideLayouts/slideLayout1.xml')?.async('string');
      if (slideLayoutXml) {
        const xmlTree = PptParseUtils.parseXmlToTree(slideLayoutXml);
        const cSld = xmlTree.children.find(node => node.tag === 'cSld');
        const sldSz = cSld?.children.find(node => node.tag === 'sldSz');
        if (sldSz) {
          const rect = PptParseUtils.parseXmlRect(sldSz.attrs);
          props.width = rect.width;
          props.height = rect.height;
        }
      }
      const coreProps = await zip.file('docProps/core.xml')?.async('string');
      if (coreProps) {
        const xmlTree = PptParseUtils.parseXmlToTree(coreProps);
        props.title = xmlTree.children.find(node => node.tag.includes('title'))?.text || '未命名PPT';
      }
    } catch (e) {
      console.warn('PPT配置解析失败，使用默认配置', e);
    }
    return props;
  }

  async function parseXmlSlides(zip: JSZip): Promise<XmlSlide[]> {
    const slides: XmlSlide[] = [];
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

  async function parseSlide(zip: JSZip, xmlSlide: XmlSlide): Promise<PptSlide> {
    const { xml, layout } = xmlSlide;
    const xmlTree = PptParseUtils.parseXmlToTree(xml);
    const elements: PptElement[] = [];

    const spTree = xmlTree.children.find(node => node.tag === 'spTree');
    spTree?.children.forEach(node => {
      const element = parseXmlElement(node);
      if (element) elements.push(element);
    });

    const bgColor = parseSlideBgColor(xmlTree);

    return {
      id: xmlSlide.slideId,
      title: `幻灯片${slides.length + 1}`,
      bgColor,
      elements,
      props: {
        width: 1280,
        height: 720,
        slideLayout: layout,
      },
    };
  }

  function parseXmlElement(xmlNode: XmlElement): PptElement | null {
    const { tag, attrs, children, text } = xmlNode;
    const rect = PptParseUtils.parseXmlRect(attrs);
    const style = PptParseUtils.parseXmlStyle(attrs);
    const id = PptParseUtils.generateId('ppt-element');

    switch (tag) {
      case 'txBody':
        return { id, type: 'text', rect, style, content: text || '', props: { lineHeight: 1.5 } };
      case 'sp':
        return { id, type: 'shape', rect, style, content: attrs['shapeType'] || 'rect', props: { borderRadius: attrs['rx'] ? parseInt(attrs['rx']) / 100 : 0 } };
      case 'pic':
        const blip = children.find(node => node.tag === 'blip');
        const imgId = blip?.attrs['embed'] || '';
        return { id, type: 'image', rect, style, content: `ppt-image-${imgId}`, props: { imgId, alt: 'PPT图片' } };
      case 'tbl':
        const rows = children.filter(node => node.tag === 'tr');
        const tableData = rows.map(row => {
          const cells = row.children.filter(node => node.tag === 'tc');
          return cells.map(cell => cell.text || '');
        });
        return { id, type: 'table', rect, style, content: tableData, props: { rowCount: rows.length, colCount: rows[0]?.children.length || 0 } };
      case 'chart':
        return { id, type: 'chart', rect, style, content: { type: attrs['chartType'] || 'bar', data: [] }, props: { legend: true, grid: true } };
      default:
        return null;
    }
  }

  function parseSlideBgColor(xmlTree: XmlElement): string {
    const bgNode = xmlTree.children.find(node => node.tag === 'bg');
    if (!bgNode) return '#ffffff';
    const fillNode = bgNode.children.find(node => node.tag === 'fill');
    return fillNode?.attrs['color'] || '#ffffff';
  }
}

/**
 * 【核心】PPT序列化算法 - 纯TS编写
 * 能力：前端结构化 PptDocument 对象 -> 标准 .pptx 文件（可下载）
 * @param pptDoc 前端编辑后的PPT结构化JSON数据
 * @returns Blob PPTX二进制文件流，可直接下载
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

  async function writeSlides(zip: JSZip, slides: PptSlide[]): Promise<void> {
    slides.forEach(async (slide, index) => {
      const slideXml = slideToXml(slide);
      zip.file(`ppt/slides/slide${index + 1}.xml`, slideXml);
    });
  }

  function slideToXml(slide: PptSlide): string {
    const { elements, bgColor, props } = slide;
    const emuWidth = PptParseUtils.px2emu(props.width);
    const emuHeight = PptParseUtils.px2emu(props.height);
    const elementsXml = elements.map(el => elementToXml(el)).join('\n');

    return `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:grpSp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emuWidth}" cy="${emuHeight}"/></a:xfrm></p:spPr>${elementsXml}</p:grpSp></p:spTree></p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bgColor.replace('#', '')}"/></a:solidFill></p:bgPr></p:bg></p:sld>`
      .replace(/\s+/g, ' ').trim();
  }

  function elementToXml(el: PptElement): string {
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
}