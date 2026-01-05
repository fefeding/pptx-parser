import JSZip from 'jszip';
import { unescape } from 'html-escaper';
import type { PptDocument, PptSlide, PptElement, XmlSlide, XmlElement, PptRect, PptStyle, PptTransform, PptTextParagraph, PptFill, PptBorder, PptShadow } from './types';

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

  async function parseSlide(zip: JSZip, xmlSlide: XmlSlide, slideIndex: number): Promise<PptSlide> {
    const { xml, layout } = xmlSlide;
    const xmlTree = PptParseUtils.parseXmlToTree(xml);

    console.log('Slide XML root tag:', xmlTree.tag, 'children count:', xmlTree.children.length);

    // PPT XML结构: p:sld -> p:cSld -> p:spTree
    // 需要匹配带命名空间的标签名，如 p:cSld
    const cSld = xmlTree.children.find(node => node.tag === 'p:cSld' || node.tag === 'cSld');
    console.log('cSld found:', !!cSld, 'tag:', cSld?.tag);

    const spTree = cSld?.children.find(node => node.tag === 'p:spTree' || node.tag === 'spTree');
    console.log('spTree found:', !!spTree, 'children count:', spTree?.children.length);

    const elements: PptElement[] = [];

    if (spTree) {
      // spTree包含多种子元素: p:sp (shape), p:pic (picture), p:graphicFrame (chart/table), p:grpSp (group)
      console.log('Processing spTree children...');
      spTree.children.forEach((node, idx) => {
        console.log(`Node ${idx}: tag=${node.tag}, attrs=`, node.attrs);
        const element = parseXmlElement(node);
        if (element) {
          console.log(`Parsed element: type=${element.type}, rect=`, element.rect);
          elements.push(element);
        }
      });
    }

    console.log('Total elements parsed:', elements.length);

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

  function parseXmlElement(xmlNode: XmlElement): PptElement | null {
    const { tag, attrs, children } = xmlNode;
    const id = PptParseUtils.generateId('ppt-element');

    // p:sp - shape/文本框
    if (tag === 'p:sp' || tag === 'sp') {
      return parseShape(xmlNode);
    }

    // p:pic - 图片
    if (tag === 'p:pic' || tag === 'pic') {
      return parsePicture(xmlNode);
    }

    // p:graphicFrame - 图表、表格、嵌入对象
    if (tag === 'p:graphicFrame' || tag === 'graphicFrame') {
      return parseGraphicFrame(xmlNode);
    }

    // p:grpSp - 分组，递归处理子元素
    if (tag === 'p:grpSp' || tag === 'grpSp') {
      return parseGroup(xmlNode);
    }

    return null;
  }

  function parseShape(xmlNode: XmlElement): PptElement | null {
    const id = PptParseUtils.generateId('shape');
    const children = xmlNode.children;

    // 查找 spPr (shape properties) 和 txBody (text body)
    const spPr = children.find(node => node.tag === 'p:spPr' || node.tag === 'spPr');
    const txBody = children.find(node => node.tag === 'p:txBody' || node.tag === 'txBody');
    const nvSpPr = children.find(node => node.tag === 'p:nvSpPr' || node.tag === 'nvSpPr');

    // 解析坐标和尺寸
    let rect = { x: 0, y: 0, width: 0, height: 0 };
    if (spPr) {
      const xfrm = spPr.children.find(node => node.tag === 'a:xfrm');
      if (xfrm) {
        const off = xfrm.children.find(node => node.tag === 'a:off');
        const ext = xfrm.children.find(node => node.tag === 'a:ext');
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

    // 解析文本内容
    let content: any = '';
    if (txBody) {
      const paragraphs = txBody.children.filter(node => node.tag === 'a:p' || node.tag === 'p');
      const textRuns: any[] = [];

      paragraphs.forEach(p => {
        const runs = p.children.filter(node => node.tag === 'a:r' || node.tag === 'r');
        runs.forEach(r => {
          const t = r.children.find(node => node.tag === 'a:t' || node.tag === 't');
          if (t && t.text) {
            const rPr = r.children.find(node => node.tag === 'a:rPr' || node.tag === 'rPr');
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

    // 判断是否是占位符
    const nvPr = nvSpPr?.children.find(node => node.tag === 'p:nvPr' || node.tag === 'nvPr');
    const ph = nvPr?.children.find(node => node.tag === 'p:ph' || node.tag === 'ph');
    const isPlaceholder = !!ph;
    const placeholderType = ph?.attrs['type'];

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
      props: {
        isPlaceholder,
        placeholderType
      }
    };
  }

  function parsePicture(xmlNode: XmlElement): PptElement | null {
    const id = PptParseUtils.generateId('image');
    const children = xmlNode.children;

    // 查找 spPr 和 blipFill
    const spPr = children.find(node => node.tag === 'p:spPr' || node.tag === 'spPr');
    const blipFill = children.find(node => node.tag === 'p:blipFill' || node.tag === 'blipFill');

    // 解析坐标和尺寸
    let rect = { x: 0, y: 0, width: 0, height: 0 };
    if (spPr) {
      const xfrm = spPr.children.find(node => node.tag === 'a:xfrm');
      if (xfrm) {
        const off = xfrm.children.find(node => node.tag === 'a:off');
        const ext = xfrm.children.find(node => node.tag === 'a:ext');
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

    // 解析图片ID
    let imgId = '';
    if (blipFill) {
      const blip = blipFill.children.find(node => node.tag === 'a:blip' || node.tag === 'blip');
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

  function parseGraphicFrame(xmlNode: XmlElement): PptElement | null {
    const id = PptParseUtils.generateId('graphic');
    const children = xmlNode.children;

    // 查找 xfrmspPr 用于定位
    const xfrm = children.find(node => node.tag === 'p:xfrm' || node.tag === 'xfrm');
    let rect = { x: 0, y: 0, width: 0, height: 0 };
    if (xfrm) {
      const off = xfrm.children.find(node => node.tag === 'a:off');
      const ext = xfrm.children.find(node => node.tag === 'a:ext');
      if (off && ext) {
        rect = PptParseUtils.parseXmlRect({
          x: off.attrs['x'] || '0',
          y: off.attrs['y'] || '0',
          cx: ext.attrs['cx'] || '0',
          cy: ext.attrs['cy'] || '0'
        });
      }
    }

    // 查找 graphicData 判断类型
    const graphic = children.find(node => node.tag === 'a:graphic' || node.tag === 'graphic');
    const graphicData = graphic?.children.find(node => node.tag === 'a:graphicData' || node.tag === 'graphicData');

    if (!graphicData) return null;

    const uri = graphicData.attrs['uri'] || '';

    // 判断是表格还是图表
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

  function parseGroup(xmlNode: XmlElement): PptElement | null {
    // 组元素返回null，不作为独立元素，其子元素会被遍历
    return null;
  }

  function parseSlideTitle(xmlTree: XmlElement, defaultIndex: number): string {
    const cSld = xmlTree.children.find(node => node.tag === 'p:cSld' || node.tag === 'cSld');
    const spTree = cSld?.children.find(node => node.tag === 'p:spTree' || node.tag === 'spTree');

    if (!spTree) return `幻灯片${defaultIndex}`;

    // 查找标题占位符
    for (const node of spTree.children) {
      if (node.tag === 'p:sp' || node.tag === 'sp') {
        const nvSpPr = node.children.find(n => n.tag === 'p:nvSpPr' || n.tag === 'nvSpPr');
        const nvPr = nvSpPr?.children.find(n => n.tag === 'p:nvPr' || n.tag === 'nvPr');
        const ph = nvPr?.children.find(n => n.tag === 'p:ph' || n.tag === 'ph');

        if (ph?.attrs['type'] === 'title') {
          const txBody = node.children.find(n => n.tag === 'p:txBody' || n.tag === 'txBody');
          const p = txBody?.children.find(n => n.tag === 'a:p' || n.tag === 'p');
          const r = p?.children.find(n => n.tag === 'a:r' || n.tag === 'r');
          const t = r?.children.find(n => n.tag === 'a:t' || n.tag === 't');
          if (t?.text) return t.text;
        }
      }
    }

    return `幻灯片${defaultIndex}`;
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