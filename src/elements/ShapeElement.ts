/**
 * 形状元素类
 * 支持文本框、自定义形状、占位符等
 * 对齐 PPTXjs 的完整文本解析能力
 */

import { BaseElement } from './BaseElement';
import { getFirstChildByTagNS, parseTextContent, parseTextWithStyle, getAttrSafe, getBoolAttr, emu2px } from '../utils';
import { NS } from '../constants';
import type { ParsedShapeElement, RelsMap } from '../types';

/**
 * 渐变停止点
 */
interface GradientStop {
  color: string;
  position: number; // 0-100
}

/**
 * 渐变填充配置
 */
interface GradientFill {
  type: 'linear' | 'radial' | 'path';
  stops: GradientStop[];
  angle?: number; // 线性渐变角度（度）
  direction?: string; // 径向/路径渐变方向
}

/**
 * 文本运行样式
 */
export interface TextRun {
  text: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strike?: boolean;
  color?: string;
  backgroundColor?: string;
  highlight?: string;
}

/**
 * 项目符号类型
 */
export interface BulletStyle {
  type?: 'none' | 'char' | 'blip' | 'autoNum';
  char?: string;
  imageSrc?: string;
  autoNumType?: string;
  level?: number;
  color?: string;
  size?: number;
  font?: string;
}

/**
 * 形状元素类
 */
export class ShapeElement extends BaseElement {
  type: 'shape' | 'text' = 'shape';

  /** 形状类型（矩形、圆形等） */
  shapeType?: string;

  /** 文本内容 */
  text?: string;

  /** 文本样式（运行级别） */
  textStyle?: TextRun[];

  /** 文本段落样式 */
  paragraphStyle?: {
    align?: 'left' | 'center' | 'right' | 'justify';
    indent?: number;
    lineSpacing?: number;
    spaceBefore?: number;
    spaceAfter?: number;
    rtl?: boolean;
  };

  /** 项目符号样式 */
  bulletStyle?: BulletStyle;

  /** 是否占位符 */
  isPlaceholder?: boolean;

  /** 占位符类型 */
  placeholderType?: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';

  /** 超链接 */
  hyperlink?: {
    id?: string;
    url?: string;
    tooltip?: string;
  };

  /** 旋转角度（度） */
  rotation?: number;

  /** 是否水平翻转 */
  flipH?: boolean;

  /** 是否垂直翻转 */
  flipV?: boolean;

  /**
   * 从XML节点创建形状元素
   */
  static fromNode(node: Element, relsMap: RelsMap): ShapeElement | null {
    try {
      const element = new ShapeElement('', 'shape', { x: 0, y: 0, width: 0, height: 0 }, {}, {}, relsMap);

      // 解析ID和名称
      const nvSpPr = getFirstChildByTagNS(node, 'nvSpPr', NS.p);
      const cNvPr = nvSpPr ? getFirstChildByTagNS(nvSpPr, 'cNvPr', NS.p) : null;

      element.id = getAttrSafe(cNvPr, 'id', element.generateId());
      element.name = getAttrSafe(cNvPr, 'name', '');
      element.hidden = getBoolAttr(cNvPr, 'hidden');

      // 检查是否是占位符
      const nvPr = nvSpPr ? getFirstChildByTagNS(nvSpPr, 'nvPr', NS.p) : null;
      const ph = nvPr ? getFirstChildByTagNS(nvPr, 'ph', NS.p) : null;
      element.isPlaceholder = !!ph;
      const phType = ph?.getAttribute('type');
      if (phType) {
        // 映射可能的占位符类型
        const validTypes: ('title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other')[] = ['title', 'body', 'dateTime', 'slideNumber', 'footer', 'other'];
        if (validTypes.includes(phType as any)) {
          element.placeholderType = phType as 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';
        } else {
          element.placeholderType = 'other';
        }
      }

      // 解析位置尺寸和变换
      const spPr = getFirstChildByTagNS(node, 'spPr', NS.p);
      element.parseShapeProperties(spPr, node);

      // 解析文本内容
      const txBody = getFirstChildByTagNS(node, 'txBody', NS.p);
      if (txBody) {
        element.parseTextBody(txBody, node, relsMap);
      }

      element.shapeType = element.detectShapeType(spPr);

      // 判断元素类型
      element.type = (element.text || element.textStyle) ? 'text' : 'shape';

      element.rawNode = node;

      return element;
    } catch (error) {
      console.error('Failed to parse shape element:', error);
      return null;
    }
  }

  /**
   * 解析形状属性
   */
  private parseShapeProperties(spPr: Element | null, node: Element): void {
    if (!spPr) return;

      const xfrm = getFirstChildByTagNS(spPr, 'xfrm', NS.a);
      if (xfrm) {
        const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
        const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

        if (off) {
          this.rect.x = emu2px(off.getAttribute('x') || '0');
          this.rect.y = emu2px(off.getAttribute('y') || '0');
        }
        if (ext) {
          this.rect.width = emu2px(ext.getAttribute('cx') || '0');
          this.rect.height = emu2px(ext.getAttribute('cy') || '0');
        }

      // 旋转
      const rot = xfrm.getAttribute('rot');
      if (rot !== null) {
        this.rotation = parseInt(rot) / 60000; // 60000 EMU = 90度
      }

      // 翻转
      this.flipH = xfrm.getAttribute('flipH') === '1';
      this.flipV = xfrm.getAttribute('flipV') === '1';
    }

    // 背景填充
    this.parseFill(spPr);
  }

  /**
   * 解析填充
   */
  private parseFill(spPr: Element): void {
    const solidFill = getFirstChildByTagNS(spPr, 'solidFill', NS.a);
    if (solidFill) {
      const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
      if (srgbClr?.getAttribute('val')) {
        this.style.backgroundColor = `#${srgbClr.getAttribute('val')}`;
      }

      // 主题颜色
      const schemeClr = getFirstChildByTagNS(solidFill, 'schemeClr', NS.a);
      if (schemeClr?.getAttribute('val')) {
        this.style.backgroundColor = schemeClr.getAttribute('val') || 'transparent';
      }
    }

    // 渐变填充
    const gradFill = getFirstChildByTagNS(spPr, 'gradFill', NS.a);
    if (gradFill) {
      const gradient = this.parseGradientFill(gradFill);
      if (gradient) {
        this.style.background = this.generateGradientCSS(gradient);
      }
    }
  }

  /**
   * 解析渐变填充
   */
  private parseGradientFill(gradFill: Element): GradientFill | null {
    const gradient: GradientFill = {
      type: 'linear',
      stops: []
    };

    // 解析渐变类型
    const lin = getFirstChildByTagNS(gradFill, 'lin', NS.a);
    const path = getFirstChildByTagNS(gradFill, 'path', NS.a);
    const tileRect = getFirstChildByTagNS(gradFill, 'tileRect', NS.a);

    if (lin) {
      gradient.type = 'linear';
      
      // 解析角度
      const ang = lin.getAttribute('ang');
      if (ang) {
        gradient.angle = (parseInt(ang) / 60000) % 360; // 60000 EMU = 1度
      } else {
        gradient.angle = 0; // 默认从左到右
      }
    } else if (path) {
      gradient.type = 'path';
      
      // 解析路径方向
      const pathAttr = path.getAttribute('path');
      if (pathAttr) {
        gradient.direction = pathAttr;
      }
    } else {
      // 默认为径向渐变
      gradient.type = 'radial';
    }

    // 解析渐变停止点
    const gsLst = getFirstChildByTagNS(gradFill, 'gsLst', NS.a);
    if (gsLst) {
      const stops = Array.from(gsLst.children).filter(
        child => child.tagName === 'a:gs' || child.tagName.includes(':gs')
      );

      for (const stop of stops) {
        const color = this.parseGradientStopColor(stop);
        const pos = stop.getAttribute('pos');
        
        if (color) {
          gradient.stops.push({
            color,
            position: pos ? parseInt(pos) / 100000 : 0 // 位置是万分比
          });
        }
      }
    }

    // 如果没有停止点，添加默认的黑白渐变
    if (gradient.stops.length === 0) {
      gradient.stops.push(
        { color: '#000000', position: 0 },
        { color: '#ffffff', position: 100 }
      );
    }

    return gradient;
  }

  /**
   * 解析渐变停止点颜色
   */
  private parseGradientStopColor(stop: Element): string | null {
    const solidFill = getFirstChildByTagNS(stop, 'solidFill', NS.a);
    if (!solidFill) return null;

    // RGB颜色
    const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
    if (srgbClr?.getAttribute('val')) {
      return `#${srgbClr.getAttribute('val')}`;
    }

    // 主题颜色
    const schemeClr = getFirstChildByTagNS(solidFill, 'schemeClr', NS.a);
    if (schemeClr?.getAttribute('val')) {
      return schemeClr.getAttribute('val') || null;
    }

    // 系统颜色
    const sysClr = getFirstChildByTagNS(solidFill, 'sysClr', NS.a);
    if (sysClr?.getAttribute('val')) {
      return sysClr.getAttribute('val') || null;
    }

    return null;
  }

  /**
   * 生成CSS渐变字符串
   */
  private generateGradientCSS(gradient: GradientFill): string {
    const sortedStops = [...gradient.stops].sort((a, b) => a.position - b.position);
    const stopStrings = sortedStops.map(stop => `${stop.color} ${stop.position}%`).join(', ');

    switch (gradient.type) {
      case 'linear':
        const angle = gradient.angle !== undefined ? gradient.angle : 0;
        return `linear-gradient(${angle}deg, ${stopStrings})`;

      case 'radial':
        return `radial-gradient(circle, ${stopStrings})`;

      case 'path':
        // CSS不支持path渐变，回退到径向渐变
        const direction = gradient.direction || 'circle';
        return `radial-gradient(${direction}, ${stopStrings})`;

      default:
        return `linear-gradient(${stopStrings})`;
    }
  }

  /**
   * 解析文本主体
   */
  private parseTextBody(txBody: Element, shapeNode: Element, relsMap: RelsMap): void {
    // 解析段落
    const paragraphs = Array.from(txBody.children).filter(
      child => child.tagName === 'a:p' || child.tagName.includes(':p')
    );

    this.textStyle = paragraphs.flatMap(p => this.parseParagraph(p, shapeNode, relsMap));
    this.text = this.textStyle.map(t => t.text).join('');

    // 解析段落属性
    const firstParagraph = txBody.querySelector('a\\:p, p\\:p a\\:p');
    const pPr = firstParagraph ? getFirstChildByTagNS(firstParagraph, 'pPr', NS.a) : null;
    if (pPr) {
      this.paragraphStyle = {
        align: pPr.getAttribute('algn') as any || undefined,
        indent: parseInt(pPr.getAttribute('indent') || '0') / 100,
        lineSpacing: parseInt(pPr.getAttribute('lnSpc') || '0') / 100,
        spaceBefore: parseInt(getAttrSafe(
          getFirstChildByTagNS(pPr, 'spcBef', NS.a),
          'spcPts',
          '0'
        )) / 100,
        spaceAfter: parseInt(getAttrSafe(
          getFirstChildByTagNS(pPr, 'spcAft', NS.a),
          'spcPts',
          '0'
        )) / 100,
        rtl: pPr.getAttribute('rtl') === '1'
      };
    }
  }

  /**
   * 解析段落
   */
  private parseParagraph(paragraph: Element, shapeNode: Element, relsMap: RelsMap): TextRun[] {
    const runs: TextRun[] = [];

    // 解析项目符号
    this.bulletStyle = this.parseBulletStyle(paragraph);

    // 解析文本运行
    const textRuns = Array.from(paragraph.children).filter(
      child => child.tagName === 'a:r' || child.tagName.includes(':r')
    );

    for (const r of textRuns) {
      const text = this.parseTextRun(r, relsMap);
      if (text) runs.push(text);
    }

    // 解析超链接
    const hyperlink = getFirstChildByTagNS(paragraph, 'hlinkClick', NS.a);
    if (hyperlink) {
      const rId = hyperlink.getAttributeNS(NS.r, 'id') || hyperlink.getAttribute('r:id') || '';
      if (rId && relsMap[rId]) {
        this.hyperlink = {
          id: rId,
          url: relsMap[rId].target,
          tooltip: hyperlink.getAttribute('tooltip') || undefined
        };
      }
    }

    return runs;
  }

  /**
   * 解析项目符号
   */
  private parseBulletStyle(paragraph: Element): BulletStyle {
    const pPr = getFirstChildByTagNS(paragraph, 'pPr', NS.a);
    if (!pPr) return {};

    // 检查无项目符号
    const buNone = getFirstChildByTagNS(pPr, 'buNone', NS.a);
    if (buNone) return { type: 'none' };

    // 字符项目符号
    const buChar = getFirstChildByTagNS(pPr, 'buChar', NS.a);
    if (buChar) {
      const char = buChar.getAttribute('char') || '•';
      const font = buChar.getAttribute('font') || undefined;
      const color = this.parseColor(buChar);
      const size = parseInt(buChar.getAttribute('sz') || '100') / 100;
      return {
        type: 'char',
        char,
        font,
        color,
        size
      };
    }

    // 图片项目符号
    const buBlip = getFirstChildByTagNS(pPr, 'buBlip', NS.a);
    if (buBlip) {
      const rId = buBlip.getAttributeNS(NS.r, 'embed') || buBlip.getAttribute('r:embed') || '';
      return {
        type: 'blip',
        imageSrc: rId // 需要从relsMap解析
      };
    }

    // 自动编号
    const buAutoNum = getFirstChildByTagNS(pPr, 'buAutoNum', NS.a);
    if (buAutoNum) {
      return {
        type: 'autoNum',
        autoNumType: buAutoNum.getAttribute('type') || 'arabic',
        level: parseInt(buAutoNum.getAttribute('startAt') || '1')
      };
    }

    return { type: 'none' };
  }

  /**
   * 解析文本运行
   */
  private parseTextRun(run: Element, relsMap: RelsMap): TextRun | null {
    const rPr = getFirstChildByTagNS(run, 'rPr', NS.a);
    const textElem = getFirstChildByTagNS(run, 't', NS.a);

    if (!textElem) return null;

    const text = textElem.textContent || '';

    const result: TextRun = { text };

    if (rPr) {
      // 字体大小
      const sz = rPr.getAttribute('sz');
      if (sz) {
        result.fontSize = parseInt(sz) / 100; // 单位是百分之一磅
      }

      // 字体家族
      const latin = getFirstChildByTagNS(rPr, 'latin', NS.a);
      const ea = getFirstChildByTagNS(rPr, 'ea', NS.a);
      const cs = getFirstChildByTagNS(rPr, 'cs', NS.a);
      const latinTypeface = latin?.getAttribute('typeface');
      const eaTypeface = ea?.getAttribute('typeface');
      const csTypeface = cs?.getAttribute('typeface');
      if (latinTypeface) {
        result.fontFamily = latinTypeface;
      } else if (eaTypeface) {
        result.fontFamily = eaTypeface;
      } else if (csTypeface) {
        result.fontFamily = csTypeface;
      }

      // 加粗
      if (rPr.getAttribute('b') === '1') {
        result.bold = true;
        this.style.fontWeight = 'bold';
      }

      // 斜体
      if (rPr.getAttribute('i') === '1') {
        result.italic = true;
      }

      // 下划线
      const u = rPr.getAttribute('u');
      if (u) {
        result.underline = u === 'none' ? 'none' : 'underline';
      }

      // 删除线
      if (rPr.getAttribute('strike') === '1') {
        result.strike = true;
      }

      // 颜色
      const solidFill = getFirstChildByTagNS(rPr, 'solidFill', NS.a);
      if (solidFill) {
        const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
        if (srgbClr?.getAttribute('val')) {
          result.color = `#${srgbClr.getAttribute('val')}`;
          this.style.color = result.color;
        }
      }

      // 高亮
      const highlight = getFirstChildByTagNS(rPr, 'highlight', NS.a);
      if (highlight) {
        const srgbClr = getFirstChildByTagNS(highlight, 'srgbClr', NS.a);
        if (srgbClr?.getAttribute('val')) {
          result.backgroundColor = `#${srgbClr.getAttribute('val')}`;
        }
      }
    }

    return result;
  }

  /**
   * 解析颜色
   */
  private parseColor(element: Element): string | undefined {
    const solidFill = getFirstChildByTagNS(element, 'solidFill', NS.a);
    if (solidFill) {
      const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
      if (srgbClr?.getAttribute('val')) {
        return `#${srgbClr.getAttribute('val')}`;
      }
    }
    return undefined;
  }

  /**
   * 检测形状类型
   */
  private detectShapeType(spPr: Element | null): string {
    if (!spPr) return 'rectangle';

    const prstGeom = getFirstChildByTagNS(spPr, 'prstGeom', NS.a);
    if (prstGeom) {
      const prst = prstGeom.getAttribute('prst') || '';
      // 常见形状映射
      const shapeMap: Record<string, string> = {
        'rect': 'rectangle',
        'ellipse': 'ellipse',
        'roundRect': 'roundedRectangle',
        'triangle': 'triangle',
        'diamond': 'diamond',
        'pentagon': 'pentagon',
        'hexagon': 'hexagon',
        'octagon': 'octagon',
        'star4': 'star',
        'star5': '5-point star',
        'star6': '6-point star',
        'star8': '8-point star',
        'star10': '10-point star',
        'star12': '12-point star',
        'star16': '16-point star',
        'star24': '24-point star',
        'star32': '32-point star',
        'heart': 'heart',
        'lightning': 'lightning',
        'sun': 'sun',
        'moon': 'moon',
        'cloud': 'cloud',
        'arrow': 'arrow',
        'bentArrow': 'bent arrow',
        'chevron': 'chevron',
        'home': 'home',
        'cube': 'cube',
        'bevel': 'bevel',
        'donut': 'donut',
        'noSmoking': 'noSmoking',
        'blockArc': 'blockArc',
        'foldedCorner': 'foldedCorner',
        'smileyFace': 'smileyFace'
      };
      return shapeMap[prst] || prst;
    }

    const custGeom = getFirstChildByTagNS(spPr, 'custGeom', NS.a);
    if (custGeom) {
      return 'custom';
    }

    return 'rectangle';
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const style = this.getContainerStyle();
    const textStyle = this.getTextStyle();
    const rotationStyle = this.getRotationStyle();
    const dataAttrs = this.formatDataAttributes();

    const combinedStyle = [...this.parseStyleString(style), ...this.parseStyleString(rotationStyle)].join('; ');

    // 处理图片类型
    if (this.type === 'image' && this.rawData) {
      const blipFill = this.rawData.querySelector('p\:blipFill, blipFill');
      if (blipFill) {
        const blip = blipFill.querySelector('a\:blip, blip');
        if (blip) {
          const rEmbed = blip.getAttribute('r:embed') || blip.getAttribute('embed');
          if (rEmbed && this.relsMap) {
            const target = this.relsMap[rEmbed]?.target;
            if (target) {
              // 构建正确的图片路径
              let imageSrc = target;
              // 如果是相对路径，确保路径正确
              if (!imageSrc.startsWith('http') && !imageSrc.startsWith('/')) {
                // 根据 rEmbed 判断是 slide 还是 layout 的图片
                if (rEmbed.startsWith('rId')) {
                  // 检查是否在 layout 中
                  if (this.rawData.closest('p\:sldLayout, sldLayout')) {
                    imageSrc = `ppt/slideLayouts/${target}`;
                  } else {
                    imageSrc = `ppt/slides/${target}`;
                  }
                }
              }
              
              return `<img ${dataAttrs} style="${combinedStyle}" src="${imageSrc}" data-rel-id="${rEmbed}" data-file="${imageSrc}" />`;
            }
          }
        }
      }
    }

    if (this.type === 'text' && this.text) {
      // 文本框
      return `<div ${dataAttrs} style="${combinedStyle}">
        <div style="${textStyle}">${this.escapeHtml(this.text)}</div>
      </div>`;
    } else {
      // 形状（矩形、圆形等）
      const shapeStyle = this.getShapeStyle();
      const finalStyle = [...this.parseStyleString(style), ...this.parseStyleString(shapeStyle), ...this.parseStyleString(rotationStyle)].join('; ');
      return `<div ${dataAttrs} style="${finalStyle}"></div>`;
    }
  }

  /**
   * 获取文本样式
   */
  private getTextStyle(): string {
    const styles = [
      `display: flex`,
      `align-items: center`,
      `justify-content: this.textStyleFromAlign(this.paragraphStyle?.align)`,
      `padding: 10px`,
      `color: ${this.style.color}`,
      `font-size: ${this.textStyleFromFontSize()}px`
    ];

    if (this.style.fontWeight === 'bold') {
      styles.push(`font-weight: bold`);
    }

    if (this.style.backgroundColor && this.style.backgroundColor !== 'transparent') {
      styles.push(`background-color: ${this.style.backgroundColor}`);
    }

    if (this.style.borderWidth && this.style.borderWidth > 0) {
      styles.push(`border: ${this.style.borderWidth}px solid ${this.style.borderColor}`);
    }

    if (this.textStyle?.[0]?.fontFamily) {
      styles.push(`font-family: ${this.textStyle[0].fontFamily}`);
    }

    return styles.join('; ');
  }

  /**
   * 获取旋转样式
   */
  private getRotationStyle(): string {
    if (this.rotation === undefined || this.rotation === 0) return '';

    return `transform: rotate(${this.rotation}deg); transform-origin: center;`;
  }

  /**
   * 获取形状样式
   */
  private getShapeStyle(): string {
    const styles = [
      `background-color: ${this.style.backgroundColor || '#ffffff'}`,
      `border: ${this.style.borderWidth}px solid ${this.style.borderColor}`
    ];

    return styles.join('; ');
  }

  /**
   * 文本对齐转换
   */
  private textStyleFromAlign(align?: string): string {
    switch (align) {
      case 'right': return 'flex-end';
      case 'center': return 'center';
      case 'justify': return 'center';
      default: return 'flex-start';
    }
  }

  /**
   * 文本大小转换
   */
  private textStyleFromFontSize(): number {
    if (this.textStyle?.[0]?.fontSize) return this.textStyle[0].fontSize;
    return this.style.fontSize || 14;
  }

  /**
   * HTML转义
   */
  private escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  /**
   * 解析样式字符串
   */
  private parseStyleString(style: string): string[] {
    return style.split(';').filter(s => s.trim());
  }

  /**
   * 转换为ParsedShapeElement格式
   */
  toParsedElement(): ParsedShapeElement {
    return {
      id: this.id,
      type: 'shape',
      rect: this.rect,
      style: this.style,
      content: this.content,
      props: {
        isPlaceholder: this.isPlaceholder,
        placeholderType: this.placeholderType,
        shapeType: this.shapeType,
        textStyle: this.textStyle,
        paragraphStyle: this.paragraphStyle,
        bulletStyle: this.bulletStyle,
        hyperlink: this.hyperlink,
        rotation: this.rotation
      },
      name: this.name,
      hidden: this.hidden,
      text: this.text,
      textStyle: this.textStyle,
      attrs: this.getAttributes(this.rawNode!),
      rawNode: this.rawNode
    };
  }
}
