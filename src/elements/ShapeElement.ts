/**
 * 形状元素类
 * 支持文本框、自定义形状、占位符等
 * 对齐 PPTXjs 的完整文本解析能力
 */

import { BaseElement } from './BaseElement';
import { getFirstChildByTagNS, getChildrenByTagNS, parseTextContent, parseTextWithStyle, getAttrSafe, getBoolAttr, emu2px } from '../utils';
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
  letterSpacing?: number;
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
    marginLeft?: number;
    marginRight?: number;
    paddingTop?: number;
    paddingBottom?: number;
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

      // 解析索引（idx）
      element.idx = cNvPr ? parseInt(cNvPr.getAttribute('idx') || '0') : undefined;

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
    const paragraphs = getChildrenByTagNS(txBody, 'p', NS.a);

    this.textStyle = paragraphs.flatMap(p => this.parseParagraph(p, shapeNode, relsMap));
    this.text = this.textStyle.map((t: TextRun) => t.text).join('');

    // 解析段落属性
    const firstParagraph = txBody.querySelector('a\\:p, p\\:p a\\:p');
    const pPr = firstParagraph ? getFirstChildByTagNS(firstParagraph, 'pPr', NS.a) : null;
    if (pPr) {
      // 解析段前间距（spcBef）
      const spcBef = getFirstChildByTagNS(pPr, 'spcBef', NS.a);
      const spaceBefore = spcBef ? parseInt(getAttrSafe(spcBef, 'spcPts', '0')) / 100 : 0;

      // 解析段后间距（spcAft）
      const spcAft = getFirstChildByTagNS(pPr, 'spcAft', NS.a);
      const spaceAfter = spcAft ? parseInt(getAttrSafe(spcAft, 'spcPts', '0')) / 100 : 0;

      // 解析行距（lnSpc）
      const lnSpc = getFirstChildByTagNS(pPr, 'lnSpc', NS.a);
      const lineSpacing = lnSpc ? parseInt(getAttrSafe(lnSpc, 'spcPts', '0')) / 100 : 0;

      // 解析左右边距（marL, marR）
      const marL = getFirstChildByTagNS(pPr, 'marL', NS.a);
      const marginLeft = marL ? parseInt(marL.getAttribute('val') || '0') / 100 : 0;

      const marR = getFirstChildByTagNS(pPr, 'marR', NS.a);
      const marginRight = marR ? parseInt(marR.getAttribute('val') || '0') / 100 : 0;

      // 解析上下内边距（insT, insB）
      const insT = getFirstChildByTagNS(pPr, 'insT', NS.a);
      const paddingTop = insT ? parseInt(insT.getAttribute('val') || '0') / 100 : 0;

      const insB = getFirstChildByTagNS(pPr, 'insB', NS.a);
      const paddingBottom = insB ? parseInt(insB.getAttribute('val') || '0') / 100 : 0;

      // 解析首行缩进（indent）
      const indent = pPr.getAttribute('indent');
      const indentPx = indent ? parseInt(indent) / 100 : 0;

      this.paragraphStyle = {
        align: pPr.getAttribute('algn') as any || undefined,
        indent: indentPx,
        lineSpacing: lineSpacing,
        spaceBefore: spaceBefore,
        spaceAfter: spaceAfter,
        marginLeft: marginLeft,
        marginRight: marginRight,
        paddingTop: paddingTop,
        paddingBottom: paddingBottom,
        rtl: pPr.getAttribute('rtl') === '1'
      };

      // 将段落样式保存到this.style中供HTML渲染使用
      this.style.spaceBefore = spaceBefore;
      this.style.spaceAfter = spaceAfter;
      this.style.paddingTop = paddingTop;
      this.style.paddingBottom = paddingBottom;
      this.style.marginLeft = marginLeft;
      this.style.marginRight = marginRight;
    }
  }

  /**
   * 解析段落
   */
  private parseParagraph(paragraph: Element, shapeNode: Element, relsMap: RelsMap): TextRun[] {
    const runs: TextRun[] = [];

    // 解析项目符号
    this.bulletStyle = this.parseBulletStyle(paragraph);

    // 解析段落的默认运行样式（defRPr）
    const pPr = getFirstChildByTagNS(paragraph, 'pPr', NS.a);
    const defRPr = pPr ? getFirstChildByTagNS(pPr, 'defRPr', NS.a) : null;
    const defaultStyle = defRPr ? this.parseRunProperties(defRPr) : {};

    // 解析文本运行
    const textRuns = getChildrenByTagNS(paragraph, 'r', NS.a);

    for (const r of textRuns) {
      const text = this.parseTextRun(r, relsMap, defaultStyle);
      if (text) runs.push(text);
    }

    // 解析段落末尾运行属性（endParaRPr）
    const endParaRPr = getFirstChildByTagNS(paragraph, 'endParaRPr', NS.a);
    if (endParaRPr) {
      // endParaRPr中的样式会影响段落中所有文本
      const endStyle = this.parseRunProperties(endParaRPr);
      // 将endParaRPr的样式应用到所有没有该属性的文本运行
      runs.forEach(run => {
        if (!run.fontSize && endStyle.fontSize) run.fontSize = endStyle.fontSize;
        if (!run.fontFamily && endStyle.fontFamily) run.fontFamily = endStyle.fontFamily;
        if (!run.color && endStyle.color) run.color = endStyle.color;
      });
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
  private parseTextRun(run: Element, relsMap: RelsMap, defaultStyle: Partial<TextRun> = {}): TextRun | null {
    const rPr = getFirstChildByTagNS(run, 'rPr', NS.a);
    const textElem = getFirstChildByTagNS(run, 't', NS.a);

    if (!textElem) return null;

    const text = textElem.textContent || '';

    // 从defRPr继承的默认样式
    const result: TextRun = {
      text,
      ...defaultStyle
    };

    // 如果有直接的rPr，覆盖默认样式
    if (rPr) {
      const runStyle = this.parseRunProperties(rPr);
      Object.assign(result, runStyle);
    }

    return result;
  }

  /**
   * 解析运行属性（rPr或defRPr）
   */
  private parseRunProperties(rPr: Element): Partial<TextRun> {
    const style: Partial<TextRun> = {};

  // 字体大小
  let fontSizePt: number | undefined;
  const szAttr = rPr.getAttribute('sz');
  if (szAttr) {
    // PPTX中sz单位是百分之一磅（1/100 pt）
    fontSizePt = parseInt(szAttr) / 100;
  } else {
    // 检查 <a:sz> 元素
    const szElem = getFirstChildByTagNS(rPr, 'sz', NS.a);
    if (szElem) {
      const val = szElem.getAttribute('val');
      if (val) {
        fontSizePt = parseInt(val) / 100;
      }
    }
  }
  if (fontSizePt !== undefined) {
    // 需要将磅转换为像素：1 pt = 4/3 px（96 DPI下）
    style.fontSize = fontSizePt * (4 / 3);
  }

    // 字体家族
    const latin = getFirstChildByTagNS(rPr, 'latin', NS.a);
    const ea = getFirstChildByTagNS(rPr, 'ea', NS.a);
    const cs = getFirstChildByTagNS(rPr, 'cs', NS.a);
    const latinTypeface = latin?.getAttribute('typeface');
    const eaTypeface = ea?.getAttribute('typeface');
    const csTypeface = cs?.getAttribute('typeface');
    if (latinTypeface) {
      style.fontFamily = latinTypeface;
    } else if (eaTypeface) {
      style.fontFamily = eaTypeface;
    } else if (csTypeface) {
      style.fontFamily = csTypeface;
    }

    // 加粗
    if (rPr.getAttribute('b') === '1') {
      style.bold = true;
    }

    // 斜体
    if (rPr.getAttribute('i') === '1') {
      style.italic = true;
    }

    // 下划线
    const u = rPr.getAttribute('u');
    if (u) {
      style.underline = u === 'none' ? 'none' : 'underline';
    }

    // 删除线
    if (rPr.getAttribute('strike') === '1') {
      style.strike = true;
    }

    // 颜色
    const solidFill = getFirstChildByTagNS(rPr, 'solidFill', NS.a);
    if (solidFill) {
      const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
      if (srgbClr?.getAttribute('val')) {
        const color = `#${srgbClr.getAttribute('val')}`;
        style.color = color;
        this.style.color = color;
      }
    }

    // 高亮
    const highlight = getFirstChildByTagNS(rPr, 'highlight', NS.a);
    if (highlight) {
      const srgbClr = getFirstChildByTagNS(highlight, 'srgbClr', NS.a);
      if (srgbClr?.getAttribute('val')) {
        style.backgroundColor = `#${srgbClr.getAttribute('val')}`;
      }
    }

    // 字间距
    const kern = getFirstChildByTagNS(rPr, 'kern', NS.a);
    if (kern) {
      const kernVal = kern.getAttribute('val');
      if (kernVal) {
        style.letterSpacing = parseInt(kernVal) / 100; // 单位是百分之一磅
      }
    }

    return style;
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
   * 完全复刻 PPTXjs 的 DOM 结构
   */
  toHTML(): string {
    // 构建 data-* 属性字符串
    const dataAttrs = this.formatDataAttributes();
    
    // 构建 block 样式
    const blockStyle = this.generateBlockStyle();
    
    // 构建 block 类名
    const blockClasses = this.generateBlockClasses();
    
    // 构建内部结构
    const innerHTML = this.generateInnerHTML();
    
    return `<div class="${blockClasses.join(' ')}" ${dataAttrs} style="${blockStyle}">
      ${innerHTML}
    </div>`;
  }
  
  /**
   * 生成 block 样式
   */
  private generateBlockStyle(): string {
    const { x, y, width, height } = this.rect;
    
    const styles = [
      `position: absolute`,
      `top: ${y}px`,
      `left: ${x}px`,
      `width: ${width}px`,
      `height: ${height}px`
    ];
    
    // 边框样式 - 根据实际的边框宽度和颜色设置
    if (this.style.borderWidth && parseFloat(this.style.borderWidth as any) > 0 && this.style.borderColor) {
      styles.push(`border: ${this.style.borderWidth}px solid ${this.style.borderColor}`);
    } else {
      // 文本占位符通常没有边框，或者边框透明
      styles.push(`border: none`);
    }
    
    // 背景颜色 - 文本占位符通常透明背景
    if (this.style.backgroundColor && this.style.backgroundColor !== 'transparent') {
      styles.push(`background-color: ${this.style.backgroundColor}`);
    } else {
      // 文本占位符应该是透明的，不继承父级背景
      styles.push(`background-color: transparent`);
    }
    
    // z-index
    if (this.zIndex !== undefined) {
      styles.push(`z-index: ${this.zIndex}`);
    } else {
      styles.push(`z-index: 1`);
    }
    
    // 旋转
    if (this.rotation) {
      styles.push(`transform: rotate(${this.rotation}deg)`);
    } else {
      styles.push(`transform: rotate(0deg)`);
    }
    
    return styles.join('; ');
  }
  
  /**
   * 生成 block 类名
   */
  private generateBlockClasses(): string[] {
    const classes = ['block'];
    
    // 垂直对齐类（根据形状位置）
    // 简化处理：如果 y 坐标较小，可能是 v-up
    if (this.rect.y < 100) {
      classes.push('v-up');
    }
    
    // 内容类（如果是文本或占位符）
    if (this.type === 'text' || this.isPlaceholder) {
      classes.push('content');
    }
    
    return classes;
  }
  
  /**
   * 生成内部 HTML 结构
   */
  private generateInnerHTML(): string {
    if (this.type === 'text' && this.text) {
      // 文本框结构
      const { width, height } = this.rect;
      
      // 对齐类
      const alignClass = this.getAlignClass();
      
      // 文本容器样式
      const textContainerStyle = [
        `height: 100%`,
        `direction: ${this.paragraphStyle?.rtl ? 'rtl' : 'initial'}`,
        `overflow-wrap: break-word`,
        `word-wrap: break-word`,
        `width: ${width}px`
      ].join('; ');
      
      // 文本内容
      const textContent = this.renderTextContentPPTXjs();
      
      return `<div style="display: flex; width: ${width}px;" class="slide-prgrph ${alignClass} pregraph-ltr">
        <div style="${textContainerStyle}; background: transparent;">
          ${textContent}
        </div>
      </div>`;
    } else {
      // 形状元素
      return '';
    }
  }
  

  
  /**
   * 渲染文本内容（PPTXjs 风格）
   * 生成 <span class="text-block">文本</span> 结构
   */
  private renderTextContentPPTXjs(): string {
    if (!this.textStyle || this.textStyle.length === 0) {
      // 单个文本运行
      return `<span class="text-block" style="${this.generateTextSpanStyle()}">${this.escapeHtml(this.text || '')}</span>`;
    }
    
    // 多个文本运行
    return this.textStyle.map(run => {
      const runStyles = this.generateTextRunStyle(run);
      const styleStr = runStyles.join('; ');
      return `<span class="text-block" style="${styleStr}">${this.escapeHtml(run.text)}</span>`;
    }).join('');
  }

  /**
   * 生成文本 span 样式
   */
  private generateTextSpanStyle(): string {
    const styles = [
      `font-size: ${this.style.fontSize || 14}px`,
      `font-family: ${this.style.fontFamily || 'inherit'}`,
      `font-weight: ${this.style.fontWeight || 'inherit'}`,
      `font-style: ${this.style.fontStyle || 'inherit'}`,
      `text-decoration: ${this.style.textDecoration || 'inherit'}`,
      `text-align: ${this.paragraphStyle?.align || 'left'}`,
      `vertical-align: baseline`
    ];
    
    if (this.style.color) {
      styles.push(`color: ${this.style.color}`);
    }
    
    return styles.join('; ');
  }

  /**
   * 生成文本运行样式
   */
  private generateTextRunStyle(run: TextRun): string[] {
    const styles: string[] = [];
    
    // 字体大小
    if (run.fontSize && run.fontSize > 0) {
      styles.push(`font-size: ${run.fontSize}px`);
    }
    
    // 字体家族
    if (run.fontFamily) {
      styles.push(`font-family: ${run.fontFamily}`);
    }
    
    // 加粗
    if (run.bold !== undefined) {
      styles.push(`font-weight: ${run.bold ? 'bold' : 'normal'}`);
    } else if (run.fontWeight) {
      styles.push(`font-weight: ${run.fontWeight}`);
    }
    
    // 斜体
    if (run.italic !== undefined) {
      styles.push(`font-style: ${run.italic ? 'italic' : 'normal'}`);
    } else if (run.fontStyle) {
      styles.push(`font-style: ${run.fontStyle}`);
    }
    
    // 下划线
    if (run.underline && run.underline !== 'none') {
      styles.push(`text-decoration: underline`);
    } else if (run.textDecoration) {
      styles.push(`text-decoration: ${run.textDecoration}`);
    }
    
    // 字体颜色
    if (run.color) {
      styles.push(`color: ${run.color}`);
    }
    
    // 对齐（使用段落对齐）
    styles.push(`text-align: ${this.paragraphStyle?.align || 'left'}`);
    
    // 垂直对齐
    styles.push(`vertical-align: baseline`);
    
    return styles;
  }





  /**
   * 渲染文本内容（PPTXjs 风格）
   * 生成 <span class="text-block">文本</span> 结构
   */
  private renderTextContent(): string {
    if (!this.textStyle || this.textStyle.length === 0) {
      return `<span class="text-block" style="font-size:${this.style.fontSize || 14}px;">${this.escapeHtml(this.text || '')}</span>`;
    }

    // 为每个文本运行生成 span 标签
    return this.textStyle.map(run => {
      const runStyles: string[] = [];

      // 字体大小
      if (run.fontSize && run.fontSize > 0) {
        runStyles.push(`font-size: ${run.fontSize}px`);
      }

      // 字体家族
      if (run.fontFamily) {
        runStyles.push(`font-family: ${run.fontFamily}`);
      }

      // 加粗
      if (run.bold !== undefined) {
        runStyles.push(`font-weight: ${run.bold ? 'bold' : 'normal'}`);
      } else {
        runStyles.push(`font-weight: inherit`);
      }

      // 斜体
      if (run.italic !== undefined) {
        runStyles.push(`font-style: ${run.italic ? 'italic' : 'normal'}`);
      } else {
        runStyles.push(`font-style: inherit`);
      }

      // 下划线
      if (run.underline && run.underline !== 'none') {
        runStyles.push(`text-decoration: underline`);
      } else {
        runStyles.push(`text-decoration: inherit`);
      }

      // 字体颜色
      if (run.color) {
        runStyles.push(`color: ${run.color}`);
      }

      // 对齐
      runStyles.push(`text-align: ${this.paragraphStyle?.align || 'left'}`);

      // 垂直对齐
      runStyles.push(`vertical-align: baseline`);

      const styleStr = runStyles.join(';');
      return `<span class="text-block" style="${styleStr}">${this.escapeHtml(run.text)}</span>`;
    }).join('');
  }

  /**
   * 获取文本样式
   */
  private getTextStyle(): string {
    const styles = [
      `display: flex`,
      `align-items: center`,
      `justify-content: ${this.textStyleFromAlign(this.paragraphStyle?.align)}`,
      `width: 100%`,
      `height: 100%`,
      `box-sizing: border-box`
    ];

    // 段前间距
    if (this.style.spaceBefore) {
      styles.push(`padding-top: ${this.style.spaceBefore}px`);
    }

    // 段后间距
    if (this.style.spaceAfter) {
      styles.push(`padding-bottom: ${this.style.spaceAfter}px`);
    }

    // 上边距
    if (this.style.paddingTop) {
      styles.push(`margin-top: ${this.style.paddingTop}px`);
    }

    // 下边距
    if (this.style.paddingBottom) {
      styles.push(`margin-bottom: ${this.style.paddingBottom}px`);
    }

    // 文本内边距（从 bodyPr 获取或使用段落属性）
    if (this.style.padding) {
      styles.push(`padding-left: ${this.style.padding}`);
      styles.push(`padding-right: ${this.style.padding}`);
    } else {
      // 使用左右边距
      const paddingLeft = this.style.marginLeft || 10;
      const paddingRight = this.style.marginRight || 10;
      styles.push(`padding-left: ${paddingLeft}px`);
      styles.push(`padding-right: ${paddingRight}px`);
    }

    // 处理每个文本运行，使用内联样式支持混合样式
    let htmlContent = this.text || '';
    if (this.textStyle && this.textStyle.length > 0) {
      // 如果有多个文本运行，使用 span 标签
      htmlContent = this.textStyle.map(run => {
        const runStyles: string[] = [];

        // 字体大小
        if (run.fontSize && run.fontSize > 0) {
          runStyles.push(`font-size: ${run.fontSize}px`);
        }

        // 字体颜色
        if (run.color) {
          runStyles.push(`color: ${run.color}`);
        }

        // 字体家族
        if (run.fontFamily) {
          runStyles.push(`font-family: ${run.fontFamily}`);
        }

        // 加粗
        if (run.bold) {
          runStyles.push(`font-weight: bold`);
        }

        // 斜体
        if (run.italic) {
          runStyles.push(`font-style: italic`);
        }

        // 下划线
        if (run.underline && run.underline !== 'none') {
          runStyles.push(`text-decoration: underline`);
        }

        // 删除线
        if (run.strike) {
          runStyles.push(`text-decoration: line-through`);
        }

        // 背景颜色/高亮
        if (run.backgroundColor) {
          runStyles.push(`background-color: ${run.backgroundColor}`);
        }

        // 字间距
        if (run.letterSpacing) {
          runStyles.push(`letter-spacing: ${run.letterSpacing}px`);
        }

        const styleStr = runStyles.length > 0 ? ` style="${runStyles.join('; ')}"` : '';
        return `<span${styleStr}>${this.escapeHtml(run.text)}</span>`;
      }).join('');
    }

    // 返回容器样式和内容
    return styles.join('; ') + ';' + `>CONTENT`;
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
   * 获取对齐类（PPTXjs 风格）
   */
  private getAlignClass(): string {
    const align = this.paragraphStyle?.align || 'left';
    switch (align) {
      case 'center': return 'h-mid';
      case 'right': return 'h-right';
      default: return 'h-left';
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
      attrs: this.getAttributes(this.rawNode!),
      rawNode: this.rawNode
    };
  }
}
