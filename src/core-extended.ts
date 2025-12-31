/**
 * PPT解析扩展功能 - 基于 PPTXjs 的高级特性
 * 支持渐变填充、项目符号、超链接、阴影效果、变换等
 */
import type {
  PptStyle,
  PptFill,
  PptBorder,
  PptShadow,
  PptTransform,
  PptTextParagraph,
  XmlElement,
} from './types';

/**
 * 扩展工具函数
 */
export const PptParseUtilsExtended = {
  /** 解析填充效果 */
  parseXmlFill: (fillNode: XmlElement | null): PptFill => {
    if (!fillNode) return { type: 'none' };

    const solidFill = fillNode.children.find(n => n.tag === 'a:solidFill');
    const gradFill = fillNode.children.find(n => n.tag === 'a:gradFill');
    const blipFill = fillNode.children.find(n => n.tag === 'a:blipFill');
    const pattFill = fillNode.children.find(n => n.tag === 'a:pattFill');

    if (solidFill) {
      const srgbClr = solidFill.children.find(n => n.tag === 'a:srgbClr');
      const schemeClr = solidFill.children.find(n => n.tag === 'a:schemeClr');
      return {
        type: 'solid',
        color: srgbClr?.attrs['val'] || schemeClr?.attrs['val'] || '#ffffff',
      };
    }

    if (gradFill) {
      const gsLst = gradFill.children.find(n => n.tag === 'a:gsLst');
      const stops = gsLst?.children.map(gs => {
        const srgbClr = gs.children.find(n => n.tag === 'a:srgbClr');
        const schemeClr = gs.children.find(n => n.tag === 'a:schemeClr');
        return {
          position: parseInt(gs.attrs['pos'] || '0') / 100000,
          color: srgbClr?.attrs['val'] || schemeClr?.attrs['val'] || '#ffffff',
        };
      }) || [];

      const lin = gradFill.children.find(n => n.tag === 'a:lin');
      const angle = lin ? parseInt(lin.attrs['ang'] || '0') : 0;

      return {
        type: 'gradient',
        gradientStops: stops,
        gradientDirection: (angle + 90) % 360,
      };
    }

    if (blipFill) {
      const blip = blipFill.children.find(n => n.tag === 'a:blip');
      return {
        type: 'picture',
        image: blip?.attrs['r:embed'] || '',
      };
    }

    if (pattFill) {
      const fgClr = pattFill.children.find(n => n.tag === 'a:fgClr');
      const bgClr = pattFill.children.find(n => n.tag === 'a:bgClr');
      const pattern = pattFill.children.find(n => n.tag === 'a:prstPatt');

      return {
        type: 'pattern',
        color: fgClr?.children.find(n => n.tag === 'a:srgbClr')?.attrs['val'] || '#000000',
        opacity: parseInt(pattFill.attrs['opacity'] || '100000') / 100000,
      };
    }

    return { type: 'none' };
  },

  /** 解析边框样式 */
  parseXmlBorder: (lnNode: XmlElement | null): PptBorder => {
    if (!lnNode) return { color: '#000000', width: 0, style: 'solid' };

    const solidFill = lnNode.children.find(n => n.tag === 'a:solidFill');
    const srgbClr = solidFill?.children.find(n => n.tag === 'a:srgbClr');

    const prstDash = lnNode.children.find(n => n.tag === 'a:prstDash');
    const dashStyle = prstDash?.attrs['val'] || '';

    const noFill = lnNode.children.find(n => n.tag === 'a:noFill');

    return {
      color: noFill ? 'transparent' : (srgbClr?.attrs['val'] || '#000000'),
      width: noFill ? 0 : (parseInt(lnNode.attrs['w'] || '0') / 12700),
      style: lnNode.attrs['cmpd'] === 'dbl'
        ? 'double'
        : dashStyle
        ? 'dashed'
        : (prstDash && prstDash.attrs['val'] === 'sysDot')
        ? 'dotted'
        : 'solid',
      dashStyle,
    };
  },

  /** 解析阴影效果 */
  parseXmlShadow: (effectLstNode: XmlElement | null): PptShadow | undefined => {
    if (!effectLstNode) return undefined;

    const outerShdw = effectLstNode.children.find(n => n.tag === 'a:outerShdw');
    const innerShdw = effectLstNode.children.find(n => n.tag === 'a:innerShdw');
    const shdwNode = outerShdw || innerShdw;

    if (!shdwNode) return undefined;

    const srgbClr = shdwNode.children.find(n => n.tag === 'a:srgbClr');
    const schemeClr = shdwNode.children.find(n => n.tag === 'a:schemeClr');

    const alpha = parseInt(shdwNode.attrs['alpha'] || '100000');
    const blurRad = parseInt(shdwNode.attrs['blurRad'] || '0');
    const dist = parseInt(shdwNode.attrs['dist'] || '0');
    const dir = parseInt(shdwNode.attrs['dir'] || '0');

    return {
      color: srgbClr?.attrs['val'] || schemeClr?.attrs['val'] || '#000000',
      blur: blurRad / 12700,
      offsetX: (dist * Math.cos((dir * Math.PI) / 180)) / 12700,
      offsetY: (dist * Math.sin((dir * Math.PI) / 180)) / 12700,
      opacity: alpha / 100000,
    };
  },

  /** 解析变换效果 */
  parseXmlTransform: (xfrmNode: XmlElement | null): PptTransform => {
    if (!xfrmNode) return {};

    const rot = parseInt(xfrmNode.attrs['rot'] || '0');
    const flipH = xfrmNode.attrs['flipH'] === '1';
    const flipV = xfrmNode.attrs['flipV'] === '1';

    return {
      rotate: rot / 60000,
      flipH,
      flipV,
    };
  },

  /** 解析文本段落（支持项目符号和超链接） */
  parseXmlTextParagraphs: (pNodes: XmlElement[]): PptTextParagraph[] => {
    return pNodes.map(pNode => {
      const pPr = pNode.children.find(n => n.tag === 'a:pPr');

      // 项目符号
      let bullet: PptTextParagraph['bullet'] = undefined;
      if (pPr) {
        const buChar = pPr.children.find(n => n.tag === 'a:buChar');
        const buAutoNum = pPr.children.find(n => n.tag === 'a:buAutoNum');
        const buNone = pPr.children.find(n => n.tag === 'a:buNone');

        if (!buNone) {
          const lvl = parseInt(pPr.attrs['lvl'] || '0');

          if (buChar) {
            const char = buChar.children.find(n => n.tag === 'a:char');
            bullet = {
              type: 'bullet' as const,
              char: char?.attrs['val'] || '•',
              level: lvl,
            };
          } else if (buAutoNum) {
            bullet = {
              type: 'numbered' as const,
              level: lvl,
            };
          }
        }
      }

      // 超链接
      const hlinkClick = pNode.children.find(n => n.tag === 'a:hlinkClick');
      const hyperlink = hlinkClick
        ? {
            url: hlinkClick.attrs['r:id'] || '',
            tooltip: hlinkClick.attrs['tooltip'],
          }
        : undefined;

      // 文本内容
      const rNodes = pNode.children.filter(n => n.tag === 'a:r');
      const textSegments = rNodes.map(rNode => {
        const rPr = rNode.children.find(n => n.tag === 'a:rPr');
        const tNode = rNode.children.find(n => n.tag === 'a:t');

        const style: Partial<PptStyle> = {};
        if (rPr) {
          const sz = rPr.children.find(n => n.tag === 'a:sz');
          const b = rPr.children.find(n => n.tag === 'a:b');
          const i = rPr.children.find(n => n.tag === 'a:i');
          const u = rPr.children.find(n => n.tag === 'a:u');
          const strike = rPr.children.find(n => n.tag === 'a:strike');

          if (sz) style.fontSize = parseInt(sz.attrs['val'] || '0') / 100;
          if (b) style.fontWeight = b.attrs['val'] === '1' ? 'bold' : 'normal';
          if (i) style.fontStyle = i.attrs['val'] === '1' ? 'italic' : 'normal';
          if (u) style.textDecoration = 'underline';
          if (strike) style.textDecoration = 'line-through';

          const solidFill = rPr.children.find(n => n.tag === 'a:solidFill');
          const srgbClr = solidFill?.children.find(n => n.tag === 'a:srgbClr');
          if (srgbClr) style.color = srgbClr.attrs['val'];
        }

        return {
          text: tNode?.text || '',
          style,
        };
      });

      // 段落对齐
      const align = pPr?.attrs['algn'] || 'left';
      const textVerticalAlign = pPr?.attrs['vert'] || 'baseline';

      return {
        text: textSegments.map(s => s.text).join(''),
        style: {
          ...textSegments[0]?.style,
          textAlign: align as any,
          textVerticalAlign: textVerticalAlign as any,
        },
        bullet,
        hyperlink,
      };
    });
  },

  /** 解析形状类型 */
  parseShapeType: (spPrNode: XmlElement | null): string => {
    if (!spPrNode) return 'rectangle';

    const prstGeom = spPrNode.children.find(n => n.tag === 'a:prstGeom');
    if (prstGeom) {
      return prstGeom.attrs['prst'] || 'rectangle';
    }

    const custGeom = spPrNode.children.find(n => n.tag === 'a:custGeom');
    if (custGeom) {
      return 'custom';
    }

    return 'rectangle';
  },

  /** 解析形状圆角 */
  parseShapeRadius: (spPrNode: XmlElement | null): number => {
    if (!spPrNode) return 0;

    const xfrm = spPrNode.children.find(n => n.tag === 'a:xfrm');
    if (xfrm && xfrm.children) {
      const off = xfrm.children.find(n => n.tag === 'a:off');
      const ext = xfrm.children.find(n => n.tag === 'a:ext');

      if (off && ext) {
        const ox = parseInt(off.attrs['x'] || '0');
        const oy = parseInt(off.attrs['y'] || '0');
        const ex = parseInt(ext.attrs['cx'] || '0');
        const ey = parseInt(ext.attrs['cy'] || '0');

        // 计算圆角半径
        return Math.min(Math.abs(ex - ox), Math.abs(ey - oy)) / 2;
      }
    }

    return 0;
  },

  /** 获取主题颜色映射 */
  parseThemeColors: (themeNode: XmlElement | null): Record<string, string> => {
    const colors: Record<string, string> = {};
    if (!themeNode) return colors;

    const colorSchemes = themeNode.children.find(n => n.tag === 'a:clrScheme');
    if (colorSchemes) {
      const mapping: Record<string, string> = {
        'a:dk1': 'background',
        'a:lt1': 'text',
        'a:accent1': 'accent1',
        'a:accent2': 'accent2',
        'a:accent3': 'accent3',
        'a:accent4': 'accent4',
        'a:accent5': 'accent5',
        'a:accent6': 'accent6',
      };

      colorSchemes.children.forEach(node => {
        const colorName = mapping[node.tag];
        if (colorName) {
          const srgbClr = node.children.find(n => n.tag === 'a:srgbClr');
          if (srgbClr) {
            colors[colorName] = srgbClr.attrs['val'];
          }
        }
      });
    }

    return colors;
  },

  /** 解析关系映射（用于图片、媒体资源） */
  parseRelationships: (relsXml: string): Record<string, { type: string; target: string }> => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(relsXml, 'application/xml');
    const relationships: Record<string, { type: string; target: string }> = {};

    Array.from(doc.querySelectorAll('Relationship')).forEach(rel => {
      const id = rel.getAttribute('Id');
      const type = rel.getAttribute('Type');
      const target = rel.getAttribute('Target');
      if (id && type && target) {
        relationships[id] = { type, target };
      }
    });

    return relationships;
  },
};
