/**
 * PPTXjs颜色处理工具 - TypeScript转译版
 * 对齐PPTXjs的颜色处理逻辑
 */

/**
 * 颜色类型枚举
 */
export enum ColorType {
  SOLID = 'solid',
  GRADIENT = 'gradient',
  PATTERN = 'pattern',
  NONE = 'none',
}

/**
 * 颜色值结构
 */
export interface ColorValue {
  type: ColorType;
  color?: string;
  alpha?: number;
  stops?: Array<{
    position: number;
    color: string;
    alpha?: number;
  }>;
}

/**
 * 颜色映射结构
 */
export interface ColorMap {
  bg1?: string;
  tx1?: string;
  bg2?: string;
  tx2?: string;
  accent1?: string;
  accent2?: string;
  accent3?: string;
  accent4?: string;
  accent5?: string;
  accent6?: string;
  hlink?: string;
  folHlink?: string;
}

/**
 * 主题颜色定义 - 对齐PPTXjs的主题颜色系统
 */
export const THEME_COLORS: Record<string, string> = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '1F497D',
  lt2: 'EEECE1',
  accent1: '4F81BD',
  accent2: 'C0504D',
  accent3: '9BBB59',
  accent4: '8064A2',
  accent5: '4BACC6',
  accent6: 'F79646',
  hlink: '0000FF',
  folHlink: '800080',
};

/**
 * 获取颜色值 - 对齐PPTXjs的颜色获取逻辑
 * 支持从多种来源获取颜色：十六进制、主题颜色、系统颜色等
 */
export function getColorValue(colorNode: any): string | null {
  if (!colorNode) return null;

  // 1. 尝试获取十六进制颜色值
  const srgbClr = getTextByPathList(colorNode, ['a:srgbClr', 'attrs', 'val']);
  if (srgbClr) {
    return '#' + srgbClr;
  }

  // 2. 尝试获取方案颜色（主题颜色）
  const schemeClr = getTextByPathList(colorNode, ['a:schemeClr', 'attrs', 'val']);
  if (schemeClr) {
    return getThemeColor(schemeClr);
  }

  // 3. 尝试获取系统颜色
  const sysClr = getTextByPathList(colorNode, ['a:sysClr', 'attrs', 'lastClr']);
  if (sysClr) {
    return '#' + sysClr;
  }

  // 4. 尝试获取预设颜色
  const prstClr = getTextByPathList(colorNode, ['a:prstClr', 'attrs', 'val']);
  if (prstClr) {
    return getPresetColor(prstClr);
  }

  return null;
}

/**
 * 获取主题颜色
 */
export function getThemeColor(schemeColor: string): string {
  const themeColor = THEME_COLORS[schemeColor];
  return themeColor ? '#' + themeColor : '#000000';
}

/**
 * 获取预设颜色 - 对齐PPTXjs的预设颜色映射
 */
export function getPresetColor(presetColor: string): string {
  const presetColors: Record<string, string> = {
    'aliceBlue': 'F0F8FF',
    'antiqueWhite': 'FAEBD7',
    'aqua': '00FFFF',
    'aquamarine': '7FFFD4',
    'azure': 'F0FFFF',
    'beige': 'F5F5DC',
    'bisque': 'FFE4C4',
    'black': '000000',
    'blanchedAlmond': 'FFEBCD',
    'blue': '0000FF',
    'blueViolet': '8A2BE2',
    'brown': 'A52A2A',
    'burlywood': 'DEB887',
    'cadetBlue': '5F9EA0',
    'chartreuse': '7FFF00',
    'chocolate': 'D2691E',
    'coral': 'FF7F50',
    'cornflowerBlue': '6495ED',
    'cornsilk': 'FFF8DC',
    'crimson': 'DC143C',
    'cyan': '00FFFF',
    'darkBlue': '00008B',
    'darkCyan': '008B8B',
    'darkGoldenrod': 'B8860B',
    'darkGray': 'A9A9A9',
    'darkGreen': '006400',
    'darkGrey': 'A9A9A9',
    'darkKhaki': 'BDB76B',
    'darkMagenta': '8B008B',
    'darkOliveGreen': '556B2F',
    'darkOrange': 'FF8C00',
    'darkOrchid': '9932CC',
    'darkRed': '8B0000',
    'darkSalmon': 'E9967A',
    'darkSeaGreen': '8FBC8F',
    'darkSlateBlue': '483D8B',
    'darkSlateGray': '2F4F4F',
    'darkSlateGrey': '2F4F4F',
    'darkTurquoise': '00CED1',
    'darkViolet': '9400D3',
    'deepPink': 'FF1493',
    'deepSkyBlue': '00BFFF',
    'dimGray': '696969',
    'dimGrey': '696969',
    'dkBlue': '00008B',
    'dkCyan': '008B8B',
    'dkGoldenrod': 'B8860B',
    'dkGray': 'A9A9A9',
    'dkGreen': '006400',
    'dkKhaki': 'BDB76B',
    'dkMagenta': '8B008B',
    'dkOliveGreen': '556B2F',
    'dkOrange': 'FF8C00',
    'dkOrchid': '9932CC',
    'dkRed': '8B0000',
    'dkSalmon': 'E9967A',
    'dkSeaGreen': '8FBC8F',
    'dkSlateBlue': '483D8B',
    'dkSlateGray': '2F4F4F',
    'dkSlateGrey': '2F4F4F',
    'dkTurquoise': '00CED1',
    'dkViolet': '9400D3',
    'dodgerBlue': '1E90FF',
    'firebrick': 'B22222',
    'floralWhite': 'FFFAF0',
    'forestGreen': '228B22',
    'fuchsia': 'FF00FF',
    'gainsboro': 'DCDCDC',
    'ghostWhite': 'F8F8FF',
    'gold': 'FFD700',
    'goldenrod': 'DAA520',
    'gray': '808080',
    'green': '008000',
    'greenYellow': 'ADFF2F',
    'grey': '808080',
    'honeydew': 'F0FFF0',
    'hotPink': 'FF69B4',
    'indianRed': 'CD5C5C',
    'indigo': '4B0082',
    'ivory': 'FFFFF0',
    'khaki': 'F0E68C',
    'lavender': 'E6E6FA',
    'lavenderBlush': 'FFF0F5',
    'lawnGreen': '7CFC00',
    'lemonChiffon': 'FFFACD',
    'lightBlue': 'ADD8E6',
    'lightCoral': 'F08080',
    'lightCyan': 'E0FFFF',
    'lightGoldenrodYellow': 'FAFAD2',
    'lightGray': 'D3D3D3',
    'lightGreen': '90EE90',
    'lightGrey': 'D3D3D3',
    'lightPink': 'FFB6C1',
    'lightSalmon': 'FFA07A',
    'lightSeaGreen': '20B2AA',
    'lightSkyBlue': '87CEFA',
    'lightSlateGray': '778899',
    'lightSlateGrey': '778899',
    'lightSteelBlue': 'B0C4DE',
    'lightYellow': 'FFFFE0',
    'lime': '00FF00',
    'limeGreen': '32CD32',
    'linen': 'FAF0E6',
    'ltBlue': 'ADD8E6',
    'ltCoral': 'F08080',
    'ltCyan': 'E0FFFF',
    'ltGoldenrodYellow': 'FAFAD2',
    'ltGray': 'D3D3D3',
    'ltGreen': '90EE90',
    'ltGrey': 'D3D3D3',
    'ltPink': 'FFB6C1',
    'ltSalmon': 'FFA07A',
    'ltSeaGreen': '20B2AA',
    'ltSkyBlue': '87CEFA',
    'ltSlateGray': '778899',
    'ltSlateGrey': '778899',
    'ltSteelBlue': 'B0C4DE',
    'ltYellow': 'FFFFE0',
    'magenta': 'FF00FF',
    'maroon': '800000',
    'mediumAquamarine': '66CDAA',
    'mediumBlue': '0000CD',
    'mediumOrchid': 'BA55D3',
    'mediumPurple': '9370DB',
    'mediumSeaGreen': '3CB371',
    'mediumSlateBlue': '7B68EE',
    'mediumSpringGreen': '00FA9A',
    'mediumTurquoise': '48D1CC',
    'mediumVioletRed': 'C71585',
    'midnightBlue': '191970',
    'mintCream': 'F5FFFA',
    'mistyRose': 'FFE4E1',
    'moccasin': 'FFE4B5',
    'navajoWhite': 'FFDEAD',
    'navy': '000080',
    'oldLace': 'FDF5E6',
    'olive': '808000',
    'oliveDrab': '6B8E23',
    'orange': 'FFA500',
    'orangeRed': 'FF4500',
    'orchid': 'DA70D6',
    'paleGoldenrod': 'EEE8AA',
    'paleGreen': '98FB98',
    'paleTurquoise': 'AFEEEE',
    'paleVioletRed': 'DB7093',
    'papayaWhip': 'FFEFD5',
    'peachPuff': 'FFDAB9',
    'peru': 'CD853F',
    'pink': 'FFC0CB',
    'plum': 'DDA0DD',
    'powderBlue': 'B0E0E6',
    'purple': '800080',
    'red': 'FF0000',
    'rosyBrown': 'BC8F8F',
    'royalBlue': '4169E1',
    'saddleBrown': '8B4513',
    'salmon': 'FA8072',
    'sandyBrown': 'F4A460',
    'seaGreen': '2E8B57',
    'seaShell': 'FFF5EE',
    'sienna': 'A0522D',
    'silver': 'C0C0C0',
    'skyBlue': '87CEEB',
    'slateBlue': '6A5ACD',
    'slateGray': '708090',
    'slateGrey': '708090',
    'snow': 'FFFAFA',
    'springGreen': '00FF7F',
    'steelBlue': '4682B4',
    'tan': 'D2B48C',
    'teal': '008080',
    'thistle': 'D8BFD8',
    'tomato': 'FF6347',
    'turquoise': '40E0D0',
    'violet': 'EE82EE',
    'wheat': 'F5DEB3',
    'white': 'FFFFFF',
    'whiteSmoke': 'F5F5F5',
    'yellow': 'FFFF00',
    'yellowGreen': '9ACD32',
  };

  const hex = presetColors[presetColor];
  return hex ? '#' + hex : '#000000';
}

/**
 * 获取Alpha通道值
 */
export function getAlphaValue(colorNode: any): number {
  if (!colorNode) return 1;

  // 尝试获取alpha值
  const alphaNode = getTextByPathList(colorNode, ['a:alpha', 'attrs', 'val']);
  if (alphaNode) {
    return parseInt(alphaNode) / 100000;
  }

  // 尝试获取lumMod值
  const lumMod = getTextByPathList(colorNode, ['a:lumMod', 'attrs', 'val']);
  if (lumMod) {
    return parseInt(lumMod) / 100000;
  }

  return 1;
}

/**
 * 应用颜色映射 - 对齐PPTXjs的颜色映射逻辑
 */
export function applyColorMap(
  color: string,
  colorMapOvr?: ColorMap
): string {
  if (!colorMapOvr) return color;

  // 移除#前缀
  const colorKey = color.replace('#', '').toLowerCase();

  // 查找映射
  for (const [key, value] of Object.entries(colorMapOvr)) {
    if (value && value.replace('#', '').toLowerCase() === colorKey) {
      const mappedColor = THEME_COLORS[key];
      if (mappedColor) {
        return '#' + mappedColor;
      }
    }
  }

  return color;
}

/**
 * 解析颜色填充 - 对齐PPTXjs的颜色填充解析
 */
export function parseColorFill(fillNode: any): ColorValue | null {
  if (!fillNode) return null;

  // 1. 纯色填充
  if (fillNode['a:solidFill']) {
    const color = getColorValue(fillNode['a:solidFill']);
    const alpha = getAlphaValue(fillNode['a:solidFill']);
    
    if (color) {
      return {
        type: ColorType.SOLID,
        color,
        alpha,
      };
    }
  }

  // 2. 渐变填充
  if (fillNode['a:gradFill']) {
    const stops: Array<{ position: number; color: string; alpha?: number }> = [];
    const gsLst = fillNode['a:gradFill']['a:gsLst'];

    if (gsLst) {
      const gsNodes = Array.isArray(gsLst['a:gs']) ? gsLst['a:gs'] : [gsLst['a:gs']];
      
      for (const gs of gsNodes) {
        const position = parseInt(gs.attrs.pos) / 100000;
        const color = getColorValue(gs);
        const alpha = getAlphaValue(gs);
        
        if (color) {
          stops.push({ position, color, alpha });
        }
      }
    }

    if (stops.length > 0) {
      return {
        type: ColorType.GRADIENT,
        stops: stops.sort((a, b) => a.position - b.position),
      };
    }
  }

  // 3. 图案填充
  if (fillNode['a:pattFill']) {
    const fgColor = getColorValue(fillNode['a:pattFill']['a:fgClr']);
    const bgColor = getColorValue(fillNode['a:pattFill']['a:bgClr']);
    
    if (fgColor) {
      return {
        type: ColorType.PATTERN,
        color: fgColor,
      };
    }
  }

  return null;
}

/**
 * 生成CSS颜色值
 */
export function generateCssColor(colorValue: ColorValue): string {
  if (!colorValue || colorValue.type === ColorType.NONE) {
    return 'transparent';
  }

  if (colorValue.type === ColorType.SOLID) {
    let color = colorValue.color || '#000000';
    
    // 添加alpha通道
    if (colorValue.alpha !== undefined && colorValue.alpha < 1) {
      color = hexToRgba(color, colorValue.alpha);
    }
    
    return color;
  }

  if (colorValue.type === ColorType.GRADIENT && colorValue.stops) {
    const stops = colorValue.stops
      .map(stop => {
        let color = stop.color;
        if (stop.alpha !== undefined && stop.alpha < 1) {
          color = hexToRgba(color, stop.alpha);
        }
        return `${color} ${Math.round(stop.position * 100)}%`;
      })
      .join(', ');
    
    return `linear-gradient(to bottom, ${stops})`;
  }

  return colorValue.color || '#000000';
}

/**
 * 十六进制转RGBA
 */
export function hexToRgba(hex: string, alpha: number): string {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

/**
 * 解析颜色映射覆盖 - 对齐PPTXjs的颜色映射覆盖逻辑
 */
export function parseColorMapOverride(
  slideContent: any,
  slideLayoutContent: any,
  slideMasterContent: any
): {
  slide?: ColorMap;
  layout?: ColorMap;
  master?: ColorMap;
} {
  const result: {
    slide?: ColorMap;
    layout?: ColorMap;
    master?: ColorMap;
  } = {};

  // 解析slide的颜色映射
  const slideClrMapOvr = getTextByPathList(slideContent, ['p:sld', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs']);
  if (slideClrMapOvr) {
    result.slide = parseColorMapAttrs(slideClrMapOvr);
  }

  // 解析layout的颜色映射
  const layoutClrMapOvr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs']);
  if (layoutClrMapOvr) {
    result.layout = parseColorMapAttrs(layoutClrMapOvr);
  }

  // 解析master的颜色映射
  const masterClrMapOvr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs']);
  if (masterClrMapOvr) {
    result.master = parseColorMapAttrs(masterClrMapOvr);
  }

  return result;
}

/**
 * 解析颜色映射属性
 */
function parseColorMapAttrs(attrs: any): ColorMap {
  const colorMap: ColorMap = {};

  const colorKeys = [
    'bg1', 'tx1', 'bg2', 'tx2',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
    'hlink', 'folHlink'
  ];

  for (const key of colorKeys) {
    const value = attrs[key];
    if (value) {
      (colorMap as any)[key] = getThemeColor(value);
    }
  }

  return colorMap;
}

/**
 * 获取路径文本值 - 本地实现（避免循环依赖）
 */
export function getTextByPathList(obj: any, pathList: string[]): any {
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
 * 导出函数
 */
