/**
 * 主题解析器
 * 解析PPTX主题文件，包括颜色方案、字体方案、效果方案
 */

import JSZip from 'jszip';
import { log, getFirstChildByTagNS } from '../utils/index';

export interface ThemeColors {
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

export interface ThemeResult {
  colors: ThemeColors;
}

/**
 * 解析主题文件
 * @param zip JSZip对象
 * @param themePath 主题文件路径（默认 ppt/theme/theme1.xml）
 * @returns 主题解析结果
 */
export async function parseTheme(
  zip: JSZip,
  themePath: string = 'ppt/theme/theme1.xml'
): Promise<ThemeResult | null> {
  try {
    const themeXml = await zip.file(themePath)?.async('string');
    if (!themeXml) {
      log('warn', `Theme file not found: ${themePath}`);
      return null;
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(themeXml, 'application/xml');
    const root = doc.documentElement;

    // 解析颜色方案
    const colors = parseColorScheme(root);

    log('info', `Parsed theme from ${themePath}`);

    return { colors };
  } catch (error) {
    log('error', `Failed to parse theme: ${themePath}`, error);
    return null;
  }
}

/**
 * 解析颜色方案 <a:themeElements><a:clrScheme>
 * @param root 主题根元素
 * @returns 颜色方案对象
 */
function parseColorScheme(root: Element): ThemeColors {
  const colors: ThemeColors = {};

  // 查找 a:clrScheme
  const themeElements = getFirstChildByTagNS(root, 'themeElements', 
    'http://schemas.openxmlformats.org/drawingml/2006/main');
  
  if (!themeElements) return colors;

  const clrScheme = getFirstChildByTagNS(themeElements, 'clrScheme', 
    'http://schemas.openxmlformats.org/drawingml/2006/main');
  
  if (!clrScheme) return colors;

  // 解析各个颜色值
  const colorMap: Record<string, keyof ThemeColors> = {
    'bg1': 'bg1',
    'tx1': 'tx1',
    'bg2': 'bg2',
    'tx2': 'tx2',
    'accent1': 'accent1',
    'accent2': 'accent2',
    'accent3': 'accent3',
    'accent4': 'accent4',
    'accent5': 'accent5',
    'accent6': 'accent6',
    'hlink': 'hlink',
    'folHlink': 'folHlink'
  };

  // 遍历 clrScheme 的子元素
  Array.from(clrScheme.children).forEach(child => {
    if (child.nodeType !== 1) return;

    const localName = child.localName || child.tagName.split(':').pop();
    if (localName && colorMap[localName]) {
      colors[colorMap[localName] as keyof ThemeColors] = parseColorValue(child);
    }
  });

  return colors;
}

/**
 * 解析单个颜色值
 * @param colorEl 颜色元素
 * @returns 颜色值 (十六进制字符串)
 */
function parseColorValue(colorEl: Element): string {
  // 检查 srgbClr
  const srgbClr = getFirstChildByTagNS(colorEl, 'srgbClr', 
    'http://schemas.openxmlformats.org/drawingml/2006/main');
  
  if (srgbClr) {
    const val = srgbClr.getAttribute('val');
    if (val) {
      return `#${val}`;
    }
  }

  // 检查 sysClr
  const sysClr = getFirstChildByTagNS(colorEl, 'sysClr', 
    'http://schemas.openxmlformats.org/drawingml/2006/main');
  
  if (sysClr) {
    const lastClr = sysClr.getAttribute('lastClr');
    if (lastClr) {
      return `#${lastClr}`;
    }
  }

  // 其他颜色类型暂不支持，返回默认值
  return '#ffffff';
}

/**
 * 解析颜色映射 override
 * @param root slide/slideMaster/slideLayout根元素
 * @returns 颜色映射对象
 */
export function parseColorMap(root: Element): Record<string, string> {
  const clrMapOvr = getFirstChildByTagNS(root, 'clrMapOvr', 
    'http://schemas.openxmlformats.org/presentationml/2006/main');
  
  const clrMap: Record<string, string> = {};

  if (!clrMapOvr) {
    return clrMap;
  }

  // 解析 masterClrMapping
  const masterClrMapping = getFirstChildByTagNS(clrMapOvr, 'masterClrMapping', 
    'http://schemas.openxmlformats.org/drawingml/2006/main');
  
  if (masterClrMapping) {
    // 读取所有属性作为映射
    Array.from(masterClrMapping.attributes).forEach(attr => {
      clrMap[attr.name] = attr.value;
    });
  }

  return clrMap;
}

/**
 * 将方案颜色名称解析为实际颜色值
 * @param schemeColor 方案颜色名称 (如 'bg1', 'accent1')
 * @param themeColors 主题颜色方案
 * @param colorMap 颜色映射覆盖
 * @returns 实际颜色值
 */
export function resolveSchemeColor(
  schemeColor: string,
  themeColors: ThemeColors,
  colorMap: Record<string, string> = {}
): string {
  // 检查是否有颜色映射覆盖
  if (colorMap[schemeColor]) {
    // 解析映射值 (格式: "light" 或 "dark")
    const mappedColor = colorMap[schemeColor];
    // 转换映射值 (如 bg1->lt1, tx1->dk1)
    const resolvedColor = resolveColorMapping(mappedColor);
    // 返回主题中的颜色
    return themeColors[resolvedColor as keyof ThemeColors] || '#ffffff';
  }

  // 直接返回主题中的颜色
  return themeColors[schemeColor as keyof ThemeColors] || '#ffffff';
}

/**
 * 解析颜色映射名称
 * @param mappedColor 映射值
 * @returns 解析后的颜色键名
 */
function resolveColorMapping(mappedColor: string): string {
  const mapping: Record<string, string> = {
    'light': 'lt1',
    'dark': 'dk1',
    // 可以添加更多映射
  };

  return mapping[mappedColor] || mappedColor;
}
