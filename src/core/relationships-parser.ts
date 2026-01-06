/**
 * 关联关系解析器
 * 解析 rels 文件，管理资源引用关系
 */

import JSZip from 'jszip';
import { PATHS } from '../constants';
import { parseRels, log } from '../utils/index';
import type { RelsMap, Relationship } from './types';

/**
 * 解析全局关联关系
 * @param zip JSZip对象
 * @returns 关联关系映射表
 */
export async function parseGlobalRels(zip: JSZip): Promise<RelsMap> {
  try {
    const relsXml = await zip.file('_rels/.rels')?.async('string');
    if (!relsXml) {
      return {};
    }

    return parseRels(relsXml);
  } catch (error) {
    log('warn', 'Failed to parse global relationships', error);
    return {};
  }
}

/**
 * 解析幻灯片关联关系
 * @param zip JSZip对象
 * @param slideNumber 幻灯片编号
 * @param relsBasePath 关联关系的基础路径（默认是slides）
 * @returns 关联关系映射表
 */
export async function parseSlideRels(
  zip: JSZip,
  slideNumber: string,
  relsBasePath: string = PATHS.SLIDE_RELS
): Promise<RelsMap> {
  try {
    const relsPath = `${relsBasePath}${slideNumber}.xml.rels`;
    const relsXml = await zip.file(relsPath)?.async('string');

    if (!relsXml) {
      return {};
    }

    return parseRels(relsXml);
  } catch (error) {
    log('warn', `Failed to parse ${relsBasePath}${slideNumber} relationships`, error);
    return {};
  }
}

/**
 * 从关联关系中获取幻灯片布局的引用
 * @param relsMap 关联关系映射表
 * @returns 布局文件路径，如果没有找到则返回undefined
 */
export function getSlideLayoutRef(relsMap: RelsMap): string | undefined {
  for (const [relId, rel] of Object.entries(relsMap)) {
    if (rel.type.includes('slideLayout')) {
      // 提取 layout 文件名，例如 "../slideLayouts/slideLayout1.xml" -> "slideLayout1"
      const match = rel.target.match(/slideLayout(\d+)\.xml/);
      if (match) {
        return `slideLayout${match[1]}`;
      }
    }
  }
  return undefined;
}
