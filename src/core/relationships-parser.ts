/**
 * 关联关系解析器
 * 解析 rels 文件，管理资源引用关系
 */

import { PATHS } from '../constants';
import { parseRels, log } from '../utils';
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
 * @returns 关联关系映射表
 */
export async function parseSlideRels(zip: JSZip, slideNumber: string): Promise<RelsMap> {
  try {
    const relsPath = `${PATHS.SLIDE_RELS}slide${slideNumber}.xml.rels`;
    const relsXml = await zip.file(relsPath)?.async('string');

    if (!relsXml) {
      return {};
    }

    return parseRels(relsXml);
  } catch (error) {
    log('warn', `Failed to parse slide ${slideNumber} relationships`, error);
    return {};
  }
}
