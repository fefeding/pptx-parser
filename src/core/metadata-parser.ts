/**
 * 元数据解析器
 * 解析 docProps/core.xml 等元数据文件
 */

import JSZip from 'jszip';
import { PATHS } from '../constants';
import { parseMetadata, log, emu2px } from '../utils';
import type { Metadata, SlideSize } from './types';

/**
 * 解析核心属性（docProps/core.xml）
 * @param zip JSZip对象
 * @returns 元数据对象
 */
export async function parseCoreProperties(zip: JSZip): Promise<Metadata> {
  try {
    const coreXml = await zip.file(`${PATHS.DOCPROPS}core.xml`)?.async('string');
    if (!coreXml) {
      log('warn', 'core.xml not found');
      return {};
    }

    return parseMetadata(coreXml);
  } catch (error) {
    log('warn', 'Failed to parse core properties', error);
    return {};
  }
}

/**
 * 解析幻灯片尺寸
 * @param zip JSZip对象
 * @returns 尺寸对象（像素）
 */
export async function parseSlideLayoutSize(zip: JSZip): Promise<SlideSize> {
  try {
    // 尝试从 presentation.xml 解析
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
    if (presentationXml) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(presentationXml, 'application/xml');
      const NS = doc.documentElement.namespaceURI || 'http://schemas.openxmlformats.org/presentationml/2006/main';

      const sldSz = doc.getElementsByTagNameNS(NS, 'sldSz')[0];
      if (sldSz) {
        const cx = sldSz.getAttribute('cx');
        const cy = sldSz.getAttribute('cy');

        if (cx && cy) {
          return {
            width: emu2px(cx),
            height: emu2px(cy)
          };
        }
      }
    }

    // 默认尺寸（16:9）
    return { width: 1280, height: 720 };
  } catch (error) {
    log('warn', 'Failed to parse slide size', error);
    return { width: 1280, height: 720 };
  }
}

/**
 * 推断页面比例类型
 * @param ratio 宽高比
 * @returns 页面类型
 */
export function inferPageSize(ratio: number): '4:3' | '16:9' | '16:10' | 'custom' {
  const epsilon = 0.01;

  if (Math.abs(ratio - 1.33333) < epsilon) return '4:3';
  if (Math.abs(ratio - 1.77778) < epsilon) return '16:9';
  if (Math.abs(ratio - 1.6) < epsilon) return '16:10';
  return 'custom';
}
